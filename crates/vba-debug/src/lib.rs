//! Debugger for the VBA runtime, modelled on the VBE debugger.
//!
//! Execution modes:
//! - **Run**: no hook installed, full speed.
//! - **Debug run**: runs until a breakpoint, `Stop`, failed `Debug.Assert`,
//!   a watch that fires, an error (per [`ErrorBreak`]) or an async pause
//!   request (Ctrl+Break, see [`Debugger::pause_handle`]).
//! - **Stepping**: Step Into (F8), Step Over (Shift+F8), Step Out
//!   (Ctrl+Shift+F8), Run to Cursor (Ctrl+F8).
//!
//! Breakpoints can be conditional, have hit counts, or be logpoints that print
//! a message instead of stopping. Watches come in the three VBE flavours:
//! plain expression, *Break When Value Is True*, *Break When Value Changes*.
//!
//! A [`Frontend`] decides what to do when the program pauses. Implement it
//! directly for an in-process UI, or use [`ChannelFrontend`] to drive the
//! debugger from another thread (an IDE, a DAP server).

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{Receiver, Sender};
use std::sync::Arc;

use vba_runtime::{DebugAction, DebugContext, DebugHook, Location, StopKind, Value, VbaError};

pub type BreakpointId = u32;
pub type WatchId = u32;

#[derive(Debug, Clone, PartialEq)]
pub enum HitCondition {
    /// Break when the hit count equals N.
    Equals(u32),
    /// Break when the hit count is at least N.
    AtLeast(u32),
    /// Break every N-th hit.
    Multiple(u32),
}

#[derive(Debug, Clone)]
pub struct Breakpoint {
    pub id: BreakpointId,
    /// Module name; `None` matches any module.
    pub module: Option<String>,
    pub line: u32,
    pub enabled: bool,
    /// VBA expression; the breakpoint fires only when it is True.
    pub condition: Option<String>,
    pub hit_condition: Option<HitCondition>,
    /// Logpoint: print this message (with `{expr}` interpolation) instead of pausing.
    pub log_message: Option<String>,
    pub hits: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum WatchKind {
    Expression,
    BreakWhenTrue,
    BreakWhenChanged,
}

#[derive(Debug, Clone)]
pub struct Watch {
    pub id: WatchId,
    pub expression: String,
    pub kind: WatchKind,
    /// Restrict evaluation to one procedure (the VBE "Context" setting).
    pub procedure: Option<String>,
    last_value: Option<String>,
}

/// A watch evaluated at a pause.
#[derive(Debug, Clone, PartialEq)]
pub struct WatchValue {
    pub id: WatchId,
    pub expression: String,
    pub value: Result<String, String>,
    pub type_name: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ErrorBreak {
    Never,
    /// VBE default: break only if no `On Error` handler is active.
    #[default]
    Unhandled,
    All,
}

#[derive(Debug, Clone, PartialEq)]
pub enum PauseReason {
    /// Paused before the first statement (`Debugger::stop_on_entry`).
    Entry,
    Step,
    Breakpoint(BreakpointId),
    Watch(WatchId),
    Stop,
    AssertFailed,
    Error(VbaError),
    PauseRequested,
}

/// Snapshot sent to the frontend. Contains only `Send` data, so it can cross
/// threads.
#[derive(Debug, Clone)]
pub struct PauseEvent {
    pub reason: PauseReason,
    pub location: Location,
    pub call_stack: Vec<Location>,
    pub watches: Vec<WatchValue>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Command {
    Continue,
    StepInto,
    StepOver,
    StepOut,
    RunToLine {
        module: Option<String>,
        line: u32,
    },
    /// Stop the program (VBE "Reset").
    Abort,
}

#[derive(Debug, Clone, PartialEq)]
enum Mode {
    Run,
    StepInto,
    StepOver(usize),
    StepOut(usize),
    RunTo(Option<String>, u32),
}

/// Breakpoints, watches and settings; reachable from the frontend while paused.
#[derive(Debug, Default)]
pub struct DebugState {
    pub breakpoints: Vec<Breakpoint>,
    pub watches: Vec<Watch>,
    pub error_break: ErrorBreak,
    next_id: u32,
}

impl DebugState {
    fn next_id(&mut self) -> u32 {
        self.next_id += 1;
        self.next_id
    }

    pub fn add_breakpoint(&mut self, module: Option<&str>, line: u32) -> BreakpointId {
        let id = self.next_id();
        self.breakpoints.push(Breakpoint {
            id,
            module: module.map(str::to_string),
            line,
            enabled: true,
            condition: None,
            hit_condition: None,
            log_message: None,
            hits: 0,
        });
        id
    }

    pub fn breakpoint_mut(&mut self, id: BreakpointId) -> Option<&mut Breakpoint> {
        self.breakpoints.iter_mut().find(|b| b.id == id)
    }

    /// Toggles a breakpoint on a line (F9). Returns `true` if one was added.
    pub fn toggle_breakpoint(&mut self, module: Option<&str>, line: u32) -> bool {
        let before = self.breakpoints.len();
        self.breakpoints
            .retain(|b| !(b.line == line && module_matches(b.module.as_deref(), module)));
        if self.breakpoints.len() == before {
            self.add_breakpoint(module, line);
            true
        } else {
            false
        }
    }

    pub fn remove_breakpoint(&mut self, id: BreakpointId) -> bool {
        let before = self.breakpoints.len();
        self.breakpoints.retain(|b| b.id != id);
        before != self.breakpoints.len()
    }

    pub fn clear_breakpoints(&mut self) {
        self.breakpoints.clear();
    }

    pub fn add_watch(&mut self, expression: &str, kind: WatchKind) -> WatchId {
        let id = self.next_id();
        self.watches.push(Watch {
            id,
            expression: expression.to_string(),
            kind,
            procedure: None,
            last_value: None,
        });
        id
    }

    pub fn watch_mut(&mut self, id: WatchId) -> Option<&mut Watch> {
        self.watches.iter_mut().find(|w| w.id == id)
    }

    pub fn remove_watch(&mut self, id: WatchId) -> bool {
        let before = self.watches.len();
        self.watches.retain(|w| w.id != id);
        before != self.watches.len()
    }
}

fn module_matches(a: Option<&str>, b: Option<&str>) -> bool {
    match (a, b) {
        (Some(a), Some(b)) => a.eq_ignore_ascii_case(b),
        _ => true,
    }
}

/// What the frontend can do while the program is paused.
pub struct Session<'s, 'c> {
    pub ctx: &'s mut DebugContext<'c>,
    pub state: &'s mut DebugState,
}

impl Session<'_, '_> {
    /// Evaluates in a frame (0 = innermost) and formats like the Watch window.
    pub fn evaluate(&mut self, frame: usize, expr: &str) -> Result<String, String> {
        self.ctx.eval_in_frame(frame, expr).map(|v| display(&v))
    }

    /// Locals window: name, value, type.
    pub fn locals(&self, frame: usize) -> Vec<(String, String, String)> {
        self.ctx
            .locals(frame)
            .into_iter()
            .map(|(n, v)| (n, display(&v), v.type_name()))
            .collect()
    }
}

/// Formats a value for debugger display (strings quoted, like the VBE).
pub fn display(v: &Value) -> String {
    match v {
        Value::String(_) | Value::Empty => format!("{v:?}"),
        Value::Array(a) => {
            let items: Vec<String> = a.data.iter().take(10).map(display).collect();
            let more = if a.data.len() > 10 { ", ..." } else { "" };
            format!("{:?} [{}{more}]", a.bounds, items.join(", "))
        }
        Value::Object(Some(o)) => {
            let props: Vec<String> = o
                .inspect()
                .iter()
                .take(8)
                .map(|(n, v)| format!("{n}: {}", display(v)))
                .collect();
            if props.is_empty() {
                format!("<{}>", o.type_name())
            } else {
                format!("<{}> {{ {} }}", o.type_name(), props.join(", "))
            }
        }
        Value::Udt(u) => {
            let fields: Vec<String> = u
                .fields
                .iter()
                .map(|f| format!("{}: {}", f.name, display(&f.value)))
                .collect();
            format!("<{}> {{ {} }}", u.type_name, fields.join(", "))
        }
        other => other.to_string(),
    }
}

pub trait Frontend {
    /// Called when execution pauses. Inspect or change state through
    /// `session`, then return how to resume.
    fn paused(&mut self, event: &PauseEvent, session: &mut Session<'_, '_>) -> Command;

    /// Output of logpoints.
    fn log(&mut self, message: &str, ctx: &mut DebugContext<'_>) {
        ctx.interpreter().host_mut().debug_print(message);
    }
}

/// The debugger. Install it with `Engine::set_debug_hook`.
pub struct Debugger<F: Frontend> {
    pub state: DebugState,
    frontend: F,
    mode: Mode,
    entry_pending: bool,
    pause_flag: Arc<AtomicBool>,
}

impl<F: Frontend> Debugger<F> {
    pub fn new(frontend: F) -> Self {
        Debugger {
            state: DebugState::default(),
            frontend,
            mode: Mode::Run,
            entry_pending: false,
            pause_flag: Arc::default(),
        }
    }

    /// Pause before the first statement (like pressing F8 to start a macro).
    pub fn stop_on_entry(mut self) -> Self {
        self.mode = Mode::StepInto;
        self.entry_pending = true;
        self
    }

    /// A thread-safe handle: storing `true` pauses at the next statement (Ctrl+Break).
    pub fn pause_handle(&self) -> Arc<AtomicBool> {
        self.pause_flag.clone()
    }

    pub fn frontend_mut(&mut self) -> &mut F {
        &mut self.frontend
    }

    fn pause(&mut self, ctx: &mut DebugContext<'_>, reason: PauseReason) -> DebugAction {
        let location = ctx.location();
        let procedure = location.procedure.clone();
        let watches = self
            .state
            .watches
            .iter()
            .map(|w| {
                let in_scope = w
                    .procedure
                    .as_deref()
                    .is_none_or(|p| p.eq_ignore_ascii_case(&procedure));
                let (value, type_name) = if !in_scope {
                    (Err("<Out of context>".to_string()), String::new())
                } else {
                    match ctx.eval(&w.expression) {
                        Ok(v) => (Ok(display(&v)), v.type_name()),
                        Err(e) => (Err(e), String::new()),
                    }
                };
                WatchValue {
                    id: w.id,
                    expression: w.expression.clone(),
                    value,
                    type_name,
                }
            })
            .collect();
        let event = PauseEvent {
            reason,
            location,
            call_stack: ctx.call_stack(),
            watches,
        };
        let command = self.frontend.paused(
            &event,
            &mut Session {
                ctx,
                state: &mut self.state,
            },
        );
        let depth = ctx.depth();
        self.mode = match command {
            Command::Continue => Mode::Run,
            Command::StepInto => Mode::StepInto,
            Command::StepOver => Mode::StepOver(depth),
            Command::StepOut => Mode::StepOut(depth),
            Command::RunToLine { module, line } => Mode::RunTo(module, line),
            Command::Abort => return DebugAction::Abort,
        };
        DebugAction::Continue
    }

    fn check_breakpoints(
        &mut self,
        ctx: &mut DebugContext<'_>,
        loc: &Location,
    ) -> Option<BreakpointId> {
        let mut fired = None;
        for i in 0..self.state.breakpoints.len() {
            let bp = &self.state.breakpoints[i];
            if !bp.enabled
                || bp.line != loc.line
                || !module_matches(bp.module.as_deref(), Some(&loc.module))
            {
                continue;
            }
            if let Some(cond) = bp.condition.clone() {
                // A condition that fails to evaluate breaks, so the user notices.
                if matches!(ctx.eval(&cond), Ok(v) if !v.to_bool().unwrap_or(true)) {
                    continue;
                }
            }
            let bp = &mut self.state.breakpoints[i];
            bp.hits += 1;
            let hit_ok = match bp.hit_condition {
                None => true,
                Some(HitCondition::Equals(n)) => bp.hits == n,
                Some(HitCondition::AtLeast(n)) => bp.hits >= n,
                Some(HitCondition::Multiple(n)) => n > 0 && bp.hits.is_multiple_of(n),
            };
            if !hit_ok {
                continue;
            }
            if let Some(msg) = bp.log_message.clone() {
                let text = interpolate(&msg, ctx);
                self.frontend.log(&text, ctx);
                continue;
            }
            fired.get_or_insert(bp.id);
        }
        fired
    }

    fn check_watches(&mut self, ctx: &mut DebugContext<'_>, procedure: &str) -> Option<WatchId> {
        let mut fired = None;
        for w in self
            .state
            .watches
            .iter_mut()
            .filter(|w| w.kind != WatchKind::Expression)
        {
            if !w
                .procedure
                .as_deref()
                .is_none_or(|p| p.eq_ignore_ascii_case(procedure))
            {
                continue;
            }
            let Ok(v) = ctx.eval(&w.expression) else {
                continue;
            };
            match w.kind {
                WatchKind::BreakWhenTrue => {
                    if v.to_bool().unwrap_or(false) {
                        fired.get_or_insert(w.id);
                    }
                }
                WatchKind::BreakWhenChanged => {
                    let now = format!("{}:{}", v.type_name(), display(&v));
                    if w.last_value.as_ref().is_some_and(|prev| *prev != now) {
                        fired.get_or_insert(w.id);
                    }
                    w.last_value = Some(now);
                }
                WatchKind::Expression => {}
            }
        }
        fired
    }
}

/// Replaces `{expr}` in a logpoint message with evaluated values.
fn interpolate(msg: &str, ctx: &mut DebugContext<'_>) -> String {
    let mut out = String::new();
    let mut rest = msg;
    while let Some(start) = rest.find('{') {
        out.push_str(&rest[..start]);
        let Some(end) = rest[start..].find('}') else {
            out.push_str(&rest[start..]);
            return out;
        };
        let expr = &rest[start + 1..start + end];
        match ctx.eval(expr) {
            Ok(v) => out.push_str(&v.to_string()),
            Err(e) => out.push_str(&format!("<{e}>")),
        }
        rest = &rest[start + end + 1..];
    }
    out.push_str(rest);
    out
}

impl<F: Frontend> DebugHook for Debugger<F> {
    fn on_statement(&mut self, ctx: &mut DebugContext<'_>) -> DebugAction {
        let loc = ctx.location();
        let depth = ctx.depth();
        let first = std::mem::replace(&mut self.mode, Mode::Run);
        let step_hit = match &first {
            Mode::Run => false,
            Mode::StepInto => true,
            Mode::StepOver(d) => depth <= *d,
            Mode::StepOut(d) => depth < *d,
            Mode::RunTo(m, line) => {
                loc.line == *line && module_matches(m.as_deref(), Some(&loc.module))
            }
        };
        if !step_hit {
            self.mode = first.clone();
        }
        let bp = self.check_breakpoints(ctx, &loc);
        let watch = if self.state.watches.is_empty() {
            None
        } else {
            self.check_watches(ctx, &loc.procedure)
        };
        let reason = if self.pause_flag.swap(false, Ordering::SeqCst) {
            Some(PauseReason::PauseRequested)
        } else if let Some(id) = bp {
            Some(PauseReason::Breakpoint(id))
        } else if let Some(id) = watch {
            Some(PauseReason::Watch(id))
        } else if step_hit {
            Some(if std::mem::take(&mut self.entry_pending) {
                PauseReason::Entry
            } else {
                PauseReason::Step
            })
        } else {
            None
        };
        match reason {
            Some(r) => self.pause(ctx, r),
            None => DebugAction::Continue,
        }
    }

    fn on_stop(&mut self, ctx: &mut DebugContext<'_>, kind: StopKind) -> DebugAction {
        let reason = match kind {
            StopKind::Stop => PauseReason::Stop,
            StopKind::AssertFailed => PauseReason::AssertFailed,
        };
        self.pause(ctx, reason)
    }

    fn on_error(
        &mut self,
        ctx: &mut DebugContext<'_>,
        err: &VbaError,
        handled: bool,
    ) -> DebugAction {
        let stop = match self.state.error_break {
            ErrorBreak::Never => false,
            ErrorBreak::Unhandled => !handled,
            ErrorBreak::All => true,
        };
        if stop {
            self.pause(ctx, PauseReason::Error(err.clone()))
        } else {
            DebugAction::Continue
        }
    }
}

// ---- cross-thread frontend ------------------------------------------------

/// A request from a debugger client running on another thread.
pub enum Request {
    Resume(Command),
    Evaluate {
        frame: usize,
        expression: String,
        reply: Sender<Result<String, String>>,
    },
    Execute {
        statements: String,
        reply: Sender<Result<(), String>>,
    },
    Locals {
        frame: usize,
        reply: Sender<Vec<(String, String, String)>>,
    },
    SetBreakpoint {
        module: Option<String>,
        line: u32,
        condition: Option<String>,
        reply: Sender<BreakpointId>,
    },
    RemoveBreakpoint(BreakpointId),
    AddWatch {
        expression: String,
        kind: WatchKind,
        reply: Sender<WatchId>,
    },
    RemoveWatch(WatchId),
    SetErrorBreak(ErrorBreak),
}

/// Forwards pauses to another thread and serves its requests until it
/// resumes. If the client hangs up, the program is aborted.
pub struct ChannelFrontend {
    events: Sender<PauseEvent>,
    requests: Receiver<Request>,
}

impl ChannelFrontend {
    /// Returns the frontend plus the client's ends of the channels.
    pub fn new() -> (Self, Receiver<PauseEvent>, Sender<Request>) {
        let (ev_tx, ev_rx) = std::sync::mpsc::channel();
        let (rq_tx, rq_rx) = std::sync::mpsc::channel();
        (
            ChannelFrontend {
                events: ev_tx,
                requests: rq_rx,
            },
            ev_rx,
            rq_tx,
        )
    }
}

impl Frontend for ChannelFrontend {
    fn paused(&mut self, event: &PauseEvent, s: &mut Session<'_, '_>) -> Command {
        if self.events.send(event.clone()).is_err() {
            return Command::Abort;
        }
        while let Ok(req) = self.requests.recv() {
            match req {
                Request::Resume(cmd) => return cmd,
                Request::Evaluate {
                    frame,
                    expression,
                    reply,
                } => {
                    let _ = reply.send(s.evaluate(frame, &expression));
                }
                Request::Execute { statements, reply } => {
                    let _ = reply.send(s.ctx.execute(&statements));
                }
                Request::Locals { frame, reply } => {
                    let _ = reply.send(s.locals(frame));
                }
                Request::SetBreakpoint {
                    module,
                    line,
                    condition,
                    reply,
                } => {
                    let id = s.state.add_breakpoint(module.as_deref(), line);
                    s.state.breakpoint_mut(id).unwrap().condition = condition;
                    let _ = reply.send(id);
                }
                Request::RemoveBreakpoint(id) => {
                    s.state.remove_breakpoint(id);
                }
                Request::AddWatch {
                    expression,
                    kind,
                    reply,
                } => {
                    let _ = reply.send(s.state.add_watch(&expression, kind));
                }
                Request::RemoveWatch(id) => {
                    s.state.remove_watch(id);
                }
                Request::SetErrorBreak(mode) => s.state.error_break = mode,
            }
        }
        Command::Abort
    }
}

/// A frontend driven by a closure; useful for tests and scripted sessions.
pub struct FnFrontend<T: FnMut(&PauseEvent, &mut Session<'_, '_>) -> Command>(pub T);

impl<T: FnMut(&PauseEvent, &mut Session<'_, '_>) -> Command> Frontend for FnFrontend<T> {
    fn paused(&mut self, event: &PauseEvent, session: &mut Session<'_, '_>) -> Command {
        (self.0)(event, session)
    }
}

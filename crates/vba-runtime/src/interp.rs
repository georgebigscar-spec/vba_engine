//! Tree-walking interpreter.

use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::HashMap;
use std::rc::Rc;

use vba_syntax::ast::*;

use crate::builtins;
use crate::debug::{DebugAction, DebugContext, DebugHook, Location, StopKind};
use crate::errobj::ErrObject;
use crate::error::{VbaError, VbaResult};
use crate::host::Host;
use crate::object::{same_object, Args, Collection, Invoke, ObjectRef};
use crate::value::{self, Value, VarType, VbArray};

const MAX_DEPTH: usize = 1000;

/// Storage for one variable. Shared (`Rc`) so `ByRef` arguments alias the
/// caller's variable.
#[derive(Debug, Clone)]
pub struct Variable {
    pub value: Value,
    pub ty: VarType,
    pub is_const: bool,
}

pub type Slot = Rc<RefCell<Variable>>;

fn slot(value: Value, ty: VarType) -> Slot {
    Rc::new(RefCell::new(Variable {
        value,
        ty,
        is_const: false,
    }))
}

#[derive(Debug, Clone, PartialEq)]
enum ErrMode {
    Off,
    ResumeNext,
    GoTo(String),
}

struct Frame {
    procedure: Rc<Procedure>,
    module: Rc<ModuleInfo>,
    locals: Vec<(String, Slot)>,
    line: u32,
    with_stack: Vec<Value>,
    on_error: ErrMode,
    in_handler: bool,
    /// Block (by address) and statement index to come back to with `Resume`.
    resume_target: Option<(usize, usize)>,
}

struct ModuleInfo {
    name: String,
    explicit: bool,
    text_compare: bool,
}

enum Flow {
    Normal,
    Exit(ExitKind),
    GoTo(String),
    Resume(ResumeKind),
}

pub struct Interpreter {
    procs: HashMap<String, (Rc<Procedure>, Rc<ModuleInfo>)>,
    globals: Vec<(String, Slot)>,
    frames: Vec<Frame>,
    host: Box<dyn Host>,
    hook: Option<Box<dyn DebugHook>>,
    err: Rc<ErrObject>,
    steps: u64,
    /// Set while the debugger evaluates a watch: unknown names are errors
    /// instead of implicit declarations, so inspection has no side effects.
    inspecting: bool,
}

impl Interpreter {
    pub fn new(host: impl Host + 'static) -> Self {
        Interpreter {
            procs: HashMap::new(),
            globals: Vec::new(),
            frames: Vec::new(),
            host: Box::new(host),
            hook: None,
            err: Rc::new(ErrObject::default()),
            steps: 0,
            inspecting: false,
        }
    }

    pub fn host_mut(&mut self) -> &mut dyn Host {
        self.host.as_mut()
    }

    pub fn set_debug_hook(
        &mut self,
        hook: Option<Box<dyn DebugHook>>,
    ) -> Option<Box<dyn DebugHook>> {
        std::mem::replace(&mut self.hook, hook)
    }

    /// Loads a parsed module: registers its procedures and initializes its
    /// module-level variables and constants.
    pub fn load_module(&mut self, module: Module) -> VbaResult<()> {
        let info = Rc::new(ModuleInfo {
            name: module.name.clone(),
            explicit: module.options.explicit,
            text_compare: module.options.compare_text,
        });
        for p in module.procedures {
            self.procs
                .insert(p.name.to_ascii_lowercase(), (Rc::new(p), info.clone()));
        }
        for decl in &module.declarations {
            match &decl.kind {
                StmtKind::Dim(vars) => {
                    for v in vars {
                        let s = self.new_var_slot(v)?;
                        self.globals.push((v.name.clone(), s));
                    }
                }
                StmtKind::Const(items) => {
                    for (name, e) in items {
                        let value = self.eval(e)?;
                        let s = slot(value, VarType::Variant);
                        s.borrow_mut().is_const = true;
                        self.globals.push((name.clone(), s));
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    pub fn procedure_names(&self) -> Vec<String> {
        let mut v: Vec<String> = self.procs.values().map(|(p, _)| p.name.clone()).collect();
        v.sort();
        v
    }

    pub fn procedure(&self, name: &str) -> Option<Rc<Procedure>> {
        self.procs
            .get(&name.to_ascii_lowercase())
            .map(|(p, _)| p.clone())
    }

    /// Runs a `Sub` or `Function` by name with the given arguments.
    pub fn run(&mut self, name: &str, args: Vec<Value>) -> VbaResult<Value> {
        let (proc, module) = self
            .procs
            .get(&name.to_ascii_lowercase())
            .cloned()
            .ok_or_else(|| VbaError::sub_not_defined(name))?;
        self.frames.clear();
        let slots = args.into_iter().map(ArgSource::Value).collect();
        let r = self.call_procedure(proc, module, slots, Vec::new());
        self.frames.clear();
        match r {
            Err(e) if e.is_end() => Ok(Value::Empty),
            other => other,
        }
    }

    // ---- debugger support ---------------------------------------------------

    pub fn depth(&self) -> usize {
        self.frames.len()
    }

    fn frame_index(&self, from_top: usize) -> Option<usize> {
        self.frames.len().checked_sub(from_top + 1)
    }

    pub fn location(&self, frame: usize) -> Option<Location> {
        let f = &self.frames[self.frame_index(frame)?];
        Some(Location {
            module: f.module.name.clone(),
            procedure: f.procedure.name.clone(),
            line: f.line,
        })
    }

    pub fn locals(&self, frame: usize) -> Vec<(String, Value)> {
        self.frame_index(frame)
            .map(|i| {
                self.frames[i]
                    .locals
                    .iter()
                    .map(|(n, s)| (n.clone(), s.borrow().value.clone()))
                    .collect()
            })
            .unwrap_or_default()
    }

    pub fn module_variables(&self) -> Vec<(String, Value)> {
        self.globals
            .iter()
            .map(|(n, s)| (n.clone(), s.borrow().value.clone()))
            .collect()
    }

    /// Runs `f` with frame `frame` (0 = innermost) temporarily on top.
    fn with_frame<T>(&mut self, frame: usize, f: impl FnOnce(&mut Self) -> T) -> Option<T> {
        let idx = self.frame_index(frame)?;
        let tail = self.frames.split_off(idx + 1);
        let hook = self.hook.take();
        let r = f(self);
        self.hook = hook;
        self.frames.extend(tail);
        Some(r)
    }

    pub fn eval_in_frame(&mut self, frame: usize, src: &str) -> Result<Value, String> {
        let expr = vba_syntax::parse_expression(src).map_err(|e| e.to_string())?;
        let was = std::mem::replace(&mut self.inspecting, true);
        let r = if self.frames.is_empty() {
            Some(self.eval(&expr))
        } else {
            self.with_frame(frame, |me| me.eval(&expr))
        };
        self.inspecting = was;
        r.ok_or_else(|| "no such frame".to_string())?
            .map_err(|e| e.description)
    }

    pub fn execute_in_frame(&mut self, frame: usize, src: &str) -> Result<(), String> {
        let wrapped = format!("Sub Immediate__()\n{src}\nEnd Sub\n");
        let module = vba_syntax::parse_module(&wrapped, "Immediate").map_err(|e| e.to_string())?;
        let body = module
            .procedures
            .into_iter()
            .next()
            .map(|p| p.body)
            .unwrap_or_default();
        let run = |me: &mut Self| me.exec_block(&body).map(|_| ()).map_err(|e| e.to_string());
        if self.frames.is_empty() {
            return run(self);
        }
        self.with_frame(frame, run)
            .ok_or_else(|| "no such frame".to_string())?
    }

    fn with_hook(
        &mut self,
        f: impl FnOnce(&mut dyn DebugHook, &mut DebugContext<'_>) -> DebugAction,
    ) -> VbaResult<()> {
        let Some(mut hook) = self.hook.take() else {
            return Ok(());
        };
        let action = f(hook.as_mut(), &mut DebugContext { interp: self });
        self.hook = Some(hook);
        match action {
            DebugAction::Continue => Ok(()),
            DebugAction::Abort => Err(VbaError::interrupted()),
        }
    }

    fn before_statement(&mut self, line: u32) -> VbaResult<()> {
        if let Some(f) = self.frames.last_mut() {
            f.line = line;
        }
        self.steps += 1;
        if self.steps.is_multiple_of(1024) && !self.host.poll() {
            return Err(VbaError::interrupted());
        }
        self.with_hook(|h, ctx| h.on_statement(ctx))
    }

    // ---- procedures ---------------------------------------------------------

    fn call_procedure(
        &mut self,
        proc: Rc<Procedure>,
        module: Rc<ModuleInfo>,
        positional: Vec<ArgSource>,
        named: Vec<(String, ArgSource)>,
    ) -> VbaResult<Value> {
        if self.frames.len() >= MAX_DEPTH {
            return Err(VbaError::new(28, "Out of stack space"));
        }
        if positional.len() > proc.params.len() {
            return Err(VbaError::wrong_arg_count());
        }
        let mut positional: Vec<Option<ArgSource>> = positional.into_iter().map(Some).collect();
        let mut named = named;
        let mut locals = Vec::with_capacity(proc.params.len() + 1);
        for (i, p) in proc.params.iter().enumerate() {
            let ty =
                p.ty.as_ref()
                    .map_or(VarType::Variant, |t| VarType::from_name(&t.name));
            let src = named
                .iter()
                .position(|(n, _)| n.eq_ignore_ascii_case(&p.name))
                .map(|k| named.remove(k).1)
                .or_else(|| positional.get_mut(i).and_then(Option::take))
                .filter(|s| !matches!(s, ArgSource::Value(Value::Missing)));
            let s = match src {
                Some(ArgSource::Slot(s)) if !p.by_val => s,
                Some(ArgSource::Slot(s)) => {
                    let v = s.borrow().value.clone();
                    slot(ty.coerce(v)?, ty)
                }
                Some(ArgSource::Value(v)) => slot(ty.coerce(v)?, ty),
                None if p.optional => {
                    let v = match &p.default {
                        Some(e) => ty.coerce(self.eval(e)?)?,
                        None if ty == VarType::Variant => Value::Missing,
                        None => ty.default_value(),
                    };
                    slot(v, ty)
                }
                None => {
                    return Err(VbaError::new(
                        449,
                        format!("Argument not optional: {}", p.name),
                    ))
                }
            };
            locals.push((p.name.clone(), s));
        }
        if let Some((n, _)) = named.first() {
            return Err(VbaError::new(448, format!("Named argument not found: {n}")));
        }
        if proc.kind == ProcKind::Function {
            let ty = proc
                .return_type
                .as_ref()
                .map_or(VarType::Variant, |t| VarType::from_name(&t.name));
            locals.push((proc.name.clone(), slot(ty.default_value(), ty)));
        }
        self.frames.push(Frame {
            procedure: proc.clone(),
            module,
            locals,
            line: proc.span.line,
            with_stack: Vec::new(),
            on_error: ErrMode::Off,
            in_handler: false,
            resume_target: None,
        });
        let _ = self.with_hook(|h, ctx| {
            h.on_enter(ctx);
            DebugAction::Continue
        });
        let mut result = self.exec_block(&proc.body).and_then(|flow| match flow {
            Flow::GoTo(l) => Err(VbaError::label_not_defined(&l)),
            Flow::Resume(_) => Err(VbaError::new(20, "Resume without error")),
            _ => Ok(()),
        });
        if result.is_ok() {
            result = self.before_statement(proc.end_line);
        }
        let _ = self.with_hook(|h, ctx| {
            h.on_exit(ctx);
            DebugAction::Continue
        });
        let frame = self.frames.pop().expect("frame");
        if result.is_ok() && (frame.on_error != ErrMode::Off || frame.in_handler) {
            // Exit Sub / End Sub from a procedure with a handler resets Err.
            self.err.clear();
        }
        result?;
        Ok(match proc.kind {
            ProcKind::Function => frame
                .locals
                .last()
                .map(|(_, s)| s.borrow().value.clone())
                .unwrap_or(Value::Empty),
            ProcKind::Sub => Value::Empty,
        })
    }

    // ---- statements ---------------------------------------------------------

    fn frame(&self) -> Option<&Frame> {
        self.frames.last()
    }

    fn frame_mut(&mut self) -> Option<&mut Frame> {
        self.frames.last_mut()
    }

    fn handlers_active(&self) -> bool {
        self.frames
            .iter()
            .any(|f| f.on_error != ErrMode::Off && !f.in_handler)
    }

    fn exec_block(&mut self, stmts: &[Stmt]) -> VbaResult<Flow> {
        let block_id = stmts.as_ptr() as usize;
        let mut pc = 0;
        while pc < stmts.len() {
            match self.exec_stmt(&stmts[pc]) {
                Ok(Flow::Normal) => pc += 1,
                Ok(Flow::GoTo(label)) => match find_label(stmts, &label) {
                    Some(i) => pc = i,
                    None => return Ok(Flow::GoTo(label)),
                },
                Ok(Flow::Resume(kind)) => {
                    let target = self.frame().and_then(|f| f.resume_target);
                    match target {
                        Some((id, at)) if id == block_id => {
                            pc = match &kind {
                                ResumeKind::Same => at,
                                ResumeKind::Next => at + 1,
                                ResumeKind::Label(l) => find_label(stmts, l)
                                    .ok_or_else(|| VbaError::label_not_defined(l))?,
                            };
                            self.err.clear();
                            if let Some(f) = self.frame_mut() {
                                f.in_handler = false;
                                f.resume_target = None;
                            }
                        }
                        _ => return Ok(Flow::Resume(kind)),
                    }
                }
                Ok(flow) => return Ok(flow),
                Err(e) if !e.trappable => return Err(e),
                Err(mut e) => {
                    if !e.debugger_notified {
                        // Report each error to the debugger once, at the innermost block.
                        let handled = self.handlers_active();
                        self.with_hook(|h, ctx| h.on_error(ctx, &e, handled))?;
                        e.debugger_notified = true;
                    }
                    let Some(frame) = self.frames.last_mut() else {
                        return Err(e);
                    };
                    match frame.on_error.clone() {
                        ErrMode::ResumeNext => {
                            self.err.set(&e);
                            pc += 1;
                        }
                        ErrMode::GoTo(label) if !frame.in_handler => {
                            match find_label(stmts, &label) {
                                Some(i) => {
                                    frame.in_handler = true;
                                    frame.resume_target = Some((block_id, pc));
                                    self.err.set(&e);
                                    pc = i;
                                }
                                None => return Err(e),
                            }
                        }
                        _ => return Err(e),
                    }
                }
            }
        }
        Ok(Flow::Normal)
    }

    fn exec_stmt(&mut self, stmt: &Stmt) -> VbaResult<Flow> {
        let line = stmt.span.line;
        if !matches!(
            stmt.kind,
            StmtKind::Label(_) | StmtKind::Dim(_) | StmtKind::Const(_)
        ) {
            self.before_statement(line)?;
        } else if let Some(f) = self.frame_mut() {
            f.line = line;
        }
        self.exec_stmt_kind(&stmt.kind)
            .map_err(|e| e.with_line(line))
    }

    fn exec_stmt_kind(&mut self, kind: &StmtKind) -> VbaResult<Flow> {
        match kind {
            StmtKind::Label(_) => {}
            StmtKind::Dim(vars) => {
                for v in vars {
                    let s = self.new_var_slot(v)?;
                    match self.frame_mut() {
                        Some(f) => f.locals.push((v.name.clone(), s)),
                        None => self.globals.push((v.name.clone(), s)),
                    }
                }
            }
            StmtKind::Const(items) => {
                for (name, e) in items {
                    let v = self.eval(e)?;
                    let s = slot(v, VarType::Variant);
                    s.borrow_mut().is_const = true;
                    if let Some(f) = self.frame_mut() {
                        f.locals.push((name.clone(), s));
                    }
                }
            }
            StmtKind::ReDim { preserve, decls } => {
                for d in decls {
                    let bounds = self.eval_bounds(d.bounds.as_deref().unwrap_or(&[]))?;
                    let existing = self.find_slot(&d.name);
                    let elem_ty = d.ty.as_ref().map(|t| VarType::from_name(&t.name));
                    match existing {
                        Some(s) => {
                            let mut var = s.borrow_mut();
                            match (&mut var.value, *preserve) {
                                (Value::Array(a), true) => a.redim_preserve(bounds)?,
                                (Value::Array(a), false) => {
                                    **a = VbArray::new(bounds, elem_ty.unwrap_or(a.elem_type))
                                }
                                (v, _) => {
                                    *v = Value::Array(Box::new(VbArray::new(
                                        bounds,
                                        elem_ty.unwrap_or(VarType::Variant),
                                    )))
                                }
                            }
                        }
                        None => {
                            let arr = Value::Array(Box::new(VbArray::new(
                                bounds,
                                elem_ty.unwrap_or(VarType::Variant),
                            )));
                            self.declare_implicit(&d.name, arr)?;
                        }
                    }
                }
            }
            StmtKind::Assign { target, value } => {
                let v = self.eval(value)?;
                self.assign(target, v, false)?;
            }
            StmtKind::SetAssign { target, value } => {
                let v = self.eval(value)?;
                if !matches!(v, Value::Object(_)) {
                    return Err(VbaError::object_required());
                }
                self.assign(target, v, true)?;
            }
            StmtKind::Call { callee, args } => return self.exec_call(callee, args),
            StmtKind::If {
                branches,
                else_body,
            } => {
                for (cond, body) in branches {
                    let c = self.eval(cond)?;
                    if truthy(&c)? {
                        return self.exec_block(body);
                    }
                }
                if let Some(body) = else_body {
                    return self.exec_block(body);
                }
            }
            StmtKind::For {
                var,
                from,
                to,
                step,
                body,
            } => {
                let start = self.eval(from)?;
                let end = self.eval(to)?.into_scalar()?;
                let step = match step {
                    Some(s) => self.eval(s)?.into_scalar()?,
                    None => Value::Integer(1),
                };
                let counter = match self.find_slot(var) {
                    Some(s) => s,
                    None => self.declare_implicit(var, Value::Empty)?,
                };
                {
                    let mut c = counter.borrow_mut();
                    c.value = c.ty.coerce(start.into_scalar()?)?;
                }
                let descending = step.to_f64()? < 0.0;
                loop {
                    let cur = counter.borrow().value.clone();
                    let ord = value::compare(&cur, &end, false)?;
                    if (!descending && ord == Some(Ordering::Greater))
                        || (descending && ord == Some(Ordering::Less))
                    {
                        break;
                    }
                    match self.exec_block(body)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    let mut c = counter.borrow_mut();
                    let next = value::add(&c.value, &step)?;
                    c.value = c.ty.coerce(next)?;
                }
            }
            StmtKind::ForEach {
                var,
                collection,
                body,
            } => {
                let items = match self.eval(collection)? {
                    Value::Array(a) => a.data,
                    Value::Object(Some(o)) => o.enumerate()?,
                    Value::Object(None) => return Err(VbaError::object_not_set()),
                    _ => return Err(VbaError::new(92, "For loop not initialized")),
                };
                let target = Expr::Ident(var.clone());
                for item in items {
                    let is_obj = matches!(item, Value::Object(_));
                    self.assign(&target, item, is_obj)?;
                    match self.exec_block(body)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                }
            }
            StmtKind::DoLoop { pre, post, body } => loop {
                if let Some(c) = pre {
                    if !self.loop_cond(c)? {
                        break;
                    }
                }
                match self.exec_block(body)? {
                    Flow::Normal => {}
                    Flow::Exit(ExitKind::Do) => break,
                    other => return Ok(other),
                }
                if let Some(c) = post {
                    if !self.loop_cond(c)? {
                        break;
                    }
                }
            },
            StmtKind::Select {
                subject,
                cases,
                else_body,
            } => {
                let subj = self.eval(subject)?.into_scalar()?;
                let text = self.text_compare();
                for (clauses, body) in cases {
                    for clause in clauses {
                        let hit = match clause {
                            CaseClause::Value(e) => {
                                let v = self.eval(e)?;
                                value::compare(&subj, &v, text)? == Some(Ordering::Equal)
                            }
                            CaseClause::Range(lo, hi) => {
                                let (lo, hi) = (self.eval(lo)?, self.eval(hi)?);
                                matches!(
                                    value::compare(&subj, &lo, text)?,
                                    Some(Ordering::Greater | Ordering::Equal)
                                ) && matches!(
                                    value::compare(&subj, &hi, text)?,
                                    Some(Ordering::Less | Ordering::Equal)
                                )
                            }
                            CaseClause::Is(op, e) => {
                                let v = self.eval(e)?;
                                compare_op(*op, value::compare(&subj, &v, text)?)
                            }
                        };
                        if hit {
                            return self.exec_block(body);
                        }
                    }
                }
                if let Some(body) = else_body {
                    return self.exec_block(body);
                }
            }
            StmtKind::With { object, body } => {
                let obj = self.eval(object)?;
                if let Some(f) = self.frame_mut() {
                    f.with_stack.push(obj);
                }
                let r = self.exec_block(body);
                if let Some(f) = self.frame_mut() {
                    f.with_stack.pop();
                }
                return r;
            }
            StmtKind::Exit(k) => return Ok(Flow::Exit(*k)),
            StmtKind::OnError(mode) => {
                self.err.clear();
                if let Some(f) = self.frame_mut() {
                    f.on_error = match mode {
                        OnError::ResumeNext => ErrMode::ResumeNext,
                        OnError::GoTo(l) => ErrMode::GoTo(l.clone()),
                        OnError::GoToZero => ErrMode::Off,
                    };
                    if matches!(mode, OnError::GoToZero) {
                        f.in_handler = false;
                    }
                }
            }
            StmtKind::Resume(k) => return Ok(Flow::Resume(k.clone())),
            StmtKind::GoTo(l) => return Ok(Flow::GoTo(l.clone())),
            StmtKind::End => return Err(VbaError::end()),
        }
        Ok(Flow::Normal)
    }

    fn loop_cond(&mut self, c: &LoopCond) -> VbaResult<bool> {
        let v = self.eval(&c.cond)?;
        Ok(truthy(&v)? != c.until)
    }

    fn exec_call(&mut self, callee: &Expr, args: &[Arg]) -> VbaResult<Flow> {
        // Debug.Print / Debug.Assert
        if let Expr::Member(obj, member) = callee {
            if matches!(&**obj, Expr::Ident(n) if n.eq_ignore_ascii_case("debug"))
                && self.find_slot("debug").is_none()
            {
                if member.eq_ignore_ascii_case("print") {
                    let mut parts = Vec::new();
                    for a in args {
                        parts.push(match &a.value {
                            Some(e) => match self.eval(e)? {
                                Value::Null => "Null".to_string(),
                                v => v.into_scalar()?.to_string(),
                            },
                            None => String::new(),
                        });
                    }
                    self.host.debug_print(&parts.join(" "));
                    return Ok(Flow::Normal);
                }
                if member.eq_ignore_ascii_case("assert") {
                    let cond = match args.first().and_then(|a| a.value.as_ref()) {
                        Some(e) => self.eval(e)?,
                        None => return Err(VbaError::wrong_arg_count()),
                    };
                    if !truthy(&cond)? {
                        self.with_hook(|h, ctx| h.on_stop(ctx, StopKind::AssertFailed))?;
                    }
                    return Ok(Flow::Normal);
                }
            }
        }
        if let Expr::Ident(name) = callee {
            if name.eq_ignore_ascii_case("stop")
                && args.is_empty()
                && !self.procs.contains_key("stop")
            {
                self.with_hook(|h, ctx| h.on_stop(ctx, StopKind::Stop))?;
                return Ok(Flow::Normal);
            }
        }
        self.eval_call(callee, args, Invoke::Method)?;
        Ok(Flow::Normal)
    }

    // ---- variables ----------------------------------------------------------

    fn find_slot(&self, name: &str) -> Option<Slot> {
        if let Some(f) = self.frame() {
            if let Some((_, s)) = f
                .locals
                .iter()
                .rev()
                .find(|(n, _)| n.eq_ignore_ascii_case(name))
            {
                return Some(s.clone());
            }
        }
        self.globals
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, s)| s.clone())
    }

    fn explicit(&self) -> bool {
        self.frame().is_some_and(|f| f.module.explicit)
    }

    fn text_compare(&self) -> bool {
        self.frame().is_some_and(|f| f.module.text_compare)
    }

    fn declare_implicit(&mut self, name: &str, value: Value) -> VbaResult<Slot> {
        if self.explicit() || self.inspecting {
            return Err(VbaError::variable_not_defined(name));
        }
        let s = slot(value, VarType::Variant);
        match self.frame_mut() {
            Some(f) => f.locals.push((name.to_string(), s.clone())),
            None => self.globals.push((name.to_string(), s.clone())),
        }
        Ok(s)
    }

    fn new_var_slot(&mut self, v: &VarDecl) -> VbaResult<Slot> {
        let ty =
            v.ty.as_ref()
                .map_or(VarType::Variant, |t| VarType::from_name(&t.name));
        let value = match (&v.bounds, &v.ty) {
            (Some(bounds), _) => {
                Value::Array(Box::new(VbArray::new(self.eval_bounds(bounds)?, ty)))
            }
            (None, Some(t)) if t.is_new => self.new_object(&t.name)?,
            _ => ty.default_value(),
        };
        let ty = if v.bounds.is_some() {
            VarType::Variant
        } else {
            ty
        };
        Ok(slot(value, ty))
    }

    fn eval_bounds(&mut self, bounds: &[(Option<Expr>, Expr)]) -> VbaResult<Vec<(i64, i64)>> {
        bounds
            .iter()
            .map(|(lo, hi)| {
                let lo = match lo {
                    Some(e) => self.eval(e)?.to_i64()?,
                    None => 0,
                };
                Ok((lo, self.eval(hi)?.to_i64()?))
            })
            .collect()
    }

    fn new_object(&mut self, class: &str) -> VbaResult<Value> {
        if class.eq_ignore_ascii_case("collection") || class.eq_ignore_ascii_case("vba.collection")
        {
            return Ok(Value::from(Collection::new()));
        }
        self.host.new_object(class)
    }

    fn with_object(&self) -> VbaResult<ObjectRef> {
        self.frame()
            .and_then(|f| f.with_stack.last())
            .ok_or_else(|| VbaError::new(-1, "Invalid or unqualified reference"))?
            .as_object()
    }

    fn assign(&mut self, target: &Expr, v: Value, is_set: bool) -> VbaResult<()> {
        let kind = |v: Value| {
            if is_set {
                Invoke::Set(v)
            } else {
                Invoke::Let(v)
            }
        };
        match target {
            Expr::Ident(name) => {
                let s = match self.find_slot(name) {
                    Some(s) => s,
                    None => {
                        // `ActiveCell = 5`: a property of the host's global object.
                        if let Some(app) = self.host_global_object() {
                            match app.invoke(name, kind(v.clone()), Args::default()) {
                                Err(e) if e.number == 438 => {}
                                other => return other.map(|_| ()),
                            }
                        }
                        self.declare_implicit(name, Value::Empty)?
                    }
                };
                let mut var = s.borrow_mut();
                if var.is_const {
                    return Err(VbaError::new(-1, "Assignment to constant not permitted"));
                }
                if is_set {
                    if !matches!(var.ty, VarType::Variant | VarType::Object) {
                        return Err(VbaError::type_mismatch());
                    }
                    var.value = v;
                } else if let Value::Object(Some(o)) = &var.value {
                    // Let on an object variable assigns its default member.
                    let o = o.clone();
                    drop(var);
                    o.invoke("", Invoke::Let(v.into_scalar()?), Args::default())?;
                } else {
                    var.value = var.ty.coerce(v.into_scalar()?)?;
                }
                Ok(())
            }
            Expr::Member(obj, name) => {
                let o = self.eval(obj)?.as_object()?;
                let v = if is_set { v } else { v.into_scalar()? };
                o.invoke(name, kind(v), Args::default()).map(|_| ())
            }
            Expr::WithMember(name) => {
                let o = self.with_object()?;
                let v = if is_set { v } else { v.into_scalar()? };
                o.invoke(name, kind(v), Args::default()).map(|_| ())
            }
            Expr::Index(callee, args) => {
                let args = self.eval_args(args)?;
                let v = if is_set { v } else { v.into_scalar()? };
                match &**callee {
                    Expr::Ident(name) => {
                        if let Some(s) = self.find_slot(name) {
                            let mut var = s.borrow_mut();
                            return match &mut var.value {
                                Value::Array(a) => {
                                    let idx = indices(&args)?;
                                    let off = a.offset(&idx)?;
                                    a.data[off] = a.elem_type.coerce(v)?;
                                    Ok(())
                                }
                                Value::Object(Some(o)) => {
                                    let o = o.clone();
                                    drop(var);
                                    o.invoke("", kind(v), args).map(|_| ())
                                }
                                _ => Err(VbaError::type_mismatch()),
                            };
                        }
                        let app = self
                            .host_global_object()
                            .ok_or_else(|| VbaError::sub_not_defined(name))?;
                        app.invoke(name, kind(v), args).map(|_| ())
                    }
                    Expr::Member(obj, name) => {
                        let o = self.eval(obj)?.as_object()?;
                        o.invoke(name, kind(v), args).map(|_| ())
                    }
                    Expr::WithMember(name) => {
                        self.with_object()?.invoke(name, kind(v), args).map(|_| ())
                    }
                    other => {
                        let o = self.eval(other)?.as_object()?;
                        o.invoke("", kind(v), args).map(|_| ())
                    }
                }
            }
            _ => Err(VbaError::new(-1, "Invalid assignment target")),
        }
    }

    fn host_global_object(&self) -> Option<ObjectRef> {
        self.host.global_object()
    }

    // ---- expressions --------------------------------------------------------

    pub fn eval(&mut self, e: &Expr) -> VbaResult<Value> {
        Ok(match e {
            Expr::Int(v) => match *v {
                v if (i16::MIN as i64..=i16::MAX as i64).contains(&v) => Value::Integer(v as i16),
                v if (i32::MIN as i64..=i32::MAX as i64).contains(&v) => Value::Long(v as i32),
                v => Value::Double(v as f64),
            },
            Expr::Float(v) => Value::Double(*v),
            Expr::Str(s) => Value::String(s.clone()),
            Expr::Date(s) => {
                Value::Date(crate::date::parse_date(s).ok_or_else(VbaError::type_mismatch)?)
            }
            Expr::Bool(b) => Value::Boolean(*b),
            Expr::Nothing => Value::Object(None),
            Expr::Null => Value::Null,
            Expr::Empty => Value::Empty,
            Expr::New(class) => self.new_object(class)?,
            Expr::Ident(_) | Expr::Member(..) | Expr::WithMember(_) => {
                self.eval_call(e, &[], Invoke::Get)?
            }
            Expr::Index(callee, args) => self.eval_call(callee, args, Invoke::Get)?,
            Expr::Unary(UnOp::Neg, a) => value::neg(&self.eval(a)?.into_scalar()?)?,
            Expr::Unary(UnOp::Not, a) => value::not(&self.eval(a)?.into_scalar()?)?,
            Expr::Binary(op, a, b) => {
                let (a, b) = (self.eval(a)?, self.eval(b)?);
                self.binary(*op, a, b)?
            }
        })
    }

    fn binary(&mut self, op: BinOp, a: Value, b: Value) -> VbaResult<Value> {
        if op == BinOp::Is {
            return Ok(Value::Boolean(match (&a, &b) {
                (Value::Object(None), Value::Object(None)) => true,
                (Value::Object(Some(x)), Value::Object(Some(y))) => same_object(x, y),
                (Value::Object(_), Value::Object(_)) => false,
                _ => return Err(VbaError::object_required()),
            }));
        }
        let (a, b) = (a.into_scalar()?, b.into_scalar()?);
        let text = self.text_compare();
        Ok(match op {
            BinOp::Add => value::add(&a, &b)?,
            BinOp::Sub => value::sub(&a, &b)?,
            BinOp::Mul => value::mul(&a, &b)?,
            BinOp::Div => value::div(&a, &b)?,
            BinOp::IntDiv => value::int_div(&a, &b)?,
            BinOp::Mod => value::modulo(&a, &b)?,
            BinOp::Pow => value::pow(&a, &b)?,
            BinOp::Concat => value::concat(&a, &b)?,
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                match value::compare(&a, &b, text)? {
                    None => Value::Null,
                    ord => Value::Boolean(compare_op(op, ord)),
                }
            }
            BinOp::Like => {
                if matches!(a, Value::Null) || matches!(b, Value::Null) {
                    Value::Null
                } else {
                    Value::Boolean(value::like(
                        &a.to_string_value()?,
                        &b.to_string_value()?,
                        text,
                    ))
                }
            }
            BinOp::And => value::logical(|x, y| x & y, &a, &b)?,
            BinOp::Or => value::logical(|x, y| x | y, &a, &b)?,
            BinOp::Xor => value::logical(|x, y| x ^ y, &a, &b)?,
            BinOp::Eqv => value::logical(|x, y| !(x ^ y), &a, &b)?,
            BinOp::Imp => value::logical(|x, y| !x | y, &a, &b)?,
            BinOp::Is => unreachable!(),
        })
    }

    fn eval_args(&mut self, args: &[Arg]) -> VbaResult<Args> {
        let mut out = Args::default();
        for a in args {
            let v = match &a.value {
                Some(e) => self.eval(e)?,
                None => Value::Missing,
            };
            match &a.name {
                Some(n) => out.named.push((n.clone(), v)),
                None => out.positional.push(v),
            }
        }
        Ok(out)
    }

    fn arg_sources(
        &mut self,
        args: &[Arg],
    ) -> VbaResult<(Vec<ArgSource>, Vec<(String, ArgSource)>)> {
        let mut positional = Vec::new();
        let mut named = Vec::new();
        for a in args {
            let src = match &a.value {
                Some(Expr::Ident(n)) => match self.find_slot(n) {
                    Some(s) => ArgSource::Slot(s),
                    None => ArgSource::Value(self.eval(a.value.as_ref().unwrap())?),
                },
                Some(e) => ArgSource::Value(self.eval(e)?),
                None => ArgSource::Value(Value::Missing),
            };
            match &a.name {
                Some(n) => named.push((n.clone(), src)),
                None => positional.push(src),
            }
        }
        Ok((positional, named))
    }

    /// Evaluates `callee(args)`: array indexing, procedure call, built-in,
    /// object member, or host global — resolved at runtime like VBA does.
    fn eval_call(&mut self, callee: &Expr, args: &[Arg], kind: Invoke) -> VbaResult<Value> {
        match callee {
            Expr::Ident(name) => {
                let current_fn = self
                    .frame()
                    .is_some_and(|f| f.procedure.name.eq_ignore_ascii_case(name));
                if let Some(s) = self.find_slot(name) {
                    // Inside a Function, `Fn(x)` with arguments is a recursive call.
                    if !(current_fn && !args.is_empty()) {
                        if args.is_empty() {
                            return Ok(s.borrow().value.clone());
                        }
                        let evaluated = self.eval_args(args)?;
                        let var = s.borrow();
                        return match &var.value {
                            Value::Array(a) => Ok(a.data[a.offset(&indices(&evaluated)?)?].clone()),
                            Value::Object(Some(o)) => {
                                let o = o.clone();
                                drop(var);
                                o.invoke("", kind, evaluated)
                            }
                            _ => Err(VbaError::type_mismatch()),
                        };
                    }
                }
                if let Some((proc, module)) = self.procs.get(&name.to_ascii_lowercase()).cloned() {
                    let (positional, named) = self.arg_sources(args)?;
                    return self.call_procedure(proc, module, positional, named);
                }
                if name.eq_ignore_ascii_case("err") {
                    return self.index_value(Value::from(self.err.clone()), args, kind);
                }
                let evaluated = self.eval_args(args)?;
                let text = self.text_compare();
                if let Some(r) = builtins::call(name, &evaluated, self.host.as_mut(), text) {
                    return r;
                }
                if evaluated.is_empty() {
                    if let Some(c) = builtins::constant(name) {
                        return Ok(c);
                    }
                }
                if let Some(v) = self.host.global(name) {
                    return self.index_value_evaluated(v, evaluated, kind);
                }
                if let Some(app) = self.host_global_object() {
                    match app.invoke(name, kind.clone(), evaluated.clone()) {
                        Err(e) if e.number == 438 => {}
                        other => return other,
                    }
                }
                if self.explicit()
                    || self.inspecting
                    || !evaluated.is_empty()
                    || matches!(kind, Invoke::Method)
                {
                    if evaluated.is_empty() && !matches!(kind, Invoke::Method) {
                        return Err(VbaError::variable_not_defined(name));
                    }
                    return Err(VbaError::sub_not_defined(name));
                }
                // Implicitly declared variable, read before assignment.
                self.declare_implicit(name, Value::Empty)?;
                Ok(Value::Empty)
            }
            Expr::Member(obj, member) => {
                if let Expr::Ident(q) = &**obj {
                    if q.eq_ignore_ascii_case("vba") && self.find_slot(q).is_none() {
                        return self.eval_call(&Expr::Ident(member.clone()), args, kind);
                    }
                }
                let o = self.eval(obj)?.as_object()?;
                let evaluated = self.eval_args(args)?;
                o.invoke(member, kind, evaluated)
            }
            Expr::WithMember(member) => {
                let o = self.with_object()?;
                let evaluated = self.eval_args(args)?;
                o.invoke(member, kind, evaluated)
            }
            other => {
                let v = self.eval(other)?;
                self.index_value(v, args, kind)
            }
        }
    }

    fn index_value(&mut self, v: Value, args: &[Arg], kind: Invoke) -> VbaResult<Value> {
        if args.is_empty() {
            return Ok(v);
        }
        let evaluated = self.eval_args(args)?;
        self.index_value_evaluated(v, evaluated, kind)
    }

    fn index_value_evaluated(&mut self, v: Value, args: Args, kind: Invoke) -> VbaResult<Value> {
        if args.is_empty() {
            return Ok(v);
        }
        match v {
            Value::Array(a) => Ok(a.data[a.offset(&indices(&args)?)?].clone()),
            Value::Object(Some(o)) => o.invoke("", kind, args),
            Value::Object(None) => Err(VbaError::object_not_set()),
            _ => Err(VbaError::type_mismatch()),
        }
    }
}

enum ArgSource {
    Slot(Slot),
    Value(Value),
}

fn truthy(v: &Value) -> VbaResult<bool> {
    match v {
        Value::Null => Ok(false),
        v => v.to_bool(),
    }
}

fn compare_op(op: BinOp, ord: Option<Ordering>) -> bool {
    let Some(ord) = ord else { return false };
    match op {
        BinOp::Eq => ord == Ordering::Equal,
        BinOp::Ne => ord != Ordering::Equal,
        BinOp::Lt => ord == Ordering::Less,
        BinOp::Le => ord != Ordering::Greater,
        BinOp::Gt => ord == Ordering::Greater,
        BinOp::Ge => ord != Ordering::Less,
        _ => false,
    }
}

fn indices(args: &Args) -> VbaResult<Vec<i64>> {
    args.positional.iter().map(|v| v.to_i64()).collect()
}

fn find_label(stmts: &[Stmt], label: &str) -> Option<usize> {
    stmts
        .iter()
        .position(|s| matches!(&s.kind, StmtKind::Label(l) if l.eq_ignore_ascii_case(label)))
}

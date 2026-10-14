//! `vba` — run a VBA module or debug it interactively.
//!
//! ```text
//! vba run   Module1.bas [--entry Main]
//! vba debug Module1.bas [--entry Main] [--break-on never|unhandled|all]
//! ```

use std::io::{self, BufRead, Write};
use std::process::ExitCode;

use vba_debug::{
    Command, Debugger, ErrorBreak, Frontend, PauseEvent, PauseReason, Session, WatchKind,
};
use vba_runtime::{Engine, StdHost};

const HELP: &str = "\
Commands (VBE shortcut in brackets):
  s            step into            [F8]
  n            step over            [Shift+F8]
  o            step out             [Ctrl+Shift+F8]
  c            continue             [F5]
  r <line>     run to line          [Ctrl+F8]
  b <line> [if <cond>]  toggle / add conditional breakpoint  [F9]
  log <line> <message with {expr}>   add a logpoint
  bl           list breakpoints
  bd <id>      delete breakpoint
  w <expr>     add watch
  wt <expr>    watch: break when value is True
  wc <expr>    watch: break when value changes
  wd <id>      delete watch
  ? <expr>     evaluate (Quick Watch)
  ! <stmt>     execute in the current frame (Immediate window)
  l [frame]    locals
  bt           call stack           [Ctrl+L]
  list         show source around the current line
  q            stop the program     [Reset]";

struct Console {
    source: Vec<String>,
}

impl Console {
    fn show_source(&self, line: u32, radius: u32) {
        let from = line.saturating_sub(radius).max(1);
        let to = (line + radius).min(self.source.len() as u32);
        for l in from..=to {
            let marker = if l == line { "=>" } else { "  " };
            eprintln!("{marker}{l:4} | {}", self.source[(l - 1) as usize]);
        }
    }
}

impl Frontend for Console {
    fn paused(&mut self, ev: &PauseEvent, s: &mut Session<'_, '_>) -> Command {
        let why = match &ev.reason {
            PauseReason::Entry => "entry".to_string(),
            PauseReason::Step => "step".to_string(),
            PauseReason::Breakpoint(id) => format!("breakpoint #{id}"),
            PauseReason::Watch(id) => format!("watch #{id}"),
            PauseReason::Stop => "Stop statement".to_string(),
            PauseReason::AssertFailed => "Debug.Assert failed".to_string(),
            PauseReason::Error(e) => e.to_string(),
            PauseReason::PauseRequested => "paused".to_string(),
        };
        eprintln!(
            "-- {} [{}:{}] ({why})",
            ev.location.procedure, ev.location.module, ev.location.line
        );
        self.show_source(ev.location.line, 0);
        for w in &ev.watches {
            match &w.value {
                Ok(v) => eprintln!(
                    "   watch #{} {} = {v} : {}",
                    w.id, w.expression, w.type_name
                ),
                Err(e) => eprintln!("   watch #{} {} = <{e}>", w.id, w.expression),
            }
        }
        let stdin = io::stdin();
        loop {
            eprint!("(vba) ");
            let _ = io::stderr().flush();
            let mut input = String::new();
            if stdin.lock().read_line(&mut input).unwrap_or(0) == 0 {
                return Command::Abort;
            }
            let input = input.trim();
            let (cmd, rest) = input
                .split_once(' ')
                .map_or((input, ""), |(c, r)| (c, r.trim()));
            match cmd {
                "" | "s" => return Command::StepInto,
                "n" => return Command::StepOver,
                "o" => return Command::StepOut,
                "c" => return Command::Continue,
                "q" => return Command::Abort,
                "r" => match rest.parse() {
                    Ok(line) => return Command::RunToLine { module: None, line },
                    Err(_) => eprintln!("usage: r <line>"),
                },
                "b" => {
                    let (line, cond) = rest
                        .split_once(" if ")
                        .map_or((rest, None), |(l, c)| (l, Some(c.trim().to_string())));
                    match line.trim().parse() {
                        Ok(line) if cond.is_some() => {
                            let id = s.state.add_breakpoint(None, line);
                            s.state.breakpoint_mut(id).unwrap().condition = cond;
                            eprintln!("breakpoint #{id} at line {line}");
                        }
                        Ok(line) => {
                            let added = s.state.toggle_breakpoint(None, line);
                            eprintln!(
                                "breakpoint at line {line} {}",
                                if added { "set" } else { "cleared" }
                            );
                        }
                        Err(_) => eprintln!("usage: b <line> [if <condition>]"),
                    }
                }
                "log" => match rest
                    .split_once(' ')
                    .and_then(|(l, m)| Some((l.parse().ok()?, m)))
                {
                    Some((line, msg)) => {
                        let id = s.state.add_breakpoint(None, line);
                        s.state.breakpoint_mut(id).unwrap().log_message = Some(msg.to_string());
                        eprintln!("logpoint #{id} at line {line}");
                    }
                    None => eprintln!("usage: log <line> <message>"),
                },
                "bl" => {
                    for b in &s.state.breakpoints {
                        let extra = b
                            .condition
                            .as_ref()
                            .map(|c| format!(" if {c}"))
                            .unwrap_or_default();
                        let log = b
                            .log_message
                            .as_ref()
                            .map(|m| format!(" log \"{m}\""))
                            .unwrap_or_default();
                        eprintln!("#{} line {}{extra}{log} (hits: {})", b.id, b.line, b.hits);
                    }
                }
                "bd" => match rest.parse() {
                    Ok(id) if s.state.remove_breakpoint(id) => eprintln!("deleted #{id}"),
                    _ => eprintln!("no such breakpoint"),
                },
                "w" | "wt" | "wc" => {
                    let kind = match cmd {
                        "wt" => WatchKind::BreakWhenTrue,
                        "wc" => WatchKind::BreakWhenChanged,
                        _ => WatchKind::Expression,
                    };
                    let id = s.state.add_watch(rest, kind);
                    eprintln!(
                        "watch #{id}: {rest} = {}",
                        s.evaluate(0, rest).unwrap_or_else(|e| format!("<{e}>"))
                    );
                }
                "wd" => match rest.parse() {
                    Ok(id) if s.state.remove_watch(id) => eprintln!("deleted watch #{id}"),
                    _ => eprintln!("no such watch"),
                },
                "?" | "p" => match s.evaluate(0, rest) {
                    Ok(v) => eprintln!("{v}"),
                    Err(e) => eprintln!("error: {e}"),
                },
                "!" => {
                    if let Err(e) = s.ctx.execute(rest) {
                        eprintln!("error: {e}");
                    }
                }
                "l" => {
                    let frame = rest.parse().unwrap_or(0);
                    for (name, value, ty) in s.locals(frame) {
                        eprintln!("  {name} = {value} : {ty}");
                    }
                }
                "bt" => {
                    for (i, loc) in ev.call_stack.iter().enumerate() {
                        eprintln!("  #{i} {}.{} line {}", loc.module, loc.procedure, loc.line);
                    }
                }
                "list" => self.show_source(ev.location.line, 5),
                "h" | "help" => eprintln!("{HELP}"),
                other => {
                    // A bare `?expr` (no space) as in the Immediate window.
                    if let Some(expr) = other.strip_prefix('?') {
                        let expr = format!("{expr} {rest}");
                        match s.evaluate(0, expr.trim()) {
                            Ok(v) => eprintln!("{v}"),
                            Err(e) => eprintln!("error: {e}"),
                        }
                    } else {
                        eprintln!("unknown command '{other}', try 'help'");
                    }
                }
            }
        }
    }
}

struct Options {
    debug: bool,
    file: String,
    entry: String,
    error_break: ErrorBreak,
}

fn parse_args() -> Result<Options, String> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let usage = "usage: vba <run|debug> <file.bas> [--entry Main] [--break-on never|unhandled|all]";
    let debug = match args.first().map(String::as_str) {
        Some("run") => false,
        Some("debug") => true,
        _ => return Err(usage.into()),
    };
    let file = args.get(1).cloned().ok_or(usage)?;
    let mut opts = Options {
        debug,
        file,
        entry: "Main".into(),
        error_break: ErrorBreak::Unhandled,
    };
    let mut i = 2;
    while i < args.len() {
        let val = args.get(i + 1).cloned().ok_or(usage)?;
        match args[i].as_str() {
            "--entry" => opts.entry = val,
            "--break-on" => {
                opts.error_break = match val.as_str() {
                    "never" => ErrorBreak::Never,
                    "all" => ErrorBreak::All,
                    _ => ErrorBreak::Unhandled,
                }
            }
            _ => return Err(usage.into()),
        }
        i += 2;
    }
    Ok(opts)
}

fn run(opts: Options) -> Result<(), String> {
    let source = std::fs::read_to_string(&opts.file).map_err(|e| format!("{}: {e}", opts.file))?;
    let module_name = std::path::Path::new(&opts.file)
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("Module1")
        .to_string();
    let mut engine = Engine::new(StdHost);
    engine
        .load_source(&source, &module_name)
        .map_err(|e| e.to_string())?;
    if opts.debug {
        eprintln!("Debugging {} — type 'help' for commands.", opts.file);
        let mut dbg = Debugger::new(Console {
            source: source.lines().map(str::to_string).collect(),
        })
        .stop_on_entry();
        dbg.state.error_break = opts.error_break;
        engine.set_debug_hook(dbg);
    }
    engine
        .run(&opts.entry, vec![])
        .map(|_| ())
        .map_err(|e| e.to_string())
}

fn main() -> ExitCode {
    let opts = match parse_args() {
        Ok(o) => o,
        Err(e) => {
            eprintln!("{e}");
            return ExitCode::from(2);
        }
    };
    // Deep VBA recursion maps onto Rust recursion; give it room.
    let worker = std::thread::Builder::new()
        .stack_size(256 << 20)
        .spawn(move || run(opts));
    match worker
        .map_err(|e| e.to_string())
        .and_then(|h| h.join().map_err(|_| "interpreter panicked".to_string()))
    {
        Ok(Ok(())) => ExitCode::SUCCESS,
        Ok(Err(e)) | Err(e) => {
            eprintln!("{e}");
            ExitCode::FAILURE
        }
    }
}

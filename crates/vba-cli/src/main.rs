//! `vba` — run a VBA module or debug it interactively.
//!
//! ```text
//! vba run   Module1.bas [Class1.cls ...] [--entry Main]
//! vba debug Module1.bas [Class1.cls ...] [--entry Main] [--break-on never|unhandled|all]
//! ```
//!
//! Excel macros run against an in-memory workbook: `--excel` starts with an
//! empty `Book1`, `--workbook in.xlsx` opens a file, `--save out.xlsx` writes
//! the active workbook after the macro finishes.

use std::io::{self, BufRead, Write};
use std::process::ExitCode;

use vba_debug::{
    Command, Debugger, ErrorBreak, Frontend, PauseEvent, PauseReason, Session, WatchKind,
};
use vba_excel::ExcelHost;
use vba_runtime::{Engine, StdHost};

const HELP: &str = "\
Commands (VBE shortcut in brackets):
  s            step into            [F8]
  n            step over            [Shift+F8]
  o            step out             [Ctrl+Shift+F8]
  c            continue             [F5]
  r <loc>      run to line          [Ctrl+F8]
  b <loc> [if <cond>]  toggle / add conditional breakpoint  [F9]
  log <loc> <message with {expr}>   add a logpoint
               <loc> is a line in the current module, or Module:line
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
    /// Module name and source lines.
    sources: Vec<(String, Vec<String>)>,
}

impl Console {
    fn show_source(&self, module: &str, line: u32, radius: u32) {
        let Some((_, source)) = self
            .sources
            .iter()
            .find(|(m, _)| m.eq_ignore_ascii_case(module))
        else {
            return;
        };
        let from = line.saturating_sub(radius).max(1);
        let to = (line + radius).min(source.len() as u32);
        for l in from..=to {
            let marker = if l == line { "=>" } else { "  " };
            eprintln!("{marker}{l:4} | {}", source[(l - 1) as usize]);
        }
    }
}

/// `12` (current module) or `Module:12`.
fn parse_loc(text: &str, current: &str) -> Option<(Option<String>, u32)> {
    match text.trim().split_once(':') {
        Some((m, l)) => Some((Some(m.trim().to_string()), l.trim().parse().ok()?)),
        None => Some((Some(current.to_string()), text.trim().parse().ok()?)),
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
        self.show_source(&ev.location.module, ev.location.line, 0);
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
                "r" => match parse_loc(rest, &ev.location.module) {
                    Some((module, line)) => return Command::RunToLine { module, line },
                    None => eprintln!("usage: r <line> | r <Module:line>"),
                },
                "b" => {
                    let (line, cond) = rest
                        .split_once(" if ")
                        .map_or((rest, None), |(l, c)| (l, Some(c.trim().to_string())));
                    match parse_loc(line, &ev.location.module) {
                        Some((module, line)) if cond.is_some() => {
                            let id = s.state.add_breakpoint(module.as_deref(), line);
                            s.state.breakpoint_mut(id).unwrap().condition = cond;
                            eprintln!("breakpoint #{id} at line {line}");
                        }
                        Some((module, line)) => {
                            let added = s.state.toggle_breakpoint(module.as_deref(), line);
                            eprintln!(
                                "breakpoint at line {line} {}",
                                if added { "set" } else { "cleared" }
                            );
                        }
                        None => eprintln!("usage: b <line> [if <condition>]"),
                    }
                }
                "log" => match rest
                    .split_once(' ')
                    .and_then(|(l, m)| Some((parse_loc(l, &ev.location.module)?, m)))
                {
                    Some(((module, line), msg)) => {
                        let id = s.state.add_breakpoint(module.as_deref(), line);
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
                "list" => self.show_source(&ev.location.module, ev.location.line, 5),
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
    files: Vec<String>,
    entry: String,
    error_break: ErrorBreak,
    excel: bool,
    workbook: Option<String>,
    save: Option<String>,
}

fn parse_args() -> Result<Options, String> {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let usage = "usage: vba <run|debug> <file.bas|file.cls>... [--entry Main] [--break-on never|unhandled|all]\n\
                 \x20      [--excel] [--workbook in.xlsx] [--save out.xlsx]";
    let debug = match args.first().map(String::as_str) {
        Some("run") => false,
        Some("debug") => true,
        _ => return Err(usage.into()),
    };
    let mut opts = Options {
        debug,
        files: Vec::new(),
        entry: "Main".into(),
        error_break: ErrorBreak::Unhandled,
        excel: false,
        workbook: None,
        save: None,
    };
    let mut i = 1;
    while i < args.len() && !args[i].starts_with("--") {
        opts.files.push(args[i].clone());
        i += 1;
    }
    if opts.files.is_empty() {
        return Err(usage.into());
    }
    while i < args.len() {
        if args[i] == "--excel" {
            opts.excel = true;
            i += 1;
            continue;
        }
        let val = args.get(i + 1).cloned().ok_or(usage)?;
        match args[i].as_str() {
            "--workbook" => opts.workbook = Some(val),
            "--save" => opts.save = Some(val),
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
    let excel = match &opts.workbook {
        Some(path) => Some(ExcelHost::open(path).map_err(|e| e.to_string())?),
        None if opts.excel || opts.save.is_some() => Some(ExcelHost::new()),
        None => None,
    }
    .map(|h| h.with_echo(true));
    let mut engine = match &excel {
        Some(host) => Engine::new(host.clone()),
        None => Engine::new(StdHost),
    };
    let mut sources = Vec::new();
    for file in &opts.files {
        let source = std::fs::read_to_string(file).map_err(|e| format!("{file}: {e}"))?;
        let path = std::path::Path::new(file);
        let stem = path
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("Module1");
        let is_class = path
            .extension()
            .is_some_and(|e| e.eq_ignore_ascii_case("cls"));
        let parsed = if is_class {
            vba_syntax::parse_class_module(&source, stem)
        } else {
            vba_syntax::parse_module(&source, stem)
        };
        let module = parsed.map_err(|e| format!("{file}: {e}"))?;
        sources.push((
            module.name.clone(),
            source.lines().map(str::to_string).collect(),
        ));
        engine
            .interpreter
            .load_module(module)
            .map_err(|e| format!("{file}: {e}"))?;
    }
    if opts.debug {
        eprintln!(
            "Debugging {} — type 'help' for commands.",
            opts.files.join(", ")
        );
        let mut dbg = Debugger::new(Console { sources }).stop_on_entry();
        dbg.state.error_break = opts.error_break;
        engine.set_debug_hook(dbg);
    }
    engine.run(&opts.entry, vec![]).map_err(|e| e.to_string())?;
    if let (Some(host), Some(path)) = (&excel, &opts.save) {
        host.save_as(path).map_err(|e| e.to_string())?;
        eprintln!("saved {path}");
    }
    Ok(())
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

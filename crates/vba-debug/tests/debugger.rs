use std::cell::RefCell;
use std::rc::Rc;

use vba_debug::*;
use vba_runtime::{CaptureHost, Engine};

const SRC: &str = r#"Function Square(x)
    Square = x * x
End Function

Sub Main()
    Dim i As Long, total As Long
    For i = 1 To 3
        total = total + Square(i)
    Next i
    Debug.Print total
End Sub
"#;

/// Runs SRC with a scripted frontend; returns (reason, line, procedure) per pause.
fn session(
    setup: impl FnOnce(&mut DebugState),
    entry: bool,
    mut script: impl FnMut(&PauseEvent, &mut Session<'_, '_>) -> Command + 'static,
) -> (Vec<(PauseReason, u32, String)>, Vec<String>) {
    let log = Rc::new(RefCell::new(Vec::new()));
    let log2 = log.clone();
    let frontend = FnFrontend(move |ev: &PauseEvent, s: &mut Session<'_, '_>| {
        log2.borrow_mut().push((
            ev.reason.clone(),
            ev.location.line,
            ev.location.procedure.clone(),
        ));
        script(ev, s)
    });
    let mut dbg = Debugger::new(frontend);
    if entry {
        dbg = dbg.stop_on_entry();
    }
    setup(&mut dbg.state);
    let host = CaptureHost::new();
    let mut engine = Engine::new(host.clone());
    engine.load_source(SRC, "Module1").unwrap();
    engine.set_debug_hook(dbg);
    let _ = engine.run("Main", vec![]);
    let pauses = log.borrow().clone();
    (pauses, host.lines())
}

fn lines(p: &[(PauseReason, u32, String)]) -> Vec<u32> {
    p.iter().map(|(_, l, _)| *l).collect()
}

#[test]
fn step_into_enters_functions() {
    let mut n = 0;
    let (p, _) = session(
        |_| {},
        true,
        move |_, _| {
            n += 1;
            if n < 6 {
                Command::StepInto
            } else {
                Command::Continue
            }
        },
    );
    assert_eq!(p[0].0, PauseReason::Entry);
    // For header, body line, into Square, its End Function, back to Next... (line 8 again for i=2)
    assert_eq!(lines(&p), vec![7, 8, 2, 3, 8, 2]);
}

#[test]
fn step_over_stays_in_main() {
    let (p, out) = session(|_| {}, true, |_, _| Command::StepOver);
    assert_eq!(lines(&p), vec![7, 8, 8, 8, 10, 11]);
    assert!(p.iter().all(|(_, _, proc)| proc == "Main"));
    assert_eq!(out, vec!["14"]);
}

#[test]
fn step_out_returns_to_caller() {
    let mut n = 0;
    let (p, _) = session(
        |s| {
            s.add_breakpoint(None, 2);
        },
        false,
        move |_, _| {
            n += 1;
            if n == 1 {
                Command::StepOut
            } else {
                Command::Abort
            }
        },
    );
    assert_eq!(p[0].0, PauseReason::Breakpoint(1));
    assert_eq!((p[1].1, p[1].2.as_str()), (8, "Main"));
}

#[test]
fn conditional_breakpoint_and_eval() {
    let seen = Rc::new(RefCell::new(Vec::new()));
    let seen2 = seen.clone();
    let (p, _) = session(
        |s| {
            let id = s.add_breakpoint(Some("Module1"), 8);
            s.breakpoint_mut(id).unwrap().condition = Some("i = 2".into());
        },
        false,
        move |_, s| {
            seen2
                .borrow_mut()
                .push(s.evaluate(0, "total & \"/\" & i").unwrap());
            Command::Continue
        },
    );
    assert_eq!(lines(&p), vec![8]);
    assert_eq!(*seen.borrow(), vec!["\"1/2\""]);
}

#[test]
fn logpoint_prints_without_pausing() {
    let (p, out) = session(
        |s| {
            let id = s.add_breakpoint(None, 8);
            s.breakpoint_mut(id).unwrap().log_message = Some("i={i}".into());
        },
        false,
        |_, _| Command::Continue,
    );
    assert!(p.is_empty());
    assert_eq!(out, vec!["i=1", "i=2", "i=3", "14"]);
}

#[test]
fn watch_break_when_changed_and_true() {
    let (p, _) = session(
        |s| {
            s.add_watch("total", WatchKind::BreakWhenChanged);
        },
        false,
        |_, _| Command::Continue,
    );
    // total changes after each loop body; pause is on the statement after the change.
    assert_eq!(
        p.iter()
            .filter(|(r, _, _)| matches!(r, PauseReason::Watch(_)))
            .count(),
        3
    );

    let (p, _) = session(
        |s| {
            s.add_watch("total > 4", WatchKind::BreakWhenTrue);
        },
        false,
        |_, _| Command::Abort,
    );
    assert_eq!(p.len(), 1);
}

#[test]
fn immediate_window_changes_state() {
    let (_, out) = session(
        |s| {
            s.add_breakpoint(None, 10);
        },
        false,
        |_, s| {
            s.ctx.execute("total = 1000").unwrap();
            let locals = s.locals(0);
            assert!(locals
                .iter()
                .any(|(n, v, t)| n == "total" && v == "1000" && t == "Long"));
            Command::Continue
        },
    );
    assert_eq!(out, vec!["1000"]);
}

#[test]
fn breaks_on_unhandled_error_and_stop() {
    let src = "Sub Main()\n  Stop\n  x = 1 / 0\nEnd Sub\n";
    let reasons = Rc::new(RefCell::new(Vec::new()));
    let r2 = reasons.clone();
    let dbg = Debugger::new(FnFrontend(
        move |ev: &PauseEvent, _: &mut Session<'_, '_>| {
            r2.borrow_mut().push((ev.reason.clone(), ev.location.line));
            Command::Continue
        },
    ));
    let mut engine = Engine::new(CaptureHost::new());
    engine.load_source(src, "M").unwrap();
    engine.set_debug_hook(dbg);
    assert!(engine.run("Main", vec![]).is_err());
    let r = reasons.borrow();
    assert_eq!(r[0], (PauseReason::Stop, 2));
    assert!(matches!(&r[1], (PauseReason::Error(e), 3) if e.number == 11));
}

#[test]
fn channel_frontend_across_threads() {
    let (frontend, events, requests) = ChannelFrontend::new();
    let worker = std::thread::spawn(move || {
        let host = CaptureHost::new();
        let mut engine = Engine::new(host.clone());
        engine.load_source(SRC, "Module1").unwrap();
        let mut dbg = Debugger::new(frontend);
        dbg.state.add_breakpoint(None, 10);
        engine.set_debug_hook(dbg);
        engine.run("Main", vec![]).unwrap();
        host.lines()
    });
    let ev = events.recv().unwrap();
    assert_eq!(ev.location.line, 10);
    let (tx, rx) = std::sync::mpsc::channel();
    requests
        .send(Request::Evaluate {
            frame: 0,
            expression: "total * 2".into(),
            reply: tx,
        })
        .unwrap();
    assert_eq!(rx.recv().unwrap().unwrap(), "28");
    requests.send(Request::Resume(Command::Continue)).unwrap();
    assert_eq!(worker.join().unwrap(), vec!["14"]);
}

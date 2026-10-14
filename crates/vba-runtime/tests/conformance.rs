//! Conformance suite written in VBA itself. Each `tests/vba/*.bas` file has a
//! `Main` that calls the helpers in `Assert.bas`; a failed check prints a line
//! starting with "FAIL". Expected values mostly come from the examples in
//! Microsoft's VBA language reference (MicrosoftDocs/VBA-Docs, MIT licensed
//! code samples), see `tests/vba/README.md`.

use vba_runtime::{CaptureHost, Engine};

const ASSERT: &str = include_str!("vba/Assert.bas");

fn conformance(name: &str, src: &str) {
    let host = CaptureHost::new();
    let mut engine = Engine::new(host.clone());
    engine.load_source(ASSERT, "Assert").unwrap();
    engine
        .load_source(src, name)
        .unwrap_or_else(|e| panic!("{name}: {e}"));
    let result = engine.run("Main", vec![]);
    let lines = host.lines();
    let failures: Vec<&String> = lines.iter().filter(|l| l.starts_with("FAIL")).collect();
    if let Err(e) = result {
        panic!(
            "{name}: runtime error {e}\noutput so far:\n{}",
            lines.join("\n")
        );
    }
    assert!(
        failures.is_empty(),
        "{name}: {} failed checks\n{}",
        failures.len(),
        failures
            .iter()
            .map(|s| s.as_str())
            .collect::<Vec<_>>()
            .join("\n")
    );
}

macro_rules! vba_suite {
    ($($name:ident),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                conformance(stringify!($name), include_str!(concat!("vba/", stringify!($name), ".bas")));
            }
        )*
    };
}

vba_suite!(
    smoke,
    strings,
    conversion,
    math,
    dates,
    format,
    operators,
    control_flow,
);

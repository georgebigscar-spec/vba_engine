# vba_engine

An embeddable Rust engine that runs VBA (Visual Basic for Applications)
compatible code for office automation, with a VBE-style debugger.

```rust
use vba_runtime::{Engine, StdHost};

let mut engine = Engine::new(StdHost);
engine.load_source(r#"
Sub Main()
    Dim i As Long
    For i = 1 To 3
        Debug.Print "Hello #" & i
    Next
End Sub
"#, "Module1")?;
engine.run("Main", vec![])?;
```

## Crates

| Crate | What |
|---|---|
| `vba-syntax` | lexer, parser, AST |
| `vba-runtime` | Variant type system, host object model (`Host`, `Object`), standard library, interpreter, debug hooks |
| `vba-debug` | breakpoints (conditional, hit counts, logpoints), step into/over/out, run to line, watches (break when true / changed), break on errors, cross-thread frontend |
| `vba-excel` | in-memory Excel object model as a `Host`: `Application`, `Workbooks`, `Worksheets`, `Range` (`Cells`, `End`, `Offset`, `Resize`, `CurrentRegion`, array `Value`, `Font`, `Interior`, `NumberFormat`), `WorksheetFunction`, `xl*` constants; `.xlsx` load/save |
| `vba-cli` | `vba run file.bas`, `vba debug file.bas`, `--excel`, `--workbook in.xlsx`, `--save out.xlsx` |

## Try it

```sh
cargo run -p vba-cli -- run examples/Invoice.bas
cargo run -p vba-cli -- run examples/Bank.bas examples/Account.cls   # classes, Type, Enum
cargo run -p vba-cli -- debug examples/Invoice.bas   # then: help
cargo run -p vba-cli -- run examples/SalesReport.bas --excel --save report.xlsx
cargo run -p vba-cli -- run MyMacro.bas --workbook data.xlsx --save out.xlsx
cargo test
```

Running Excel macros from Rust:

```rust
use vba_excel::ExcelHost;
use vba_runtime::Engine;

let host = ExcelHost::open("data.xlsx")?;
let mut engine = Engine::new(host.clone());
engine.load_source(macro_source, "Module1")?;
engine.run("Main", vec![])?;
host.save_as("out.xlsx")?;
```

Exposing an application to VBA code means implementing `Host` (globals such
as `Application`, `ActiveSheet`) and `Object` (an `IDispatch`-like `invoke`)
for its objects. See `docs/ARCHITECTURE.md` (in Russian) for the design and roadmap.

## License

MIT OR Apache-2.0

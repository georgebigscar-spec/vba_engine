# VBA conformance suite

Each `*.bas` file here has a `Main` that calls `Check`, `CheckTyped` and
`CheckErr` from `Assert.bas`; `tests/conformance.rs` runs every file as its own
Rust test. A failed check prints a `FAIL ...` line.

Expected values describe real VBA 7 (Office, en-US locale). Most of them are
ported from the examples in Microsoft's VBA language reference
(https://github.com/MicrosoftDocs/VBA-Docs, code samples under the MIT
license, Copyright (c) Microsoft Corporation); check labels start with the
docs page they come from, e.g. `mid-function: FirstWord`. The rest are edge
cases of documented VBA behaviour.

To add a file: create `name.bas` with `Attribute VB_Name = "name"` and a
`Sub Main`, then add `name` to the `vba_suite!` list in `tests/conformance.rs`.

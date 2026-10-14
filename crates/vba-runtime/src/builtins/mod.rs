//! VBA standard library: `VBA.Strings`, `VBA.Math`, `VBA.Conversion`,
//! `VBA.DateTime`, `VBA.Information`, `VBA.Interaction` and constants.

use std::cell::Cell;

use crate::date;
use crate::error::{VbaError, VbaResult};
use crate::host::Host;
use crate::object::{Args, Collection};
use crate::value::{self, parse_number, Value, VarType, VbArray};

mod conversion;
mod datetime;
mod format;
mod math;
mod strings;

pub use format::format_value;

/// Built-in constants (`vbCrLf`, `vbYesNo`, ...).
pub fn constant(name: &str) -> Option<Value> {
    let lname = name.to_ascii_lowercase();
    let v = match lname.as_str() {
        "vbcrlf" | "vbnewline" => Value::str("\r\n"),
        "vbcr" => Value::str("\r"),
        "vblf" => Value::str("\n"),
        "vbtab" => Value::str("\t"),
        "vbnullstring" => Value::str(""),
        "vbnullchar" => Value::str("\0"),
        "vbokonly" | "vbbinarycompare" | "vbempty" => Value::Integer(0),
        "vbok" | "vbokcancel" | "vbtextcompare" | "vbnull" => Value::Integer(1),
        "vbcancel" | "vbabortretryignore" => Value::Integer(2),
        "vbabort" => Value::Integer(3),
        "vbcritical" => Value::Integer(16),
        "vbretry" => Value::Integer(4),
        "vbignore" => Value::Integer(5),
        "vbyes" => Value::Integer(6),
        "vbno" => Value::Integer(7),
        "vbyesnocancel" => Value::Integer(3),
        "vbyesno" => Value::Integer(4),
        "vbretrycancel" => Value::Integer(5),
        "vbquestion" => Value::Integer(32),
        "vbexclamation" => Value::Integer(48),
        "vbinformation" => Value::Integer(64),
        "vbinteger" => Value::Integer(2),
        "vblong" => Value::Integer(3),
        "vbsingle" => Value::Integer(4),
        "vbdouble" => Value::Integer(5),
        "vbcurrency" => Value::Integer(6),
        "vbdate" => Value::Integer(7),
        "vbstring" => Value::Integer(8),
        "vbobject" => Value::Integer(9),
        "vberror" => Value::Integer(10),
        "vbboolean" => Value::Integer(11),
        "vbvariant" => Value::Integer(12),
        "vbbyte" => Value::Integer(17),
        "vbarray" => Value::Integer(8192),
        "vbobjecterror" => Value::Long(-2147221504),
        "vbblack" => Value::Long(0),
        "vbred" => Value::Long(0xFF),
        "vbgreen" => Value::Long(0xFF00),
        "vbblue" => Value::Long(0xFF0000),
        "vbwhite" => Value::Long(0xFFFFFF),
        "vbsunday" => Value::Integer(1),
        "vbmonday" => Value::Integer(2),
        _ => {
            return strings::constant(&lname)
                .or_else(|| conversion::constant(&lname))
                .or_else(|| math::constant(&lname))
                .or_else(|| datetime::constant(&lname))
                .or_else(|| format::constant(&lname))
        }
    };
    Some(v)
}

fn arg(args: &Args, i: usize) -> VbaResult<&Value> {
    args.positional
        .get(i)
        .filter(|v| !matches!(v, Value::Missing))
        .ok_or_else(VbaError::wrong_arg_count)
}

fn opt(args: &Args, i: usize) -> Option<&Value> {
    args.positional
        .get(i)
        .filter(|v| !matches!(v, Value::Missing))
}

fn s(args: &Args, i: usize) -> VbaResult<String> {
    arg(args, i)?.to_string_value()
}

fn n(args: &Args, i: usize) -> VbaResult<f64> {
    arg(args, i)?.to_f64()
}

fn int(args: &Args, i: usize) -> VbaResult<i64> {
    arg(args, i)?.to_i64()
}

fn chars(s: &str) -> Vec<char> {
    s.chars().collect()
}

fn null_passthrough(args: &Args) -> bool {
    matches!(args.positional.first(), Some(Value::Null))
}

thread_local! {
    static RND_STATE: Cell<u32> = const { Cell::new(0x50000) };
}

/// Calls a built-in function. Returns `None` if `name` is not a built-in.
/// Built-ins that look at an object itself rather than its default value.
pub fn takes_objects(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "typename"
            | "isobject"
            | "vartype"
            | "isnull"
            | "isempty"
            | "isarray"
            | "ismissing"
            | "objptr"
            | "callbyname"
            | "iserror"
    )
}

pub fn call(
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    text_compare: bool,
) -> Option<VbaResult<Value>> {
    let lname = name.to_ascii_lowercase();
    let parts: [Part; 6] = [
        strings::call,
        conversion::call,
        math::call,
        datetime::call,
        format::call,
        interaction,
    ];
    for part in parts {
        match part(&lname, name, args, host, text_compare) {
            Err(e) if e.number == 35 && e.description.ends_with(name) => continue,
            other => return Some(other),
        }
    }
    None
}

type Part = fn(&str, &str, &Args, &mut dyn Host, bool) -> VbaResult<Value>;

fn interaction(
    _lname: &str,
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match _lname {
        // ---- interaction ----
        "msgbox" => {
            let buttons = opt(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0) as i32;
            let title = opt(args, 2)
                .map(|v| v.to_string_value())
                .transpose()?
                .unwrap_or_default();
            Value::Integer(host.msgbox(&s(args, 0)?, buttons, &title)? as i16)
        }
        "inputbox" => {
            let title = opt(args, 1)
                .map(|v| v.to_string_value())
                .transpose()?
                .unwrap_or_default();
            let default = opt(args, 2)
                .map(|v| v.to_string_value())
                .transpose()?
                .unwrap_or_default();
            Value::String(
                host.inputbox(&s(args, 0)?, &title, &default)?
                    .unwrap_or_default(),
            )
        }
        "createobject" => match s(args, 0)?.to_ascii_lowercase().as_str() {
            "vba.collection" => Value::from(Collection::new()),
            prog_id => host.create_object(prog_id)?,
        },
        "doevents" => Value::Integer(0),
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::host::CaptureHost;

    fn call_s(name: &str, args: Vec<Value>) -> String {
        call(name, &Args::new(args), &mut CaptureHost::new(), false)
            .unwrap()
            .unwrap()
            .to_string()
    }

    #[test]
    fn string_functions() {
        assert_eq!(call_s("Mid", vec!["Hello World".into(), 7.into()]), "World");
        assert_eq!(call_s("InStr", vec!["Hello".into(), "l".into()]), "3");
        assert_eq!(
            call_s("InStr", vec![4.into(), "Hello".into(), "l".into()]),
            "4"
        );
        assert_eq!(
            call_s("Replace", vec!["a-b-c".into(), "-".into(), "+".into()]),
            "a+b+c"
        );
        assert_eq!(call_s("Val", vec!["  12.5abc".into()]), "12.5");
    }

    #[test]
    fn format_function() {
        assert_eq!(
            call_s("Format", vec![1234.5.into(), "#,##0.00".into()]),
            "1,234.50"
        );
        assert_eq!(call_s("Format", vec![0.256.into(), "0.0%".into()]), "25.6%");
        let d = Value::Date(date::serial(2024, 3, 5) + 0.5);
        assert_eq!(
            call_s("Format", vec![d, "yyyy-mm-dd hh:nn".into()]),
            "2024-03-05 12:00"
        );
    }

    #[test]
    fn unknown_is_none() {
        assert!(call("NoSuchFn", &Args::default(), &mut CaptureHost::new(), false).is_none());
    }
}

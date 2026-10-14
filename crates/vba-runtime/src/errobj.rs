//! The global `Err` object.

use std::any::Any;
use std::cell::{Cell, RefCell};

use crate::error::{VbaError, VbaResult};
use crate::object::{Args, Invoke, Object};
use crate::value::Value;

/// The message VBA gives a trappable error number (`Error 11`,
/// `Err.Raise 9`); other numbers are application-defined.
pub fn standard_description(number: i32) -> &'static str {
    match number {
        3 => "Return without GoSub",
        5 => "Invalid procedure call or argument",
        6 => "Overflow",
        7 => "Out of memory",
        9 => "Subscript out of range",
        10 => "This array is fixed or temporarily locked",
        11 => "Division by zero",
        13 => "Type mismatch",
        14 => "Out of string space",
        16 => "Expression too complex",
        17 => "Can't perform requested operation",
        18 => "User interrupt occurred",
        20 => "Resume without error",
        28 => "Out of stack space",
        35 => "Sub or Function not defined",
        47 => "Too many DLL application clients",
        48 => "Error in loading DLL",
        49 => "Bad DLL calling convention",
        51 => "Internal error",
        52 => "Bad file name or number",
        53 => "File not found",
        54 => "Bad file mode",
        55 => "File already open",
        57 => "Device I/O error",
        58 => "File already exists",
        59 => "Bad record length",
        61 => "Disk full",
        62 => "Input past end of file",
        63 => "Bad record number",
        67 => "Too many files",
        68 => "Device unavailable",
        70 => "Permission denied",
        71 => "Disk not ready",
        74 => "Can't rename with different drive",
        75 => "Path/File access error",
        76 => "Path not found",
        91 => "Object variable or With block variable not set",
        92 => "For loop not initialized",
        93 => "Invalid pattern string",
        94 => "Invalid use of Null",
        424 => "Object required",
        429 => "ActiveX component can't create object",
        438 => "Object doesn't support this property or method",
        440 => "Automation error",
        445 => "Object doesn't support this action",
        446 => "Object doesn't support named arguments",
        448 => "Named argument not found",
        449 => "Argument not optional",
        450 => "Wrong number of arguments or invalid property assignment",
        457 => "This key is already associated with an element of this collection",
        _ => "Application-defined or object-defined error",
    }
}

/// The error `Error n` / `Err.Raise n` raises: number 0 is itself an
/// invalid argument (error 5).
pub fn standard_error(number: i32) -> VbaError {
    let number = if number == 0 { 5 } else { number };
    VbaError::new(number, standard_description(number))
}

#[derive(Default)]
pub struct ErrObject {
    state: RefCell<(i32, String, String)>,
    /// `Erl`: the last line number passed before the error.
    erl: Cell<i64>,
}

impl ErrObject {
    pub fn set(&self, e: &VbaError) {
        *self.state.borrow_mut() = (e.number, e.description.clone(), e.source.clone());
    }

    pub fn set_erl(&self, line: i64) {
        self.erl.set(line);
    }

    pub fn erl(&self) -> i64 {
        self.erl.get()
    }

    pub fn clear(&self) {
        *self.state.borrow_mut() = (0, String::new(), String::new());
        self.erl.set(0);
    }

    pub fn number(&self) -> i32 {
        self.state.borrow().0
    }
}

impl Object for ErrObject {
    fn type_name(&self) -> &str {
        "ErrObject"
    }

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        let mut st = self.state.borrow_mut();
        match (member.to_ascii_lowercase().as_str(), kind) {
            ("" | "number", Invoke::Let(v)) => st.0 = v.to_i64()? as i32,
            ("" | "number", _) => return Ok(Value::Long(st.0)),
            ("description", Invoke::Let(v)) => st.1 = v.to_string_value()?,
            ("description", _) => return Ok(Value::String(st.1.clone())),
            ("source", Invoke::Let(v)) => st.2 = v.to_string_value()?,
            ("source", _) => return Ok(Value::String(st.2.clone())),
            ("clear", _) => {
                *st = (0, String::new(), String::new());
                self.erl.set(0);
            }
            ("raise", _) => {
                let number = args.require(0, "Number")?.to_i64()? as i32;
                let standard = standard_error(number);
                let source = args
                    .get(1, "Source")
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                let description = args
                    .get(2, "Description")
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or(standard.description);
                let mut e = VbaError::new(standard.number, description);
                e.source = source;
                return Err(e);
            }
            _ => return Err(VbaError::not_supported(member)),
        }
        Ok(Value::Empty)
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        let st = self.state.borrow();
        vec![
            ("Number".into(), Value::Long(st.0)),
            ("Description".into(), Value::String(st.1.clone())),
            ("Source".into(), Value::String(st.2.clone())),
        ]
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

//! The global `Err` object.

use std::any::Any;
use std::cell::RefCell;

use crate::error::{VbaError, VbaResult};
use crate::object::{Args, Invoke, Object};
use crate::value::Value;

#[derive(Default)]
pub struct ErrObject {
    state: RefCell<(i32, String, String)>,
}

impl ErrObject {
    pub fn set(&self, e: &VbaError) {
        *self.state.borrow_mut() = (e.number, e.description.clone(), e.source.clone());
    }

    pub fn clear(&self) {
        *self.state.borrow_mut() = (0, String::new(), String::new());
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
            ("clear", _) => *st = (0, String::new(), String::new()),
            ("raise", _) => {
                let number = args.require(0, "Number")?.to_i64()? as i32;
                let source = args
                    .get(1, "Source")
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                let description = args
                    .get(2, "Description")
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_else(|| "Application-defined or object-defined error".into());
                let mut e = VbaError::new(number, description);
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

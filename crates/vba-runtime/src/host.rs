//! The host interface: how an Office-like application plugs into the runtime.

use std::cell::RefCell;
use std::rc::Rc;

use crate::error::{VbaError, VbaResult};
use crate::object::ObjectRef;
use crate::value::Value;

/// Implemented by the embedding application (Excel-, Word-like hosts, tests).
pub trait Host {
    /// Resolves a global name not defined in VBA code: `Application`,
    /// `ActiveSheet`, `ThisWorkbook`, host constants like `xlUp`.
    fn global(&self, _name: &str) -> Option<Value> {
        None
    }

    /// The application object whose members are in scope unqualified, so that
    /// `Range("A1")` means `Application.Range("A1")`.
    fn global_object(&self) -> Option<ObjectRef> {
        None
    }

    /// Output of `Debug.Print` (the Immediate window).
    fn debug_print(&mut self, text: &str);

    /// `MsgBox`; returns the clicked button (`vbOK` = 1).
    fn msgbox(&mut self, prompt: &str, _buttons: i32, _title: &str) -> VbaResult<i32> {
        self.debug_print(&format!("[MsgBox] {prompt}"));
        Ok(1)
    }

    /// `InputBox`; `None` means Cancel.
    fn inputbox(
        &mut self,
        _prompt: &str,
        _title: &str,
        default: &str,
    ) -> VbaResult<Option<String>> {
        Ok(Some(default.to_string()))
    }

    /// `CreateObject("Scripting.Dictionary")` and friends.
    fn create_object(&mut self, prog_id: &str) -> VbaResult<Value> {
        Err(VbaError::new(
            429,
            format!("ActiveX component can't create object: {prog_id}"),
        ))
    }

    /// `New ClassName` for host-provided classes.
    fn new_object(&mut self, class: &str) -> VbaResult<Value> {
        Err(VbaError::new(429, format!("Unknown class: {class}")))
    }

    /// Current local time as an OLE date (`Now`).
    fn now(&self) -> f64 {
        let secs = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs_f64())
            .unwrap_or(0.0);
        25_569.0 + secs / 86_400.0
    }

    /// Called periodically during execution — the place to honour `DoEvents`
    /// or a host-side cancel button. Returning `false` interrupts the macro.
    fn poll(&mut self) -> bool {
        true
    }
}

/// Writes `Debug.Print` / `MsgBox` to stdout.
#[derive(Default)]
pub struct StdHost;

impl Host for StdHost {
    fn debug_print(&mut self, text: &str) {
        println!("{text}");
    }
}

/// Collects output in memory; handy for tests.
#[derive(Default, Clone)]
pub struct CaptureHost {
    pub output: Rc<RefCell<Vec<String>>>,
    pub globals: Vec<(String, Value)>,
}

impl CaptureHost {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn lines(&self) -> Vec<String> {
        self.output.borrow().clone()
    }
}

impl Host for CaptureHost {
    fn global(&self, name: &str) -> Option<Value> {
        self.globals
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v)| v.clone())
    }

    fn debug_print(&mut self, text: &str) {
        self.output.borrow_mut().push(text.to_string());
    }
}

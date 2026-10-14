//! Debugger hooks. The interpreter calls a [`DebugHook`] before every
//! executable statement and on procedure entry/exit, `Stop`, `Debug.Assert`
//! and runtime errors. The hook gets a [`DebugContext`] for inspecting and
//! changing program state. `vba-debug` builds breakpoints, stepping and
//! watches on top of this.

use crate::error::VbaError;
use crate::interp::Interpreter;
use crate::value::Value;

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Location {
    pub module: String,
    pub procedure: String,
    pub line: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DebugAction {
    Continue,
    /// Stop the program (the VBE "Reset" button). Raises a non-trappable error.
    Abort,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StopKind {
    /// A `Stop` statement.
    Stop,
    /// `Debug.Assert` with a false condition.
    AssertFailed,
}

pub trait DebugHook {
    /// Before each executable statement, and once on the `End Sub` /
    /// `End Function` line before a procedure returns.
    fn on_statement(&mut self, ctx: &mut DebugContext<'_>) -> DebugAction;

    /// After a new frame is pushed (arguments are bound).
    fn on_enter(&mut self, _ctx: &mut DebugContext<'_>) {}

    /// Before a frame is popped.
    fn on_exit(&mut self, _ctx: &mut DebugContext<'_>) {}

    fn on_stop(&mut self, ctx: &mut DebugContext<'_>, kind: StopKind) -> DebugAction;

    /// A runtime error was raised. `handled` tells whether an `On Error`
    /// handler is active anywhere on the call stack, which is what VBA's
    /// "Break on Unhandled Errors" option looks at.
    fn on_error(
        &mut self,
        _ctx: &mut DebugContext<'_>,
        _err: &VbaError,
        _handled: bool,
    ) -> DebugAction {
        DebugAction::Continue
    }
}

/// A live view of the paused program. Frame index 0 is the innermost frame.
pub struct DebugContext<'a> {
    pub(crate) interp: &'a mut Interpreter,
}

impl<'a> DebugContext<'a> {
    pub fn location(&self) -> Location {
        self.interp.location(0).unwrap_or(Location {
            module: String::new(),
            procedure: String::new(),
            line: 0,
        })
    }

    /// Number of active procedure frames.
    pub fn depth(&self) -> usize {
        self.interp.depth()
    }

    /// Innermost frame first, like the VBE Call Stack dialog.
    pub fn call_stack(&self) -> Vec<Location> {
        (0..self.depth())
            .filter_map(|i| self.interp.location(i))
            .collect()
    }

    /// Locals window contents for a frame.
    pub fn locals(&self, frame: usize) -> Vec<(String, Value)> {
        self.interp.locals(frame)
    }

    /// Module-level variables.
    pub fn module_variables(&self) -> Vec<(String, Value)> {
        self.interp.module_variables()
    }

    /// Evaluates an expression in the innermost frame (Watch / Quick Watch / `? expr`).
    pub fn eval(&mut self, expr: &str) -> Result<Value, String> {
        self.interp.eval_in_frame(0, expr)
    }

    pub fn eval_in_frame(&mut self, frame: usize, expr: &str) -> Result<Value, String> {
        self.interp.eval_in_frame(frame, expr)
    }

    /// Executes statements in the innermost frame (the Immediate window),
    /// e.g. `x = 10` or `Debug.Print Len(s)`.
    pub fn execute(&mut self, statements: &str) -> Result<(), String> {
        self.interp.execute_in_frame(0, statements)
    }

    pub fn interpreter(&mut self) -> &mut Interpreter {
        self.interp
    }
}

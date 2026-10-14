//! Runtime of the VBA engine: the `Variant` type system, the host object
//! model, the standard library and a tree-walking interpreter with
//! debugger hooks.
//!
//! ```
//! use vba_runtime::{Engine, CaptureHost};
//! let host = CaptureHost::new();
//! let mut engine = Engine::new(host.clone());
//! engine.load_source("Sub Main()\n  Debug.Print \"Sum: \" & (2 + 3)\nEnd Sub\n", "Module1").unwrap();
//! engine.run("Main", vec![]).unwrap();
//! assert_eq!(host.lines(), vec!["Sum: 5"]);
//! ```

pub mod builtins;
pub mod class;
pub mod date;
pub mod debug;
pub mod errobj;
pub mod error;
pub mod host;
pub mod interp;
pub mod object;
pub mod value;

pub use debug::{DebugAction, DebugContext, DebugHook, Location, StopKind};
pub use error::{VbaError, VbaResult};
pub use host::{CaptureHost, Host, StdHost};
pub use interp::Interpreter;
pub use object::{Args, Collection, DynObject, Invoke, Object, ObjectRef};
pub use value::{UdtValue, Value, VarType, VbArray};

use std::fmt;

/// Either a parse error or a runtime error.
#[derive(Debug, Clone)]
pub enum EngineError {
    Syntax(vba_syntax::SyntaxError),
    Runtime(VbaError),
}

impl fmt::Display for EngineError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            EngineError::Syntax(e) => e.fmt(f),
            EngineError::Runtime(e) => e.fmt(f),
        }
    }
}

impl std::error::Error for EngineError {}

impl From<vba_syntax::SyntaxError> for EngineError {
    fn from(e: vba_syntax::SyntaxError) -> Self {
        EngineError::Syntax(e)
    }
}

impl From<VbaError> for EngineError {
    fn from(e: VbaError) -> Self {
        EngineError::Runtime(e)
    }
}

/// Convenience wrapper: parse + load + run.
pub struct Engine {
    pub interpreter: Interpreter,
}

impl Engine {
    pub fn new(host: impl Host + 'static) -> Self {
        Engine {
            interpreter: Interpreter::new(host),
        }
    }

    pub fn load_source(&mut self, src: &str, module_name: &str) -> Result<(), EngineError> {
        let module = vba_syntax::parse_module(src, module_name)?;
        self.interpreter.load_module(module)?;
        Ok(())
    }

    pub fn run(&mut self, procedure: &str, args: Vec<Value>) -> Result<Value, EngineError> {
        Ok(self.interpreter.run(procedure, args)?)
    }

    pub fn set_debug_hook(&mut self, hook: impl DebugHook + 'static) {
        self.interpreter.set_debug_hook(Some(Box::new(hook)));
    }
}

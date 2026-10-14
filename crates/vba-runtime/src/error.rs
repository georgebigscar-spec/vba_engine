//! Runtime errors, numbered like VBA's `Err.Number`.

use std::fmt;

pub type VbaResult<T> = Result<T, VbaError>;

#[derive(Debug, Clone, PartialEq)]
pub struct VbaError {
    pub number: i32,
    pub description: String,
    pub source: String,
    /// Source line where the error was raised (filled in by the interpreter).
    pub line: Option<u32>,
    /// Non-trappable errors (the debugger's Reset, `End`) bypass `On Error`.
    pub trappable: bool,
    /// Set once the debugger hook has seen this error.
    #[doc(hidden)]
    pub debugger_notified: bool,
}

impl VbaError {
    pub fn new(number: i32, description: impl Into<String>) -> Self {
        VbaError {
            number,
            description: description.into(),
            source: String::new(),
            line: None,
            trappable: true,
            debugger_notified: false,
        }
    }

    pub fn with_line(mut self, line: u32) -> Self {
        self.line.get_or_insert(line);
        self
    }

    pub fn overflow() -> Self {
        Self::new(6, "Overflow")
    }
    pub fn division_by_zero() -> Self {
        Self::new(11, "Division by zero")
    }
    pub fn subscript_out_of_range() -> Self {
        Self::new(9, "Subscript out of range")
    }
    pub fn type_mismatch() -> Self {
        Self::new(13, "Type mismatch")
    }
    pub fn invalid_call() -> Self {
        Self::new(5, "Invalid procedure call or argument")
    }
    pub fn invalid_use_of_null() -> Self {
        Self::new(94, "Invalid use of Null")
    }
    pub fn object_required() -> Self {
        Self::new(424, "Object required")
    }
    pub fn object_not_set() -> Self {
        Self::new(91, "Object variable or With block variable not set")
    }
    pub fn wrong_arg_count() -> Self {
        Self::new(
            450,
            "Wrong number of arguments or invalid property assignment",
        )
    }
    pub fn not_supported(member: &str) -> Self {
        Self::new(
            438,
            format!("Object doesn't support this property or method: {member}"),
        )
    }
    pub fn sub_not_defined(name: &str) -> Self {
        Self::new(35, format!("Sub or Function not defined: {name}"))
    }
    pub fn variable_not_defined(name: &str) -> Self {
        Self::new(-1, format!("Variable not defined: {name}"))
    }
    pub fn label_not_defined(name: &str) -> Self {
        Self::new(-1, format!("Label not defined: {name}"))
    }

    /// Raised when the debugger stops execution (VBA's "Reset").
    pub fn interrupted() -> Self {
        VbaError {
            trappable: false,
            ..Self::new(18, "User interrupt occurred")
        }
    }

    /// Raised by the `End` statement.
    pub fn end() -> Self {
        VbaError {
            trappable: false,
            ..Self::new(0, "End")
        }
    }

    pub fn is_end(&self) -> bool {
        !self.trappable && self.number == 0
    }
}

impl fmt::Display for VbaError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(f, "Run-time error '{}': {}", self.number, self.description)?;
        if let Some(line) = self.line {
            write!(f, " (line {line})")?;
        }
        Ok(())
    }
}

impl std::error::Error for VbaError {}

//! Abstract syntax tree for a VBA standard module.
//!
//! Every statement carries a [`Span`] so the runtime and debugger can map
//! execution back to source lines.

use crate::Span;

#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    pub name: String,
    pub options: ModuleOptions,
    /// Module-level `Dim` / `Private` / `Public` / `Const` declarations.
    pub declarations: Vec<Stmt>,
    pub procedures: Vec<Procedure>,
}

#[derive(Debug, Clone, Default, PartialEq)]
pub struct ModuleOptions {
    pub explicit: bool,
    pub base: u8,
    pub compare_text: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProcKind {
    Sub,
    Function,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Visibility {
    Public,
    Private,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Procedure {
    pub name: String,
    pub kind: ProcKind,
    pub visibility: Visibility,
    pub params: Vec<Param>,
    pub return_type: Option<TypeName>,
    pub body: Vec<Stmt>,
    pub span: Span,
    /// Line of the closing `End Sub` / `End Function`.
    pub end_line: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    pub name: String,
    pub by_val: bool,
    pub optional: bool,
    pub ty: Option<TypeName>,
    pub default: Option<Expr>,
}

/// A declared type, kept as written (`Long`, `Range`, `Excel.Worksheet`).
#[derive(Debug, Clone, PartialEq)]
pub struct TypeName {
    pub name: String,
    pub is_new: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    pub name: String,
    pub ty: Option<TypeName>,
    /// `Dim a(1 To 10, 5)` → bounds; `Dim a()` → `Some(vec![])` (dynamic).
    pub bounds: Option<Vec<(Option<Expr>, Expr)>>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Stmt {
    pub kind: StmtKind,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub enum StmtKind {
    Dim(Vec<VarDecl>),
    Const(Vec<(String, Expr)>),
    ReDim {
        preserve: bool,
        decls: Vec<VarDecl>,
    },
    /// `x = e` / `Let x = e`
    Assign {
        target: Expr,
        value: Expr,
    },
    /// `Set x = e`
    SetAssign {
        target: Expr,
        value: Expr,
    },
    /// `Foo a, b`, `Call Foo(a, b)`, `obj.Method a`
    Call {
        callee: Expr,
        args: Vec<Arg>,
    },
    If {
        branches: Vec<(Expr, Vec<Stmt>)>,
        else_body: Option<Vec<Stmt>>,
    },
    For {
        var: String,
        from: Expr,
        to: Expr,
        step: Option<Expr>,
        body: Vec<Stmt>,
    },
    ForEach {
        var: String,
        collection: Expr,
        body: Vec<Stmt>,
    },
    /// Covers `Do While/Until ... Loop`, `Do ... Loop While/Until` and `While ... Wend`.
    DoLoop {
        pre: Option<LoopCond>,
        post: Option<LoopCond>,
        body: Vec<Stmt>,
    },
    Select {
        subject: Expr,
        cases: Vec<(Vec<CaseClause>, Vec<Stmt>)>,
        else_body: Option<Vec<Stmt>>,
    },
    With {
        object: Expr,
        body: Vec<Stmt>,
    },
    Exit(ExitKind),
    OnError(OnError),
    Resume(ResumeKind),
    GoTo(String),
    Label(String),
    /// `End` — terminates the program.
    End,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LoopCond {
    pub until: bool,
    pub cond: Expr,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CaseClause {
    Value(Expr),
    Range(Expr, Expr),
    Is(BinOp, Expr),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExitKind {
    Sub,
    Function,
    For,
    Do,
}

#[derive(Debug, Clone, PartialEq)]
pub enum OnError {
    ResumeNext,
    GoTo(String),
    GoToZero,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ResumeKind {
    Same,
    Next,
    Label(String),
}

#[derive(Debug, Clone, PartialEq)]
pub struct Arg {
    pub name: Option<String>,
    /// `None` for an omitted positional argument: `Foo a, , c`.
    pub value: Option<Expr>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Expr {
    Int(i64),
    Float(f64),
    Str(String),
    Date(String),
    Bool(bool),
    Nothing,
    Null,
    Empty,
    Ident(String),
    /// `.Name` inside a `With` block.
    WithMember(String),
    Member(Box<Expr>, String),
    /// `f(a, b)` — call or array index; resolved at runtime like VBA does.
    Index(Box<Expr>, Vec<Arg>),
    New(String),
    Unary(UnOp, Box<Expr>),
    Binary(BinOp, Box<Expr>, Box<Expr>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UnOp {
    Neg,
    Not,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BinOp {
    Pow,
    Mul,
    Div,
    IntDiv,
    Mod,
    Add,
    Sub,
    Concat,
    Eq,
    Ne,
    Lt,
    Le,
    Gt,
    Ge,
    Like,
    Is,
    And,
    Or,
    Xor,
    Eqv,
    Imp,
}

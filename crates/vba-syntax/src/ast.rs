//! Abstract syntax tree for VBA standard and class modules.
//!
//! Every statement carries a [`Span`] so the runtime and debugger can map
//! execution back to source lines.

use crate::Span;

#[derive(Debug, Clone, PartialEq)]
pub struct Module {
    pub name: String,
    pub kind: ModuleKind,
    pub options: ModuleOptions,
    /// Module-level `Dim` / `Private` / `Public` / `Const` declarations.
    pub declarations: Vec<Declaration>,
    pub procedures: Vec<Procedure>,
    pub types: Vec<TypeDef>,
    pub enums: Vec<EnumDef>,
    /// `Implements IShape` (class modules).
    pub implements: Vec<String>,
    /// `Public Event Changed(ByVal value As Long)` (class modules).
    pub events: Vec<EventDef>,
}

/// `Event Name(params)` declared in a class module.
#[derive(Debug, Clone, PartialEq)]
pub struct EventDef {
    pub name: String,
    pub params: Vec<Param>,
    pub span: Span,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ModuleKind {
    /// A standard module (`.bas`).
    #[default]
    Standard,
    /// A class module (`.cls`): each `New` creates an instance.
    Class,
}

/// A module-level `Dim` / `Const` statement and its visibility. `Dim` is
/// private; `Public` / `Global` make it visible outside the module (for a
/// class: a public field).
#[derive(Debug, Clone, PartialEq)]
pub struct Declaration {
    pub visibility: Visibility,
    pub stmt: Stmt,
}

/// `Type Name ... End Type`.
#[derive(Debug, Clone, PartialEq)]
pub struct TypeDef {
    pub name: String,
    pub fields: Vec<VarDecl>,
    pub span: Span,
}

/// `Enum Name ... End Enum`. Members without a value continue from the
/// previous one (starting at 0).
#[derive(Debug, Clone, PartialEq)]
pub struct EnumDef {
    pub name: String,
    pub members: Vec<(String, Option<Expr>)>,
    pub span: Span,
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
    PropertyGet,
    PropertyLet,
    PropertySet,
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
    /// `Attribute Name.VB_UserMemId = n` inside the body: 0 marks the default
    /// member, -4 the `For Each` enumerator.
    pub user_mem_id: Option<i64>,
    /// `Static Sub` / `Static Function`: every local keeps its value
    /// between calls.
    pub is_static: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Param {
    pub name: String,
    pub by_val: bool,
    pub optional: bool,
    /// `ParamArray args()`: collects the remaining arguments into an array.
    pub param_array: bool,
    pub ty: Option<TypeName>,
    pub default: Option<Expr>,
}

/// A declared type, kept as written (`Long`, `Range`, `Excel.Worksheet`).
#[derive(Debug, Clone, PartialEq)]
pub struct TypeName {
    pub name: String,
    pub is_new: bool,
    /// `String * 10`: the length of a fixed-length string.
    pub fixed_len: Option<Box<Expr>>,
    /// `Function F() As Long()`: a dynamic array of this type.
    pub array: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct VarDecl {
    pub name: String,
    /// `Private WithEvents src As Source`: receives `src_EventName` calls.
    pub with_events: bool,
    pub ty: Option<TypeName>,
    /// `Dim a(1 To 10, 5)` → bounds; `Dim a()` → `Some(vec![])` (dynamic).
    pub bounds: Option<Vec<(Option<Expr>, Expr)>>,
}

/// An item of a `Debug.Print` list.
#[derive(Debug, Clone, PartialEq)]
pub enum PrintItem {
    Expr(Expr),
    /// `;`: no space between items.
    Semicolon,
    /// `,`: move to the next 14-column print zone.
    Comma,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Stmt {
    pub kind: StmtKind,
    pub span: Span,
}

#[derive(Debug, Clone, PartialEq)]
pub enum StmtKind {
    Dim(Vec<VarDecl>),
    /// `Static x As Long`: a local that keeps its value between calls.
    Static(Vec<VarDecl>),
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
    /// `Debug.Print a; b, c;` — kept apart from calls because `;` and `,`
    /// mean different things here.
    DebugPrint(Vec<PrintItem>),
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
    /// `GoSub Label`: jumps to a label of the procedure; `Return` comes back.
    GoSub(String),
    Return,
    /// `On n GoTo a, b` / `On n GoSub a, b`: jumps to the n-th label.
    OnGoTo {
        index: Expr,
        labels: Vec<String>,
        gosub: bool,
    },
    /// `Error n`: raises run-time error n with its standard description.
    Error(Expr),
    /// `Name:` or a line number (`100 x = 1`).
    Label(String),
    /// `End` — terminates the program.
    End,
    /// `RaiseEvent Name(args)` inside a class module.
    RaiseEvent {
        name: String,
        args: Vec<Arg>,
    },
    /// `Open path For Input As #1`
    Open {
        path: Expr,
        mode: OpenMode,
        file: Expr,
    },
    /// `Close #1, #2`; no files closes all.
    Close(Vec<Expr>),
    /// `Print #1, a; b` — same separators as `Debug.Print`.
    PrintFile {
        file: Expr,
        items: Vec<PrintItem>,
    },
    /// `Write #1, a, b` — comma-delimited, strings quoted.
    WriteFile {
        file: Expr,
        items: Vec<PrintItem>,
    },
    /// `Line Input #1, s`
    LineInput {
        file: Expr,
        target: Expr,
    },
    /// `Input #1, a, b`
    InputFile {
        file: Expr,
        targets: Vec<Expr>,
    },
    /// `Erase a, b`
    Erase(Vec<Expr>),
    /// `LSet s = v` / `RSet s = v`: align within the string's current length.
    AlignSet {
        right: bool,
        target: Expr,
        value: Expr,
    },
    /// `Mid(s, start[, length]) = v`
    MidAssign {
        target: Expr,
        start: Expr,
        length: Option<Expr>,
        value: Expr,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OpenMode {
    Input,
    Output,
    Append,
    Binary,
    Random,
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
    Property,
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
    /// Literal with an explicit type: `%` Integer, `&` Long, `^` LongLong,
    /// `!` Single, `#` Double, `@` Currency.
    Typed(f64, char),
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
    /// `TypeOf obj Is ClassName`
    TypeOf(Box<Expr>, String),
    Unary(UnOp, Box<Expr>),
    /// `(e)`: a parenthesised argument is passed by value.
    Paren(Box<Expr>),
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

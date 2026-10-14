//! Front end of the VBA engine: tokenizer, parser and AST.
//!
//! ```
//! let module = vba_syntax::parse_module("Sub Main()\n  x = 1 + 2\nEnd Sub\n", "Module1").unwrap();
//! assert_eq!(module.procedures[0].name, "Main");
//! ```

pub mod ast;
pub mod lexer;
pub mod parser;

use std::fmt;

/// 1-based source position.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default, Hash)]
pub struct Span {
    pub line: u32,
    pub col: u32,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SyntaxError {
    pub message: String,
    pub span: Span,
}

impl SyntaxError {
    pub fn new(message: impl Into<String>, span: Span) -> Self {
        SyntaxError {
            message: message.into(),
            span,
        }
    }
}

impl fmt::Display for SyntaxError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "syntax error at {}:{}: {}",
            self.span.line, self.span.col, self.message
        )
    }
}

impl std::error::Error for SyntaxError {}

/// Parses a standard module (`.bas`).
pub fn parse_module(src: &str, name: &str) -> Result<ast::Module, SyntaxError> {
    let toks = lexer::tokenize(src)?;
    parser::Parser::new(toks).parse_module(name)
}

/// Parses a single expression, e.g. a watch or a breakpoint condition.
pub fn parse_expression(src: &str) -> Result<ast::Expr, SyntaxError> {
    let toks = lexer::tokenize(src)?;
    let mut p = parser::Parser::new(toks);
    let e = p.parse_expr()?;
    Ok(e)
}

#[cfg(test)]
mod tests {
    use super::ast::*;
    use super::*;

    #[test]
    fn parses_control_flow() {
        let src = r#"
Option Explicit

Private counter As Long

Public Function Fact(ByVal n As Long) As Long
    If n <= 1 Then
        Fact = 1
    Else
        Fact = n * Fact(n - 1)
    End If
End Function

Sub Main()
    Dim i As Long, s As String
    For i = 1 To 10 Step 2
        s = s & i & ","
    Next i
    Do While i > 0: i = i - 1: Loop
    Select Case i
        Case 0, 1 To 3: Debug.Print "low"
        Case Is > 3
            Debug.Print "high"
        Case Else
    End Select
    With Application
        .StatusBar = "done"
    End With
    MsgBox "Result: " & Fact(5), vbInformation
End Sub
"#;
        let m = parse_module(src, "Module1").unwrap();
        assert!(m.options.explicit);
        assert_eq!(m.declarations.len(), 1);
        assert_eq!(m.procedures.len(), 2);
        let main = &m.procedures[1];
        assert_eq!(main.body.len(), 6);
        assert!(matches!(main.body[1].kind, StmtKind::For { .. }));
        assert!(matches!(&main.body[5].kind, StmtKind::Call { args, .. } if args.len() == 2));
    }

    #[test]
    fn operator_precedence() {
        // -2 ^ 2 = -(2 ^ 2); a + b * c; Not a = b → Not (a = b)
        let e = parse_expression("-2 ^ 2").unwrap();
        assert!(matches!(e, Expr::Unary(UnOp::Neg, _)));
        let e = parse_expression("a + b * c").unwrap();
        assert!(matches!(e, Expr::Binary(BinOp::Add, _, _)));
        let e = parse_expression("Not a = b").unwrap();
        assert!(
            matches!(e, Expr::Unary(UnOp::Not, ref inner) if matches!(**inner, Expr::Binary(BinOp::Eq, _, _)))
        );
    }

    #[test]
    fn statement_lines_are_tracked() {
        let m = parse_module("Sub A()\n  x = 1\n\n  y = 2\nEnd Sub\n", "M").unwrap();
        let lines: Vec<u32> = m.procedures[0].body.iter().map(|s| s.span.line).collect();
        assert_eq!(lines, vec![2, 4]);
        assert_eq!(m.procedures[0].end_line, 5);
    }
}

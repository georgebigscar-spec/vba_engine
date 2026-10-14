//! Line-oriented, case-insensitive VBA tokenizer.
//!
//! Handles `'` and `Rem` comments, ` _` line continuations, `""` escapes in
//! string literals, `&H` / `&O` integer literals and type-declaration suffixes.

use crate::{Span, SyntaxError};

#[derive(Debug, Clone, PartialEq)]
pub enum Tok {
    Ident(String),
    Keyword(Kw),
    Int(i64),
    Float(f64),
    Str(String),
    Date(String),
    Op(&'static str),
    Newline,
    Eof,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kw {
    And,
    As,
    ByRef,
    ByVal,
    Call,
    Case,
    Const,
    Dim,
    Do,
    Each,
    Else,
    ElseIf,
    End,
    Eqv,
    Exit,
    False,
    For,
    Function,
    GoTo,
    If,
    Imp,
    In,
    Is,
    Like,
    Loop,
    Mod,
    New,
    Next,
    Not,
    Nothing,
    Null,
    On,
    Optional,
    Or,
    Private,
    Public,
    ReDim,
    Resume,
    Select,
    Set,
    Step,
    Sub,
    Then,
    To,
    True,
    Until,
    Wend,
    While,
    With,
    Xor,
    Empty,
    Error,
    Option,
    Let,
    Preserve,
}

impl Kw {
    fn lookup(word: &str) -> Option<Kw> {
        use Kw::*;
        Some(match word.to_ascii_lowercase().as_str() {
            "and" => And,
            "as" => As,
            "byref" => ByRef,
            "byval" => ByVal,
            "call" => Call,
            "case" => Case,
            "const" => Const,
            "dim" => Dim,
            "do" => Do,
            "each" => Each,
            "else" => Else,
            "elseif" => ElseIf,
            "end" => End,
            "eqv" => Eqv,
            "exit" => Exit,
            "false" => False,
            "for" => For,
            "function" => Function,
            "goto" => GoTo,
            "if" => If,
            "imp" => Imp,
            "in" => In,
            "is" => Is,
            "like" => Like,
            "loop" => Loop,
            "mod" => Mod,
            "new" => New,
            "next" => Next,
            "not" => Not,
            "nothing" => Nothing,
            "null" => Null,
            "on" => On,
            "optional" => Optional,
            "or" => Or,
            "private" => Private,
            "public" => Public,
            "redim" => ReDim,
            "resume" => Resume,
            "select" => Select,
            "set" => Set,
            "step" => Step,
            "sub" => Sub,
            "then" => Then,
            "to" => To,
            "true" => True,
            "until" => Until,
            "wend" => Wend,
            "while" => While,
            "with" => With,
            "xor" => Xor,
            "empty" => Empty,
            "error" => Error,
            "option" => Option,
            "let" => Let,
            "preserve" => Preserve,
            _ => return None,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Token {
    pub tok: Tok,
    pub span: Span,
}

const OPS: &[&str] = &[
    ":=", "<>", "<=", ">=", "+", "-", "*", "/", "\\", "^", "&", "=", "<", ">", "(", ")", ",", ".",
    ":", "!", ";",
];

pub fn tokenize(src: &str) -> Result<Vec<Token>, SyntaxError> {
    let chars: Vec<char> = src.chars().collect();
    let mut out = Vec::new();
    let (mut i, mut line, mut col) = (0usize, 1u32, 1u32);

    macro_rules! push {
        ($tok:expr, $l:expr, $c:expr) => {
            out.push(Token {
                tok: $tok,
                span: Span { line: $l, col: $c },
            })
        };
    }

    while i < chars.len() {
        let c = chars[i];
        let (sl, sc) = (line, col);
        match c {
            ' ' | '\t' => {
                // Line continuation: whitespace, `_`, optional whitespace, newline.
                if i + 1 < chars.len() && chars[i + 1] == '_' {
                    let mut j = i + 2;
                    while j < chars.len() && (chars[j] == ' ' || chars[j] == '\t') {
                        j += 1;
                    }
                    if j < chars.len() && (chars[j] == '\n' || chars[j] == '\r') {
                        if chars[j] == '\r' && j + 1 < chars.len() && chars[j + 1] == '\n' {
                            j += 1;
                        }
                        i = j + 1;
                        line += 1;
                        col = 1;
                        continue;
                    }
                }
                i += 1;
                col += 1;
            }
            '\r' => {
                i += 1;
            }
            '\n' => {
                push!(Tok::Newline, sl, sc);
                i += 1;
                line += 1;
                col = 1;
            }
            '\'' => {
                while i < chars.len() && chars[i] != '\n' {
                    i += 1;
                }
            }
            '"' => {
                let mut s = String::new();
                i += 1;
                col += 1;
                loop {
                    match chars.get(i) {
                        None | Some('\n') => {
                            return Err(SyntaxError::new(
                                "unterminated string literal",
                                Span { line: sl, col: sc },
                            ))
                        }
                        Some('"') if chars.get(i + 1) == Some(&'"') => {
                            s.push('"');
                            i += 2;
                            col += 2;
                        }
                        Some('"') => {
                            i += 1;
                            col += 1;
                            break;
                        }
                        Some(&ch) => {
                            s.push(ch);
                            i += 1;
                            col += 1;
                        }
                    }
                }
                push!(Tok::Str(s), sl, sc);
            }
            '#' if out
                .last()
                .is_none_or(|t: &Token| !matches!(t.tok, Tok::Ident(_))) =>
            {
                let start = i + 1;
                let mut j = start;
                while j < chars.len() && chars[j] != '#' && chars[j] != '\n' {
                    j += 1;
                }
                if chars.get(j) != Some(&'#') {
                    return Err(SyntaxError::new(
                        "unterminated date literal",
                        Span { line: sl, col: sc },
                    ));
                }
                let text: String = chars[start..j].iter().collect();
                col += (j + 1 - i) as u32;
                i = j + 1;
                push!(Tok::Date(text.trim().to_string()), sl, sc);
            }
            '&' if matches!(chars.get(i + 1), Some('H' | 'h' | 'O' | 'o')) => {
                let radix = if matches!(chars[i + 1], 'H' | 'h') {
                    16
                } else {
                    8
                };
                let mut j = i + 2;
                while j < chars.len() && chars[j].is_digit(radix) {
                    j += 1;
                }
                let digits: String = chars[i + 2..j].iter().collect();
                let v = i64::from_str_radix(&digits, radix).map_err(|_| {
                    SyntaxError::new("invalid radix literal", Span { line: sl, col: sc })
                })?;
                if matches!(chars.get(j), Some('&' | '%')) {
                    j += 1;
                }
                col += (j - i) as u32;
                i = j;
                push!(Tok::Int(v), sl, sc);
            }
            c if c.is_ascii_digit()
                || (c == '.' && chars.get(i + 1).is_some_and(|d| d.is_ascii_digit())) =>
            {
                let mut j = i;
                let mut is_float = false;
                while j < chars.len() && (chars[j].is_ascii_digit() || chars[j] == '.') {
                    is_float |= chars[j] == '.';
                    j += 1;
                }
                if matches!(chars.get(j), Some('e' | 'E')) {
                    let mut k = j + 1;
                    if matches!(chars.get(k), Some('+' | '-')) {
                        k += 1;
                    }
                    if chars.get(k).is_some_and(|d| d.is_ascii_digit()) {
                        is_float = true;
                        j = k;
                        while j < chars.len() && chars[j].is_ascii_digit() {
                            j += 1;
                        }
                    }
                }
                let text: String = chars[i..j].iter().collect();
                // Type suffixes: % & ^ (integral) and ! # @ (floating).
                match chars.get(j) {
                    Some('!' | '#' | '@') => {
                        is_float = true;
                        j += 1;
                    }
                    Some('%' | '&' | '^') if !is_float => j += 1,
                    _ => {}
                }
                let tok = if is_float {
                    Tok::Float(text.parse().map_err(|_| {
                        SyntaxError::new("invalid number", Span { line: sl, col: sc })
                    })?)
                } else {
                    match text.parse::<i64>() {
                        Ok(v) => Tok::Int(v),
                        Err(_) => Tok::Float(text.parse().unwrap_or(f64::INFINITY)),
                    }
                };
                col += (j - i) as u32;
                i = j;
                push!(tok, sl, sc);
            }
            c if c.is_alphabetic() || c == '_' => {
                let mut j = i;
                while j < chars.len() && (chars[j].is_alphanumeric() || chars[j] == '_') {
                    j += 1;
                }
                let word: String = chars[i..j].iter().collect();
                // Identifier type suffixes ($ % # @) are accepted and dropped
                // (`Left$` == `Left`). `&` and `!` stay operators.
                if matches!(chars.get(j), Some('$' | '%' | '#' | '@')) {
                    j += 1;
                }
                col += (j - i) as u32;
                i = j;
                if word.eq_ignore_ascii_case("rem") {
                    while i < chars.len() && chars[i] != '\n' {
                        i += 1;
                    }
                    continue;
                }
                match Kw::lookup(&word) {
                    Some(kw) => push!(Tok::Keyword(kw), sl, sc),
                    None => push!(Tok::Ident(word), sl, sc),
                }
            }
            _ => {
                let rest: String = chars[i..chars.len().min(i + 2)].iter().collect();
                let op = OPS
                    .iter()
                    .find(|op| rest.starts_with(**op))
                    .ok_or_else(|| {
                        SyntaxError::new(
                            format!("unexpected character '{c}'"),
                            Span { line: sl, col: sc },
                        )
                    })?;
                i += op.len();
                col += op.len() as u32;
                push!(Tok::Op(op), sl, sc);
            }
        }
    }
    out.push(Token {
        tok: Tok::Newline,
        span: Span { line, col },
    });
    out.push(Token {
        tok: Tok::Eof,
        span: Span { line, col },
    });
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn toks(src: &str) -> Vec<Tok> {
        tokenize(src).unwrap().into_iter().map(|t| t.tok).collect()
    }

    #[test]
    fn strings_comments_and_continuations() {
        let t = toks("x = \"a\"\"b\" ' comment\ny = 1 + _\n  2");
        assert_eq!(
            t,
            vec![
                Tok::Ident("x".into()),
                Tok::Op("="),
                Tok::Str("a\"b".into()),
                Tok::Newline,
                Tok::Ident("y".into()),
                Tok::Op("="),
                Tok::Int(1),
                Tok::Op("+"),
                Tok::Int(2),
                Tok::Newline,
                Tok::Eof,
            ]
        );
    }

    #[test]
    fn keywords_are_case_insensitive() {
        assert_eq!(toks("DIM")[0], Tok::Keyword(Kw::Dim));
        assert_eq!(toks("&HFF")[0], Tok::Int(255));
        assert_eq!(toks("1.5E2")[0], Tok::Float(150.0));
    }
}

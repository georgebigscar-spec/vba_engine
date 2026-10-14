//! Recursive-descent parser producing [`ast::Module`](crate::ast::Module).

use crate::ast::*;
use crate::lexer::{Kw, Tok, Token};
use crate::{Span, SyntaxError};

type PResult<T> = Result<T, SyntaxError>;

pub struct Parser {
    toks: Vec<Token>,
    pos: usize,
}

impl Parser {
    pub fn new(toks: Vec<Token>) -> Self {
        Parser { toks, pos: 0 }
    }

    // ---- token helpers -------------------------------------------------

    fn peek(&self) -> &Tok {
        &self.toks[self.pos].tok
    }

    fn peek_at(&self, n: usize) -> &Tok {
        &self.toks[(self.pos + n).min(self.toks.len() - 1)].tok
    }

    fn span(&self) -> Span {
        self.toks[self.pos].span
    }

    fn bump(&mut self) -> Tok {
        let t = self.toks[self.pos].tok.clone();
        if self.pos < self.toks.len() - 1 {
            self.pos += 1;
        }
        t
    }

    fn is_kw(&self, kw: Kw) -> bool {
        *self.peek() == Tok::Keyword(kw)
    }

    fn is_op(&self, op: &str) -> bool {
        matches!(self.peek(), Tok::Op(o) if *o == op)
    }

    fn eat_kw(&mut self, kw: Kw) -> bool {
        if self.is_kw(kw) {
            self.bump();
            true
        } else {
            false
        }
    }

    fn eat_op(&mut self, op: &str) -> bool {
        if self.is_op(op) {
            self.bump();
            true
        } else {
            false
        }
    }

    fn is_ident(&self, word: &str) -> bool {
        matches!(self.peek(), Tok::Ident(w) if w.eq_ignore_ascii_case(word))
    }

    fn err<T>(&self, msg: impl Into<String>) -> PResult<T> {
        Err(SyntaxError::new(msg, self.span()))
    }

    fn expect_kw(&mut self, kw: Kw) -> PResult<()> {
        if self.eat_kw(kw) {
            Ok(())
        } else {
            self.err(format!(
                "expected '{kw:?}', found {}",
                describe(self.peek())
            ))
        }
    }

    fn expect_op(&mut self, op: &str) -> PResult<()> {
        if self.eat_op(op) {
            Ok(())
        } else {
            self.err(format!("expected '{op}', found {}", describe(self.peek())))
        }
    }

    fn ident(&mut self) -> PResult<String> {
        match self.bump() {
            Tok::Ident(s) => Ok(s),
            other => {
                self.pos -= 1;
                self.err(format!("expected identifier, found {}", describe(&other)))
            }
        }
    }

    /// Identifier or keyword used as a member name (`rng.End`, `ws.Select`).
    fn member_name(&mut self) -> PResult<String> {
        match self.bump() {
            Tok::Ident(s) => Ok(s),
            Tok::Keyword(k) => Ok(format!("{k:?}")),
            other => {
                self.pos -= 1;
                self.err(format!("expected member name, found {}", describe(&other)))
            }
        }
    }

    fn at_stmt_end(&self) -> bool {
        matches!(self.peek(), Tok::Newline | Tok::Eof) || self.is_op(":") || self.is_kw(Kw::Else)
    }

    fn end_stmt(&mut self) -> PResult<()> {
        if self.eat_op(":") || matches!(self.peek(), Tok::Newline) {
            if matches!(self.peek(), Tok::Newline) {
                self.bump();
            }
            Ok(())
        } else if matches!(self.peek(), Tok::Eof) || self.is_kw(Kw::Else) {
            Ok(())
        } else {
            self.err(format!(
                "expected end of statement, found {}",
                describe(self.peek())
            ))
        }
    }

    fn skip_separators(&mut self) {
        while matches!(self.peek(), Tok::Newline) || self.is_op(":") {
            self.bump();
        }
    }

    /// `End <kw>` lookahead used for block terminators.
    fn at_end_of(&self, kw: Kw) -> bool {
        self.is_kw(Kw::End) && *self.peek_at(1) == Tok::Keyword(kw)
    }

    // ---- module level --------------------------------------------------

    pub fn parse_module(&mut self, name: &str) -> PResult<Module> {
        let mut module = Module {
            name: name.to_string(),
            kind: ModuleKind::Standard,
            options: ModuleOptions::default(),
            declarations: Vec::new(),
            procedures: Vec::new(),
            types: Vec::new(),
            enums: Vec::new(),
        };
        loop {
            self.skip_separators();
            if matches!(self.peek(), Tok::Eof) {
                break;
            }
            // Exported `.cls` header: `VERSION 1.0 CLASS` / `BEGIN ... END`.
            if self.is_ident("VERSION") {
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    if self.is_ident("CLASS") {
                        module.kind = ModuleKind::Class;
                    }
                    self.bump();
                }
                continue;
            }
            if self.is_ident("BEGIN") && matches!(self.peek_at(1), Tok::Newline) {
                while !(self.is_kw(Kw::End) && matches!(self.peek_at(1), Tok::Newline | Tok::Eof)) {
                    if matches!(self.peek(), Tok::Eof) {
                        return self.err("unterminated BEGIN block");
                    }
                    self.bump();
                }
                self.bump();
                continue;
            }
            if self.is_ident("Attribute") {
                self.bump();
                if self.is_ident("VB_Name") {
                    self.bump();
                    self.expect_op("=")?;
                    if let Tok::Str(s) = self.bump() {
                        module.name = s;
                    }
                }
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    self.bump();
                }
                continue;
            }
            if self.eat_kw(Kw::Option) {
                if self.is_ident("Explicit") {
                    self.bump();
                    module.options.explicit = true;
                } else if self.is_ident("Base") {
                    self.bump();
                    module.options.base = match self.bump() {
                        Tok::Int(1) => 1,
                        _ => 0,
                    };
                } else if self.is_ident("Compare") {
                    self.bump();
                    module.options.compare_text = self.is_ident("Text");
                    self.bump();
                } else if self.is_ident("Private") || self.is_kw(Kw::Private) {
                    self.bump();
                    self.ident()?; // Option Private Module
                }
                self.end_stmt()?;
                continue;
            }
            let start = self.span();
            // Procedures are public by default, variables private.
            let explicit_vis = if self.eat_kw(Kw::Private) {
                Some(Visibility::Private)
            } else if self.eat_kw(Kw::Public) {
                Some(Visibility::Public)
            } else if self.is_ident("Global") || self.is_ident("Friend") {
                self.bump();
                Some(Visibility::Public)
            } else {
                None
            };
            if self.is_ident("Static") {
                self.bump();
            }
            if self.is_kw(Kw::Sub) || self.is_kw(Kw::Function) || self.at_property() {
                let proc =
                    self.parse_procedure(explicit_vis.unwrap_or(Visibility::Public), start)?;
                module.procedures.push(proc);
            } else if self.is_ident("Type") {
                self.bump();
                module.types.push(self.parse_type_def(start)?);
            } else if self.is_ident("Enum") {
                self.bump();
                module.enums.push(self.parse_enum_def(start)?);
            } else if self.is_ident("Declare") {
                return self.err("Declare statements are not supported");
            } else if self.eat_kw(Kw::Const) {
                let c = self.parse_const_list()?;
                module.declarations.push(Declaration {
                    visibility: explicit_vis.unwrap_or(Visibility::Private),
                    stmt: Stmt {
                        kind: c,
                        span: start,
                    },
                });
                self.end_stmt()?;
            } else {
                // `Dim x`, `Private x As Long`, `Public x`
                self.eat_kw(Kw::Dim);
                let decls = self.parse_var_decls()?;
                module.declarations.push(Declaration {
                    visibility: explicit_vis.unwrap_or(Visibility::Private),
                    stmt: Stmt {
                        kind: StmtKind::Dim(decls),
                        span: start,
                    },
                });
                self.end_stmt()?;
            }
        }
        Ok(module)
    }

    /// `Property Get|Let|Set` lookahead.
    fn at_property(&self) -> bool {
        self.is_ident("Property")
            && (matches!(self.peek_at(1), Tok::Ident(w) if w.eq_ignore_ascii_case("Get"))
                || matches!(self.peek_at(1), Tok::Keyword(Kw::Let | Kw::Set)))
    }

    /// `End <word>` where the word is not a keyword (`End Type`, `End Property`).
    fn at_end_word(&self, word: &str) -> bool {
        self.is_kw(Kw::End)
            && matches!(self.peek_at(1), Tok::Ident(w) if w.eq_ignore_ascii_case(word))
    }

    fn parse_type_def(&mut self, span: Span) -> PResult<TypeDef> {
        let name = self.ident()?;
        self.end_stmt()?;
        let mut fields = Vec::new();
        loop {
            self.skip_separators();
            if self.at_end_word("Type") {
                self.bump();
                self.bump();
                self.end_stmt()?;
                return Ok(TypeDef { name, fields, span });
            }
            if matches!(self.peek(), Tok::Eof) {
                return self.err("expected 'End Type'");
            }
            fields.extend(self.parse_var_decls()?);
            self.end_stmt()?;
        }
    }

    fn parse_enum_def(&mut self, span: Span) -> PResult<EnumDef> {
        let name = self.ident()?;
        self.end_stmt()?;
        let mut members = Vec::new();
        loop {
            self.skip_separators();
            if self.at_end_word("Enum") {
                self.bump();
                self.bump();
                self.end_stmt()?;
                return Ok(EnumDef {
                    name,
                    members,
                    span,
                });
            }
            if matches!(self.peek(), Tok::Eof) {
                return self.err("expected 'End Enum'");
            }
            let member = self.member_name()?;
            let value = if self.eat_op("=") {
                Some(self.parse_expr()?)
            } else {
                None
            };
            members.push((member, value));
            self.end_stmt()?;
        }
    }

    fn parse_procedure(&mut self, visibility: Visibility, span: Span) -> PResult<Procedure> {
        let kind = if self.eat_kw(Kw::Sub) {
            ProcKind::Sub
        } else if self.eat_kw(Kw::Function) {
            ProcKind::Function
        } else {
            self.bump(); // Property
            match self.bump() {
                Tok::Keyword(Kw::Let) => ProcKind::PropertyLet,
                Tok::Keyword(Kw::Set) => ProcKind::PropertySet,
                _ => ProcKind::PropertyGet,
            }
        };
        let name = self.ident()?;
        let mut params = Vec::new();
        if self.eat_op("(") {
            if !self.is_op(")") {
                loop {
                    params.push(self.parse_param()?);
                    if !self.eat_op(",") {
                        break;
                    }
                }
            }
            self.expect_op(")")?;
        }
        let return_type = if self.eat_kw(Kw::As) {
            Some(self.parse_type()?)
        } else {
            None
        };
        self.end_stmt()?;
        let body = match kind {
            ProcKind::Sub => self.parse_block(&|p| p.at_end_of(Kw::Sub))?,
            ProcKind::Function => self.parse_block(&|p| p.at_end_of(Kw::Function))?,
            _ => self.parse_block(&|p| p.at_end_word("Property"))?,
        };
        let end_line = self.span().line;
        self.expect_kw(Kw::End)?;
        self.bump();
        self.end_stmt()?;
        Ok(Procedure {
            name,
            kind,
            visibility,
            params,
            return_type,
            body,
            span,
            end_line,
        })
    }

    fn parse_param(&mut self) -> PResult<Param> {
        let optional = self.eat_kw(Kw::Optional);
        let mut by_val = false;
        if self.eat_kw(Kw::ByVal) {
            by_val = true;
        } else {
            self.eat_kw(Kw::ByRef);
        }
        let param_array = self.is_ident("ParamArray");
        if param_array {
            self.bump();
        }
        let name = self.ident()?;
        if self.eat_op("(") {
            self.expect_op(")")?;
        }
        let ty = if self.eat_kw(Kw::As) {
            Some(self.parse_type()?)
        } else {
            None
        };
        let default = if self.eat_op("=") {
            Some(self.parse_expr()?)
        } else {
            None
        };
        Ok(Param {
            name,
            by_val,
            optional,
            param_array,
            ty,
            default,
        })
    }

    fn parse_type(&mut self) -> PResult<TypeName> {
        let is_new = self.eat_kw(Kw::New);
        let mut name = self.member_name()?;
        while self.eat_op(".") {
            name.push('.');
            name.push_str(&self.member_name()?);
        }
        // `String * 10` fixed-length strings: length is ignored for now.
        if self.eat_op("*") {
            self.parse_primary()?;
        }
        Ok(TypeName { name, is_new })
    }

    fn parse_var_decls(&mut self) -> PResult<Vec<VarDecl>> {
        let mut decls = Vec::new();
        loop {
            if self.is_ident("WithEvents") {
                self.bump();
            }
            let name = self.ident()?;
            let mut bounds = None;
            if self.eat_op("(") {
                let mut dims = Vec::new();
                if !self.is_op(")") {
                    loop {
                        let first = self.parse_expr()?;
                        if self.eat_kw(Kw::To) {
                            dims.push((Some(first), self.parse_expr()?));
                        } else {
                            dims.push((None, first));
                        }
                        if !self.eat_op(",") {
                            break;
                        }
                    }
                }
                self.expect_op(")")?;
                bounds = Some(dims);
            }
            let ty = if self.eat_kw(Kw::As) {
                Some(self.parse_type()?)
            } else {
                None
            };
            decls.push(VarDecl { name, ty, bounds });
            if !self.eat_op(",") {
                break;
            }
        }
        Ok(decls)
    }

    fn parse_const_list(&mut self) -> PResult<StmtKind> {
        let mut items = Vec::new();
        loop {
            let name = self.ident()?;
            if self.eat_kw(Kw::As) {
                self.parse_type()?;
            }
            self.expect_op("=")?;
            items.push((name, self.parse_expr()?));
            if !self.eat_op(",") {
                break;
            }
        }
        Ok(StmtKind::Const(items))
    }

    // ---- statements ----------------------------------------------------

    fn parse_block(&mut self, is_end: &dyn Fn(&Parser) -> bool) -> PResult<Vec<Stmt>> {
        let mut out = Vec::new();
        loop {
            self.skip_separators();
            if is_end(self) {
                return Ok(out);
            }
            if matches!(self.peek(), Tok::Eof) {
                return self.err("unexpected end of file inside block");
            }
            out.push(self.parse_statement()?);
            self.end_stmt()?;
        }
    }

    fn parse_statement(&mut self) -> PResult<Stmt> {
        let span = self.span();
        let kind = self.parse_statement_kind()?;
        Ok(Stmt { kind, span })
    }

    fn parse_statement_kind(&mut self) -> PResult<StmtKind> {
        // Labels: `Name:` or a line number at the start of a statement.
        if let Tok::Ident(name) = self.peek().clone() {
            if matches!(self.peek_at(1), Tok::Op(":")) && !matches!(self.peek_at(2), Tok::Op("=")) {
                self.bump();
                return Ok(StmtKind::Label(name));
            }
        }
        if let Tok::Int(n) = *self.peek() {
            self.bump();
            return Ok(StmtKind::Label(n.to_string()));
        }
        if self.is_ident("Static") {
            self.bump();
            return Ok(StmtKind::Dim(self.parse_var_decls()?));
        }
        let Tok::Keyword(kw) = *self.peek() else {
            return self.parse_simple_statement();
        };
        match kw {
            Kw::Dim | Kw::Private | Kw::Public => {
                self.bump();
                Ok(StmtKind::Dim(self.parse_var_decls()?))
            }
            Kw::Const => {
                self.bump();
                self.parse_const_list()
            }
            Kw::ReDim => {
                self.bump();
                let preserve = self.eat_kw(Kw::Preserve);
                Ok(StmtKind::ReDim {
                    preserve,
                    decls: self.parse_var_decls()?,
                })
            }
            Kw::Set => {
                self.bump();
                let target = self.parse_postfix()?;
                self.expect_op("=")?;
                Ok(StmtKind::SetAssign {
                    target,
                    value: self.parse_expr()?,
                })
            }
            Kw::Let => {
                self.bump();
                let target = self.parse_postfix()?;
                self.expect_op("=")?;
                Ok(StmtKind::Assign {
                    target,
                    value: self.parse_expr()?,
                })
            }
            Kw::Call => {
                self.bump();
                match self.parse_postfix()? {
                    Expr::Index(callee, args) => Ok(StmtKind::Call {
                        callee: *callee,
                        args,
                    }),
                    callee => Ok(StmtKind::Call {
                        callee,
                        args: vec![],
                    }),
                }
            }
            Kw::If => self.parse_if(),
            Kw::For => self.parse_for(),
            Kw::Do => self.parse_do(),
            Kw::While => {
                self.bump();
                let cond = self.parse_expr()?;
                let body = self.parse_block(&|p| p.is_kw(Kw::Wend))?;
                self.bump();
                Ok(StmtKind::DoLoop {
                    pre: Some(LoopCond { until: false, cond }),
                    post: None,
                    body,
                })
            }
            Kw::Select => self.parse_select(),
            Kw::With => {
                self.bump();
                let object = self.parse_expr()?;
                let body = self.parse_block(&|p| p.at_end_of(Kw::With))?;
                self.bump();
                self.bump();
                Ok(StmtKind::With { object, body })
            }
            Kw::Exit => {
                self.bump();
                let k = match self.bump() {
                    Tok::Keyword(Kw::Sub) => ExitKind::Sub,
                    Tok::Keyword(Kw::Function) => ExitKind::Function,
                    Tok::Ident(w) if w.eq_ignore_ascii_case("Property") => ExitKind::Property,
                    Tok::Keyword(Kw::For) => ExitKind::For,
                    Tok::Keyword(Kw::Do) => ExitKind::Do,
                    other => {
                        return self.err(format!("unsupported Exit target {}", describe(&other)))
                    }
                };
                Ok(StmtKind::Exit(k))
            }
            Kw::On => {
                self.bump();
                self.expect_kw(Kw::Error)?;
                if self.eat_kw(Kw::Resume) {
                    self.expect_kw(Kw::Next)?;
                    Ok(StmtKind::OnError(OnError::ResumeNext))
                } else {
                    self.expect_kw(Kw::GoTo)?;
                    match self.bump() {
                        Tok::Int(0) => Ok(StmtKind::OnError(OnError::GoToZero)),
                        Tok::Int(n) => Ok(StmtKind::OnError(OnError::GoTo(n.to_string()))),
                        Tok::Ident(l) => Ok(StmtKind::OnError(OnError::GoTo(l))),
                        other => self.err(format!("expected label, found {}", describe(&other))),
                    }
                }
            }
            Kw::Resume => {
                self.bump();
                if self.eat_kw(Kw::Next) {
                    Ok(StmtKind::Resume(ResumeKind::Next))
                } else if self.at_stmt_end() {
                    Ok(StmtKind::Resume(ResumeKind::Same))
                } else {
                    Ok(StmtKind::Resume(ResumeKind::Label(self.label_name()?)))
                }
            }
            Kw::GoTo => {
                self.bump();
                Ok(StmtKind::GoTo(self.label_name()?))
            }
            Kw::End => {
                self.bump();
                Ok(StmtKind::End)
            }
            _ => self.parse_simple_statement(),
        }
    }

    fn label_name(&mut self) -> PResult<String> {
        match self.bump() {
            Tok::Ident(s) => Ok(s),
            Tok::Int(n) => Ok(n.to_string()),
            other => self.err(format!("expected label, found {}", describe(&other))),
        }
    }

    /// Assignment or call statement starting with an expression.
    fn parse_simple_statement(&mut self) -> PResult<StmtKind> {
        let target = self.parse_postfix()?;
        if self.eat_op("=") {
            return Ok(StmtKind::Assign {
                target,
                value: self.parse_expr()?,
            });
        }
        if self.at_stmt_end() {
            return Ok(match target {
                Expr::Index(callee, args) => StmtKind::Call {
                    callee: *callee,
                    args,
                },
                callee => StmtKind::Call {
                    callee,
                    args: vec![],
                },
            });
        }
        let args = self.parse_args(&|p| p.at_stmt_end())?;
        Ok(StmtKind::Call {
            callee: target,
            args,
        })
    }

    fn parse_if(&mut self) -> PResult<StmtKind> {
        self.expect_kw(Kw::If)?;
        let cond = self.parse_expr()?;
        self.expect_kw(Kw::Then)?;
        if !matches!(self.peek(), Tok::Newline) {
            // Single-line form: If c Then a: b Else c: d
            let mut then_body = Vec::new();
            while !matches!(self.peek(), Tok::Newline | Tok::Eof) && !self.is_kw(Kw::Else) {
                then_body.push(self.parse_statement()?);
                self.eat_op(":");
            }
            let else_body = if self.eat_kw(Kw::Else) {
                let mut body = Vec::new();
                while !matches!(self.peek(), Tok::Newline | Tok::Eof) {
                    body.push(self.parse_statement()?);
                    self.eat_op(":");
                }
                Some(body)
            } else {
                None
            };
            return Ok(StmtKind::If {
                branches: vec![(cond, then_body)],
                else_body,
            });
        }
        let is_branch_end =
            |p: &Parser| p.is_kw(Kw::ElseIf) || p.is_kw(Kw::Else) || p.at_end_of(Kw::If);
        let mut branches = vec![(cond, self.parse_block(&is_branch_end)?)];
        let mut else_body = None;
        loop {
            if self.eat_kw(Kw::ElseIf) {
                let c = self.parse_expr()?;
                self.expect_kw(Kw::Then)?;
                branches.push((c, self.parse_block(&is_branch_end)?));
            } else if self.eat_kw(Kw::Else) {
                else_body = Some(self.parse_block(&|p| p.at_end_of(Kw::If))?);
            } else {
                self.bump();
                self.bump();
                break;
            }
        }
        Ok(StmtKind::If {
            branches,
            else_body,
        })
    }

    fn parse_for(&mut self) -> PResult<StmtKind> {
        self.expect_kw(Kw::For)?;
        if self.eat_kw(Kw::Each) {
            let var = self.ident()?;
            self.expect_kw(Kw::In)?;
            let collection = self.parse_expr()?;
            let body = self.parse_next_block()?;
            return Ok(StmtKind::ForEach {
                var,
                collection,
                body,
            });
        }
        let var = self.ident()?;
        self.expect_op("=")?;
        let from = self.parse_expr()?;
        self.expect_kw(Kw::To)?;
        let to = self.parse_expr()?;
        let step = if self.eat_kw(Kw::Step) {
            Some(self.parse_expr()?)
        } else {
            None
        };
        let body = self.parse_next_block()?;
        Ok(StmtKind::For {
            var,
            from,
            to,
            step,
            body,
        })
    }

    fn parse_next_block(&mut self) -> PResult<Vec<Stmt>> {
        let body = self.parse_block(&|p| p.is_kw(Kw::Next))?;
        self.bump();
        if matches!(self.peek(), Tok::Ident(_)) {
            self.bump();
        }
        Ok(body)
    }

    fn parse_do(&mut self) -> PResult<StmtKind> {
        self.expect_kw(Kw::Do)?;
        let pre = self.parse_loop_cond()?;
        let body = self.parse_block(&|p| p.is_kw(Kw::Loop))?;
        self.bump();
        let post = self.parse_loop_cond()?;
        Ok(StmtKind::DoLoop { pre, post, body })
    }

    fn parse_loop_cond(&mut self) -> PResult<Option<LoopCond>> {
        if self.eat_kw(Kw::While) {
            Ok(Some(LoopCond {
                until: false,
                cond: self.parse_expr()?,
            }))
        } else if self.eat_kw(Kw::Until) {
            Ok(Some(LoopCond {
                until: true,
                cond: self.parse_expr()?,
            }))
        } else {
            Ok(None)
        }
    }

    fn parse_select(&mut self) -> PResult<StmtKind> {
        self.expect_kw(Kw::Select)?;
        self.expect_kw(Kw::Case)?;
        let subject = self.parse_expr()?;
        let mut cases = Vec::new();
        let mut else_body = None;
        let is_case_end = |p: &Parser| p.is_kw(Kw::Case) || p.at_end_of(Kw::Select);
        self.skip_separators();
        while self.eat_kw(Kw::Case) {
            if self.eat_kw(Kw::Else) {
                else_body = Some(self.parse_block(&is_case_end)?);
                continue;
            }
            let mut clauses = Vec::new();
            loop {
                if self.eat_kw(Kw::Is) {
                    let op = match self.bump() {
                        Tok::Op("=") => BinOp::Eq,
                        Tok::Op("<>") => BinOp::Ne,
                        Tok::Op("<") => BinOp::Lt,
                        Tok::Op("<=") => BinOp::Le,
                        Tok::Op(">") => BinOp::Gt,
                        Tok::Op(">=") => BinOp::Ge,
                        other => {
                            return self
                                .err(format!("expected comparison, found {}", describe(&other)))
                        }
                    };
                    clauses.push(CaseClause::Is(op, self.parse_expr()?));
                } else {
                    let first = self.parse_expr()?;
                    if self.eat_kw(Kw::To) {
                        clauses.push(CaseClause::Range(first, self.parse_expr()?));
                    } else {
                        clauses.push(CaseClause::Value(first));
                    }
                }
                if !self.eat_op(",") {
                    break;
                }
            }
            cases.push((clauses, self.parse_block(&is_case_end)?));
        }
        if !self.at_end_of(Kw::Select) {
            return self.err("expected 'End Select'");
        }
        self.bump();
        self.bump();
        Ok(StmtKind::Select {
            subject,
            cases,
            else_body,
        })
    }

    // ---- expressions ---------------------------------------------------

    fn parse_args(&mut self, is_end: &dyn Fn(&Parser) -> bool) -> PResult<Vec<Arg>> {
        let mut args = Vec::new();
        if is_end(self) {
            return Ok(args);
        }
        loop {
            if self.is_op(",") || self.is_op(";") || is_end(self) {
                args.push(Arg {
                    name: None,
                    value: None,
                });
            } else if matches!(self.peek(), Tok::Ident(_))
                && matches!(self.peek_at(1), Tok::Op(":="))
            {
                let name = self.ident()?;
                self.bump();
                args.push(Arg {
                    name: Some(name),
                    value: Some(self.parse_expr()?),
                });
            } else {
                if self.eat_kw(Kw::ByVal) {
                    // `Foo ByVal x` is accepted and treated as a plain argument.
                }
                args.push(Arg {
                    name: None,
                    value: Some(self.parse_expr()?),
                });
            }
            if !(self.eat_op(",") || self.eat_op(";")) {
                break;
            }
        }
        Ok(args)
    }

    pub fn parse_expr(&mut self) -> PResult<Expr> {
        self.parse_binary(0)
    }

    fn parse_binary(&mut self, level: usize) -> PResult<Expr> {
        // Lowest to highest precedence; unary operators are handled in between.
        const LEVELS: &[&[(Tok2, BinOp)]] = &[
            &[(Tok2::K(Kw::Imp), BinOp::Imp)],
            &[(Tok2::K(Kw::Eqv), BinOp::Eqv)],
            &[(Tok2::K(Kw::Xor), BinOp::Xor)],
            &[(Tok2::K(Kw::Or), BinOp::Or)],
            &[(Tok2::K(Kw::And), BinOp::And)],
            // level 5 = Not (unary)
            &[
                (Tok2::O("="), BinOp::Eq),
                (Tok2::O("<>"), BinOp::Ne),
                (Tok2::O("<"), BinOp::Lt),
                (Tok2::O("<="), BinOp::Le),
                (Tok2::O(">"), BinOp::Gt),
                (Tok2::O(">="), BinOp::Ge),
                (Tok2::K(Kw::Like), BinOp::Like),
                (Tok2::K(Kw::Is), BinOp::Is),
            ],
            &[(Tok2::O("&"), BinOp::Concat)],
            &[(Tok2::O("+"), BinOp::Add), (Tok2::O("-"), BinOp::Sub)],
            &[(Tok2::K(Kw::Mod), BinOp::Mod)],
            &[(Tok2::O("\\"), BinOp::IntDiv)],
            &[(Tok2::O("*"), BinOp::Mul), (Tok2::O("/"), BinOp::Div)],
        ];
        if level == 5 && self.eat_kw(Kw::Not) {
            let e = self.parse_binary(5)?;
            return Ok(Expr::Unary(UnOp::Not, Box::new(e)));
        }
        if level == LEVELS.len() {
            return self.parse_unary_minus();
        }
        let mut lhs = self.parse_binary(level + 1)?;
        'outer: loop {
            for (t, op) in LEVELS[level] {
                let hit = match t {
                    Tok2::K(k) => self.is_kw(*k),
                    Tok2::O(o) => self.is_op(o),
                };
                if hit {
                    self.bump();
                    let rhs = self.parse_binary(level + 1)?;
                    lhs = Expr::Binary(*op, Box::new(lhs), Box::new(rhs));
                    continue 'outer;
                }
            }
            return Ok(lhs);
        }
    }

    fn parse_unary_minus(&mut self) -> PResult<Expr> {
        if self.eat_op("-") {
            let e = self.parse_unary_minus()?;
            return Ok(Expr::Unary(UnOp::Neg, Box::new(e)));
        }
        if self.eat_op("+") {
            return self.parse_unary_minus();
        }
        self.parse_pow()
    }

    fn parse_pow(&mut self) -> PResult<Expr> {
        let mut lhs = self.parse_postfix()?;
        while self.eat_op("^") {
            let rhs = if self.eat_op("-") {
                Expr::Unary(UnOp::Neg, Box::new(self.parse_postfix()?))
            } else {
                self.parse_postfix()?
            };
            lhs = Expr::Binary(BinOp::Pow, Box::new(lhs), Box::new(rhs));
        }
        Ok(lhs)
    }

    fn parse_postfix(&mut self) -> PResult<Expr> {
        let mut e = self.parse_primary()?;
        loop {
            if self.eat_op(".") {
                e = Expr::Member(Box::new(e), self.member_name()?);
            } else if self.eat_op("!") {
                let key = self.member_name()?;
                e = Expr::Index(
                    Box::new(e),
                    vec![Arg {
                        name: None,
                        value: Some(Expr::Str(key)),
                    }],
                );
            } else if self.is_op("(") {
                self.bump();
                let args = self.parse_args(&|p| p.is_op(")"))?;
                self.expect_op(")")?;
                e = Expr::Index(Box::new(e), args);
            } else {
                return Ok(e);
            }
        }
    }

    fn parse_primary(&mut self) -> PResult<Expr> {
        let span = self.span();
        Ok(match self.bump() {
            Tok::Int(v) => Expr::Int(v),
            Tok::Float(v) => Expr::Float(v),
            Tok::Str(s) => Expr::Str(s),
            Tok::Date(s) => Expr::Date(s),
            Tok::Ident(s) if s.eq_ignore_ascii_case("TypeOf") => {
                let e = self.parse_postfix()?;
                self.expect_kw(Kw::Is)?;
                Expr::TypeOf(Box::new(e), self.parse_type()?.name)
            }
            Tok::Ident(s) => Expr::Ident(s),
            Tok::Keyword(Kw::True) => Expr::Bool(true),
            Tok::Keyword(Kw::False) => Expr::Bool(false),
            Tok::Keyword(Kw::Nothing) => Expr::Nothing,
            Tok::Keyword(Kw::Null) => Expr::Null,
            Tok::Keyword(Kw::Empty) => Expr::Empty,
            Tok::Keyword(Kw::New) => Expr::New(self.parse_type()?.name),
            // `Error(n)` / `Error$` as a function name.
            Tok::Keyword(Kw::Error) => Expr::Ident("Error".into()),
            Tok::Op(".") => Expr::WithMember(self.member_name()?),
            Tok::Op("(") => {
                let e = self.parse_expr()?;
                self.expect_op(")")?;
                e
            }
            other => {
                return Err(SyntaxError::new(
                    format!("expected expression, found {}", describe(&other)),
                    span,
                ))
            }
        })
    }
}

enum Tok2 {
    K(Kw),
    O(&'static str),
}

fn describe(t: &Tok) -> String {
    match t {
        Tok::Ident(s) => format!("identifier '{s}'"),
        Tok::Keyword(k) => format!("keyword '{k:?}'"),
        Tok::Int(v) => format!("number {v}"),
        Tok::Float(v) => format!("number {v}"),
        Tok::Str(s) => format!("string \"{s}\""),
        Tok::Date(s) => format!("date #{s}#"),
        Tok::Op(o) => format!("'{o}'"),
        Tok::Newline => "end of line".into(),
        Tok::Eof => "end of file".into(),
    }
}

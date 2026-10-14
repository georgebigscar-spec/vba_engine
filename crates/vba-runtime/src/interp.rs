//! Tree-walking interpreter.

use std::cell::RefCell;
use std::cmp::Ordering;
use std::collections::HashMap;
use std::rc::Rc;

use vba_syntax::ast::*;

use crate::builtins;
use crate::class::{ClassDef, ClassInstance, ProcSet};
use crate::debug::{DebugAction, DebugContext, DebugHook, Location, StopKind};
use crate::errobj::ErrObject;
use crate::error::{VbaError, VbaResult};
use crate::host::Host;
use crate::object::{same_object, Args, Collection, Invoke, ObjectRef};
use crate::value::{self, UdtField, UdtValue, Value, VarType, VbArray};

const MAX_DEPTH: usize = 1000;

/// Storage for one variable. Shared (`Rc`) so `ByRef` arguments alias the
/// caller's variable.
#[derive(Debug, Clone)]
pub struct Variable {
    pub value: Value,
    pub ty: VarType,
    pub is_const: bool,
}

pub type Slot = Rc<RefCell<Variable>>;

fn slot(value: Value, ty: VarType) -> Slot {
    Rc::new(RefCell::new(Variable {
        value,
        ty,
        is_const: false,
    }))
}

#[derive(Debug, Clone, PartialEq)]
enum ErrMode {
    Off,
    ResumeNext,
    GoTo(String),
}

struct Frame {
    procedure: Rc<Procedure>,
    module: Rc<ModuleInfo>,
    locals: Vec<(String, Slot)>,
    /// The instance a class method runs on (`Me`).
    me: Option<ObjectRef>,
    line: u32,
    with_stack: Vec<WithTarget>,
    on_error: ErrMode,
    in_handler: bool,
    /// Block (by address) and statement index to come back to with `Resume`.
    resume_target: Option<(usize, usize)>,
}

/// The object of a `With` block. A user-defined type is kept as the place
/// expression it came from, so `.Field = x` writes back to the variable.
enum WithTarget {
    Value(Value),
    Place(Expr),
}

pub(crate) struct ModuleInfo {
    name: String,
    explicit: bool,
    text_compare: bool,
}

enum Flow {
    Normal,
    Exit(ExitKind),
    GoTo(String),
    Resume(ResumeKind),
}

pub struct Interpreter {
    procs: HashMap<String, (Rc<Procedure>, Rc<ModuleInfo>)>,
    classes: HashMap<String, Rc<ClassDef>>,
    types: HashMap<String, Rc<TypeDef>>,
    enums: HashMap<String, Vec<(String, Value)>>,
    globals: Vec<(String, Slot)>,
    frames: Vec<Frame>,
    host: Box<dyn Host>,
    hook: Option<Box<dyn DebugHook>>,
    err: Rc<ErrObject>,
    steps: u64,
    /// Set while the debugger evaluates a watch: unknown names are errors
    /// instead of implicit declarations, so inspection has no side effects.
    inspecting: bool,
    /// A `Debug.Print` line left open by a trailing `;` or `,`.
    print_pending: Option<String>,
}

impl Interpreter {
    pub fn new(host: impl Host + 'static) -> Self {
        Interpreter {
            procs: HashMap::new(),
            classes: HashMap::new(),
            types: HashMap::new(),
            enums: HashMap::new(),
            globals: Vec::new(),
            frames: Vec::new(),
            host: Box::new(host),
            hook: None,
            err: Rc::new(ErrObject::default()),
            steps: 0,
            inspecting: false,
            print_pending: None,
        }
    }

    pub fn host_mut(&mut self) -> &mut dyn Host {
        self.host.as_mut()
    }

    pub fn set_debug_hook(
        &mut self,
        hook: Option<Box<dyn DebugHook>>,
    ) -> Option<Box<dyn DebugHook>> {
        std::mem::replace(&mut self.hook, hook)
    }

    /// Loads a parsed module: registers its procedures and initializes its
    /// module-level variables and constants.
    pub fn load_module(&mut self, module: Module) -> VbaResult<()> {
        let info = Rc::new(ModuleInfo {
            name: module.name.clone(),
            explicit: module.options.explicit,
            text_compare: module.options.compare_text,
        });
        for t in module.types {
            self.types.insert(t.name.to_ascii_lowercase(), Rc::new(t));
        }
        for e in &module.enums {
            let mut next = Value::Long(0);
            let mut members = Vec::new();
            for (name, value) in &e.members {
                let v = match value {
                    Some(expr) => VarType::Long.coerce(self.eval(expr)?)?,
                    None => next.clone(),
                };
                next = value::add(&v, &Value::Long(1))?;
                let s = slot(v.clone(), VarType::Long);
                s.borrow_mut().is_const = true;
                self.globals.push((name.clone(), s));
                members.push((name.clone(), v));
            }
            self.enums.insert(e.name.to_ascii_lowercase(), members);
        }
        if module.kind == ModuleKind::Class {
            let mut members: HashMap<String, ProcSet> = HashMap::new();
            for p in module.procedures {
                members
                    .entry(p.name.to_ascii_lowercase())
                    .or_default()
                    .insert(Rc::new(p));
            }
            let mut fields = Vec::new();
            let mut consts = Vec::new();
            for d in module.declarations {
                match d.stmt.kind {
                    StmtKind::Dim(vars) => {
                        fields.extend(vars.into_iter().map(|v| (d.visibility, v)))
                    }
                    StmtKind::Const(items) => consts.extend(items),
                    _ => {}
                }
            }
            let class = ClassDef {
                name: module.name.clone(),
                info,
                members,
                fields,
                consts,
            };
            self.classes
                .insert(module.name.to_ascii_lowercase(), Rc::new(class));
            return Ok(());
        }
        for p in module.procedures {
            self.procs
                .insert(p.name.to_ascii_lowercase(), (Rc::new(p), info.clone()));
        }
        for decl in &module.declarations {
            match &decl.stmt.kind {
                StmtKind::Dim(vars) => {
                    for v in vars {
                        let s = self.new_var_slot(v)?;
                        self.globals.push((v.name.clone(), s));
                    }
                }
                StmtKind::Const(items) => {
                    for (name, e) in items {
                        let value = self.eval(e)?;
                        let s = slot(value, VarType::Variant);
                        s.borrow_mut().is_const = true;
                        self.globals.push((name.clone(), s));
                    }
                }
                _ => {}
            }
        }
        Ok(())
    }

    pub fn procedure_names(&self) -> Vec<String> {
        let mut v: Vec<String> = self.procs.values().map(|(p, _)| p.name.clone()).collect();
        v.sort();
        v
    }

    pub fn procedure(&self, name: &str) -> Option<Rc<Procedure>> {
        self.procs
            .get(&name.to_ascii_lowercase())
            .map(|(p, _)| p.clone())
    }

    /// Runs a `Sub` or `Function` by name with the given arguments.
    pub fn run(&mut self, name: &str, args: Vec<Value>) -> VbaResult<Value> {
        let (proc, module) = self
            .procs
            .get(&name.to_ascii_lowercase())
            .cloned()
            .ok_or_else(|| VbaError::sub_not_defined(name))?;
        self.frames.clear();
        let slots = args.into_iter().map(ArgSource::Value).collect();
        let r = self.call_procedure(proc, module, None, slots, Vec::new());
        self.frames.clear();
        if let Some(line) = self.print_pending.take() {
            self.host.debug_print(&line);
        }
        match r {
            Err(e) if e.is_end() => Ok(Value::Empty),
            other => other,
        }
    }

    // ---- debugger support ---------------------------------------------------

    pub fn depth(&self) -> usize {
        self.frames.len()
    }

    fn frame_index(&self, from_top: usize) -> Option<usize> {
        self.frames.len().checked_sub(from_top + 1)
    }

    pub fn location(&self, frame: usize) -> Option<Location> {
        let f = &self.frames[self.frame_index(frame)?];
        Some(Location {
            module: f.module.name.clone(),
            procedure: f.procedure.name.clone(),
            line: f.line,
        })
    }

    pub fn locals(&self, frame: usize) -> Vec<(String, Value)> {
        self.frame_index(frame)
            .map(|i| {
                let f = &self.frames[i];
                let me =
                    f.me.iter()
                        .map(|o| ("Me".to_string(), Value::Object(Some(o.clone()))));
                me.chain(
                    f.locals
                        .iter()
                        .map(|(n, s)| (n.clone(), s.borrow().value.clone())),
                )
                .collect()
            })
            .unwrap_or_default()
    }

    pub fn module_variables(&self) -> Vec<(String, Value)> {
        self.globals
            .iter()
            .map(|(n, s)| (n.clone(), s.borrow().value.clone()))
            .collect()
    }

    /// Runs `f` with frame `frame` (0 = innermost) temporarily on top.
    fn with_frame<T>(&mut self, frame: usize, f: impl FnOnce(&mut Self) -> T) -> Option<T> {
        let idx = self.frame_index(frame)?;
        let tail = self.frames.split_off(idx + 1);
        let hook = self.hook.take();
        let r = f(self);
        self.hook = hook;
        self.frames.extend(tail);
        Some(r)
    }

    pub fn eval_in_frame(&mut self, frame: usize, src: &str) -> Result<Value, String> {
        let expr = vba_syntax::parse_expression(src).map_err(|e| e.to_string())?;
        let was = std::mem::replace(&mut self.inspecting, true);
        let r = if self.frames.is_empty() {
            Some(self.eval(&expr))
        } else {
            self.with_frame(frame, |me| me.eval(&expr))
        };
        self.inspecting = was;
        r.ok_or_else(|| "no such frame".to_string())?
            .map_err(|e| e.description)
    }

    pub fn execute_in_frame(&mut self, frame: usize, src: &str) -> Result<(), String> {
        let wrapped = format!("Sub Immediate__()\n{src}\nEnd Sub\n");
        let module = vba_syntax::parse_module(&wrapped, "Immediate").map_err(|e| e.to_string())?;
        let body = module
            .procedures
            .into_iter()
            .next()
            .map(|p| p.body)
            .unwrap_or_default();
        let run = |me: &mut Self| me.exec_block(&body).map(|_| ()).map_err(|e| e.to_string());
        if self.frames.is_empty() {
            return run(self);
        }
        self.with_frame(frame, run)
            .ok_or_else(|| "no such frame".to_string())?
    }

    fn with_hook(
        &mut self,
        f: impl FnOnce(&mut dyn DebugHook, &mut DebugContext<'_>) -> DebugAction,
    ) -> VbaResult<()> {
        let Some(mut hook) = self.hook.take() else {
            return Ok(());
        };
        let action = f(hook.as_mut(), &mut DebugContext { interp: self });
        self.hook = Some(hook);
        match action {
            DebugAction::Continue => Ok(()),
            DebugAction::Abort => Err(VbaError::interrupted()),
        }
    }

    fn before_statement(&mut self, line: u32) -> VbaResult<()> {
        if let Some(f) = self.frames.last_mut() {
            f.line = line;
        }
        self.steps += 1;
        if self.steps.is_multiple_of(1024) && !self.host.poll() {
            return Err(VbaError::interrupted());
        }
        self.with_hook(|h, ctx| h.on_statement(ctx))
    }

    // ---- procedures ---------------------------------------------------------

    fn call_procedure(
        &mut self,
        proc: Rc<Procedure>,
        module: Rc<ModuleInfo>,
        me: Option<ObjectRef>,
        positional: Vec<ArgSource>,
        named: Vec<(String, ArgSource)>,
    ) -> VbaResult<Value> {
        if self.frames.len() >= MAX_DEPTH {
            return Err(VbaError::new(28, "Out of stack space"));
        }
        let has_param_array = proc.params.last().is_some_and(|p| p.param_array);
        if positional.len() > proc.params.len() && !has_param_array {
            return Err(VbaError::wrong_arg_count());
        }
        let mut positional: Vec<Option<ArgSource>> = positional.into_iter().map(Some).collect();
        let mut named = named;
        let mut locals = Vec::with_capacity(proc.params.len() + 1);
        for (i, p) in proc.params.iter().enumerate() {
            if p.param_array {
                let rest: Vec<Value> = positional
                    .iter_mut()
                    .skip(i)
                    .filter_map(Option::take)
                    .map(ArgSource::into_value)
                    .collect();
                let arr = Value::Array(Box::new(VbArray::from_values(rest, 0)));
                locals.push((p.name.clone(), slot(arr, VarType::Variant)));
                break;
            }
            let ty = self.var_type_of(p.ty.as_ref());
            let src = named
                .iter()
                .position(|(n, _)| n.eq_ignore_ascii_case(&p.name))
                .map(|k| named.remove(k).1)
                .or_else(|| positional.get_mut(i).and_then(Option::take))
                .filter(|s| !matches!(s, ArgSource::Value(Value::Missing)));
            let s = match src {
                Some(ArgSource::Slot(s)) if !p.by_val => s,
                Some(ArgSource::Slot(s)) => {
                    let v = s.borrow().value.clone();
                    slot(ty.coerce(v)?, ty)
                }
                Some(ArgSource::Value(v)) => slot(ty.coerce(v)?, ty),
                None if p.optional => {
                    let v = match &p.default {
                        Some(e) => ty.coerce(self.eval(e)?)?,
                        None if ty == VarType::Variant => Value::Missing,
                        None => ty.default_value(),
                    };
                    slot(v, ty)
                }
                None => {
                    return Err(VbaError::new(
                        449,
                        format!("Argument not optional: {}", p.name),
                    ))
                }
            };
            locals.push((p.name.clone(), s));
        }
        if let Some((n, _)) = named.first() {
            return Err(VbaError::new(448, format!("Named argument not found: {n}")));
        }
        let returns = matches!(proc.kind, ProcKind::Function | ProcKind::PropertyGet);
        if returns {
            let (ty, v) = self.default_of(proc.return_type.as_ref())?;
            locals.push((proc.name.clone(), slot(v, ty)));
        }
        self.frames.push(Frame {
            procedure: proc.clone(),
            module,
            locals,
            me,
            line: proc.span.line,
            with_stack: Vec::new(),
            on_error: ErrMode::Off,
            in_handler: false,
            resume_target: None,
        });
        let _ = self.with_hook(|h, ctx| {
            h.on_enter(ctx);
            DebugAction::Continue
        });
        let mut result = self.exec_block(&proc.body).and_then(|flow| match flow {
            Flow::GoTo(l) => Err(VbaError::label_not_defined(&l)),
            Flow::Resume(_) => Err(VbaError::new(20, "Resume without error")),
            _ => Ok(()),
        });
        if result.is_ok() {
            result = self.before_statement(proc.end_line);
        }
        let _ = self.with_hook(|h, ctx| {
            h.on_exit(ctx);
            DebugAction::Continue
        });
        let frame = self.frames.pop().expect("frame");
        if result.is_ok() && (frame.on_error != ErrMode::Off || frame.in_handler) {
            // Exit Sub / End Sub from a procedure with a handler resets Err.
            self.err.clear();
        }
        result?;
        Ok(if returns {
            frame
                .locals
                .iter()
                .rev()
                .find(|(n, _)| n.eq_ignore_ascii_case(&proc.name))
                .map(|(_, s)| s.borrow().value.clone())
                .unwrap_or(Value::Empty)
        } else {
            Value::Empty
        })
    }

    // ---- statements ---------------------------------------------------------

    fn frame(&self) -> Option<&Frame> {
        self.frames.last()
    }

    fn frame_mut(&mut self) -> Option<&mut Frame> {
        self.frames.last_mut()
    }

    fn handlers_active(&self) -> bool {
        self.frames
            .iter()
            .any(|f| f.on_error != ErrMode::Off && !f.in_handler)
    }

    fn exec_block(&mut self, stmts: &[Stmt]) -> VbaResult<Flow> {
        let block_id = stmts.as_ptr() as usize;
        let mut pc = 0;
        while pc < stmts.len() {
            match self.exec_stmt(&stmts[pc]) {
                Ok(Flow::Normal) => pc += 1,
                Ok(Flow::GoTo(label)) => match find_label(stmts, &label) {
                    Some(i) => pc = i,
                    None => return Ok(Flow::GoTo(label)),
                },
                Ok(Flow::Resume(kind)) => {
                    let target = self.frame().and_then(|f| f.resume_target);
                    match target {
                        Some((id, at)) if id == block_id => {
                            pc = match &kind {
                                ResumeKind::Same => at,
                                ResumeKind::Next => at + 1,
                                ResumeKind::Label(l) => find_label(stmts, l)
                                    .ok_or_else(|| VbaError::label_not_defined(l))?,
                            };
                            self.err.clear();
                            if let Some(f) = self.frame_mut() {
                                f.in_handler = false;
                                f.resume_target = None;
                            }
                        }
                        _ => return Ok(Flow::Resume(kind)),
                    }
                }
                Ok(flow) => return Ok(flow),
                Err(e) if !e.trappable => return Err(e),
                Err(mut e) => {
                    if !e.debugger_notified {
                        // Report each error to the debugger once, at the innermost block.
                        let handled = self.handlers_active();
                        self.with_hook(|h, ctx| h.on_error(ctx, &e, handled))?;
                        e.debugger_notified = true;
                    }
                    let Some(frame) = self.frames.last_mut() else {
                        return Err(e);
                    };
                    match frame.on_error.clone() {
                        ErrMode::ResumeNext => {
                            self.err.set(&e);
                            pc += 1;
                        }
                        ErrMode::GoTo(label) if !frame.in_handler => {
                            match find_label(stmts, &label) {
                                Some(i) => {
                                    frame.in_handler = true;
                                    frame.resume_target = Some((block_id, pc));
                                    self.err.set(&e);
                                    pc = i;
                                }
                                None => return Err(e),
                            }
                        }
                        _ => return Err(e),
                    }
                }
            }
        }
        Ok(Flow::Normal)
    }

    fn exec_stmt(&mut self, stmt: &Stmt) -> VbaResult<Flow> {
        let line = stmt.span.line;
        if !matches!(
            stmt.kind,
            StmtKind::Label(_) | StmtKind::Dim(_) | StmtKind::Const(_)
        ) {
            self.before_statement(line)?;
        } else if let Some(f) = self.frame_mut() {
            f.line = line;
        }
        self.exec_stmt_kind(&stmt.kind)
            .map_err(|e| e.with_line(line))
    }

    fn exec_stmt_kind(&mut self, kind: &StmtKind) -> VbaResult<Flow> {
        match kind {
            StmtKind::Label(_) => {}
            StmtKind::Dim(vars) => {
                for v in vars {
                    let s = self.new_var_slot(v)?;
                    match self.frame_mut() {
                        Some(f) => f.locals.push((v.name.clone(), s)),
                        None => self.globals.push((v.name.clone(), s)),
                    }
                }
            }
            StmtKind::Const(items) => {
                for (name, e) in items {
                    let v = self.eval(e)?;
                    let s = slot(v, VarType::Variant);
                    s.borrow_mut().is_const = true;
                    if let Some(f) = self.frame_mut() {
                        f.locals.push((name.clone(), s));
                    }
                }
            }
            StmtKind::ReDim { preserve, decls } => {
                for d in decls {
                    let bounds = self.eval_bounds(d.bounds.as_deref().unwrap_or(&[]))?;
                    let existing = self.find_slot(&d.name);
                    match existing {
                        Some(s) => {
                            let is_array = matches!(s.borrow().value, Value::Array(_));
                            if *preserve && is_array {
                                if let Value::Array(a) = &mut s.borrow_mut().value {
                                    a.redim_preserve(bounds)?;
                                }
                            } else {
                                let arr = self.new_array(bounds, d.ty.as_ref())?;
                                s.borrow_mut().value = Value::Array(Box::new(arr));
                            }
                        }
                        None => {
                            let arr = self.new_array(bounds, d.ty.as_ref())?;
                            self.declare_implicit(&d.name, Value::Array(Box::new(arr)))?;
                        }
                    }
                }
            }
            StmtKind::Assign { target, value } => {
                let v = self.eval(value)?;
                self.assign(target, v, false)?;
            }
            StmtKind::SetAssign { target, value } => {
                let v = self.eval(value)?;
                if !matches!(v, Value::Object(_)) {
                    return Err(VbaError::object_required());
                }
                self.assign(target, v, true)?;
            }
            StmtKind::Call { callee, args } => return self.exec_call(callee, args),
            StmtKind::DebugPrint(items) => self.exec_print(items)?,
            StmtKind::If {
                branches,
                else_body,
            } => {
                for (cond, body) in branches {
                    let c = self.eval(cond)?;
                    if truthy(&c)? {
                        return self.exec_block(body);
                    }
                }
                if let Some(body) = else_body {
                    return self.exec_block(body);
                }
            }
            StmtKind::For {
                var,
                from,
                to,
                step,
                body,
            } => {
                let start = self.eval(from)?;
                let end = self.eval(to)?.into_scalar()?;
                let step = match step {
                    Some(s) => self.eval(s)?.into_scalar()?,
                    None => Value::Integer(1),
                };
                let counter = match self.find_slot(var) {
                    Some(s) => s,
                    None => self.declare_implicit(var, Value::Empty)?,
                };
                {
                    let mut c = counter.borrow_mut();
                    c.value = c.ty.coerce(start.into_scalar()?)?;
                }
                let descending = step.to_f64()? < 0.0;
                loop {
                    let cur = counter.borrow().value.clone();
                    let ord = value::compare(&cur, &end, false)?;
                    if (!descending && ord == Some(Ordering::Greater))
                        || (descending && ord == Some(Ordering::Less))
                    {
                        break;
                    }
                    match self.exec_block(body)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                    let mut c = counter.borrow_mut();
                    let next = value::add(&c.value, &step)?;
                    c.value = c.ty.coerce(next)?;
                }
            }
            StmtKind::ForEach {
                var,
                collection,
                body,
            } => {
                let items = match self.eval(collection)? {
                    Value::Array(a) => a.data,
                    Value::Object(Some(o)) => o.enumerate()?,
                    Value::Object(None) => return Err(VbaError::object_not_set()),
                    _ => return Err(VbaError::new(92, "For loop not initialized")),
                };
                let target = Expr::Ident(var.clone());
                for item in items {
                    let is_obj = matches!(item, Value::Object(_));
                    self.assign(&target, item, is_obj)?;
                    match self.exec_block(body)? {
                        Flow::Normal => {}
                        Flow::Exit(ExitKind::For) => break,
                        other => return Ok(other),
                    }
                }
            }
            StmtKind::DoLoop { pre, post, body } => loop {
                if let Some(c) = pre {
                    if !self.loop_cond(c)? {
                        break;
                    }
                }
                match self.exec_block(body)? {
                    Flow::Normal => {}
                    Flow::Exit(ExitKind::Do) => break,
                    other => return Ok(other),
                }
                if let Some(c) = post {
                    if !self.loop_cond(c)? {
                        break;
                    }
                }
            },
            StmtKind::Select {
                subject,
                cases,
                else_body,
            } => {
                let subj = self.eval(subject)?.into_scalar()?;
                let text = self.text_compare();
                for (clauses, body) in cases {
                    for clause in clauses {
                        let hit = match clause {
                            CaseClause::Value(e) => {
                                let v = self.eval(e)?;
                                value::compare(&subj, &v, text)? == Some(Ordering::Equal)
                            }
                            CaseClause::Range(lo, hi) => {
                                let (lo, hi) = (self.eval(lo)?, self.eval(hi)?);
                                matches!(
                                    value::compare(&subj, &lo, text)?,
                                    Some(Ordering::Greater | Ordering::Equal)
                                ) && matches!(
                                    value::compare(&subj, &hi, text)?,
                                    Some(Ordering::Less | Ordering::Equal)
                                )
                            }
                            CaseClause::Is(op, e) => {
                                let v = self.eval(e)?;
                                compare_op(*op, value::compare(&subj, &v, text)?)
                            }
                        };
                        if hit {
                            return self.exec_block(body);
                        }
                    }
                }
                if let Some(body) = else_body {
                    return self.exec_block(body);
                }
            }
            StmtKind::With { object, body } => {
                let obj = self.eval(object)?;
                let target = match obj {
                    Value::Udt(_)
                        if matches!(
                            object,
                            Expr::Ident(_) | Expr::Index(..) | Expr::Member(..)
                        ) =>
                    {
                        WithTarget::Place(object.clone())
                    }
                    v => WithTarget::Value(v),
                };
                if let Some(f) = self.frame_mut() {
                    f.with_stack.push(target);
                }
                let r = self.exec_block(body);
                if let Some(f) = self.frame_mut() {
                    f.with_stack.pop();
                }
                return r;
            }
            StmtKind::Exit(k) => return Ok(Flow::Exit(*k)),
            StmtKind::OnError(mode) => {
                self.err.clear();
                if let Some(f) = self.frame_mut() {
                    f.on_error = match mode {
                        OnError::ResumeNext => ErrMode::ResumeNext,
                        OnError::GoTo(l) => ErrMode::GoTo(l.clone()),
                        OnError::GoToZero => ErrMode::Off,
                    };
                    if matches!(mode, OnError::GoToZero) {
                        f.in_handler = false;
                    }
                }
            }
            StmtKind::Resume(k) => return Ok(Flow::Resume(k.clone())),
            StmtKind::GoTo(l) => return Ok(Flow::GoTo(l.clone())),
            StmtKind::End => return Err(VbaError::end()),
        }
        Ok(Flow::Normal)
    }

    fn loop_cond(&mut self, c: &LoopCond) -> VbaResult<bool> {
        let v = self.eval(&c.cond)?;
        Ok(truthy(&v)? != c.until)
    }

    /// `Debug.Print`: numbers get a sign space before and a space after,
    /// `;` joins, `,` moves to the next 14-column zone, and a trailing
    /// separator leaves the line open for the next `Debug.Print`.
    fn exec_print(&mut self, items: &[PrintItem]) -> VbaResult<()> {
        let mut line = self.print_pending.take().unwrap_or_default();
        let mut open = false;
        for item in items {
            match item {
                PrintItem::Expr(e) => {
                    let v = match self.eval(e) {
                        Ok(v) => v,
                        Err(err) => {
                            self.print_pending = Some(line).filter(|l| !l.is_empty());
                            return Err(err);
                        }
                    };
                    line.push_str(&print_text(v)?);
                    open = false;
                }
                PrintItem::Semicolon => open = true,
                PrintItem::Comma => {
                    let width = line.chars().count();
                    line.extend(std::iter::repeat_n(' ', (width / 14 + 1) * 14 - width));
                    open = true;
                }
            }
        }
        if open {
            self.print_pending = Some(line);
        } else {
            self.host.debug_print(&line);
        }
        Ok(())
    }

    fn exec_call(&mut self, callee: &Expr, args: &[Arg]) -> VbaResult<Flow> {
        // Debug.Print / Debug.Assert
        if let Expr::Member(obj, member) = callee {
            if matches!(&**obj, Expr::Ident(n) if n.eq_ignore_ascii_case("debug"))
                && self.find_slot("debug").is_none()
            {
                if member.eq_ignore_ascii_case("print") {
                    let mut parts = Vec::new();
                    for a in args {
                        parts.push(match &a.value {
                            Some(e) => match self.eval(e)? {
                                Value::Null => "Null".to_string(),
                                v => v.into_scalar()?.to_string(),
                            },
                            None => String::new(),
                        });
                    }
                    self.host.debug_print(&parts.join(" "));
                    return Ok(Flow::Normal);
                }
                if member.eq_ignore_ascii_case("assert") {
                    let cond = match args.first().and_then(|a| a.value.as_ref()) {
                        Some(e) => self.eval(e)?,
                        None => return Err(VbaError::wrong_arg_count()),
                    };
                    if !truthy(&cond)? {
                        self.with_hook(|h, ctx| h.on_stop(ctx, StopKind::AssertFailed))?;
                    }
                    return Ok(Flow::Normal);
                }
            }
        }
        if let Expr::Ident(name) = callee {
            if name.eq_ignore_ascii_case("stop")
                && args.is_empty()
                && !self.procs.contains_key("stop")
            {
                self.with_hook(|h, ctx| h.on_stop(ctx, StopKind::Stop))?;
                return Ok(Flow::Normal);
            }
        }
        self.eval_call(callee, args, Invoke::Method)?;
        Ok(Flow::Normal)
    }

    // ---- variables ----------------------------------------------------------

    fn find_slot(&self, name: &str) -> Option<Slot> {
        if let Some(f) = self.frame() {
            if let Some((_, s)) = f
                .locals
                .iter()
                .rev()
                .find(|(n, _)| n.eq_ignore_ascii_case(name))
            {
                return Some(s.clone());
            }
            if let Some(inst) =
                f.me.as_ref()
                    .and_then(|o| o.as_any().downcast_ref::<ClassInstance>())
            {
                if let Some((s, _)) = inst.field_slot(name) {
                    return Some(s.clone());
                }
            }
        }
        self.globals
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, s)| s.clone())
    }

    fn explicit(&self) -> bool {
        self.frame().is_some_and(|f| f.module.explicit)
    }

    fn text_compare(&self) -> bool {
        self.frame().is_some_and(|f| f.module.text_compare)
    }

    fn declare_implicit(&mut self, name: &str, value: Value) -> VbaResult<Slot> {
        if self.explicit() || self.inspecting {
            return Err(VbaError::variable_not_defined(name));
        }
        let s = slot(value, VarType::Variant);
        match self.frame_mut() {
            Some(f) => f.locals.push((name.to_string(), s.clone())),
            None => self.globals.push((name.to_string(), s.clone())),
        }
        Ok(s)
    }

    fn new_var_slot(&mut self, v: &VarDecl) -> VbaResult<Slot> {
        if let Some(bounds) = &v.bounds {
            let bounds = self.eval_bounds(bounds)?;
            let arr = self.new_array(bounds, v.ty.as_ref())?;
            return Ok(slot(Value::Array(Box::new(arr)), VarType::Variant));
        }
        let (ty, value) = self.default_of(v.ty.as_ref())?;
        Ok(slot(value, ty))
    }

    /// Storage type for a declared type name: enums are Longs, user-defined
    /// types are held as Variants.
    fn var_type_of(&self, ty: Option<&TypeName>) -> VarType {
        let Some(t) = ty else { return VarType::Variant };
        let lname = t.name.to_ascii_lowercase();
        if self.enums.contains_key(&lname) {
            VarType::Long
        } else if self.types.contains_key(&lname) {
            VarType::Variant
        } else {
            VarType::from_name(&t.name)
        }
    }

    /// Storage type and initial value for a declaration.
    fn default_of(&mut self, ty: Option<&TypeName>) -> VbaResult<(VarType, Value)> {
        let Some(t) = ty else {
            return Ok((VarType::Variant, Value::Empty));
        };
        if let Some(def) = self.types.get(&t.name.to_ascii_lowercase()).cloned() {
            return Ok((VarType::Variant, self.new_udt(&def)?));
        }
        let vt = self.var_type_of(Some(t));
        if t.is_new {
            return Ok((vt, self.new_object(&t.name)?));
        }
        Ok((vt, vt.default_value()))
    }

    fn new_udt(&mut self, def: &TypeDef) -> VbaResult<Value> {
        let mut fields = Vec::with_capacity(def.fields.len());
        for f in &def.fields {
            let s = self.new_var_slot(f)?;
            let var = s.borrow().clone();
            fields.push(UdtField {
                name: f.name.clone(),
                ty: var.ty,
                value: var.value,
            });
        }
        Ok(Value::Udt(Box::new(UdtValue {
            type_name: def.name.clone(),
            fields,
        })))
    }

    fn new_array(&mut self, bounds: Vec<(i64, i64)>, ty: Option<&TypeName>) -> VbaResult<VbArray> {
        let elem = self.var_type_of(ty);
        let mut arr = VbArray::new(bounds, elem);
        if let Some(def) = ty.and_then(|t| self.types.get(&t.name.to_ascii_lowercase()).cloned()) {
            let init = self.new_udt(&def)?;
            arr.data.iter_mut().for_each(|v| *v = init.clone());
        }
        Ok(arr)
    }

    fn eval_bounds(&mut self, bounds: &[(Option<Expr>, Expr)]) -> VbaResult<Vec<(i64, i64)>> {
        bounds
            .iter()
            .map(|(lo, hi)| {
                let lo = match lo {
                    Some(e) => self.eval(e)?.to_i64()?,
                    None => 0,
                };
                Ok((lo, self.eval(hi)?.to_i64()?))
            })
            .collect()
    }

    fn new_object(&mut self, class: &str) -> VbaResult<Value> {
        if class.eq_ignore_ascii_case("collection") || class.eq_ignore_ascii_case("vba.collection")
        {
            return Ok(Value::from(Collection::new()));
        }
        if let Some(def) = self.classes.get(&class.to_ascii_lowercase()).cloned() {
            return self.instantiate(def);
        }
        self.host.new_object(class)
    }

    fn instantiate(&mut self, def: Rc<ClassDef>) -> VbaResult<Value> {
        let mut fields = Vec::new();
        for (vis, decl) in &def.fields {
            fields.push((decl.name.clone(), self.new_var_slot(decl)?, *vis));
        }
        for (name, e) in &def.consts {
            let s = slot(self.eval(e)?, VarType::Variant);
            s.borrow_mut().is_const = true;
            fields.push((name.clone(), s, Visibility::Private));
        }
        let obj: ObjectRef = Rc::new(ClassInstance {
            class: def.clone(),
            fields,
        });
        if let Some(init) = def
            .members
            .get("class_initialize")
            .and_then(|m| m.call.clone())
        {
            self.call_procedure(
                init,
                def.info.clone(),
                Some(obj.clone()),
                Vec::new(),
                Vec::new(),
            )?;
        }
        Ok(Value::Object(Some(obj)))
    }

    fn me(&self) -> Option<ObjectRef> {
        self.frame().and_then(|f| f.me.clone())
    }

    /// Invokes a member on any object, running VBA class methods here.
    fn invoke_object(
        &mut self,
        o: &ObjectRef,
        member: &str,
        kind: Invoke,
        args: Args,
    ) -> VbaResult<Value> {
        if o.as_any().is::<ClassInstance>() {
            let positional = args.positional.into_iter().map(ArgSource::Value).collect();
            let named = args
                .named
                .into_iter()
                .map(|(n, v)| (n, ArgSource::Value(v)))
                .collect();
            return self.call_member(o, member, kind, positional, named);
        }
        o.invoke(member, kind, args)
    }

    /// Calls a member of a VBA class instance: a method, a property
    /// procedure, or a field.
    fn call_member(
        &mut self,
        obj: &ObjectRef,
        member: &str,
        kind: Invoke,
        mut positional: Vec<ArgSource>,
        named: Vec<(String, ArgSource)>,
    ) -> VbaResult<Value> {
        let inst = obj
            .as_any()
            .downcast_ref::<ClassInstance>()
            .expect("class instance");
        let class = inst.class.clone();
        let inside = self
            .me()
            .and_then(|m| {
                m.as_any()
                    .downcast_ref::<ClassInstance>()
                    .map(|i| Rc::ptr_eq(&i.class, &class))
            })
            .unwrap_or(false);
        let visible = |p: &Rc<Procedure>| inside || p.visibility == Visibility::Public;
        let set = class
            .members
            .get(&member.to_ascii_lowercase())
            .cloned()
            .unwrap_or_default();
        let proc = match &kind {
            Invoke::Let(_) => set.let_.clone(),
            Invoke::Set(_) => set.set.clone(),
            Invoke::Get | Invoke::Method => set.call.clone().or_else(|| set.get.clone()),
        };
        if let Some(p) = proc.filter(visible) {
            if let Invoke::Let(v) | Invoke::Set(v) = kind {
                positional.push(ArgSource::Value(v));
            }
            return self.call_procedure(
                p,
                class.info.clone(),
                Some(obj.clone()),
                positional,
                named,
            );
        }
        let field = inst
            .field_slot(member)
            .filter(|(_, vis)| inside || *vis == Visibility::Public)
            .map(|(s, _)| s.clone())
            .ok_or_else(|| VbaError::not_supported(member))?;
        match kind {
            Invoke::Let(v) => {
                let mut var = field.borrow_mut();
                var.value = var.ty.coerce(v)?;
                Ok(Value::Empty)
            }
            Invoke::Set(v) => {
                field.borrow_mut().value = v;
                Ok(Value::Empty)
            }
            kind => {
                let v = field.borrow().value.clone();
                let args = Args::new(positional.into_iter().map(ArgSource::into_value).collect());
                self.index_value_evaluated(v, args, kind)
            }
        }
    }

    fn with_target(&self) -> VbaResult<&WithTarget> {
        self.frame()
            .and_then(|f| f.with_stack.last())
            .ok_or_else(|| VbaError::new(-1, "Invalid or unqualified reference"))
    }

    /// `.Name` inside `With` on a user-defined type, as an ordinary member expression.
    fn with_place(&self, name: &str) -> Option<Expr> {
        match self.with_target() {
            Ok(WithTarget::Place(e)) => Some(Expr::Member(Box::new(e.clone()), name.to_string())),
            _ => None,
        }
    }

    fn with_object(&self) -> VbaResult<ObjectRef> {
        match self.with_target()? {
            WithTarget::Value(v) => v.as_object(),
            WithTarget::Place(_) => Err(VbaError::object_required()),
        }
    }

    /// Assigns through a variable, array element or user-defined-type field.
    fn assign_place(
        &mut self,
        target: &Expr,
        f: &mut dyn FnMut(&mut Value, VarType) -> VbaResult<()>,
    ) -> VbaResult<()> {
        match target {
            Expr::Ident(name) => {
                let s = match self.find_slot(name) {
                    Some(s) => s,
                    None => self.declare_implicit(name, Value::Empty)?,
                };
                let mut var = s.borrow_mut();
                if var.is_const {
                    return Err(VbaError::new(-1, "Assignment to constant not permitted"));
                }
                let ty = var.ty;
                f(&mut var.value, ty)
            }
            Expr::Index(callee, args) => {
                let idx = indices(&self.eval_args(args)?)?;
                self.assign_place(callee, &mut |v, _| match v {
                    Value::Array(a) => {
                        let off = a.offset(&idx)?;
                        let ty = a.elem_type;
                        f(&mut a.data[off], ty)
                    }
                    _ => Err(VbaError::type_mismatch()),
                })
            }
            Expr::Member(base, field) => self.assign_place(base, &mut |v, _| match v {
                Value::Udt(u) => {
                    let fl = u.field_mut(field)?;
                    f(&mut fl.value, fl.ty)
                }
                _ => Err(VbaError::object_required()),
            }),
            Expr::WithMember(name) => match self.with_place(name) {
                Some(e) => self.assign_place(&e, f),
                None => Err(VbaError::object_required()),
            },
            _ => Err(VbaError::new(-1, "Invalid assignment target")),
        }
    }

    fn assign_udt_member(&mut self, target: &Expr, v: Value, is_set: bool) -> VbaResult<()> {
        let v = if is_set { v } else { v.into_scalar()? };
        self.assign_place(target, &mut |slot, ty| {
            *slot = if is_set || matches!(v, Value::Udt(_)) {
                v.clone()
            } else {
                ty.coerce(v.clone())?
            };
            Ok(())
        })
    }

    fn assign(&mut self, target: &Expr, v: Value, is_set: bool) -> VbaResult<()> {
        let kind = |v: Value| {
            if is_set {
                Invoke::Set(v)
            } else {
                Invoke::Let(v)
            }
        };
        match target {
            Expr::Ident(name) => {
                let s = match self.find_slot(name) {
                    Some(s) => s,
                    None => {
                        // `Balance = 5` inside a class: its own Property Let.
                        if let Some(me) = self.me() {
                            let inst = me
                                .as_any()
                                .downcast_ref::<ClassInstance>()
                                .expect("class instance");
                            if inst.class.members.contains_key(&name.to_ascii_lowercase()) {
                                return self
                                    .call_member(&me, name, kind(v), Vec::new(), Vec::new())
                                    .map(|_| ());
                            }
                        }
                        // `ActiveCell = 5`: a property of the host's global object.
                        if let Some(app) = self.host_global_object() {
                            match app.invoke(name, kind(v.clone()), Args::default()) {
                                Err(e) if e.number == 438 => {}
                                other => return other.map(|_| ()),
                            }
                        }
                        self.declare_implicit(name, Value::Empty)?
                    }
                };
                let mut var = s.borrow_mut();
                if var.is_const {
                    return Err(VbaError::new(-1, "Assignment to constant not permitted"));
                }
                if is_set {
                    if !matches!(var.ty, VarType::Variant | VarType::Object) {
                        return Err(VbaError::type_mismatch());
                    }
                    var.value = v;
                } else if let Value::Object(Some(o)) = &var.value {
                    // Let on an object variable assigns its default member.
                    let o = o.clone();
                    drop(var);
                    self.invoke_object(&o, "", Invoke::Let(v.into_scalar()?), Args::default())?;
                } else if matches!(v, Value::Udt(_)) {
                    var.value = v;
                } else {
                    var.value = var.ty.coerce(v.into_scalar()?)?;
                }
                Ok(())
            }
            Expr::Member(obj, name) => {
                let base = self.eval(obj)?;
                if let Value::Udt(_) = base {
                    return self.assign_udt_member(target, v, is_set);
                }
                let o = base.as_object()?;
                let v = if is_set { v } else { v.into_scalar()? };
                self.invoke_object(&o, name, kind(v), Args::default())
                    .map(|_| ())
            }
            Expr::WithMember(name) => {
                if let Some(place) = self.with_place(name) {
                    return self.assign_udt_member(&place, v, is_set);
                }
                let o = self.with_object()?;
                let v = if is_set { v } else { v.into_scalar()? };
                self.invoke_object(&o, name, kind(v), Args::default())
                    .map(|_| ())
            }
            Expr::Index(callee, raw_args) => {
                let place_base = match &**callee {
                    Expr::Member(obj, _) => matches!(self.eval(obj)?, Value::Udt(_)),
                    Expr::WithMember(n) => self.with_place(n).is_some(),
                    _ => false,
                };
                if place_base {
                    return self.assign_udt_member(target, v, is_set);
                }
                let args = self.eval_args(raw_args)?;
                let v = if is_set { v } else { v.into_scalar()? };
                match &**callee {
                    Expr::Ident(name) => {
                        if let Some(s) = self.find_slot(name) {
                            let mut var = s.borrow_mut();
                            return match &mut var.value {
                                Value::Array(a) => {
                                    let idx = indices(&args)?;
                                    let off = a.offset(&idx)?;
                                    a.data[off] = if matches!(v, Value::Udt(_)) {
                                        v
                                    } else {
                                        a.elem_type.coerce(v)?
                                    };
                                    Ok(())
                                }
                                Value::Object(Some(o)) => {
                                    let o = o.clone();
                                    drop(var);
                                    self.invoke_object(&o, "", kind(v), args).map(|_| ())
                                }
                                _ => Err(VbaError::type_mismatch()),
                            };
                        }
                        let app = self
                            .host_global_object()
                            .ok_or_else(|| VbaError::sub_not_defined(name))?;
                        app.invoke(name, kind(v), args).map(|_| ())
                    }
                    Expr::Member(obj, name) => {
                        let o = self.eval(obj)?.as_object()?;
                        self.invoke_object(&o, name, kind(v), args).map(|_| ())
                    }
                    Expr::WithMember(name) => {
                        let o = self.with_object()?;
                        self.invoke_object(&o, name, kind(v), args).map(|_| ())
                    }
                    other => {
                        let o = self.eval(other)?.as_object()?;
                        self.invoke_object(&o, "", kind(v), args).map(|_| ())
                    }
                }
            }
            _ => Err(VbaError::new(-1, "Invalid assignment target")),
        }
    }

    fn host_global_object(&self) -> Option<ObjectRef> {
        self.host.global_object()
    }

    // ---- expressions --------------------------------------------------------

    pub fn eval(&mut self, e: &Expr) -> VbaResult<Value> {
        Ok(match e {
            Expr::Int(v) => match *v {
                v if (i16::MIN as i64..=i16::MAX as i64).contains(&v) => Value::Integer(v as i16),
                v if (i32::MIN as i64..=i32::MAX as i64).contains(&v) => Value::Long(v as i32),
                v => Value::Double(v as f64),
            },
            Expr::Float(v) => Value::Double(*v),
            Expr::Str(s) => Value::String(s.clone()),
            Expr::Date(s) => {
                Value::Date(crate::date::parse_date(s).ok_or_else(VbaError::type_mismatch)?)
            }
            Expr::Bool(b) => Value::Boolean(*b),
            Expr::Nothing => Value::Object(None),
            Expr::Null => Value::Null,
            Expr::Empty => Value::Empty,
            Expr::New(class) => self.new_object(class)?,
            Expr::TypeOf(obj, class) => match self.eval(obj)? {
                Value::Object(Some(o)) => Value::Boolean(
                    class.eq_ignore_ascii_case("Object")
                        || o.type_name().eq_ignore_ascii_case(class)
                        || class
                            .rsplit('.')
                            .next()
                            .is_some_and(|c| o.type_name().eq_ignore_ascii_case(c)),
                ),
                Value::Object(None) => Value::Boolean(false),
                _ => return Err(VbaError::object_required()),
            },
            Expr::Ident(_) | Expr::Member(..) | Expr::WithMember(_) => {
                self.eval_call(e, &[], Invoke::Get)?
            }
            Expr::Index(callee, args) => self.eval_call(callee, args, Invoke::Get)?,
            Expr::Unary(UnOp::Neg, a) => value::neg(&self.eval(a)?.into_scalar()?)?,
            Expr::Unary(UnOp::Not, a) => value::not(&self.eval(a)?.into_scalar()?)?,
            Expr::Binary(op, a, b) => {
                let (a, b) = (self.eval(a)?, self.eval(b)?);
                self.binary(*op, a, b)?
            }
        })
    }

    fn binary(&mut self, op: BinOp, a: Value, b: Value) -> VbaResult<Value> {
        if op == BinOp::Is {
            return Ok(Value::Boolean(match (&a, &b) {
                (Value::Object(None), Value::Object(None)) => true,
                (Value::Object(Some(x)), Value::Object(Some(y))) => same_object(x, y),
                (Value::Object(_), Value::Object(_)) => false,
                _ => return Err(VbaError::object_required()),
            }));
        }
        let (a, b) = (a.into_scalar()?, b.into_scalar()?);
        let text = self.text_compare();
        Ok(match op {
            BinOp::Add => value::add(&a, &b)?,
            BinOp::Sub => value::sub(&a, &b)?,
            BinOp::Mul => value::mul(&a, &b)?,
            BinOp::Div => value::div(&a, &b)?,
            BinOp::IntDiv => value::int_div(&a, &b)?,
            BinOp::Mod => value::modulo(&a, &b)?,
            BinOp::Pow => value::pow(&a, &b)?,
            BinOp::Concat => value::concat(&a, &b)?,
            BinOp::Eq | BinOp::Ne | BinOp::Lt | BinOp::Le | BinOp::Gt | BinOp::Ge => {
                match value::compare(&a, &b, text)? {
                    None => Value::Null,
                    ord => Value::Boolean(compare_op(op, ord)),
                }
            }
            BinOp::Like => {
                if matches!(a, Value::Null) || matches!(b, Value::Null) {
                    Value::Null
                } else {
                    Value::Boolean(value::like(
                        &a.to_string_value()?,
                        &b.to_string_value()?,
                        text,
                    ))
                }
            }
            BinOp::And => value::logical(|x, y| x & y, &a, &b)?,
            BinOp::Or => value::logical(|x, y| x | y, &a, &b)?,
            BinOp::Xor => value::logical(|x, y| x ^ y, &a, &b)?,
            BinOp::Eqv => value::logical(|x, y| !(x ^ y), &a, &b)?,
            BinOp::Imp => value::logical(|x, y| !x | y, &a, &b)?,
            BinOp::Is => unreachable!(),
        })
    }

    fn eval_args(&mut self, args: &[Arg]) -> VbaResult<Args> {
        let mut out = Args::default();
        for a in args {
            let v = match &a.value {
                Some(e) => self.eval(e)?,
                None => Value::Missing,
            };
            match &a.name {
                Some(n) => out.named.push((n.clone(), v)),
                None => out.positional.push(v),
            }
        }
        Ok(out)
    }

    fn arg_sources(&mut self, args: &[Arg]) -> VbaResult<ArgSources> {
        let mut positional = Vec::new();
        let mut named = Vec::new();
        for a in args {
            let src = match &a.value {
                Some(Expr::Ident(n)) => match self.find_slot(n) {
                    Some(s) => ArgSource::Slot(s),
                    None => ArgSource::Value(self.eval(a.value.as_ref().unwrap())?),
                },
                Some(e) => ArgSource::Value(self.eval(e)?),
                None => ArgSource::Value(Value::Missing),
            };
            match &a.name {
                Some(n) => named.push((n.clone(), src)),
                None => positional.push(src),
            }
        }
        Ok((positional, named))
    }

    /// Evaluates `callee(args)`: array indexing, procedure call, built-in,
    /// object member, or host global — resolved at runtime like VBA does.
    fn eval_call(&mut self, callee: &Expr, args: &[Arg], kind: Invoke) -> VbaResult<Value> {
        match callee {
            Expr::Ident(name) => {
                let current_fn = self
                    .frame()
                    .is_some_and(|f| f.procedure.name.eq_ignore_ascii_case(name));
                if let Some(s) = self.find_slot(name) {
                    // Inside a Function, `Fn(x)` with arguments is a recursive call.
                    if !(current_fn && !args.is_empty()) {
                        if args.is_empty() {
                            return Ok(s.borrow().value.clone());
                        }
                        let evaluated = self.eval_args(args)?;
                        let var = s.borrow();
                        return match &var.value {
                            Value::Array(a) => Ok(a.data[a.offset(&indices(&evaluated)?)?].clone()),
                            Value::Object(Some(o)) => {
                                let o = o.clone();
                                drop(var);
                                self.invoke_object(&o, "", kind, evaluated)
                            }
                            _ => Err(VbaError::type_mismatch()),
                        };
                    }
                }
                if name.eq_ignore_ascii_case("me") {
                    if let Some(me) = self.me() {
                        return self.index_value(Value::Object(Some(me)), args, kind);
                    }
                }
                if let Some(me) = self.me() {
                    let inst = me
                        .as_any()
                        .downcast_ref::<ClassInstance>()
                        .expect("class instance");
                    if inst.class.members.contains_key(&name.to_ascii_lowercase()) {
                        let (positional, named) = self.arg_sources(args)?;
                        return self.call_member(&me, name, kind, positional, named);
                    }
                }
                if let Some((proc, module)) = self.procs.get(&name.to_ascii_lowercase()).cloned() {
                    let (positional, named) = self.arg_sources(args)?;
                    return self.call_procedure(proc, module, None, positional, named);
                }
                if name.eq_ignore_ascii_case("err") {
                    return self.index_value(Value::from(self.err.clone()), args, kind);
                }
                let evaluated = self.eval_args(args)?;
                let text = self.text_compare();
                if let Some(r) = builtins::call(name, &evaluated, self.host.as_mut(), text) {
                    return r;
                }
                if evaluated.is_empty() {
                    if let Some(c) = builtins::constant(name) {
                        return Ok(c);
                    }
                }
                if let Some(v) = self.host.global(name) {
                    return self.index_value_evaluated(v, evaluated, kind);
                }
                if let Some(app) = self.host_global_object() {
                    match app.invoke(name, kind.clone(), evaluated.clone()) {
                        Err(e) if e.number == 438 => {}
                        other => return other,
                    }
                }
                if self.explicit()
                    || self.inspecting
                    || !evaluated.is_empty()
                    || matches!(kind, Invoke::Method)
                {
                    if evaluated.is_empty() && !matches!(kind, Invoke::Method) {
                        return Err(VbaError::variable_not_defined(name));
                    }
                    return Err(VbaError::sub_not_defined(name));
                }
                // Implicitly declared variable, read before assignment.
                self.declare_implicit(name, Value::Empty)?;
                Ok(Value::Empty)
            }
            Expr::Member(obj, member) => {
                if let Expr::Ident(q) = &**obj {
                    if q.eq_ignore_ascii_case("vba") && self.find_slot(q).is_none() {
                        return self.eval_call(&Expr::Ident(member.clone()), args, kind);
                    }
                    // `Status.Open`: a qualified enum member.
                    if let Some(members) = self.enums.get(&q.to_ascii_lowercase()) {
                        if self.find_slot(q).is_none() {
                            return members
                                .iter()
                                .find(|(n, _)| n.eq_ignore_ascii_case(member))
                                .map(|(_, v)| v.clone())
                                .ok_or_else(|| {
                                    VbaError::new(
                                        -1,
                                        format!("Method or data member not found: {member}"),
                                    )
                                });
                        }
                    }
                }
                match self.eval(obj)? {
                    Value::Udt(u) => {
                        let v = u.field(member)?.value.clone();
                        self.index_value(v, args, kind)
                    }
                    base => {
                        let o = base.as_object()?;
                        self.call_object_member(&o, member, args, kind)
                    }
                }
            }
            Expr::WithMember(member) => {
                if let Some(place) = self.with_place(member) {
                    return self.eval_call(&place, args, kind);
                }
                let o = self.with_object()?;
                self.call_object_member(&o, member, args, kind)
            }
            other => {
                let v = self.eval(other)?;
                self.index_value(v, args, kind)
            }
        }
    }

    fn call_object_member(
        &mut self,
        o: &ObjectRef,
        member: &str,
        args: &[Arg],
        kind: Invoke,
    ) -> VbaResult<Value> {
        if o.as_any().is::<ClassInstance>() {
            // VBA class methods get ByRef arguments.
            let (positional, named) = self.arg_sources(args)?;
            return self.call_member(o, member, kind, positional, named);
        }
        let evaluated = self.eval_args(args)?;
        o.invoke(member, kind, evaluated)
    }

    fn index_value(&mut self, v: Value, args: &[Arg], kind: Invoke) -> VbaResult<Value> {
        if args.is_empty() {
            return Ok(v);
        }
        let evaluated = self.eval_args(args)?;
        self.index_value_evaluated(v, evaluated, kind)
    }

    fn index_value_evaluated(&mut self, v: Value, args: Args, kind: Invoke) -> VbaResult<Value> {
        if args.is_empty() {
            return Ok(v);
        }
        match v {
            Value::Array(a) => Ok(a.data[a.offset(&indices(&args)?)?].clone()),
            Value::Object(Some(o)) => self.invoke_object(&o, "", kind, args),
            Value::Object(None) => Err(VbaError::object_not_set()),
            _ => Err(VbaError::type_mismatch()),
        }
    }
}

/// Positional and named argument sources of a call.
type ArgSources = (Vec<ArgSource>, Vec<(String, ArgSource)>);

enum ArgSource {
    Slot(Slot),
    Value(Value),
}

impl ArgSource {
    fn into_value(self) -> Value {
        match self {
            ArgSource::Slot(s) => s.borrow().value.clone(),
            ArgSource::Value(v) => v,
        }
    }
}

fn truthy(v: &Value) -> VbaResult<bool> {
    match v {
        Value::Null => Ok(false),
        v => v.to_bool(),
    }
}

fn compare_op(op: BinOp, ord: Option<Ordering>) -> bool {
    let Some(ord) = ord else { return false };
    match op {
        BinOp::Eq => ord == Ordering::Equal,
        BinOp::Ne => ord != Ordering::Equal,
        BinOp::Lt => ord == Ordering::Less,
        BinOp::Le => ord != Ordering::Greater,
        BinOp::Gt => ord == Ordering::Greater,
        BinOp::Ge => ord != Ordering::Less,
        _ => false,
    }
}

fn indices(args: &Args) -> VbaResult<Vec<i64>> {
    args.positional.iter().map(|v| v.to_i64()).collect()
}

fn find_label(stmts: &[Stmt], label: &str) -> Option<usize> {
    stmts
        .iter()
        .position(|s| matches!(&s.kind, StmtKind::Label(l) if l.eq_ignore_ascii_case(label)))
}

/// How `Print` renders a value.
fn print_text(v: Value) -> VbaResult<String> {
    let v = match v {
        Value::Null => return Ok("Null".into()),
        v => v.into_scalar()?,
    };
    Ok(match v {
        Value::Boolean(_) | Value::Date(_) | Value::Empty | Value::String(_) | Value::Error(_) => {
            v.to_string()
        }
        v if v.is_numeric_type() => {
            let s = v.to_string();
            if s.starts_with('-') {
                format!("{s} ")
            } else {
                format!(" {s} ")
            }
        }
        v => v.to_string(),
    })
}

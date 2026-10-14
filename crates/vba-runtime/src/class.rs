//! Class modules: definitions and instances.

use std::any::Any;
use std::collections::HashMap;
use std::rc::Rc;

use vba_syntax::ast::{Expr, ProcKind, Procedure, VarDecl, Visibility};

use crate::error::{VbaError, VbaResult};
use crate::interp::{ModuleInfo, Slot};
use crate::object::{Args, Invoke, Object};
use crate::value::Value;

/// The procedures sharing one member name: a `Sub`/`Function`, or the
/// `Property Get` / `Let` / `Set` group.
#[derive(Default, Clone)]
pub struct ProcSet {
    pub call: Option<Rc<Procedure>>,
    pub get: Option<Rc<Procedure>>,
    pub let_: Option<Rc<Procedure>>,
    pub set: Option<Rc<Procedure>>,
}

impl ProcSet {
    pub fn insert(&mut self, p: Rc<Procedure>) {
        match p.kind {
            ProcKind::Sub | ProcKind::Function => self.call = Some(p),
            ProcKind::PropertyGet => self.get = Some(p),
            ProcKind::PropertyLet => self.let_ = Some(p),
            ProcKind::PropertySet => self.set = Some(p),
        }
    }
}

pub struct ClassDef {
    pub name: String,
    pub(crate) info: Rc<ModuleInfo>,
    pub members: HashMap<String, ProcSet>,
    pub fields: Vec<(Visibility, VarDecl)>,
    pub consts: Vec<(String, Expr)>,
}

/// An instance created with `New`. Its fields are slots so methods can
/// update them through `Me`.
pub struct ClassInstance {
    pub class: Rc<ClassDef>,
    pub fields: Vec<(String, Slot, Visibility)>,
}

impl ClassInstance {
    pub fn field_slot(&self, name: &str) -> Option<(&Slot, Visibility)> {
        self.fields
            .iter()
            .find(|(n, _, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, s, v)| (s, *v))
    }
}

impl Object for ClassInstance {
    fn type_name(&self) -> &str {
        &self.class.name
    }

    /// Methods of VBA classes run in the interpreter, which dispatches them
    /// before reaching this; a host calling in directly can only read
    /// public fields.
    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        match (self.field_slot(member), kind) {
            (Some((s, Visibility::Public)), Invoke::Get | Invoke::Method) if args.is_empty() => {
                Ok(s.borrow().value.clone())
            }
            _ => Err(VbaError::not_supported(member)),
        }
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        self.fields
            .iter()
            .map(|(n, s, _)| (n.clone(), s.borrow().value.clone()))
            .collect()
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

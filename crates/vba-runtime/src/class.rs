//! Class modules: definitions and instances.

use std::any::Any;
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::{Rc, Weak};

use vba_syntax::ast::{Expr, ProcKind, Procedure, VarDecl, Visibility};

use crate::error::{VbaError, VbaResult};
use crate::interp::{ModuleInfo, Slot};
use crate::object::{Args, Invoke, Object, ObjectRef};
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
    pub fn is_empty(&self) -> bool {
        self.call.is_none() && self.get.is_none() && self.let_.is_none() && self.set.is_none()
    }

    /// The same set without its `Private` procedures.
    pub fn public_only(&self) -> ProcSet {
        let public =
            |p: &Option<Rc<Procedure>>| p.clone().filter(|p| p.visibility == Visibility::Public);
        ProcSet {
            call: public(&self.call),
            get: public(&self.get),
            let_: public(&self.let_),
            set: public(&self.set),
        }
    }

    /// The procedure handling an access: `Let`/`Set` go to the property
    /// procedures, reads and calls to the `Sub`/`Function` or `Property Get`.
    pub fn for_kind(&self, kind: &Invoke) -> Option<Rc<Procedure>> {
        match kind {
            Invoke::Let(_) => self.let_.clone(),
            Invoke::Set(_) => self.set.clone(),
            Invoke::Get | Invoke::Method => self.call.clone().or_else(|| self.get.clone()),
        }
    }

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
    /// Interfaces from `Implements`; `IShape_Area` implements `IShape.Area`.
    pub implements: Vec<String>,
    pub events: Vec<String>,
    /// The member marked `VB_UserMemId = 0`: `obj` / `obj(i)` mean `obj.Item(i)`.
    pub default_member: Option<String>,
    /// The member marked `VB_UserMemId = -4`, used by `For Each`.
    pub enum_member: Option<String>,
}

impl ClassDef {
    pub fn implements(&self, interface: &str) -> bool {
        self.implements.iter().any(|i| {
            i.eq_ignore_ascii_case(interface)
                || i.rsplit('.')
                    .next()
                    .is_some_and(|n| n.eq_ignore_ascii_case(interface))
        })
    }

    fn has_terminate(&self) -> bool {
        self.members
            .get("class_terminate")
            .is_some_and(|m| m.call.is_some())
    }
}

/// Instances whose last reference went away and whose `Class_Terminate`
/// is still to run; the interpreter drains it between statements.
pub(crate) type TerminateQueue = Rc<RefCell<Vec<ObjectRef>>>;

/// An instance created with `New`. Its fields are slots so methods can
/// update them through `Me`.
pub struct ClassInstance {
    pub class: Rc<ClassDef>,
    pub fields: Vec<(String, Slot, Visibility)>,
    /// Objects holding this one in a `WithEvents` field: `(sink, field name)`.
    pub(crate) listeners: RefCell<Vec<(Weak<dyn Object>, String)>>,
    terminate: Option<TerminateQueue>,
}

impl ClassInstance {
    pub(crate) fn new(
        class: Rc<ClassDef>,
        fields: Vec<(String, Slot, Visibility)>,
        queue: &TerminateQueue,
    ) -> Self {
        let terminate = class.has_terminate().then(|| queue.clone());
        ClassInstance {
            class,
            fields,
            listeners: RefCell::default(),
            terminate,
        }
    }

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

impl Drop for ClassInstance {
    /// The last reference is gone: hand the fields to a stand-in `Me` for
    /// `Class_Terminate`, which the interpreter runs at its next safe point.
    fn drop(&mut self) {
        if let Some(queue) = self.terminate.take() {
            let me = ClassInstance {
                class: self.class.clone(),
                fields: std::mem::take(&mut self.fields),
                listeners: RefCell::new(std::mem::take(self.listeners.get_mut())),
                terminate: None,
            };
            queue.borrow_mut().push(Rc::new(me));
        }
    }
}

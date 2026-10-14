//! Object model: the `IDispatch`-like interface host applications implement
//! to expose `Application`, `Workbook`, `Range`, ... to VBA code.

use std::any::Any;
use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use crate::error::{VbaError, VbaResult};
use crate::value::Value;

pub type ObjectRef = Rc<dyn Object>;

/// How a member is being accessed, mirroring `IDispatch::Invoke` flags.
#[derive(Debug, Clone)]
pub enum Invoke {
    /// Property get or function call in an expression (`x = obj.Name`).
    Get,
    /// Method called as a statement (`obj.Activate`); the result is discarded.
    Method,
    /// Property let (`obj.Value = 5`).
    Let(Value),
    /// Property set with an object (`Set obj.Font = f`).
    Set(Value),
}

/// Evaluated call arguments.
#[derive(Debug, Clone, Default)]
pub struct Args {
    pub positional: Vec<Value>,
    pub named: Vec<(String, Value)>,
}

impl Args {
    pub fn new(positional: Vec<Value>) -> Self {
        Args {
            positional,
            named: Vec::new(),
        }
    }

    pub fn is_empty(&self) -> bool {
        self.positional.is_empty() && self.named.is_empty()
    }

    pub fn len(&self) -> usize {
        self.positional.len() + self.named.len()
    }

    /// Argument by position or by name (case-insensitive); `Missing` counts as absent.
    pub fn get(&self, index: usize, name: &str) -> Option<&Value> {
        self.named
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v)| v)
            .or_else(|| self.positional.get(index))
            .filter(|v| !matches!(v, Value::Missing))
    }

    pub fn require(&self, index: usize, name: &str) -> VbaResult<&Value> {
        self.get(index, name).ok_or_else(VbaError::wrong_arg_count)
    }
}

/// A scriptable object. Member names are compared case-insensitively by
/// convention; the empty name `""` is the default member (`DISPID_VALUE`).
pub trait Object {
    fn type_name(&self) -> &str;

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value>;

    /// Value of the default member, used when an object appears where a
    /// scalar is expected (`MsgBox Range("A1")`).
    fn default_value(&self) -> VbaResult<Value> {
        self.invoke("", Invoke::Get, Args::default())
    }

    /// Items for `For Each` (`_NewEnum`).
    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        Err(VbaError::not_supported("_NewEnum"))
    }

    /// Property names shown in the debugger's Locals/Watch windows.
    fn inspect(&self) -> Vec<(String, Value)> {
        Vec::new()
    }

    fn as_any(&self) -> &dyn Any;
}

pub fn same_object(a: &ObjectRef, b: &ObjectRef) -> bool {
    std::ptr::eq(Rc::as_ptr(a) as *const (), Rc::as_ptr(b) as *const ())
}

// ---- Collection -------------------------------------------------------------

/// VBA's built-in `Collection` (1-based, optional string keys).
#[derive(Default)]
pub struct Collection {
    items: RefCell<Vec<(Option<String>, Value)>>,
}

impl Collection {
    pub fn new() -> Rc<Self> {
        Rc::new(Collection::default())
    }

    fn position(&self, key: &Value) -> VbaResult<usize> {
        let items = self.items.borrow();
        match key {
            Value::String(k) => items
                .iter()
                .position(|(ik, _)| ik.as_deref().is_some_and(|ik| ik.eq_ignore_ascii_case(k)))
                .ok_or_else(VbaError::invalid_call),
            other => {
                let i = other.to_i64()?;
                if i >= 1 && (i as usize) <= items.len() {
                    Ok(i as usize - 1)
                } else {
                    Err(VbaError::subscript_out_of_range())
                }
            }
        }
    }
}

impl Object for Collection {
    fn type_name(&self) -> &str {
        "Collection"
    }

    fn invoke(&self, member: &str, _kind: Invoke, args: Args) -> VbaResult<Value> {
        match member.to_ascii_lowercase().as_str() {
            "count" => Ok(Value::Long(self.items.borrow().len() as i32)),
            "" | "item" => {
                let i = self.position(args.require(0, "Index")?)?;
                Ok(self.items.borrow()[i].1.clone())
            }
            "add" => {
                let item = args.require(0, "Item")?.clone();
                let key = args
                    .get(1, "Key")
                    .map(|k| k.to_string_value())
                    .transpose()?;
                if let Some(k) = &key {
                    if self.position(&Value::str(k.clone())).is_ok() {
                        return Err(VbaError::new(
                            457,
                            "This key is already associated with an element of this collection",
                        ));
                    }
                }
                let at = match (args.get(2, "Before"), args.get(3, "After")) {
                    (Some(b), _) => self.position(b)?,
                    (None, Some(a)) => self.position(a)? + 1,
                    _ => self.items.borrow().len(),
                };
                self.items.borrow_mut().insert(at, (key, item));
                Ok(Value::Empty)
            }
            "remove" => {
                let i = self.position(args.require(0, "Index")?)?;
                self.items.borrow_mut().remove(i);
                Ok(Value::Empty)
            }
            // `mItems.[_NewEnum]` in a custom collection class.
            "_newenum" => Ok(Value::Object(Some(Rc::new(Enumerator(self.enumerate()?))))),
            _ => Err(VbaError::not_supported(member)),
        }
    }

    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        Ok(self.items.borrow().iter().map(|(_, v)| v.clone()).collect())
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        self.items
            .borrow()
            .iter()
            .enumerate()
            .map(|(i, (_, v))| (format!("Item {}", i + 1), v.clone()))
            .collect()
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

/// A snapshot of items returned by `[_NewEnum]`, consumed by `For Each`.
pub struct Enumerator(pub Vec<Value>);

impl Object for Enumerator {
    fn type_name(&self) -> &str {
        "Unknown"
    }

    fn invoke(&self, member: &str, _kind: Invoke, _args: Args) -> VbaResult<Value> {
        Err(VbaError::not_supported(member))
    }

    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        Ok(self.0.clone())
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

// ---- DynObject --------------------------------------------------------------

type MethodFn = Box<dyn Fn(&DynObject, Args) -> VbaResult<Value>>;

/// A convenience object for hosts and tests: a bag of properties plus
/// closure-backed methods.
///
/// ```
/// use vba_runtime::object::{DynObject, Object, Invoke, Args};
/// use vba_runtime::Value;
/// let sheet = DynObject::new("Worksheet").with_prop("Name", "Sheet1");
/// let name = sheet.invoke("name", Invoke::Get, Args::default()).unwrap();
/// assert_eq!(name.to_string(), "Sheet1");
/// ```
pub struct DynObject {
    type_name: String,
    props: RefCell<Vec<(String, Value)>>,
    methods: HashMap<String, MethodFn>,
    default_member: Option<String>,
}

impl DynObject {
    pub fn new(type_name: &str) -> Self {
        DynObject {
            type_name: type_name.into(),
            props: RefCell::default(),
            methods: HashMap::new(),
            default_member: None,
        }
    }

    pub fn with_prop(self, name: &str, value: impl Into<Value>) -> Self {
        self.props
            .borrow_mut()
            .push((name.to_string(), value.into()));
        self
    }

    pub fn with_method(
        mut self,
        name: &str,
        f: impl Fn(&DynObject, Args) -> VbaResult<Value> + 'static,
    ) -> Self {
        self.methods.insert(name.to_ascii_lowercase(), Box::new(f));
        self
    }

    pub fn with_default_member(mut self, name: &str) -> Self {
        self.default_member = Some(name.to_string());
        self
    }

    pub fn into_ref(self) -> ObjectRef {
        Rc::new(self)
    }

    pub fn get_prop(&self, name: &str) -> Option<Value> {
        self.props
            .borrow()
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v)| v.clone())
    }

    pub fn set_prop(&self, name: &str, value: Value) {
        let mut props = self.props.borrow_mut();
        match props.iter_mut().find(|(n, _)| n.eq_ignore_ascii_case(name)) {
            Some(slot) => slot.1 = value,
            None => props.push((name.to_string(), value)),
        }
    }
}

impl Object for DynObject {
    fn type_name(&self) -> &str {
        &self.type_name
    }

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        let member = if member.is_empty() {
            self.default_member
                .clone()
                .ok_or_else(|| VbaError::not_supported("<default member>"))?
        } else {
            member.to_string()
        };
        if let Some(m) = self.methods.get(&member.to_ascii_lowercase()) {
            return m(self, args);
        }
        match kind {
            Invoke::Let(v) | Invoke::Set(v) if self.get_prop(&member).is_some() => {
                self.set_prop(&member, v);
                Ok(Value::Empty)
            }
            Invoke::Get | Invoke::Method if args.is_empty() => self
                .get_prop(&member)
                .ok_or_else(|| VbaError::not_supported(&member)),
            _ => Err(VbaError::not_supported(&member)),
        }
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        self.props.borrow().clone()
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

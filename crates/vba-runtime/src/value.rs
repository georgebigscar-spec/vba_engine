//! The VBA `Variant` and its coercion / operator rules.

use std::cmp::Ordering;
use std::fmt;
use std::rc::Rc;

use crate::error::{VbaError, VbaResult};
use crate::object::ObjectRef;

/// A VBA value. Every variable is conceptually a `Variant`; declared types are
/// enforced by coercing on assignment (see [`VarType::coerce`]).
#[derive(Clone, Default)]
pub enum Value {
    #[default]
    Empty,
    Null,
    Boolean(bool),
    Byte(u8),
    Integer(i16),
    Long(i32),
    LongLong(i64),
    Single(f32),
    Double(f64),
    /// Fixed-point, scaled by 10 000.
    Currency(i64),
    /// OLE Automation date: days since 1899-12-30, fraction = time of day.
    Date(f64),
    String(String),
    /// `None` is `Nothing`.
    Object(Option<ObjectRef>),
    Array(Box<VbArray>),
    /// `CVErr(n)`.
    Error(i32),
    /// An omitted optional argument (`IsMissing` returns True).
    Missing,
    /// A value of a user-defined type (`Type ... End Type`); copied on assignment.
    Udt(Box<UdtValue>),
}

/// A value of a user-defined type.
#[derive(Clone, Debug)]
pub struct UdtValue {
    pub type_name: String,
    pub fields: Vec<UdtField>,
}

#[derive(Clone, Debug)]
pub struct UdtField {
    pub name: String,
    pub ty: VarType,
    pub value: Value,
}

impl UdtValue {
    pub fn field(&self, name: &str) -> VbaResult<&UdtField> {
        self.fields
            .iter()
            .find(|f| f.name.eq_ignore_ascii_case(name))
            .ok_or_else(|| VbaError::new(-1, format!("Method or data member not found: {name}")))
    }

    pub fn field_mut(&mut self, name: &str) -> VbaResult<&mut UdtField> {
        self.fields
            .iter_mut()
            .find(|f| f.name.eq_ignore_ascii_case(name))
            .ok_or_else(|| VbaError::new(-1, format!("Method or data member not found: {name}")))
    }
}

/// The declared type of a variable (`Dim x As Long`).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VarType {
    Variant,
    Boolean,
    Byte,
    Integer,
    Long,
    LongLong,
    Single,
    Double,
    Currency,
    Date,
    String,
    /// `String * N`: always exactly N characters.
    FixedString(u32),
    Object,
}

impl VarType {
    /// Maps a type name as written in source; unknown names are object types
    /// (`Range`, `Worksheet`, `Collection`, ...).
    pub fn from_name(name: &str) -> VarType {
        match name.to_ascii_lowercase().as_str() {
            "variant" => VarType::Variant,
            "boolean" => VarType::Boolean,
            "byte" => VarType::Byte,
            "integer" => VarType::Integer,
            "long" => VarType::Long,
            "longlong" | "longptr" => VarType::LongLong,
            "single" => VarType::Single,
            "double" => VarType::Double,
            "currency" => VarType::Currency,
            "date" => VarType::Date,
            "string" => VarType::String,
            _ => VarType::Object,
        }
    }

    pub fn default_value(self) -> Value {
        match self {
            VarType::Variant => Value::Empty,
            VarType::Boolean => Value::Boolean(false),
            VarType::Byte => Value::Byte(0),
            VarType::Integer => Value::Integer(0),
            VarType::Long => Value::Long(0),
            VarType::LongLong => Value::LongLong(0),
            VarType::Single => Value::Single(0.0),
            VarType::Double => Value::Double(0.0),
            VarType::Currency => Value::Currency(0),
            VarType::Date => Value::Date(0.0),
            VarType::String => Value::String(String::new()),
            VarType::FixedString(n) => Value::String("\0".repeat(n as usize)),
            VarType::Object => Value::Object(None),
        }
    }

    /// Converts `v` for storage in a variable of this type (`Let` semantics).
    pub fn coerce(self, v: Value) -> VbaResult<Value> {
        Ok(match self {
            VarType::Variant => v,
            VarType::Boolean => Value::Boolean(v.to_bool()?),
            VarType::Byte => Value::Byte(
                u8::try_from(round_even_i64(v.to_f64()?)?).map_err(|_| VbaError::overflow())?,
            ),
            VarType::Integer => Value::Integer(
                i16::try_from(round_even_i64(v.to_f64()?)?).map_err(|_| VbaError::overflow())?,
            ),
            VarType::Long => Value::Long(
                i32::try_from(round_even_i64(v.to_f64()?)?).map_err(|_| VbaError::overflow())?,
            ),
            VarType::LongLong => Value::LongLong(v.to_i64()?),
            VarType::Single => Value::Single(v.to_f64()? as f32),
            VarType::Double => Value::Double(v.to_f64()?),
            VarType::Currency => Value::Currency((v.to_f64()? * 10_000.0).round_ties_even() as i64),
            VarType::Date => Value::Date(v.to_date()?),
            VarType::String => Value::String(v.to_string_value()?),
            VarType::FixedString(n) => {
                let mut s: String = v.to_string_value()?.chars().take(n as usize).collect();
                let len = s.chars().count();
                s.extend(std::iter::repeat_n(' ', n as usize - len));
                Value::String(s)
            }
            VarType::Object => match v {
                Value::Object(_) => v,
                _ => return Err(VbaError::type_mismatch()),
            },
        })
    }
}

/// A (possibly multi-dimensional) VBA array. Arrays have value semantics:
/// assigning one copies it.
#[derive(Clone, Debug)]
pub struct VbArray {
    /// `(lower, upper)` bounds per dimension.
    pub bounds: Vec<(i64, i64)>,
    pub elem_type: VarType,
    /// Row-major element storage.
    pub data: Vec<Value>,
}

impl VbArray {
    pub fn new(bounds: Vec<(i64, i64)>, elem_type: VarType) -> Self {
        let len = bounds
            .iter()
            .map(|(l, u)| (u - l + 1).max(0) as usize)
            .product::<usize>();
        let len = if bounds.is_empty() { 0 } else { len };
        VbArray {
            bounds,
            elem_type,
            data: vec![elem_type.default_value(); len],
        }
    }

    pub fn from_values(values: Vec<Value>, lower: i64) -> Self {
        let upper = lower + values.len() as i64 - 1;
        VbArray {
            bounds: vec![(lower, upper)],
            elem_type: VarType::Variant,
            data: values,
        }
    }

    pub fn offset(&self, idx: &[i64]) -> VbaResult<usize> {
        if idx.len() != self.bounds.len() {
            return Err(VbaError::subscript_out_of_range());
        }
        let mut off = 0usize;
        for (&i, &(l, u)) in idx.iter().zip(&self.bounds) {
            if i < l || i > u {
                return Err(VbaError::subscript_out_of_range());
            }
            off = off * (u - l + 1) as usize + (i - l) as usize;
        }
        Ok(off)
    }

    /// `ReDim Preserve`: only the last dimension may change, data is kept.
    pub fn redim_preserve(&mut self, bounds: Vec<(i64, i64)>) -> VbaResult<()> {
        if self.bounds.is_empty() {
            *self = VbArray::new(bounds, self.elem_type);
            return Ok(());
        }
        if bounds.len() != self.bounds.len()
            || bounds[..bounds.len() - 1] != self.bounds[..self.bounds.len() - 1]
        {
            return Err(VbaError::subscript_out_of_range());
        }
        let mut next = VbArray::new(bounds, self.elem_type);
        let old_last = *self.bounds.last().unwrap();
        let new_last = *next.bounds.last().unwrap();
        let outer = self.data.len() / (old_last.1 - old_last.0 + 1).max(1) as usize;
        let keep = (old_last.1.min(new_last.1) - old_last.0 + 1).max(0) as usize;
        let (old_w, new_w) = (
            (old_last.1 - old_last.0 + 1) as usize,
            (new_last.1 - new_last.0 + 1) as usize,
        );
        for o in 0..outer {
            for k in 0..keep {
                next.data[o * new_w + k] = self.data[o * old_w + k].clone();
            }
        }
        *self = next;
        Ok(())
    }
}

impl Value {
    pub fn str(s: impl Into<String>) -> Value {
        Value::String(s.into())
    }

    pub fn nothing() -> Value {
        Value::Object(None)
    }

    pub fn is_numeric_type(&self) -> bool {
        matches!(
            self,
            Value::Byte(_)
                | Value::Integer(_)
                | Value::Long(_)
                | Value::LongLong(_)
                | Value::Single(_)
                | Value::Double(_)
                | Value::Currency(_)
                | Value::Boolean(_)
                | Value::Date(_)
                | Value::Empty
        )
    }

    /// `TypeName(v)`.
    pub fn type_name(&self) -> String {
        match self {
            Value::Empty => "Empty".into(),
            Value::Null => "Null".into(),
            Value::Boolean(_) => "Boolean".into(),
            Value::Byte(_) => "Byte".into(),
            Value::Integer(_) => "Integer".into(),
            Value::Long(_) => "Long".into(),
            Value::LongLong(_) => "LongLong".into(),
            Value::Single(_) => "Single".into(),
            Value::Double(_) => "Double".into(),
            Value::Currency(_) => "Currency".into(),
            Value::Date(_) => "Date".into(),
            Value::String(_) => "String".into(),
            Value::Object(None) => "Nothing".into(),
            Value::Object(Some(o)) => o.type_name().to_string(),
            Value::Array(a) => format!("{}()", VarType::name(a.elem_type)),
            Value::Error(_) => "Error".into(),
            Value::Missing => "Error".into(),
            Value::Udt(u) => u.type_name.clone(),
        }
    }

    /// `VarType(v)` numeric code.
    pub fn var_type_code(&self) -> i32 {
        match self {
            Value::Empty => 0,
            Value::Null => 1,
            Value::Integer(_) => 2,
            Value::Long(_) => 3,
            Value::Single(_) => 4,
            Value::Double(_) => 5,
            Value::Currency(_) => 6,
            Value::Date(_) => 7,
            Value::String(_) => 8,
            Value::Object(_) => 9,
            Value::Error(_) | Value::Missing => 10,
            Value::Udt(_) => 36,
            Value::Boolean(_) => 11,
            Value::Byte(_) => 17,
            Value::LongLong(_) => 20,
            Value::Array(_) => 8192 + 12,
        }
    }

    // ---- conversions ---------------------------------------------------

    pub fn to_f64(&self) -> VbaResult<f64> {
        Ok(match self {
            Value::Empty => 0.0,
            Value::Boolean(b) => {
                if *b {
                    -1.0
                } else {
                    0.0
                }
            }
            Value::Byte(v) => *v as f64,
            Value::Integer(v) => *v as f64,
            Value::Long(v) => *v as f64,
            Value::LongLong(v) => *v as f64,
            Value::Single(v) => *v as f64,
            Value::Double(v) | Value::Date(v) => *v,
            Value::Currency(v) => *v as f64 / 10_000.0,
            Value::String(s) => parse_number(s).ok_or_else(VbaError::type_mismatch)?,
            Value::Null => return Err(VbaError::invalid_use_of_null()),
            Value::Object(Some(o)) => return o.default_value()?.to_f64(),
            _ => return Err(VbaError::type_mismatch()),
        })
    }

    /// Integral conversion with banker's rounding (`CLng` semantics).
    pub fn to_i64(&self) -> VbaResult<i64> {
        match self {
            Value::Byte(v) => Ok(*v as i64),
            Value::Integer(v) => Ok(*v as i64),
            Value::Long(v) => Ok(*v as i64),
            Value::LongLong(v) => Ok(*v),
            other => round_even_i64(other.to_f64()?),
        }
    }

    pub fn to_bool(&self) -> VbaResult<bool> {
        match self {
            Value::Boolean(b) => Ok(*b),
            Value::String(s) if s.eq_ignore_ascii_case("true") => Ok(true),
            Value::String(s) if s.eq_ignore_ascii_case("false") => Ok(false),
            other => Ok(other.to_f64()? != 0.0),
        }
    }

    pub fn to_date(&self) -> VbaResult<f64> {
        match self {
            Value::String(s) => crate::date::parse_date(s).ok_or_else(VbaError::type_mismatch),
            other => other.to_f64(),
        }
    }

    /// `CStr` semantics.
    pub fn to_string_value(&self) -> VbaResult<String> {
        Ok(match self {
            Value::Empty => String::new(),
            Value::String(s) => s.clone(),
            Value::Null => return Err(VbaError::invalid_use_of_null()),
            Value::Object(Some(o)) => return o.default_value()?.to_string_value(),
            Value::Object(None) | Value::Array(_) | Value::Missing | Value::Udt(_) => {
                return Err(VbaError::type_mismatch())
            }
            other => other.to_string(),
        })
    }

    /// Dereferences objects to their default member (for `Let` contexts).
    pub fn into_scalar(self) -> VbaResult<Value> {
        match self {
            Value::Object(Some(o)) => o.default_value(),
            v => Ok(v),
        }
    }

    pub fn as_object(&self) -> VbaResult<ObjectRef> {
        match self {
            Value::Object(Some(o)) => Ok(o.clone()),
            Value::Object(None) => Err(VbaError::object_not_set()),
            _ => Err(VbaError::object_required()),
        }
    }

    fn numeric_rank(&self) -> u8 {
        match self {
            Value::Empty | Value::Byte(_) => 0,
            Value::Boolean(_) | Value::Integer(_) => 1,
            Value::Long(_) => 2,
            Value::LongLong(_) => 3,
            Value::Single(_) => 4,
            Value::Double(_) | Value::String(_) => 5,
            Value::Currency(_) => 6,
            Value::Date(_) => 7,
            _ => 5,
        }
    }
}

impl VarType {
    fn name(self) -> &'static str {
        match self {
            VarType::Variant => "Variant",
            VarType::Boolean => "Boolean",
            VarType::Byte => "Byte",
            VarType::Integer => "Integer",
            VarType::Long => "Long",
            VarType::LongLong => "LongLong",
            VarType::Single => "Single",
            VarType::Double => "Double",
            VarType::Currency => "Currency",
            VarType::Date => "Date",
            VarType::String | VarType::FixedString(_) => "String",
            VarType::Object => "Object",
        }
    }
}

pub(crate) fn round_even_i64(f: f64) -> VbaResult<i64> {
    let r = f.round_ties_even();
    if r.is_finite() && r >= i64::MIN as f64 && r <= i64::MAX as f64 {
        Ok(r as i64)
    } else {
        Err(VbaError::overflow())
    }
}

/// Parses a numeric string the way `Val`-less coercion does (`" 12.5 "`, `"&HFF"`).
pub fn parse_number(s: &str) -> Option<f64> {
    let t = s.trim();
    if t.is_empty() {
        return None;
    }
    if let Some(hex) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        return i64::from_str_radix(hex, 16).ok().map(|v| v as f64);
    }
    t.parse::<f64>().ok().filter(|v| v.is_finite())
}

/// Formats a double with 15 significant digits like VBA's `CStr`.
pub fn format_double(v: f64) -> String {
    if v == 0.0 {
        return "0".into();
    }
    let rounded: f64 = format!("{v:.14e}").parse().unwrap_or(v);
    let abs = rounded.abs();
    if !(1e-15..1e15).contains(&abs) {
        let s = format!("{rounded:e}");
        let (m, e) = s.split_once('e').unwrap();
        let exp: i32 = e.parse().unwrap();
        return format!("{m}E{}{:02}", if exp < 0 { '-' } else { '+' }, exp.abs());
    }
    format!("{rounded}")
}

impl fmt::Display for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::Empty => Ok(()),
            Value::Null => write!(f, "Null"),
            Value::Boolean(b) => write!(f, "{}", if *b { "True" } else { "False" }),
            Value::Byte(v) => write!(f, "{v}"),
            Value::Integer(v) => write!(f, "{v}"),
            Value::Long(v) => write!(f, "{v}"),
            Value::LongLong(v) => write!(f, "{v}"),
            Value::Single(v) => write!(f, "{}", format_double(*v as f64)),
            Value::Double(v) => write!(f, "{}", format_double(*v)),
            Value::Currency(v) => write!(f, "{}", format_double(*v as f64 / 10_000.0)),
            Value::Date(d) => write!(f, "{}", crate::date::format_date(*d)),
            Value::String(s) => write!(f, "{s}"),
            Value::Object(None) => write!(f, "Nothing"),
            Value::Object(Some(o)) => write!(f, "<{}>", o.type_name()),
            Value::Array(a) => write!(f, "<Array {:?}>", a.bounds),
            Value::Error(n) => write!(f, "Error {n}"),
            Value::Missing => write!(f, "<Missing>"),
            Value::Udt(u) => write!(f, "<{}>", u.type_name),
        }
    }
}

impl fmt::Debug for Value {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Value::String(s) => write!(f, "\"{}\"", s.replace('"', "\"\"")),
            Value::Empty => write!(f, "Empty"),
            other => write!(f, "{other}"),
        }
    }
}

impl From<bool> for Value {
    fn from(v: bool) -> Self {
        Value::Boolean(v)
    }
}
impl From<i32> for Value {
    fn from(v: i32) -> Self {
        Value::Long(v)
    }
}
impl From<f64> for Value {
    fn from(v: f64) -> Self {
        Value::Double(v)
    }
}
impl From<&str> for Value {
    fn from(v: &str) -> Self {
        Value::String(v.to_string())
    }
}
impl From<String> for Value {
    fn from(v: String) -> Self {
        Value::String(v)
    }
}
impl From<ObjectRef> for Value {
    fn from(v: ObjectRef) -> Self {
        Value::Object(Some(v))
    }
}
impl<T: crate::object::Object + 'static> From<Rc<T>> for Value {
    fn from(v: Rc<T>) -> Self {
        Value::Object(Some(v))
    }
}

// ---- operators ---------------------------------------------------------

/// Builds an integral result of the given rank, widening on overflow the way
/// Variant arithmetic does (Integer → Long → LongLong → Double).
fn int_result(v: i128, rank: u8) -> Value {
    if rank <= 1 && (i16::MIN as i128..=i16::MAX as i128).contains(&v) {
        if rank == 0 && (0..=255).contains(&v) {
            return Value::Byte(v as u8);
        }
        return Value::Integer(v as i16);
    }
    if rank <= 2 && (i32::MIN as i128..=i32::MAX as i128).contains(&v) {
        return Value::Long(v as i32);
    }
    if (i64::MIN as i128..=i64::MAX as i128).contains(&v) {
        return Value::LongLong(v as i64);
    }
    Value::Double(v as f64)
}

fn arith(
    a: &Value,
    b: &Value,
    int_op: fn(i128, i128) -> i128,
    float_op: fn(f64, f64) -> f64,
) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    let (a, b) = (a.clone().into_scalar()?, b.clone().into_scalar()?);
    let rank = a.numeric_rank().max(b.numeric_rank());
    match rank {
        0..=3 => Ok(int_result(
            int_op(a.to_i64()? as i128, b.to_i64()? as i128),
            rank,
        )),
        4 => Ok(Value::Single(float_op(a.to_f64()?, b.to_f64()?) as f32)),
        6 => Ok(Value::Currency(
            (float_op(a.to_f64()?, b.to_f64()?) * 10_000.0).round_ties_even() as i64,
        )),
        7 => Ok(Value::Date(float_op(a.to_f64()?, b.to_f64()?))),
        _ => Ok(Value::Double(float_op(a.to_f64()?, b.to_f64()?))),
    }
}

pub fn add(a: &Value, b: &Value) -> VbaResult<Value> {
    if let (Value::String(x), Value::String(y)) = (a, b) {
        return Ok(Value::String(format!("{x}{y}")));
    }
    arith(a, b, |x, y| x + y, |x, y| x + y)
}

pub fn sub(a: &Value, b: &Value) -> VbaResult<Value> {
    // Date - Date is the Double number of days between them.
    if let (Value::Date(x), Value::Date(y)) = (a, b) {
        return Ok(Value::Double(x - y));
    }
    arith(a, b, |x, y| x - y, |x, y| x - y)
}

pub fn mul(a: &Value, b: &Value) -> VbaResult<Value> {
    arith(a, b, |x, y| x * y, |x, y| x * y)
}

pub fn div(a: &Value, b: &Value) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    let d = b.to_f64()?;
    if d == 0.0 {
        return Err(VbaError::division_by_zero());
    }
    Ok(Value::Double(a.to_f64()? / d))
}

pub fn int_div(a: &Value, b: &Value) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    let (x, y) = (a.to_i64()?, b.to_i64()?);
    if y == 0 {
        return Err(VbaError::division_by_zero());
    }
    Ok(int_result(
        (x / y) as i128,
        a.numeric_rank().max(b.numeric_rank()).clamp(1, 3),
    ))
}

pub fn modulo(a: &Value, b: &Value) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    let (x, y) = (a.to_i64()?, b.to_i64()?);
    if y == 0 {
        return Err(VbaError::division_by_zero());
    }
    Ok(int_result(
        (x % y) as i128,
        a.numeric_rank().max(b.numeric_rank()).clamp(1, 3),
    ))
}

pub fn pow(a: &Value, b: &Value) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    Ok(Value::Double(a.to_f64()?.powf(b.to_f64()?)))
}

pub fn concat(a: &Value, b: &Value) -> VbaResult<Value> {
    let part = |v: &Value| -> VbaResult<String> {
        match v {
            Value::Null => Ok(String::new()),
            other => other.to_string_value(),
        }
    };
    if matches!(a, Value::Null) && matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    Ok(Value::String(part(a)? + &part(b)?))
}

pub fn neg(a: &Value) -> VbaResult<Value> {
    match a {
        Value::Null => Ok(Value::Null),
        Value::Empty => Ok(Value::Integer(0)),
        Value::Byte(v) => Ok(Value::Integer(-(*v as i16))),
        Value::Integer(v) => Ok(int_result(-(*v as i128), 1)),
        Value::Long(v) => Ok(int_result(-(*v as i128), 2)),
        Value::LongLong(v) => Ok(int_result(-(*v as i128), 3)),
        Value::Single(v) => Ok(Value::Single(-v)),
        Value::Currency(v) => Ok(Value::Currency(-v)),
        Value::Date(v) => Ok(Value::Date(-v)),
        Value::Boolean(b) => Ok(Value::Integer(if *b { 1 } else { 0 })),
        other => Ok(Value::Double(-other.to_f64()?)),
    }
}

/// `Not` / `And` / `Or` / ... are bitwise on integers; on Booleans they stay Boolean.
pub fn logical(op: fn(i64, i64) -> i64, a: &Value, b: &Value) -> VbaResult<Value> {
    if matches!(a, Value::Null) || matches!(b, Value::Null) {
        return Ok(Value::Null);
    }
    if let (Value::Boolean(x), Value::Boolean(y)) = (a, b) {
        let r = op(if *x { -1 } else { 0 }, if *y { -1 } else { 0 });
        return Ok(Value::Boolean(r != 0));
    }
    let (x, y) = (bool_or_int(a)?, bool_or_int(b)?);
    Ok(int_result(
        op(x, y) as i128,
        a.numeric_rank().max(b.numeric_rank()).clamp(1, 3),
    ))
}

fn bool_or_int(v: &Value) -> VbaResult<i64> {
    match v {
        Value::Boolean(b) => Ok(if *b { -1 } else { 0 }),
        other => other.to_i64(),
    }
}

pub fn not(a: &Value) -> VbaResult<Value> {
    match a {
        Value::Null => Ok(Value::Null),
        Value::Boolean(b) => Ok(Value::Boolean(!b)),
        other => Ok(int_result(
            !other.to_i64()? as i128,
            other.numeric_rank().clamp(1, 3),
        )),
    }
}

/// Compares two values. Returns `None` if either side is Null.
pub fn compare(a: &Value, b: &Value, text: bool) -> VbaResult<Option<Ordering>> {
    let (a, b) = (a.clone().into_scalar()?, b.clone().into_scalar()?);
    match (&a, &b) {
        (Value::Null, _) | (_, Value::Null) => Ok(None),
        (Value::String(x), Value::String(y)) => Ok(Some(if text {
            x.to_lowercase().cmp(&y.to_lowercase())
        } else {
            x.cmp(y)
        })),
        (Value::String(s), n) | (n, Value::String(s)) if n.is_numeric_type() => {
            let swapped = matches!(a, Value::String(_));
            match parse_number(s) {
                Some(sv) => {
                    let nv = n.to_f64()?;
                    let ord = if swapped {
                        sv.partial_cmp(&nv)
                    } else {
                        nv.partial_cmp(&sv)
                    };
                    Ok(ord)
                }
                None if matches!(n, Value::Empty) => {
                    let ord = String::new().cmp(s);
                    Ok(Some(if swapped { ord.reverse() } else { ord }))
                }
                None => Err(VbaError::type_mismatch()),
            }
        }
        _ => Ok(a.to_f64()?.partial_cmp(&b.to_f64()?)),
    }
}

/// `Like` operator: `*`, `?`, `#`, `[a-z]`, `[!a-z]`.
pub fn like(s: &str, pattern: &str, text: bool) -> bool {
    let (s, p): (Vec<char>, Vec<char>) = if text {
        (
            s.to_lowercase().chars().collect(),
            pattern.to_lowercase().chars().collect(),
        )
    } else {
        (s.chars().collect(), pattern.chars().collect())
    };
    fn go(s: &[char], p: &[char]) -> bool {
        match p.first() {
            None => s.is_empty(),
            Some('*') => (0..=s.len()).any(|i| go(&s[i..], &p[1..])),
            Some('?') => !s.is_empty() && go(&s[1..], &p[1..]),
            Some('#') => s.first().is_some_and(|c| c.is_ascii_digit()) && go(&s[1..], &p[1..]),
            Some('[') => {
                let Some(close) = p.iter().position(|&c| c == ']') else {
                    return false;
                };
                let Some(&c) = s.first() else { return false };
                let mut class = &p[1..close];
                let negate = class.first() == Some(&'!');
                if negate {
                    class = &class[1..];
                }
                let mut hit = false;
                let mut i = 0;
                while i < class.len() {
                    if i + 2 < class.len() && class[i + 1] == '-' {
                        hit |= class[i] <= c && c <= class[i + 2];
                        i += 3;
                    } else {
                        hit |= class[i] == c;
                        i += 1;
                    }
                }
                hit != negate && go(&s[1..], &p[close + 1..])
            }
            Some(&ch) => s.first() == Some(&ch) && go(&s[1..], &p[1..]),
        }
    }
    go(&s, &p)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn integer_arithmetic_widens() {
        let r = add(&Value::Integer(32767), &Value::Integer(1)).unwrap();
        assert!(matches!(r, Value::Long(32768)));
        assert!(
            matches!(add(&Value::str("1"), &Value::Integer(2)).unwrap(), Value::Double(v) if v == 3.0)
        );
        assert!(
            matches!(add(&Value::str("1"), &Value::str("2")).unwrap(), Value::String(s) if s == "12")
        );
    }

    #[test]
    fn bankers_rounding_and_formatting() {
        assert_eq!(
            VarType::Long
                .coerce(Value::Double(2.5))
                .unwrap()
                .to_string(),
            "2"
        );
        assert_eq!(
            VarType::Long
                .coerce(Value::Double(3.5))
                .unwrap()
                .to_string(),
            "4"
        );
        assert_eq!(Value::Double(0.1 + 0.2).to_string(), "0.3");
        assert_eq!(Value::Double(1e20).to_string(), "1E+20");
        assert!(VarType::Integer.coerce(Value::Long(40000)).is_err());
    }

    #[test]
    fn like_patterns() {
        assert!(like("Report2024.xlsx", "Report####.xls?", false));
        assert!(like("abc", "[a-c]*", false));
        assert!(!like("abc", "[!a]*", false));
        assert!(like("ABC", "abc", true));
    }
}

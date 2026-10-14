//! `VBA.Conversion` and `VBA.Information`.

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    let _ = name;
    None
}

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    _host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
        // ---- conversion ----
        "cstr" => Value::String(arg(args, 0)?.to_string_value()?),
        "cbool" => VarType::Boolean.coerce(arg(args, 0)?.clone())?,
        "cbyte" => VarType::Byte.coerce(arg(args, 0)?.clone())?,
        "cint" => VarType::Integer.coerce(arg(args, 0)?.clone())?,
        "clng" => VarType::Long.coerce(arg(args, 0)?.clone())?,
        "clnglng" => VarType::LongLong.coerce(arg(args, 0)?.clone())?,
        "csng" => VarType::Single.coerce(arg(args, 0)?.clone())?,
        "cdbl" => VarType::Double.coerce(arg(args, 0)?.clone())?,
        "ccur" => VarType::Currency.coerce(arg(args, 0)?.clone())?,
        "cdate" => VarType::Date.coerce(arg(args, 0)?.clone())?,
        "cvar" => arg(args, 0)?.clone(),
        "cverr" => Value::Error(int(args, 0)? as i32),
        "val" => Value::Double(val(&s(args, 0)?)),
        "str" => {
            let v = arg(args, 0)?;
            let t = v.to_string_value()?;
            Value::String(if v.to_f64().is_ok_and(|x| x >= 0.0) {
                format!(" {t}")
            } else {
                t
            })
        }
        "hex" => Value::String(format!("{:X}", int(args, 0)?)),
        "oct" => Value::String(format!("{:o}", int(args, 0)?)),
        // ---- information ----
        "isnumeric" => Value::Boolean(match arg(args, 0)? {
            Value::String(t) => parse_number(t).is_some(),
            Value::Date(_) => false,
            v => v.is_numeric_type(),
        }),
        "isempty" => Value::Boolean(matches!(arg(args, 0)?, Value::Empty)),
        "isnull" => Value::Boolean(matches!(arg(args, 0)?, Value::Null)),
        "isobject" => Value::Boolean(matches!(arg(args, 0)?, Value::Object(_))),
        "isarray" => Value::Boolean(matches!(arg(args, 0)?, Value::Array(_))),
        "isdate" => Value::Boolean(match arg(args, 0)? {
            Value::Date(_) => true,
            Value::String(t) => date::parse_date(t).is_some(),
            _ => false,
        }),
        "iserror" => Value::Boolean(matches!(arg(args, 0)?, Value::Error(_))),
        "ismissing" => Value::Boolean(matches!(
            args.positional.first(),
            Some(Value::Missing) | None
        )),
        "typename" => Value::String(arg(args, 0)?.type_name()),
        "vartype" => Value::Integer(arg(args, 0)?.var_type_code() as i16),
        "ubound" | "lbound" => {
            let Value::Array(a) = arg(args, 0)? else {
                return Err(VbaError::type_mismatch());
            };
            let dim = opt(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(1);
            let (l, u) = *a
                .bounds
                .get((dim - 1).max(0) as usize)
                .filter(|_| dim >= 1)
                .ok_or_else(VbaError::subscript_out_of_range)?;
            Value::Long(if lname == "ubound" { u } else { l } as i32)
        }
        "array" => Value::Array(Box::new(VbArray::from_values(args.positional.clone(), 0))),
        "iif" => {
            if arg(args, 0)?.to_bool()? {
                arg(args, 1)?.clone()
            } else {
                arg(args, 2)?.clone()
            }
        }
        "choose" => {
            let i = int(args, 0)?;
            opt(args, i.max(0) as usize).cloned().unwrap_or(Value::Null)
        }
        "switch" => {
            let mut i = 0;
            while i + 1 < args.positional.len() {
                if args.positional[i].to_bool()? {
                    return Ok(args.positional[i + 1].clone());
                }
                i += 2;
            }
            Value::Null
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

/// `Val`: parses the leading numeric part, ignoring spaces.
fn val(s: &str) -> f64 {
    let t: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    if let Some(h) = t.strip_prefix("&H").or_else(|| t.strip_prefix("&h")) {
        let digits: String = h.chars().take_while(|c| c.is_ascii_hexdigit()).collect();
        return i64::from_str_radix(&digits, 16).unwrap_or(0) as f64;
    }
    let mut end = 0;
    for i in (1..=t.len()).rev() {
        if t.is_char_boundary(i) && t[..i].parse::<f64>().is_ok() {
            end = i;
            break;
        }
    }
    t[..end].parse().unwrap_or(0.0)
}

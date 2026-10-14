//! `VBA.Conversion` and `VBA.Information`.

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    Some(match name {
        "vbdecimal" => Value::Integer(14),
        "vblonglong" => Value::Integer(20),
        _ => return None,
    })
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
        "cbyte" => to_number(VarType::Byte, arg(args, 0)?)?,
        "cint" => to_number(VarType::Integer, arg(args, 0)?)?,
        "clng" => to_number(VarType::Long, arg(args, 0)?)?,
        "clnglng" => to_number(VarType::LongLong, arg(args, 0)?)?,
        "csng" => to_number(VarType::Single, arg(args, 0)?)?,
        "cdbl" => to_number(VarType::Double, arg(args, 0)?)?,
        "ccur" => to_number(VarType::Currency, arg(args, 0)?)?,
        "cdec" => to_number(VarType::Decimal, arg(args, 0)?)?,
        "cdate" => VarType::Date.coerce(arg(args, 0)?.clone())?,
        "cvar" => arg(args, 0)?.clone(),
        "cverr" => Value::Error(int(args, 0)? as i32),
        "val" => Value::Double(val(&s(args, 0)?)?),
        "str" => match arg(args, 0)? {
            Value::Null => Value::Null,
            v => {
                let t = v.to_string_value()?;
                Value::String(if v.to_f64().is_ok_and(|x| x >= 0.0) {
                    format!(" {t}")
                } else {
                    t
                })
            }
        },
        "hex" | "oct" => match arg(args, 0)? {
            Value::Null => Value::Null,
            v => {
                let bits = radix_bits(v)?;
                Value::String(if lname == "hex" {
                    format!("{bits:X}")
                } else {
                    format!("{bits:o}")
                })
            }
        },
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
            let cond = arg(args, 0)?;
            if !matches!(cond, Value::Null) && cond.to_bool()? {
                arg(args, 1)?.clone()
            } else {
                arg(args, 2)?.clone()
            }
        }
        "choose" => {
            let i = int(args, 0)?;
            match usize::try_from(i) {
                Ok(i) if i >= 1 => opt(args, i).cloned().unwrap_or(Value::Null),
                _ => Value::Null,
            }
        }
        "switch" => {
            if !args.positional.len().is_multiple_of(2) {
                return Err(VbaError::invalid_call());
            }
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

/// `CInt`, `CLng`, ...: as assigning to a variable of type `ty`, except that
/// an error value (`CVErr(n)`) converts to its number.
fn to_number(ty: VarType, v: &Value) -> VbaResult<Value> {
    match v {
        Value::Error(n) => ty.coerce(Value::Long(*n)),
        v => ty.coerce(v.clone()),
    }
}

/// The bits `Hex` / `Oct` show: a negative number is in two's complement of
/// its own size (16 bits for an Integer, 64 for a LongLong, else 32).
fn radix_bits(v: &Value) -> VbaResult<u64> {
    let n = v.to_i64()?;
    Ok(match v {
        Value::LongLong(_) => n as u64,
        Value::Empty | Value::Boolean(_) | Value::Byte(_) | Value::Integer(_) => n as u16 as u64,
        _ => i32::try_from(n).map_or(n as u64, |n| n as u32 as u64),
    })
}

/// `Val`: the leading number of `s`, ignoring whitespace anywhere. `&H` and
/// `&O` prefixes wrap like literals; a `%` or `&` suffix after a number with
/// a fraction is a Type mismatch.
fn val(s: &str) -> VbaResult<f64> {
    let t: String = s
        .chars()
        .filter(|c| !matches!(c, ' ' | '\t' | '\n' | '\r'))
        .collect();
    if let Some(rest) = t.strip_prefix('&') {
        let (prefix, radix) = match rest.chars().next() {
            Some(c @ ('h' | 'H')) => (Some(c), 16),
            Some(c @ ('o' | 'O')) => (Some(c), 8),
            _ => (None, 8),
        };
        let rest = &rest[prefix.map_or(0, |_| 1)..];
        let digits: String = rest.chars().take_while(|c| c.is_digit(radix)).collect();
        let long = if rest[digits.len()..].starts_with('&') {
            "&"
        } else {
            ""
        };
        let lit = format!("&{}{digits}{long}", if radix == 16 { "H" } else { "O" });
        return Ok(value::parse_radix(&lit).flatten().unwrap_or(0) as f64);
    }
    let b = t.as_bytes();
    let digits = |i: &mut usize| {
        while b.get(*i).is_some_and(u8::is_ascii_digit) {
            *i += 1;
        }
    };
    let mut i = usize::from(matches!(b.first(), Some(b'+' | b'-')));
    digits(&mut i);
    let mut integral = true;
    if b.get(i) == Some(&b'.') {
        i += 1;
        digits(&mut i);
        integral = false;
    }
    let mut num = t[..i].to_string();
    if matches!(b.get(i), Some(b'e' | b'E' | b'd' | b'D')) {
        let mut j = i + 1;
        if matches!(b.get(j), Some(b'+' | b'-')) {
            j += 1;
        }
        let start = j;
        digits(&mut j);
        if j > start {
            num = format!("{num}e{}", &t[i + 1..j]);
            (i, integral) = (j, false);
        }
    }
    if !integral && matches!(b.get(i), Some(b'%' | b'&')) {
        return Err(VbaError::type_mismatch());
    }
    Ok(num.parse().unwrap_or(0.0))
}

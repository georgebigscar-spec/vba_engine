//! VBA standard library: `VBA.Strings`, `VBA.Math`, `VBA.Conversion`,
//! `VBA.DateTime`, `VBA.Information`, `VBA.Interaction` and constants.

use std::cell::Cell;

use crate::date;
use crate::error::{VbaError, VbaResult};
use crate::host::Host;
use crate::object::{Args, Collection};
use crate::value::{self, parse_number, Value, VarType, VbArray};

/// Built-in constants (`vbCrLf`, `vbYesNo`, ...).
pub fn constant(name: &str) -> Option<Value> {
    let v = match name.to_ascii_lowercase().as_str() {
        "vbcrlf" | "vbnewline" => Value::str("\r\n"),
        "vbcr" => Value::str("\r"),
        "vblf" => Value::str("\n"),
        "vbtab" => Value::str("\t"),
        "vbnullstring" => Value::str(""),
        "vbnullchar" => Value::str("\0"),
        "vbokonly" | "vbbinarycompare" | "vbempty" => Value::Integer(0),
        "vbok" | "vbokcancel" | "vbtextcompare" | "vbnull" => Value::Integer(1),
        "vbcancel" | "vbabortretryignore" => Value::Integer(2),
        "vbabort" => Value::Integer(3),
        "vbcritical" => Value::Integer(16),
        "vbretry" => Value::Integer(4),
        "vbignore" => Value::Integer(5),
        "vbyes" => Value::Integer(6),
        "vbno" => Value::Integer(7),
        "vbyesnocancel" => Value::Integer(3),
        "vbyesno" => Value::Integer(4),
        "vbretrycancel" => Value::Integer(5),
        "vbquestion" => Value::Integer(32),
        "vbexclamation" => Value::Integer(48),
        "vbinformation" => Value::Integer(64),
        "vbinteger" => Value::Integer(2),
        "vblong" => Value::Integer(3),
        "vbsingle" => Value::Integer(4),
        "vbdouble" => Value::Integer(5),
        "vbcurrency" => Value::Integer(6),
        "vbdate" => Value::Integer(7),
        "vbstring" => Value::Integer(8),
        "vbobject" => Value::Integer(9),
        "vberror" => Value::Integer(10),
        "vbboolean" => Value::Integer(11),
        "vbvariant" => Value::Integer(12),
        "vbbyte" => Value::Integer(17),
        "vbarray" => Value::Integer(8192),
        "vbobjecterror" => Value::Long(-2147221504),
        "vbblack" => Value::Long(0),
        "vbred" => Value::Long(0xFF),
        "vbgreen" => Value::Long(0xFF00),
        "vbblue" => Value::Long(0xFF0000),
        "vbwhite" => Value::Long(0xFFFFFF),
        "vbsunday" => Value::Integer(1),
        "vbmonday" => Value::Integer(2),
        _ => return None,
    };
    Some(v)
}

fn arg(args: &Args, i: usize) -> VbaResult<&Value> {
    args.positional
        .get(i)
        .filter(|v| !matches!(v, Value::Missing))
        .ok_or_else(VbaError::wrong_arg_count)
}

fn opt(args: &Args, i: usize) -> Option<&Value> {
    args.positional
        .get(i)
        .filter(|v| !matches!(v, Value::Missing))
}

fn s(args: &Args, i: usize) -> VbaResult<String> {
    arg(args, i)?.to_string_value()
}

fn n(args: &Args, i: usize) -> VbaResult<f64> {
    arg(args, i)?.to_f64()
}

fn int(args: &Args, i: usize) -> VbaResult<i64> {
    arg(args, i)?.to_i64()
}

fn chars(s: &str) -> Vec<char> {
    s.chars().collect()
}

fn null_passthrough(args: &Args) -> bool {
    matches!(args.positional.first(), Some(Value::Null))
}

thread_local! {
    static RND_STATE: Cell<u32> = const { Cell::new(0x50000) };
}

/// Calls a built-in function. Returns `None` if `name` is not a built-in.
/// Built-ins that look at an object itself rather than its default value.
pub fn takes_objects(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "typename"
            | "isobject"
            | "vartype"
            | "isnull"
            | "isempty"
            | "isarray"
            | "ismissing"
            | "objptr"
            | "callbyname"
            | "iserror"
    )
}

pub fn call(
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    text_compare: bool,
) -> Option<VbaResult<Value>> {
    let lname = name.to_ascii_lowercase();
    let r = (|| -> VbaResult<Value> {
        Ok(match lname.as_str() {
            // ---- strings ----
            "len" => match arg(args, 0)? {
                Value::Null => Value::Null,
                v => Value::Long(v.to_string_value()?.chars().count() as i32),
            },
            "left" | "right" if null_passthrough(args) => Value::Null,
            "left" => {
                let c = chars(&s(args, 0)?);
                let k = (int(args, 1)?.max(0) as usize).min(c.len());
                Value::String(c[..k].iter().collect())
            }
            "right" => {
                let c = chars(&s(args, 0)?);
                let k = (int(args, 1)?.max(0) as usize).min(c.len());
                Value::String(c[c.len() - k..].iter().collect())
            }
            "mid" => {
                if null_passthrough(args) {
                    return Ok(Value::Null);
                }
                let c = chars(&s(args, 0)?);
                let start = int(args, 1)?;
                if start < 1 {
                    return Err(VbaError::invalid_call());
                }
                let from = (start as usize - 1).min(c.len());
                let len = opt(args, 2)
                    .map(|v| v.to_i64())
                    .transpose()?
                    .map_or(c.len(), |l| l.max(0) as usize);
                Value::String(c[from..(from + len).min(c.len())].iter().collect())
            }
            "ucase" => Value::String(s(args, 0)?.to_uppercase()),
            "lcase" => Value::String(s(args, 0)?.to_lowercase()),
            "trim" => Value::String(s(args, 0)?.trim_matches(' ').to_string()),
            "ltrim" => Value::String(s(args, 0)?.trim_start_matches(' ').to_string()),
            "rtrim" => Value::String(s(args, 0)?.trim_end_matches(' ').to_string()),
            "space" => Value::String(" ".repeat(int(args, 0)?.max(0) as usize)),
            "string" => {
                let ch = match arg(args, 1)? {
                    Value::String(t) => t.chars().next().ok_or_else(VbaError::invalid_call)?,
                    v => char::from_u32(v.to_i64()? as u32).ok_or_else(VbaError::invalid_call)?,
                };
                Value::String(ch.to_string().repeat(int(args, 0)?.max(0) as usize))
            }
            "strreverse" => Value::String(s(args, 0)?.chars().rev().collect()),
            "instr" => {
                // InStr([start,] string1, string2[, compare])
                let (start, a, b, cmp) = if args.positional.len() >= 3
                    && !matches!(opt(args, 0), Some(Value::String(_)))
                {
                    (int(args, 0)?, s(args, 1)?, s(args, 2)?, opt(args, 3))
                } else {
                    (1, s(args, 0)?, s(args, 1)?, None)
                };
                let text = cmp
                    .map(|v| v.to_i64())
                    .transpose()?
                    .map_or(text_compare, |c| c == 1);
                Value::Long(instr(start, &a, &b, text)? as i32)
            }
            "instrrev" => {
                let (a, b) = (s(args, 0)?, s(args, 1)?);
                let (a, b) =
                    if opt(args, 3).map(|v| v.to_i64()).transpose()? == Some(1) || text_compare {
                        (a.to_lowercase(), b.to_lowercase())
                    } else {
                        (a, b)
                    };
                let ca = chars(&a);
                let start = opt(args, 2).map(|v| v.to_i64()).transpose()?.unwrap_or(-1);
                let end = if start < 0 {
                    ca.len()
                } else {
                    (start as usize).min(ca.len())
                };
                let hay: String = ca[..end].iter().collect();
                Value::Long(
                    hay.rfind(&b)
                        .map_or(0, |byte| hay[..byte].chars().count() + 1)
                        as i32,
                )
            }
            "replace" => {
                let (src, find, with) = (s(args, 0)?, s(args, 1)?, s(args, 2)?);
                if find.is_empty() {
                    return Ok(Value::String(src));
                }
                let text = opt(args, 5)
                    .map(|v| v.to_i64())
                    .transpose()?
                    .map_or(text_compare, |c| c == 1);
                Value::String(if text {
                    replace_ci(&src, &find, &with)
                } else {
                    src.replace(&find, &with)
                })
            }
            "strcomp" => {
                let text = opt(args, 2)
                    .map(|v| v.to_i64())
                    .transpose()?
                    .map_or(text_compare, |c| c == 1);
                let ord = value::compare(
                    &Value::String(s(args, 0)?),
                    &Value::String(s(args, 1)?),
                    text,
                )?;
                Value::Integer(ord.map_or(0, |o| o as i16))
            }
            "chr" | "chrw" => {
                let code = int(args, 0)?;
                Value::String(
                    char::from_u32(code as u32)
                        .ok_or_else(VbaError::invalid_call)?
                        .to_string(),
                )
            }
            "asc" | "ascw" => Value::Integer(
                s(args, 0)?
                    .chars()
                    .next()
                    .ok_or_else(VbaError::invalid_call)? as i16,
            ),
            "join" => {
                let delim = opt(args, 1)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_else(|| " ".into());
                match arg(args, 0)? {
                    Value::Array(a) => Value::String(
                        a.data
                            .iter()
                            .map(|v| v.to_string_value())
                            .collect::<VbaResult<Vec<_>>>()?
                            .join(&delim),
                    ),
                    _ => return Err(VbaError::type_mismatch()),
                }
            }
            "split" => {
                let src = s(args, 0)?;
                let delim = opt(args, 1)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_else(|| " ".into());
                let parts: Vec<Value> = if src.is_empty() {
                    vec![]
                } else if delim.is_empty() {
                    vec![Value::String(src)]
                } else {
                    src.split(delim.as_str()).map(Value::str).collect()
                };
                let mut arr = VbArray::from_values(parts, 0);
                arr.elem_type = VarType::String;
                Value::Array(Box::new(arr))
            }
            "format" | "format$" => {
                let fmt = opt(args, 1)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                Value::String(format_value(arg(args, 0)?, &fmt)?)
            }
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
            "cdate" | "datevalue" => VarType::Date.coerce(arg(args, 0)?.clone())?,
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
            // ---- math ----
            "abs" => match arg(args, 0)? {
                Value::Null => Value::Null,
                v if value::compare(v, &Value::Integer(0), false)?
                    == Some(std::cmp::Ordering::Less) =>
                {
                    value::neg(v)?
                }
                v => v.clone(),
            },
            "sgn" => Value::Integer(n(args, 0)?.signum() as i16 * i16::from(n(args, 0)? != 0.0)),
            "int" => Value::Double(n(args, 0)?.floor()),
            "fix" => Value::Double(n(args, 0)?.trunc()),
            "sqr" => {
                let x = n(args, 0)?;
                if x < 0.0 {
                    return Err(VbaError::invalid_call());
                }
                Value::Double(x.sqrt())
            }
            "exp" => Value::Double(n(args, 0)?.exp()),
            "log" => {
                let x = n(args, 0)?;
                if x <= 0.0 {
                    return Err(VbaError::invalid_call());
                }
                Value::Double(x.ln())
            }
            "sin" => Value::Double(n(args, 0)?.sin()),
            "cos" => Value::Double(n(args, 0)?.cos()),
            "tan" => Value::Double(n(args, 0)?.tan()),
            "atn" => Value::Double(n(args, 0)?.atan()),
            "round" => {
                let digits = opt(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0);
                let f = 10f64.powi(digits as i32);
                Value::Double((n(args, 0)? * f).round_ties_even() / f)
            }
            "rnd" => RND_STATE.with(|st| {
                // VBA's 24-bit LCG.
                let next = (st.get().wrapping_mul(0x43FD43FD).wrapping_add(0xC39EC3)) & 0xFF_FFFF;
                st.set(next);
                Value::Single(next as f32 / 16_777_216.0)
            }),
            "randomize" => Value::Empty,
            "rgb" => Value::Long(
                (int(args, 0)? & 0xFF | (int(args, 1)? & 0xFF) << 8 | (int(args, 2)? & 0xFF) << 16)
                    as i32,
            ),
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
            // ---- date / time ----
            "now" => Value::Date(host.now()),
            "date" => Value::Date(host.now().floor()),
            "time" => Value::Date(host.now().fract()),
            "timer" => Value::Single((host.now().fract() * 86_400.0) as f32),
            "year" => Value::Integer(date::ymd(arg(args, 0)?.to_date()?).0 as i16),
            "month" => Value::Integer(date::ymd(arg(args, 0)?.to_date()?).1 as i16),
            "day" => Value::Integer(date::ymd(arg(args, 0)?.to_date()?).2 as i16),
            "hour" => Value::Integer(date::hms(arg(args, 0)?.to_date()?).0 as i16),
            "minute" => Value::Integer(date::hms(arg(args, 0)?.to_date()?).1 as i16),
            "second" => Value::Integer(date::hms(arg(args, 0)?.to_date()?).2 as i16),
            "weekday" => {
                let d = arg(args, 0)?.to_date()?.floor() as i64;
                // 1899-12-30 was a Saturday (vbSaturday = 7).
                Value::Integer(((d + 6).rem_euclid(7) + 1) as i16)
            }
            "dateserial" => {
                let (y, m, d) = (int(args, 0)?, int(args, 1)?, int(args, 2)?);
                let (y, m) = (y + (m - 1).div_euclid(12), (m - 1).rem_euclid(12) + 1);
                Value::Date(date::serial(y, m, 1) + (d - 1) as f64)
            }
            "timeserial" => Value::Date(
                (int(args, 0)? * 3600 + int(args, 1)? * 60 + int(args, 2)?) as f64 / 86_400.0,
            ),
            "dateadd" => {
                let unit = s(args, 0)?.to_ascii_lowercase();
                let k = int(args, 1)?;
                let d = arg(args, 2)?.to_date()?;
                Value::Date(match unit.as_str() {
                    "d" | "y" | "w" => d + k as f64,
                    "ww" => d + 7.0 * k as f64,
                    "h" => d + k as f64 / 24.0,
                    "n" => d + k as f64 / 1440.0,
                    "s" => d + k as f64 / 86_400.0,
                    "m" | "q" | "yyyy" => {
                        let months = k * match unit.as_str() {
                            "m" => 1,
                            "q" => 3,
                            _ => 12,
                        };
                        let (y, m, dd) = date::ymd(d);
                        let total = y * 12 + (m - 1) + months;
                        let (ny, nm) = (total.div_euclid(12), total.rem_euclid(12) + 1);
                        let month_len = (date::serial(ny + nm / 12, nm % 12 + 1, 1)
                            - date::serial(ny, nm, 1))
                            as i64;
                        date::serial(ny, nm, dd.min(month_len)) + d.fract()
                    }
                    _ => return Err(VbaError::invalid_call()),
                })
            }
            "datediff" => {
                let unit = s(args, 0)?.to_ascii_lowercase();
                let (a, b) = (arg(args, 1)?.to_date()?, arg(args, 2)?.to_date()?);
                Value::LongLong(match unit.as_str() {
                    "d" | "y" => (b.floor() - a.floor()) as i64,
                    "ww" | "w" => ((b.floor() - a.floor()) / 7.0) as i64,
                    "h" => ((b - a) * 24.0).round() as i64,
                    "n" => ((b - a) * 1440.0).round() as i64,
                    "s" => ((b - a) * 86_400.0).round() as i64,
                    "m" => {
                        let ((ya, ma, _), (yb, mb, _)) = (date::ymd(a), date::ymd(b));
                        (yb * 12 + mb) - (ya * 12 + ma)
                    }
                    "yyyy" => date::ymd(b).0 - date::ymd(a).0,
                    _ => return Err(VbaError::invalid_call()),
                })
            }
            // ---- interaction ----
            "msgbox" => {
                let buttons = opt(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0) as i32;
                let title = opt(args, 2)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                Value::Integer(host.msgbox(&s(args, 0)?, buttons, &title)? as i16)
            }
            "inputbox" => {
                let title = opt(args, 1)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                let default = opt(args, 2)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_default();
                Value::String(
                    host.inputbox(&s(args, 0)?, &title, &default)?
                        .unwrap_or_default(),
                )
            }
            "createobject" => match s(args, 0)?.to_ascii_lowercase().as_str() {
                "vba.collection" => Value::from(Collection::new()),
                prog_id => host.create_object(prog_id)?,
            },
            "doevents" => Value::Integer(0),
            _ => return Err(VbaError::sub_not_defined(name)),
        })
    })();
    match r {
        Err(e) if e.number == 35 && e.description.ends_with(name) => None,
        other => Some(other),
    }
}

fn instr(start: i64, hay: &str, needle: &str, text: bool) -> VbaResult<usize> {
    if start < 1 {
        return Err(VbaError::invalid_call());
    }
    let (hay, needle) = if text {
        (hay.to_lowercase(), needle.to_lowercase())
    } else {
        (hay.to_string(), needle.to_string())
    };
    let skip = start as usize - 1;
    let byte_start = match hay.char_indices().nth(skip) {
        Some((b, _)) => b,
        None => {
            return Ok(if needle.is_empty() && skip == hay.chars().count() {
                start as usize
            } else {
                0
            })
        }
    };
    Ok(hay[byte_start..]
        .find(&needle)
        .map_or(0, |b| hay[..byte_start + b].chars().count() + 1))
}

fn replace_ci(src: &str, find: &str, with: &str) -> String {
    let lower = src.to_lowercase();
    let lfind = find.to_lowercase();
    if lower.len() != src.len() {
        return src.replace(find, with);
    }
    let mut out = String::new();
    let mut last = 0;
    for (i, _) in lower.match_indices(&lfind) {
        out.push_str(&src[last..i]);
        out.push_str(with);
        last = i + find.len();
    }
    out.push_str(&src[last..]);
    out
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

/// A pragmatic subset of `Format`: named formats, `0`/`#`/`,`/`.`/`%` number
/// masks and `yyyy mm dd hh nn ss` date masks.
pub fn format_value(v: &Value, fmt: &str) -> VbaResult<String> {
    if fmt.is_empty() {
        return v.to_string_value();
    }
    let lf = fmt.to_ascii_lowercase();
    let num = || v.to_f64();
    match lf.as_str() {
        "general number" => return Ok(value::format_double(num()?)),
        "fixed" => return Ok(format!("{:.2}", num()?)),
        "standard" => return format_number(num()?, "#,##0.00"),
        "percent" => return Ok(format!("{:.2}%", num()? * 100.0)),
        "currency" => return Ok(format!("${}", format_number(num()?, "#,##0.00")?)),
        "yes/no" => return Ok(if num()? != 0.0 { "Yes" } else { "No" }.into()),
        "true/false" => return Ok(if num()? != 0.0 { "True" } else { "False" }.into()),
        "short date" => return format_date_mask(v.to_date()?, "m/d/yyyy"),
        "long date" => return format_date_mask(v.to_date()?, "dddd, mmmm d, yyyy"),
        "short time" => return format_date_mask(v.to_date()?, "hh:nn"),
        "long time" => return format_date_mask(v.to_date()?, "h:nn:ss AM/PM"),
        _ => {}
    }
    if lf.chars().any(|c| matches!(c, 'y' | 'd' | 'h' | 'n' | 's'))
        && !lf.contains('#')
        && !lf.contains('0')
    {
        return format_date_mask(v.to_date()?, fmt);
    }
    if fmt.contains(['0', '#']) {
        return format_number(num()?, fmt);
    }
    v.to_string_value()
}

fn format_number(x: f64, fmt: &str) -> VbaResult<String> {
    let percent = fmt.contains('%');
    let x = if percent { x * 100.0 } else { x };
    let mask = fmt.trim_end_matches('%');
    let (int_mask, frac_mask) = mask.split_once('.').unwrap_or((mask, ""));
    let decimals = frac_mask.chars().filter(|c| matches!(c, '0' | '#')).count();
    let min_frac = frac_mask.chars().filter(|&c| c == '0').count();
    let min_int = int_mask.chars().filter(|&c| c == '0').count();
    let grouping = int_mask.contains(',');
    let factor = 10f64.powi(decimals as i32);
    let rounded = (x.abs() * factor).round() / factor;
    let mut s = format!("{rounded:.decimals$}");
    let (ip, fp) = match s.split_once('.') {
        Some((i, f)) => (i.to_string(), f.trim_end_matches('0').to_string()),
        None => (s.clone(), String::new()),
    };
    let mut fp = fp;
    while fp.len() < min_frac {
        fp.push('0');
    }
    let mut ip = ip.trim_start_matches('0').to_string();
    while ip.len() < min_int {
        ip.insert(0, '0');
    }
    if grouping {
        let digits: Vec<char> = ip.chars().collect();
        ip = digits
            .iter()
            .enumerate()
            .map(|(i, c)| {
                if i > 0 && (digits.len() - i).is_multiple_of(3) {
                    format!(",{c}")
                } else {
                    c.to_string()
                }
            })
            .collect();
    }
    s = if fp.is_empty() {
        ip
    } else {
        format!("{ip}.{fp}")
    };
    if x < 0.0 && rounded != 0.0 {
        s.insert(0, '-');
    }
    if percent {
        s.push('%');
    }
    Ok(s)
}

fn format_date_mask(d: f64, mask: &str) -> VbaResult<String> {
    const MONTHS: [&str; 12] = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];
    const DAYS: [&str; 7] = [
        "Sunday",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday",
        "Friday",
        "Saturday",
    ];
    let (y, m, dd) = date::ymd(d);
    let (h, mi, sec) = date::hms(d);
    let ampm = mask.to_ascii_uppercase().contains("AM/PM");
    let hour = if ampm {
        if h % 12 == 0 {
            12
        } else {
            h % 12
        }
    } else {
        h
    };
    let wd = ((d.floor() as i64 + 6).rem_euclid(7)) as usize;
    let c: Vec<char> = mask.chars().collect();
    let mut out = String::new();
    let mut i = 0;
    let mut last_was_hour = false;
    while i < c.len() {
        let ch = c[i].to_ascii_lowercase();
        let mut run = 1;
        while i + run < c.len() && c[i + run].to_ascii_lowercase() == ch {
            run += 1;
        }
        match ch {
            'y' if run >= 4 => out.push_str(&format!("{y:04}")),
            'y' => out.push_str(&format!("{:02}", y % 100)),
            // `mm` right after an hour means minutes, as in VBA.
            'm' if last_was_hour && run <= 2 => out.push_str(&format!("{mi:0run$}")),
            'm' if run >= 4 => out.push_str(MONTHS[m as usize - 1]),
            'm' if run == 3 => out.push_str(&MONTHS[m as usize - 1][..3]),
            'm' => out.push_str(&format!("{m:0run$}")),
            'd' if run >= 4 => out.push_str(DAYS[wd]),
            'd' if run == 3 => out.push_str(&DAYS[wd][..3]),
            'd' => out.push_str(&format!("{dd:0run$}")),
            'h' => out.push_str(&format!("{hour:0run$}")),
            'n' => out.push_str(&format!("{mi:0run$}")),
            's' => out.push_str(&format!("{sec:0run$}")),
            'a' if mask[i..].to_ascii_uppercase().starts_with("AM/PM") => {
                out.push_str(if h < 12 { "AM" } else { "PM" });
                run = 5;
            }
            _ => out.extend(&c[i..i + run]),
        }
        if ch != ' ' && ch != ':' {
            last_was_hour = ch == 'h';
        }
        i += run;
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::host::CaptureHost;

    fn call_s(name: &str, args: Vec<Value>) -> String {
        call(name, &Args::new(args), &mut CaptureHost::new(), false)
            .unwrap()
            .unwrap()
            .to_string()
    }

    #[test]
    fn string_functions() {
        assert_eq!(call_s("Mid", vec!["Hello World".into(), 7.into()]), "World");
        assert_eq!(call_s("InStr", vec!["Hello".into(), "l".into()]), "3");
        assert_eq!(
            call_s("InStr", vec![4.into(), "Hello".into(), "l".into()]),
            "4"
        );
        assert_eq!(
            call_s("Replace", vec!["a-b-c".into(), "-".into(), "+".into()]),
            "a+b+c"
        );
        assert_eq!(call_s("Val", vec!["  12.5abc".into()]), "12.5");
    }

    #[test]
    fn format_function() {
        assert_eq!(
            call_s("Format", vec![1234.5.into(), "#,##0.00".into()]),
            "1,234.50"
        );
        assert_eq!(call_s("Format", vec![0.256.into(), "0.0%".into()]), "25.6%");
        let d = Value::Date(date::serial(2024, 3, 5) + 0.5);
        assert_eq!(
            call_s("Format", vec![d, "yyyy-mm-dd hh:nn".into()]),
            "2024-03-05 12:00"
        );
    }

    #[test]
    fn unknown_is_none() {
        assert!(call("NoSuchFn", &Args::default(), &mut CaptureHost::new(), false).is_none());
    }
}

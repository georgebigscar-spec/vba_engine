//! `VBA.Strings`.

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
    text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
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
            let (start, a, b, cmp) =
                if args.positional.len() >= 3 && !matches!(opt(args, 0), Some(Value::String(_))) {
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
            let (a, b) = if opt(args, 3).map(|v| v.to_i64()).transpose()? == Some(1) || text_compare
            {
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
                    .map_or(0, |byte| hay[..byte].chars().count() + 1) as i32,
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
        _ => return Err(VbaError::sub_not_defined(name)),
    })
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

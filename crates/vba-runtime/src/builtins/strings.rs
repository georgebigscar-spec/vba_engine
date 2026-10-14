//! `VBA.Strings`.

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    Some(Value::Integer(match name {
        "vbuppercase" => 1,
        "vblowercase" => 2,
        "vbpropercase" => 3,
        "vbwide" => 4,
        "vbnarrow" => 8,
        "vbkatakana" => 16,
        "vbhiragana" => 32,
        "vbunicode" => 64,
        "vbfromunicode" => 128,
        _ => return None,
    }))
}

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    text_compare: bool,
) -> VbaResult<Value> {
    // `Left$` and friends: the plain function, but Null is an error.
    if let Some(base) = lname.strip_suffix('$') {
        if !matches!(
            base,
            "left"
                | "right"
                | "mid"
                | "ucase"
                | "lcase"
                | "trim"
                | "ltrim"
                | "rtrim"
                | "space"
                | "string"
                | "chr"
                | "chrw"
                | "hex"
                | "oct"
                | "str"
                | "format"
        ) {
            return Err(VbaError::sub_not_defined(name));
        }
        if base != "format" && args.positional.iter().any(|v| matches!(v, Value::Null)) {
            return Err(VbaError::invalid_use_of_null());
        }
        return super::call(base, args, host, text_compare)
            .unwrap_or_else(|| Err(VbaError::sub_not_defined(name)));
    }
    let compare = |i: usize| -> VbaResult<bool> {
        Ok(opt(args, i)
            .map(|v| v.to_i64())
            .transpose()?
            .map_or(text_compare, |c| c == 1))
    };
    Ok(match lname {
        "len" => match arg(args, 0)? {
            Value::Null => Value::Null,
            v => Value::Long(v.to_string_value()?.chars().count() as i32),
        },
        "left" | "right" => {
            let len = int(args, 1)?;
            if len < 0 {
                return Err(VbaError::invalid_call());
            }
            if null_passthrough(args) {
                return Ok(Value::Null);
            }
            let c = chars(&s(args, 0)?);
            let k = (len as usize).min(c.len());
            Value::String(if lname == "left" {
                c[..k].iter().collect()
            } else {
                c[c.len() - k..].iter().collect()
            })
        }
        "mid" => {
            let start = int(args, 1)?;
            let len = opt(args, 2).map(|v| v.to_i64()).transpose()?;
            if start < 1 || len.is_some_and(|l| l < 0) {
                return Err(VbaError::invalid_call());
            }
            if null_passthrough(args) {
                return Ok(Value::Null);
            }
            let c = chars(&s(args, 0)?);
            let from = (start as usize - 1).min(c.len());
            let len = len.map_or(c.len(), |l| l as usize);
            Value::String(c[from..(from + len).min(c.len())].iter().collect())
        }
        "ucase" | "lcase" | "trim" | "ltrim" | "rtrim" if null_passthrough(args) => Value::Null,
        "ucase" => Value::String(s(args, 0)?.to_uppercase()),
        "lcase" => Value::String(s(args, 0)?.to_lowercase()),
        "trim" => Value::String(s(args, 0)?.trim_matches(' ').to_string()),
        "ltrim" => Value::String(s(args, 0)?.trim_start_matches(' ').to_string()),
        "rtrim" => Value::String(s(args, 0)?.trim_end_matches(' ').to_string()),
        "space" => Value::String(" ".repeat(count(args, 0)?)),
        "string" => {
            let n = count(args, 0)?;
            let ch = match arg(args, 1)? {
                Value::Null => return Ok(Value::Null),
                Value::String(t) => t.chars().next().ok_or_else(VbaError::invalid_call)?,
                v => from_ansi(v.to_i64()?.rem_euclid(256) as u8),
            };
            Value::String(ch.to_string().repeat(n))
        }
        "strreverse" => Value::String(s(args, 0)?.chars().rev().collect()),
        "strconv" => {
            let conv = int(args, 1)?;
            if null_passthrough(args) {
                return Ok(Value::Null);
            }
            let src = s(args, 0)?;
            Value::String(match conv {
                1 => src.to_uppercase(),
                2 => src.to_lowercase(),
                3 => proper_case(&src),
                // Wide/narrow/kana and ANSI conversions have no effect on
                // Unicode text in a single-byte locale.
                4 | 8 | 16 | 32 | 64 | 128 => src,
                _ => return Err(VbaError::invalid_call()),
            })
        }
        "instr" => {
            // InStr([start,] string1, string2[, compare])
            let (start, a, b, cmp) = if args.positional.len() >= 3 {
                (int(args, 0)?, arg(args, 1)?, arg(args, 2)?, 3)
            } else {
                (1, arg(args, 0)?, arg(args, 1)?, 4)
            };
            let text = compare(cmp)?;
            if start < 1 {
                return Err(VbaError::invalid_call());
            }
            if matches!(a, Value::Null) || matches!(b, Value::Null) {
                return Ok(Value::Null);
            }
            let (a, b) = (chars(&a.to_string_value()?), chars(&b.to_string_value()?));
            let from = start as usize - 1;
            Value::Long(if from > a.len() {
                0
            } else {
                find(&a, &b, from, text).map_or(0, |p| p + 1) as i32
            })
        }
        "instrrev" => {
            let start = opt(args, 2).map(|v| v.to_i64()).transpose()?.unwrap_or(-1);
            let text = compare(3)?;
            if start == 0 || start < -1 {
                return Err(VbaError::invalid_call());
            }
            if null_passthrough(args) || matches!(arg(args, 1)?, Value::Null) {
                return Ok(Value::Null);
            }
            let (a, b) = (chars(&s(args, 0)?), chars(&s(args, 1)?));
            let end = if start == -1 { a.len() } else { start as usize };
            Value::Long(if a.is_empty() || end > a.len() {
                0
            } else if b.is_empty() {
                end as i32
            } else {
                (0..=end.saturating_sub(b.len()))
                    .rev()
                    .find(|&i| end >= b.len() && matches_at(&a, &b, i, text))
                    .map_or(0, |p| p + 1) as i32
            })
        }
        "replace" => {
            let start = opt(args, 3).map(|v| v.to_i64()).transpose()?.unwrap_or(1);
            let limit = opt(args, 4).map(|v| v.to_i64()).transpose()?.unwrap_or(-1);
            let text = compare(5)?;
            let (src, find_s, with) = (s(args, 0)?, s(args, 1)?, s(args, 2)?);
            if start < 1 || limit < -1 {
                return Err(VbaError::invalid_call());
            }
            let c = chars(&src);
            let from = (start as usize - 1).min(c.len());
            let tail = &c[from..];
            let needle = chars(&find_s);
            let mut out = String::new();
            let mut i = 0;
            let mut done = 0;
            while i < tail.len() {
                if !needle.is_empty()
                    && (limit < 0 || done < limit)
                    && matches_at(tail, &needle, i, text)
                {
                    out.push_str(&with);
                    i += needle.len();
                    done += 1;
                } else {
                    out.push(tail[i]);
                    i += 1;
                }
            }
            Value::String(out)
        }
        "strcomp" => {
            let text = compare(2)?;
            if null_passthrough(args) || matches!(arg(args, 1)?, Value::Null) {
                return Ok(Value::Null);
            }
            let ord = value::compare(
                &Value::String(s(args, 0)?),
                &Value::String(s(args, 1)?),
                text,
            )?;
            Value::Integer(ord.map_or(0, |o| o as i16))
        }
        "chr" => match int(args, 0)? {
            code @ 0..=255 => Value::String(from_ansi(code as u8).to_string()),
            _ => return Err(VbaError::invalid_call()),
        },
        "chrw" => match int(args, 0)? {
            code @ -32768..=65535 => Value::String(
                char::from_u32(code.rem_euclid(65536) as u32)
                    .ok_or_else(VbaError::invalid_call)?
                    .to_string(),
            ),
            _ => return Err(VbaError::invalid_call()),
        },
        "asc" => Value::Integer(to_ansi(first_char(args)?) as i16),
        "ascw" => {
            let mut buf = [0u16; 2];
            Value::Integer(first_char(args)?.encode_utf16(&mut buf)[0] as i16)
        }
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
            let src = chars(&s(args, 0)?);
            let delim = chars(
                &opt(args, 1)
                    .map(|v| v.to_string_value())
                    .transpose()?
                    .unwrap_or_else(|| " ".into()),
            );
            let limit = opt(args, 2).map(|v| v.to_i64()).transpose()?.unwrap_or(-1);
            let text = compare(3)?;
            if limit < -1 {
                return Err(VbaError::invalid_call());
            }
            let mut parts: Vec<Value> = Vec::new();
            if !src.is_empty() && limit != 0 {
                let mut from = 0;
                while limit < 0 || (parts.len() as i64) < limit - 1 {
                    let Some(p) = (!delim.is_empty())
                        .then(|| find(&src, &delim, from, text))
                        .flatten()
                    else {
                        break;
                    };
                    parts.push(Value::String(src[from..p].iter().collect()));
                    from = p + delim.len();
                }
                parts.push(Value::String(src[from..].iter().collect()));
            }
            string_array(parts)
        }
        "filter" => {
            let Value::Array(src) = arg(args, 0)? else {
                return Err(VbaError::type_mismatch());
            };
            if src.bounds.len() != 1 {
                return Err(VbaError::type_mismatch());
            }
            let needle = chars(&s(args, 1)?);
            let include = opt(args, 2)
                .map(|v| v.to_bool())
                .transpose()?
                .unwrap_or(true);
            let text = compare(3)?;
            let mut out = Vec::new();
            for v in &src.data {
                let item = v.to_string_value()?;
                if find(&chars(&item), &needle, 0, text).is_some() == include {
                    out.push(Value::String(item));
                }
            }
            string_array(out)
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

/// A non-negative count argument (`Space`, `String`).
fn count(args: &Args, i: usize) -> VbaResult<usize> {
    usize::try_from(int(args, i)?).map_err(|_| VbaError::invalid_call())
}

fn first_char(args: &Args) -> VbaResult<char> {
    s(args, 0)?
        .chars()
        .next()
        .ok_or_else(VbaError::invalid_call)
}

/// A zero-based `String()` array.
fn string_array(parts: Vec<Value>) -> Value {
    let mut arr = VbArray::from_values(parts, 0);
    arr.elem_type = VarType::String;
    Value::Array(Box::new(arr))
}

fn eq_char(a: char, b: char, text: bool) -> bool {
    a == b || (text && a.to_lowercase().eq(b.to_lowercase()))
}

fn matches_at(hay: &[char], needle: &[char], i: usize, text: bool) -> bool {
    hay.len() >= i + needle.len()
        && hay[i..i + needle.len()]
            .iter()
            .zip(needle)
            .all(|(&a, &b)| eq_char(a, b, text))
}

/// Index of the first match of `needle` at or after `from`.
fn find(hay: &[char], needle: &[char], from: usize, text: bool) -> Option<usize> {
    (from..=hay.len().saturating_sub(needle.len())).find(|&i| matches_at(hay, needle, i, text))
}

/// `StrConv(s, vbProperCase)`: first letter of each word upper case.
fn proper_case(s: &str) -> String {
    let mut out = String::new();
    let mut word_start = true;
    for c in s.chars() {
        if word_start {
            out.extend(c.to_uppercase());
        } else {
            out.extend(c.to_lowercase());
        }
        word_start = matches!(c, ' ' | '\t' | '\n' | '\r' | '\x0b' | '\x0c' | '\0');
    }
    out
}

/// Windows-1252 characters for bytes 0x80..=0x9F (undefined ones map to
/// themselves).
const CP1252_HIGH: [char; 32] = [
    '\u{20AC}', '\u{81}', '\u{201A}', '\u{192}', '\u{201E}', '\u{2026}', '\u{2020}', '\u{2021}',
    '\u{2C6}', '\u{2030}', '\u{160}', '\u{2039}', '\u{152}', '\u{8D}', '\u{17D}', '\u{8F}',
    '\u{90}', '\u{2018}', '\u{2019}', '\u{201C}', '\u{201D}', '\u{2022}', '\u{2013}', '\u{2014}',
    '\u{2DC}', '\u{2122}', '\u{161}', '\u{203A}', '\u{153}', '\u{9D}', '\u{17E}', '\u{178}',
];

/// `Chr`: the character for a Windows-1252 code.
fn from_ansi(b: u8) -> char {
    match b {
        0x80..=0x9F => CP1252_HIGH[(b - 0x80) as usize],
        _ => b as char,
    }
}

/// `Asc`: the Windows-1252 code of a character (`?` if it has none).
fn to_ansi(c: char) -> u8 {
    if let Some(i) = CP1252_HIGH.iter().position(|&h| h == c) {
        return 0x80 + i as u8;
    }
    match c as u32 {
        n @ (0..=0x7F | 0xA0..=0xFF) => n as u8,
        _ => b'?',
    }
}

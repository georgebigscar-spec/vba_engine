//! `Format` and the `FormatXxx` functions (en-US locale).

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    Some(Value::Long(match name {
        "vbgeneraldate" | "vbfalse" => 0,
        "vblongdate" => 1,
        "vbshortdate" => 2,
        "vblongtime" => 3,
        "vbshorttime" => 4,
        "vbtrue" => -1,
        "vbusedefault" => -2,
        _ => return None,
    }))
}

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    _host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
        "format" | "format$" => {
            let fmt = opt(args, 1)
                .map(|v| v.to_string_value())
                .transpose()?
                .unwrap_or_default();
            let fdow = opt_int(args, 2, 1)?;
            let fwoy = opt_int(args, 3, 1)?;
            if !(0..=7).contains(&fdow) || !(0..=3).contains(&fwoy) {
                return Err(VbaError::invalid_call());
            }
            // 0 means "system setting": Sunday and Jan 1 in en-US.
            match format_with(arg(args, 0)?, &fmt, fdow.max(1), fwoy.max(1))? {
                Some(s) => Value::String(s),
                None if lname == "format$" => return Err(VbaError::invalid_use_of_null()),
                None => Value::Null,
            }
        }
        "formatnumber" | "formatcurrency" | "formatpercent" => {
            let mut x = n(args, 0)?;
            let digits = opt_int(args, 1, -1)?;
            if digits < -1 {
                return Err(VbaError::invalid_call());
            }
            let (leading, parens, group) = (
                tristate(args, 2, true)?,
                tristate(args, 3, lname == "formatcurrency")?,
                tristate(args, 4, true)?,
            );
            let (prefix, suffix) = match lname {
                "formatcurrency" => ("$", ""),
                "formatpercent" => {
                    x *= 100.0;
                    ("", "%")
                }
                _ => ("", ""),
            };
            let digits = if digits == -1 { 2 } else { digits as usize };
            let (int, frac) = round_decimal(x, digits);
            let mut int = if group { group_digits(&int) } else { int };
            if int.is_empty() && leading {
                int.push('0');
            }
            let point = if digits > 0 { "." } else { "" };
            let body = format!("{prefix}{int}{point}{frac}{suffix}");
            let zero = !format!("{int}{frac}")
                .bytes()
                .any(|b| matches!(b, b'1'..=b'9'));
            Value::String(match (x < 0.0 && !zero, parens) {
                (false, _) => body,
                (true, true) => format!("({body})"),
                (true, false) => format!("-{body}"),
            })
        }
        "formatdatetime" => {
            let d = arg(args, 0)?.to_date()?;
            Value::String(match opt_int(args, 1, 0)? {
                0 => general_date(d),
                1 => format_date_mask(d, LONG_DATE, 1, 1),
                2 => format_date_mask(d, SHORT_DATE, 1, 1),
                3 => format_date_mask(d, LONG_TIME, 1, 1),
                4 => format_date_mask(d, "hh:nn", 1, 1),
                _ => return Err(VbaError::invalid_call()),
            })
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

fn opt_int(args: &Args, i: usize, default: i64) -> VbaResult<i64> {
    Ok(opt(args, i)
        .map(|v| v.to_i64())
        .transpose()?
        .unwrap_or(default))
}

/// A `VbTriState` argument: `vbTrue`, `vbFalse` or `vbUseDefault`.
fn tristate(args: &Args, i: usize, default: bool) -> VbaResult<bool> {
    match opt_int(args, i, -2)? {
        -2 => Ok(default),
        -1 => Ok(true),
        0 => Ok(false),
        _ => Err(VbaError::invalid_call()),
    }
}

/// `General Date`: like `CStr`, but midnight of day 0 shows as a time.
fn general_date(d: f64) -> String {
    if d == 0.0 {
        "12:00:00 AM".into()
    } else {
        date::format_date(d)
    }
}

const SHORT_DATE: &str = "m/d/yyyy";
const LONG_DATE: &str = "dddd, mmmm d, yyyy";
const LONG_TIME: &str = "h:nn:ss AM/PM";

/// `Format(v, fmt)` with default first day of week and first week of year.
/// A Null result comes back as an empty string.
pub fn format_value(v: &Value, fmt: &str) -> VbaResult<String> {
    Ok(format_with(v, fmt, 1, 1)?.unwrap_or_default())
}

/// The full `Format` engine; `None` stands for a Null result.
fn format_with(v: &Value, fmt: &str, fdow: i64, fwoy: i64) -> VbaResult<Option<String>> {
    if fmt.is_empty() {
        return Ok(match v {
            Value::Null => None,
            Value::Date(d) => Some(general_date(*d)),
            v => Some(v.to_string_value()?),
        });
    }
    if let Some(r) = named_format(v, fmt)? {
        return Ok(r);
    }
    let sections = split_sections(fmt);
    let plain: String = sections
        .iter()
        .flat_map(|s| unquoted(s))
        .map(|c| c.to_ascii_lowercase())
        .collect();
    let section = |i: usize| {
        sections
            .get(i)
            .map(String::as_str)
            .filter(|s| !s.is_empty())
    };
    let has = |set: &str| plain.contains(|c| set.contains(c));
    let datelike = has("dmyhnswqtc") || plain.contains("a/p");
    if has("@&") || (!has("0#") && !datelike && has("<>!")) {
        return Ok(match v {
            Value::Null => section(1).map(|m| format_string_mask("", m)),
            Value::String(s) if s.is_empty() => {
                Some(format_string_mask("", section(1).unwrap_or(&sections[0])))
            }
            Value::Date(d) => Some(format_string_mask(&general_date(*d), &sections[0])),
            v => Some(format_string_mask(&v.to_string_value()?, &sections[0])),
        });
    }
    if !has("0#") && datelike {
        if matches!(v, Value::Null) {
            return Ok(None);
        }
        return Ok(Some(match as_date(v) {
            Some(d) => format_date_mask(d, &sections[0], fdow, fwoy),
            None => v.to_string_value()?,
        }));
    }
    if matches!(v, Value::Null) {
        return Ok(section(3).map(|m| format_number_mask(0.0, m)));
    }
    let Some(x) = as_number(v) else {
        return Ok(Some(v.to_string_value()?));
    };
    Ok(Some(if x < 0.0 {
        match section(1) {
            Some(m) => format_number_mask(-x, m),
            None => format!("-{}", format_number_mask(-x, &sections[0])),
        }
    } else if x == 0.0 {
        format_number_mask(0.0, section(2).unwrap_or(&sections[0]))
    } else {
        format_number_mask(x, &sections[0])
    }))
}

/// Named formats such as `Currency` or `Long Date`; `None` if `fmt` is not one.
fn named_format(v: &Value, fmt: &str) -> VbaResult<Option<Option<String>>> {
    let lf = fmt.to_ascii_lowercase();
    let numeric = |f: &dyn Fn(f64) -> String| -> VbaResult<Option<Option<String>>> {
        Ok(Some(match v {
            Value::Null => None,
            v => Some(match as_number(v) {
                Some(x) => f(x),
                None => v.to_string_value()?,
            }),
        }))
    };
    let dated = |mask: &str| -> VbaResult<Option<Option<String>>> {
        Ok(Some(match v {
            Value::Null => None,
            v => Some(match as_date(v) {
                Some(d) if mask.is_empty() => general_date(d),
                Some(d) => format_date_mask(d, mask, 1, 1),
                None => v.to_string_value()?,
            }),
        }))
    };
    let word = |yes: &'static str, no: &'static str| {
        move |x: f64| if x != 0.0 { yes } else { no }.to_string()
    };
    let mask =
        |m: &'static str| move |x: f64| format_with(&Value::Double(x), m, 1, 1).unwrap().unwrap();
    match lf.as_str() {
        "general number" => numeric(&value::format_double),
        "currency" => numeric(&mask("$#,##0.00;($#,##0.00)")),
        "fixed" => numeric(&mask("0.00")),
        "standard" => numeric(&mask("#,##0.00")),
        "percent" => numeric(&mask("0.00%")),
        "scientific" => numeric(&mask("0.00E+00")),
        "yes/no" => numeric(&word("Yes", "No")),
        "true/false" => numeric(&word("True", "False")),
        "on/off" => numeric(&word("On", "Off")),
        "general date" => dated(""),
        "long date" => dated(LONG_DATE),
        "medium date" => dated("dd-mmm-yy"),
        "short date" => dated(SHORT_DATE),
        "long time" => dated(LONG_TIME),
        "medium time" => dated("hh:nn AM/PM"),
        "short time" => dated("hh:nn"),
        _ => Ok(None),
    }
}

/// A value as a number for numeric masks (numeric or date strings included).
fn as_number(v: &Value) -> Option<f64> {
    match v {
        Value::String(s) => parse_number(s).or_else(|| date::parse_date(s)),
        v => v.to_f64().ok(),
    }
}

/// A value as a date serial for date masks.
fn as_date(v: &Value) -> Option<f64> {
    match v {
        Value::String(s) => date::parse_date(s).or_else(|| parse_number(s)),
        v => v.to_f64().ok(),
    }
}

/// Splits a mask on `;` outside quotes and `\` escapes.
fn split_sections(fmt: &str) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut chars = fmt.chars();
    let mut quoted = false;
    while let Some(c) = chars.next() {
        let cur = out.last_mut().unwrap();
        match c {
            ';' if !quoted => out.push(String::new()),
            '"' => {
                quoted = !quoted;
                cur.push(c);
            }
            '\\' if !quoted => {
                cur.push(c);
                cur.extend(chars.next());
            }
            _ => cur.push(c),
        }
    }
    out
}

/// A mask piece: a literal or a format character.
enum Piece {
    Lit(String),
    Ch(char),
}

/// Splits a mask into literals (`"..."`, `\x`) and format characters.
fn pieces(mask: &str) -> Vec<Piece> {
    let mut out = Vec::new();
    let mut chars = mask.chars();
    while let Some(c) = chars.next() {
        match c {
            '"' => out.push(Piece::Lit(
                chars.by_ref().take_while(|&c| c != '"').collect(),
            )),
            '\\' => out.push(Piece::Lit(chars.next().into_iter().collect())),
            c => out.push(Piece::Ch(c)),
        }
    }
    out
}

/// The format characters of a mask, without its literals.
fn unquoted(mask: &str) -> Vec<char> {
    pieces(mask)
        .into_iter()
        .filter_map(|p| match p {
            Piece::Ch(c) => Some(c),
            Piece::Lit(_) => None,
        })
        .collect()
}

// ---- numbers ----------------------------------------------------------------

/// The 15 significant digits VBA works with: `|x| = 0.digits * 10^exp`.
fn digits15(x: f64) -> (Vec<u8>, i32) {
    let s = format!("{:.14e}", x.abs());
    let (m, e) = s.split_once('e').unwrap();
    let digits = m.bytes().filter(u8::is_ascii_digit).map(|b| b - b'0');
    (digits.collect(), e.parse::<i32>().unwrap() + 1)
}

/// Rounds `0.digits * 10^int_len` half away from zero to `decimals` places;
/// returns the integer digits (no leading zeros) and exactly `decimals`
/// fraction digits.
fn round_digits(mut digits: Vec<u8>, mut int_len: i32, decimals: usize) -> (String, String) {
    if int_len < 0 {
        digits.splice(
            0..0,
            std::iter::repeat_n(0, int_len.unsigned_abs() as usize),
        );
        int_len = 0;
    }
    let keep = int_len as usize + decimals;
    if keep < digits.len() {
        let up = digits[keep] >= 5;
        digits.truncate(keep);
        if up {
            let mut i = keep;
            loop {
                if i == 0 {
                    digits.insert(0, 1);
                    int_len += 1;
                    break;
                }
                i -= 1;
                if digits[i] == 9 {
                    digits[i] = 0;
                } else {
                    digits[i] += 1;
                    break;
                }
            }
        }
    } else {
        digits.resize(keep, 0);
    }
    let text = |d: &[u8]| d.iter().map(|d| (b'0' + d) as char).collect::<String>();
    let (int, frac) = digits.split_at(int_len as usize);
    (text(int).trim_start_matches('0').to_string(), text(frac))
}

fn round_decimal(x: f64, decimals: usize) -> (String, String) {
    let (digits, exp) = digits15(x);
    round_digits(digits, exp, decimals)
}

/// Inserts thousands separators into a digit string.
fn group_digits(int: &str) -> String {
    let n = int.len();
    let mut out = String::new();
    for (i, c) in int.chars().enumerate() {
        out.push(c);
        if i + 1 < n && (n - i - 1).is_multiple_of(3) {
            out.push(',');
        }
    }
    out
}

enum NumTok {
    Lit(String),
    Digit(char),
    Point,
    Comma,
    Exp { upper: bool, plus: bool },
}

/// Formats a non-negative number with one section of a numeric mask.
fn format_number_mask(mut x: f64, mask: &str) -> String {
    let mut toks = Vec::new();
    let mut ps = pieces(mask).into_iter().peekable();
    let (mut seen_point, mut seen_exp) = (false, false);
    while let Some(p) = ps.next() {
        toks.push(match p {
            Piece::Lit(s) => NumTok::Lit(s),
            Piece::Ch(c @ ('0' | '#')) => NumTok::Digit(c),
            Piece::Ch('.') if !seen_point && !seen_exp => {
                seen_point = true;
                NumTok::Point
            }
            Piece::Ch(',') if !seen_exp => NumTok::Comma,
            Piece::Ch('%') => {
                x *= 100.0;
                NumTok::Lit("%".into())
            }
            Piece::Ch(c @ ('E' | 'e'))
                if !seen_exp && matches!(ps.peek(), Some(Piece::Ch('+' | '-'))) =>
            {
                seen_exp = true;
                let Some(Piece::Ch(sign)) = ps.next() else {
                    unreachable!()
                };
                NumTok::Exp {
                    upper: c == 'E',
                    plus: sign == '+',
                }
            }
            Piece::Ch(c) => NumTok::Lit(c.to_string()),
        });
    }
    // Placeholders of the integer part, the fraction and the exponent.
    let end_int = toks
        .iter()
        .position(|t| matches!(t, NumTok::Point | NumTok::Exp { .. }))
        .unwrap_or(toks.len());
    let int_ph: Vec<char> = toks[..end_int]
        .iter()
        .filter_map(|t| {
            if let NumTok::Digit(c) = t {
                Some(*c)
            } else {
                None
            }
        })
        .collect();
    let exp_at = toks.iter().position(|t| matches!(t, NumTok::Exp { .. }));
    let frac_ph: Vec<char> = toks[end_int..exp_at.unwrap_or(toks.len())]
        .iter()
        .filter_map(|t| {
            if let NumTok::Digit(c) = t {
                Some(*c)
            } else {
                None
            }
        })
        .collect();
    let exp_zeros = exp_at.map_or(0, |i| {
        toks[i + 1..]
            .iter()
            .filter(|t| matches!(t, NumTok::Digit('0')))
            .count()
    });
    // Commas between integer placeholders group; commas right after the last
    // one scale by 1000 each.
    let last_int = toks[..end_int]
        .iter()
        .rposition(|t| matches!(t, NumTok::Digit(_)));
    let first_int = toks[..end_int]
        .iter()
        .position(|t| matches!(t, NumTok::Digit(_)));
    let mut grouping = false;
    let mut scale = 0;
    if let (Some(first), Some(last)) = (first_int, last_int) {
        grouping = toks[first..last].iter().any(|t| matches!(t, NumTok::Comma));
        scale = toks[last + 1..]
            .iter()
            .take_while(|t| matches!(t, NumTok::Comma))
            .count();
    }
    x /= 1000f64.powi(scale as i32);

    let (mut int, mut frac, exponent) = if exp_at.is_some() {
        let k = int_ph.len().max(1);
        let (digits, e) = digits15(x);
        let mut exponent = if x == 0.0 { 0 } else { e - k as i32 };
        let (mut int, mut frac) = round_digits(digits, k as i32, frac_ph.len());
        if int.len() > k {
            // Rounding carried into a new digit (9.99 -> 10.0).
            frac = format!("{}{frac}", &int[k..])[..frac_ph.len()].to_string();
            int.truncate(k);
            exponent += 1;
        }
        (int, frac, exponent)
    } else {
        let (int, frac) = round_decimal(x, frac_ph.len());
        (int, frac, 0)
    };
    let min_int = int_ph
        .iter()
        .position(|&c| c == '0')
        .map_or(0, |i| int_ph.len() - i);
    while int.len() < min_int {
        int.insert(0, '0');
    }
    while frac.ends_with('0') && frac_ph.get(frac.len() - 1) == Some(&'#') {
        frac.pop();
    }

    let int: Vec<char> = int.chars().collect();
    let frac: Vec<char> = frac.chars().collect();
    let offset = int_ph.len() as isize - int.len() as isize;
    let mut out = String::new();
    let push_int = |out: &mut String, from: usize, to: usize| {
        for (i, c) in int.iter().enumerate().take(to).skip(from) {
            out.push(*c);
            let rest = int.len() - i - 1;
            if grouping && rest > 0 && rest.is_multiple_of(3) {
                out.push(',');
            }
        }
    };
    let (mut ip, mut fp) = (0isize, 0usize);
    let mut in_exp = false;
    for (ti, t) in toks.iter().enumerate() {
        match t {
            NumTok::Lit(s) => out.push_str(s),
            NumTok::Comma => {}
            NumTok::Digit(_) if in_exp => {}
            NumTok::Digit(_) if ti < end_int => {
                let i = ip - offset;
                if ip == 0 && i >= 0 {
                    push_int(&mut out, 0, i as usize + 1);
                } else if i >= 0 {
                    push_int(&mut out, i as usize, i as usize + 1);
                }
                ip += 1;
            }
            NumTok::Digit(_) => {
                if let Some(c) = frac.get(fp) {
                    out.push(*c);
                }
                fp += 1;
            }
            NumTok::Point => {
                if int_ph.is_empty() {
                    push_int(&mut out, 0, int.len());
                }
                out.push('.');
            }
            NumTok::Exp { upper, plus } => {
                in_exp = true;
                out.push(if *upper { 'E' } else { 'e' });
                if exponent < 0 {
                    out.push('-');
                } else if *plus {
                    out.push('+');
                }
                out.push_str(&format!("{:0exp_zeros$}", exponent.abs()));
            }
        }
    }
    out
}

// ---- strings ----------------------------------------------------------------

/// Formats a string with one section of an `@ & ! < >` mask.
fn format_string_mask(s: &str, mask: &str) -> String {
    let ps = pieces(mask);
    let has = |ch: char| ps.iter().any(|p| matches!(p, Piece::Ch(c) if *c == ch));
    let s = if has('<') {
        s.to_lowercase()
    } else if has('>') {
        s.to_uppercase()
    } else {
        s.to_string()
    };
    let chars: Vec<char> = s.chars().collect();
    let n = ps
        .iter()
        .filter(|p| matches!(p, Piece::Ch('@' | '&')))
        .count();
    let left_to_right = has('!');
    let mut out = String::new();
    let mut j = 0isize;
    for p in &ps {
        match p {
            Piece::Lit(l) => out.push_str(l),
            Piece::Ch(c @ ('@' | '&')) => {
                let i = if left_to_right {
                    j
                } else {
                    j - (n as isize - chars.len() as isize)
                };
                let range = match (left_to_right, j) {
                    (false, 0) => 0..(i + 1).max(0) as usize,
                    (true, _) if j as usize == n - 1 => i as usize..chars.len().max(i as usize),
                    _ => i.max(0) as usize..(i + 1).max(0) as usize,
                };
                if i < 0 || i as usize >= chars.len() {
                    if *c == '@' {
                        out.push(' ');
                    }
                } else {
                    out.extend(&chars[range]);
                }
                j += 1;
            }
            Piece::Ch('<' | '>' | '!') => {}
            Piece::Ch(c) => out.push(*c),
        }
    }
    if n == 0 {
        out.push_str(&s);
    }
    out
}

// ---- dates ------------------------------------------------------------------

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

/// Day of the week (1-based) counted from `fdow` (1 = Sunday).
fn weekday(day: i64, fdow: i64) -> i64 {
    (day + 6 - (fdow - 1)).rem_euclid(7) + 1
}

/// Week of the year (`ww`) for the given first day of week / first week of year.
fn week_of_year(day: i64, fdow: i64, fwoy: i64) -> i64 {
    let jan1 = date::serial(date::ymd(day as f64).0, 1, 1) as i64;
    let lead = weekday(jan1, fdow) - 1;
    let mut w = (day - jan1 + lead) / 7 + 1;
    let partial_first = match fwoy {
        2 => lead > 3,
        3 => lead > 0,
        _ => false,
    };
    if partial_first {
        w -= 1;
        if w == 0 {
            return week_of_year(jan1 - 1, fdow, fwoy);
        }
    }
    w
}

enum DateTok {
    Lit(String),
    Field(char, usize),
    /// `AM/PM` or `A/P` as written in the mask (its case is kept), or `AMPM`.
    AmPm(String),
}

/// Formats a date serial with a date/time mask.
fn format_date_mask(d: f64, mask: &str, fdow: i64, fwoy: i64) -> String {
    let mut toks = Vec::new();
    let ps = pieces(mask);
    let mut i = 0;
    while i < ps.len() {
        let c = match &ps[i] {
            Piece::Lit(s) => {
                toks.push(DateTok::Lit(s.clone()));
                i += 1;
                continue;
            }
            Piece::Ch(c) => *c,
        };
        let ahead: String = ps[i..]
            .iter()
            .take(5)
            .map_while(|p| match p {
                Piece::Ch(c) => Some(*c),
                Piece::Lit(_) => None,
            })
            .collect();
        let la = ahead.to_ascii_lowercase();
        let marker = ["am/pm", "ampm", "a/p"]
            .into_iter()
            .find(|m| la.starts_with(m));
        if let Some(m) = marker {
            toks.push(DateTok::AmPm(ahead[..m.len()].to_string()));
            i += m.len();
            continue;
        }
        let lc = c.to_ascii_lowercase();
        if "dwmqyhnstc".contains(lc) {
            let run = ps[i..]
                .iter()
                .take_while(|p| matches!(p, Piece::Ch(x) if x.to_ascii_lowercase() == lc))
                .count();
            toks.push(DateTok::Field(lc, run));
            i += run;
        } else {
            toks.push(DateTok::Lit(c.to_string()));
            i += 1;
        }
    }

    let twelve = toks.iter().any(|t| matches!(t, DateTok::AmPm(_)));
    let (y, m, dd) = date::ymd(d);
    let (h, mi, sec) = date::hms(d);
    let day = d.floor() as i64;
    let hour = if twelve && h % 12 == 0 {
        12
    } else if twelve {
        h % 12
    } else {
        h
    };
    let field_at = |k: usize| match toks.get(k) {
        Some(DateTok::Field(c, _)) => Some(*c),
        _ => None,
    };
    let mut out = String::new();
    for (k, t) in toks.iter().enumerate() {
        match t {
            DateTok::Lit(s) => out.push_str(s),
            DateTok::AmPm(m) => {
                let pm = h >= 12;
                out.push_str(match (m.len(), pm) {
                    (4, false) => "AM",
                    (4, true) => "PM",
                    (5, false) => &m[..2],
                    (5, true) => &m[3..],
                    (_, false) => &m[..1],
                    (_, true) => &m[2..],
                });
            }
            &DateTok::Field(c, run) => {
                let pad = |v: i64, run: usize| format!("{v:0w$}", w = run.min(2));
                match c {
                    'd' => out.push_str(&match run {
                        1 | 2 => pad(dd, run),
                        3 => DAYS[weekday(day, 1) as usize - 1][..3].to_string(),
                        4 => DAYS[weekday(day, 1) as usize - 1].to_string(),
                        5 => format_date_mask(d, SHORT_DATE, fdow, fwoy),
                        _ => format_date_mask(d, LONG_DATE, fdow, fwoy),
                    }),
                    'w' if run == 1 => out.push_str(&weekday(day, fdow).to_string()),
                    'w' => out.push_str(&week_of_year(day, fdow, fwoy).to_string()),
                    'm' if run <= 2 && {
                        // `m` after an hour or before seconds means minutes.
                        let prev = toks[..k]
                            .iter()
                            .rposition(|t| !matches!(t, DateTok::Lit(_)));
                        let next = toks[k + 1..]
                            .iter()
                            .position(|t| !matches!(t, DateTok::Lit(_)))
                            .map(|p| p + k + 1);
                        prev.and_then(field_at) == Some('h') || next.and_then(field_at) == Some('s')
                    } =>
                    {
                        out.push_str(&pad(mi, run))
                    }
                    'm' => out.push_str(&match run {
                        1 | 2 => pad(m, run),
                        3 => MONTHS[m as usize - 1][..3].to_string(),
                        _ => MONTHS[m as usize - 1].to_string(),
                    }),
                    'q' => out.push_str(&((m - 1) / 3 + 1).to_string().repeat(run)),
                    'y' => {
                        let doy = day - date::serial(y, 1, 1) as i64 + 1;
                        out.push_str(&match run {
                            1 => doy.to_string(),
                            2 => format!("{:02}", y.rem_euclid(100)),
                            3 => format!("{:02}{doy}", y.rem_euclid(100)),
                            _ => format!("{y:04}"),
                        })
                    }
                    'h' => out.push_str(&pad(hour, run)),
                    'n' => out.push_str(&pad(mi, run)),
                    's' => out.push_str(&pad(sec, run)),
                    't' if run == 5 => out.push_str(&format_date_mask(d, LONG_TIME, fdow, fwoy)),
                    't' => out.push_str(&"t".repeat(run)),
                    _ => out.push_str(&general_date(d).repeat(run)),
                }
            }
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn f(v: impl Into<Value>, fmt: &str) -> String {
        format_value(&v.into(), fmt).unwrap()
    }

    #[test]
    fn numeric_masks() {
        assert_eq!(f(123456789, "000-00-0000"), "123-45-6789");
        assert_eq!(f(1.5, ".00"), "1.50");
        assert_eq!(f(9.999, "0.00"), "10.00");
        assert_eq!(f(9.99, "0.0E+00"), "1.0E+01");
        assert_eq!(f(0, "0.00E+00"), "0.00E+00");
    }

    #[test]
    fn weeks() {
        let d = Value::Date(date::serial(2021, 1, 1));
        assert_eq!(format_with(&d, "ww", 2, 2).unwrap().unwrap(), "53");
        assert_eq!(format_with(&d, "ww", 1, 1).unwrap().unwrap(), "1");
    }
}

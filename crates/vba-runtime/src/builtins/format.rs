//! `Format` and the `FormatXxx` functions.

use super::*;

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
            Value::String(format_value(arg(args, 0)?, &fmt)?)
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
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

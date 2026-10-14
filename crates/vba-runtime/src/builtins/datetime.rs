//! `VBA.DateTime`.

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    Some(Value::Integer(match name {
        "vbusesystemdayofweek" | "vbusesystem" => 0,
        "vbtuesday" => 3,
        "vbwednesday" => 4,
        "vbthursday" => 5,
        "vbfriday" => 6,
        "vbsaturday" => 7,
        "vbfirstjan1" => 1,
        "vbfirstfourdays" => 2,
        "vbfirstfullweek" => 3,
        _ => return None,
    }))
}

/// Optional `firstdayofweek` argument (`vbUseSystemDayOfWeek` means Sunday).
fn first_day(args: &Args, i: usize) -> VbaResult<i64> {
    match opt(args, i).map(|v| v.to_i64()).transpose()?.unwrap_or(1) {
        0 => Ok(1),
        d @ 1..=7 => Ok(d),
        _ => Err(VbaError::invalid_call()),
    }
}

/// Weekday of `d` counted from `first` (1 = `first` itself).
fn weekday_from(d: f64, first: i64) -> i64 {
    (date::weekday(d) - first).rem_euclid(7) + 1
}

/// Day number of the first day of the week containing `day`.
fn week_start(day: i64, first: i64) -> i64 {
    day - (weekday_from(day as f64, first) - 1)
}

/// `DatePart("ww", ...)`: week of the year under `firstweekofyear`
/// (1 = week of Jan 1, 2 = first week with four days, 3 = first full week).
fn week_of_year(d: f64, first: i64, fwoy: i64) -> i64 {
    let day = date::day_number(d);
    let week1 = |y: i64| {
        let jan1 = date::serial(y, 1, 1) as i64;
        let ws = week_start(jan1, first);
        match fwoy {
            2 if 7 - (jan1 - ws) < 4 => ws + 7,
            3 if jan1 != ws => ws + 7,
            _ => ws,
        }
    };
    let y = date::ymd(d).0;
    let mut w1 = week1(y);
    if day < w1 {
        w1 = week1(y - 1);
    }
    (week_start(day, first) - w1).div_euclid(7) + 1
}

/// Date argument of a date function, or `None` when it is Null.
fn date_arg(args: &Args, i: usize) -> VbaResult<Option<f64>> {
    match arg(args, i)? {
        Value::Null => Ok(None),
        v => v.to_date().map(Some),
    }
}

/// An `Integer` argument; out-of-range values overflow (error 6).
fn int16(args: &Args, i: usize) -> VbaResult<i64> {
    let v = int(args, i)?;
    if !(i16::MIN as i64..=i16::MAX as i64).contains(&v) {
        return Err(VbaError::overflow());
    }
    Ok(v)
}

/// Checks that a computed date lies in VBA's range (years 100 to 9999).
fn in_range(d: f64) -> VbaResult<f64> {
    if (100..=9999).contains(&date::ymd(d).0) {
        Ok(d)
    } else {
        Err(VbaError::invalid_call())
    }
}

/// `DateSerial` after the two-digit year window; months and days roll over.
fn date_serial(y: i64, m: i64, d: i64) -> f64 {
    let (y, m) = (y + (m - 1).div_euclid(12), (m - 1).rem_euclid(12) + 1);
    date::serial(y, m, 1) + (d - 1) as f64
}

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    // Year/Month/Day/...: an Integer part of a date, Null passes through.
    let part = |f: fn(f64) -> i64| -> VbaResult<Value> {
        Ok(date_arg(args, 0)?.map_or(Value::Null, |d| Value::Integer(f(d) as i16)))
    };
    Ok(match lname {
        // ---- date / time ----
        "now" => Value::Date(host.now()),
        "date" => Value::Date(host.now().floor()),
        "time" => Value::Date(host.now().fract()),
        "date$" => {
            let (y, m, d) = date::ymd(host.now());
            Value::String(format!("{m:02}-{d:02}-{y:04}"))
        }
        "time$" => {
            let (h, m, s) = date::hms(host.now());
            Value::String(format!("{h:02}:{m:02}:{s:02}"))
        }
        "timer" => Value::Single((host.now().fract() * 86_400.0) as f32),
        "year" => part(|d| date::ymd(d).0)?,
        "month" => part(|d| date::ymd(d).1)?,
        "day" => part(|d| date::ymd(d).2)?,
        "hour" => part(|d| date::hms(d).0)?,
        "minute" => part(|d| date::hms(d).1)?,
        "second" => part(|d| date::hms(d).2)?,
        "weekday" => match date_arg(args, 0)? {
            None => Value::Null,
            Some(d) => Value::Integer(weekday_from(d, first_day(args, 1)?) as i16),
        },
        "weekdayname" | "monthname" => {
            let n = int(args, 0)?;
            let abbrev = opt(args, 1)
                .map(|v| v.to_bool())
                .transpose()?
                .unwrap_or(false);
            let full = if lname == "monthname" {
                (1..=12).contains(&n).then(|| date::MONTHS[n as usize - 1])
            } else {
                let first = first_day(args, 2)?;
                (1..=7)
                    .contains(&n)
                    .then(|| date::WEEKDAYS[((n + first - 2) % 7) as usize])
            }
            .ok_or_else(VbaError::invalid_call)?;
            Value::str(if abbrev { &full[..3] } else { full })
        }
        "datevalue" => match date_arg(args, 0)? {
            None => Value::Null,
            Some(d) => Value::Date(date::day_number(d) as f64),
        },
        "timevalue" => match date_arg(args, 0)? {
            None => Value::Null,
            Some(d) => Value::Date(date::day_secs(d) as f64 / 86_400.0),
        },
        "dateserial" => {
            let (y, m, d) = (int16(args, 0)?, int16(args, 1)?, int16(args, 2)?);
            Value::Date(in_range(date_serial(date::window_year(y), m, d))?)
        }
        "timeserial" => {
            let (h, m, s) = (int16(args, 0)?, int16(args, 1)?, int16(args, 2)?);
            Value::Date((h * 3600 + m * 60 + s) as f64 / 86_400.0)
        }
        "dateadd" => {
            let unit = s(args, 0)?.to_ascii_lowercase();
            let k = n(args, 1)?.trunc();
            let Some(d) = date_arg(args, 2)? else {
                return Ok(Value::Null);
            };
            let secs = |per: f64| {
                let secs = date::linear_secs(d) as f64 + k * per;
                let day = (secs / 86_400.0).floor();
                date::combine(day as i64, (secs - day * 86_400.0) / 86_400.0)
            };
            let months = |per: i64| {
                let (y, m, dd) = date::ymd(d);
                let total = y * 12 + (m - 1) + k as i64 * per;
                let (ny, nm) = (total.div_euclid(12), total.rem_euclid(12) + 1);
                if !(100..=9999).contains(&ny) {
                    return Err(VbaError::invalid_call());
                }
                let day = date::serial(ny, nm, dd.min(date::days_in_month(ny, nm)));
                Ok(date::combine(day as i64, d.fract().abs()))
            };
            Value::Date(in_range(match unit.as_str() {
                "d" | "y" | "w" => secs(86_400.0),
                "ww" => secs(7.0 * 86_400.0),
                "h" => secs(3600.0),
                "n" => secs(60.0),
                "s" => secs(1.0),
                "m" => months(1)?,
                "q" => months(3)?,
                "yyyy" => months(12)?,
                _ => return Err(VbaError::invalid_call()),
            })?)
        }
        "datediff" => {
            let unit = s(args, 0)?.to_ascii_lowercase();
            let (Some(a), Some(b)) = (date_arg(args, 1)?, date_arg(args, 2)?) else {
                return Ok(Value::Null);
            };
            let (da, db) = (date::day_number(a), date::day_number(b));
            let (sa, sb) = (date::linear_secs(a), date::linear_secs(b));
            let ((ya, ma, _), (yb, mb, _)) = (date::ymd(a), date::ymd(b));
            let r = match unit.as_str() {
                "d" | "y" => db - da,
                "w" => (db - da) / 7,
                // Number of `firstdayofweek` days after date1 up to date2.
                "ww" => {
                    let first = first_day(args, 3)?;
                    (week_start(db, first) - week_start(da, first)) / 7
                }
                "h" => sb.div_euclid(3600) - sa.div_euclid(3600),
                "n" => sb.div_euclid(60) - sa.div_euclid(60),
                "s" => sb - sa,
                "m" => (yb * 12 + mb) - (ya * 12 + ma),
                "q" => (yb * 4 + (mb - 1) / 3) - (ya * 4 + (ma - 1) / 3),
                "yyyy" => yb - ya,
                _ => return Err(VbaError::invalid_call()),
            };
            Value::Long(i32::try_from(r).map_err(|_| VbaError::overflow())?)
        }
        "datepart" => {
            let unit = s(args, 0)?.to_ascii_lowercase();
            let Some(d) = date_arg(args, 1)? else {
                return Ok(Value::Null);
            };
            let first = first_day(args, 2)?;
            let fwoy = match opt(args, 3).map(|v| v.to_i64()).transpose()?.unwrap_or(1) {
                0 => 1,
                f @ 1..=3 => f,
                _ => return Err(VbaError::invalid_call()),
            };
            let (y, m, dd) = date::ymd(d);
            let (h, mi, sec) = date::hms(d);
            Value::Integer(match unit.as_str() {
                "yyyy" => y,
                "q" => (m - 1) / 3 + 1,
                "m" => m,
                "y" => date::day_number(d) - date::serial(y, 1, 1) as i64 + 1,
                "d" => dd,
                "w" => weekday_from(d, first),
                "ww" => week_of_year(d, first, fwoy),
                "h" => h,
                "n" => mi,
                "s" => sec,
                _ => return Err(VbaError::invalid_call()),
            } as i16)
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

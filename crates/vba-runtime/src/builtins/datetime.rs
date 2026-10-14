//! `VBA.DateTime`.

use super::*;

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
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
                        - date::serial(ny, nm, 1)) as i64;
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
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

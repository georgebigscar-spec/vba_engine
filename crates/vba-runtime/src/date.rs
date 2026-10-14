//! OLE Automation dates (days since 1899-12-30, fraction = time of day).

fn days_from_civil(y: i64, m: i64, d: i64) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = y.div_euclid(400);
    let yoe = y - era * 400;
    let mp = (m + 9) % 12;
    let doy = (153 * mp + 2) / 5 + d - 1;
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy;
    era * 146_097 + doe
}

fn civil_from_days(z: i64) -> (i64, i64, i64) {
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let d = doy - (153 * mp + 2) / 5 + 1;
    let m = if mp < 10 { mp + 3 } else { mp - 9 };
    (yoe + era * 400 + i64::from(m <= 2), m, d)
}

pub fn serial(y: i64, m: i64, d: i64) -> f64 {
    (days_from_civil(y, m, d) - days_from_civil(1899, 12, 30)) as f64
}

/// Calendar date of an OLE date. Negative dates count whole days towards the
/// past and keep a positive time of day (`-1.5` is 12/29/1899 12:00 PM).
pub fn ymd(serial: f64) -> (i64, i64, i64) {
    civil_from_days(serial.trunc() as i64 + days_from_civil(1899, 12, 30))
}

/// Seconds since midnight of an OLE date's time of day.
pub fn day_secs(serial: f64) -> i64 {
    ((serial.fract().abs() * 86_400.0).round() as i64).min(86_399)
}

pub fn hms(serial: f64) -> (i64, i64, i64) {
    let secs = day_secs(serial);
    (secs / 3600, secs / 60 % 60, secs % 60)
}

/// Days since 1899-12-30 of the date part (the day `ymd` reports).
pub fn day_number(serial: f64) -> i64 {
    serial.trunc() as i64
}

/// Builds an OLE date from a day number and a time of day in `[0, 1)`.
pub fn combine(day: i64, time: f64) -> f64 {
    if day < 0 {
        day as f64 - time
    } else {
        day as f64 + time
    }
}

/// Maps an OLE date onto a continuous time line (in days), so that
/// arithmetic and comparisons work across the negative-date discontinuity.
pub fn to_linear(serial: f64) -> f64 {
    serial.trunc() + serial.fract().abs()
}

/// Inverse of [`to_linear`].
pub fn from_linear(x: f64) -> f64 {
    let day = x.floor();
    combine(day as i64, x - day)
}

/// Whole seconds on the continuous time line.
pub fn linear_secs(serial: f64) -> i64 {
    day_number(serial) * 86_400 + day_secs(serial)
}

/// Sunday = 1 .. Saturday = 7.
pub fn weekday(serial: f64) -> i64 {
    // 1899-12-30 was a Saturday.
    (day_number(serial) + 6).rem_euclid(7) + 1
}

pub fn days_in_month(y: i64, m: i64) -> i64 {
    match m {
        2 if y % 4 == 0 && (y % 100 != 0 || y % 400 == 0) => 29,
        2 => 28,
        4 | 6 | 9 | 11 => 30,
        _ => 31,
    }
}

/// VBA's two-digit year window: 0-29 are 2000-2029, 30-99 are 1930-1999.
pub fn window_year(y: i64) -> i64 {
    match y {
        0..=29 => 2000 + y,
        30..=99 => 1900 + y,
        _ => y,
    }
}

pub const MONTHS: [&str; 12] = [
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

pub const WEEKDAYS: [&str; 7] = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
];

/// Month number of a full or three-letter English month name.
fn month_from_name(s: &str) -> Option<i64> {
    let s = s.to_ascii_lowercase();
    if s.len() < 3 {
        return None;
    }
    MONTHS
        .iter()
        .position(|m| {
            let m = m.to_ascii_lowercase();
            m == s || (s.len() == 3 && m.starts_with(&s))
        })
        .map(|i| i as i64 + 1)
}

/// Parses `h[:mm[:ss]]` with an optional AM/PM marker into a day fraction.
fn parse_time(clock: &str, ampm: Option<&str>) -> Option<f64> {
    let nums: Vec<i64> = clock
        .split(':')
        .map(|p| {
            (!p.is_empty() && p.len() <= 2 && p.bytes().all(|b| b.is_ascii_digit()))
                .then(|| p.parse().ok())
                .flatten()
        })
        .collect::<Option<_>>()?;
    if nums.len() > 3 {
        return None;
    }
    let (mut h, mi, se) = (
        nums[0],
        *nums.get(1).unwrap_or(&0),
        *nums.get(2).unwrap_or(&0),
    );
    if mi > 59 || se > 59 {
        return None;
    }
    match ampm.map(|s| s.to_ascii_uppercase()).as_deref() {
        None if h > 23 => return None,
        None => {}
        Some(_) if h > 12 => return None,
        Some("AM") if h == 12 => h = 0,
        Some("PM") if h < 12 => h += 12,
        Some("AM" | "PM") => {}
        Some(_) => return None,
    }
    Some((h * 3600 + mi * 60 + se) as f64 / 86_400.0)
}

/// Parses a date written as numbers and/or a month name, already split into
/// pieces. `dotted` marks `m.d.yy`, which VBA only reads with a short year.
fn parse_day(pieces: &[&str], dotted: bool) -> Option<i64> {
    if pieces.len() != 3 {
        return None;
    }
    let num = |p: &str| -> Option<(i64, usize)> {
        (p.len() <= 4 && p.bytes().all(|b| b.is_ascii_digit()))
            .then(|| p.parse().ok().map(|n| (n, p.len())))
            .flatten()
    };
    let named = pieces.iter().position(|p| month_from_name(p).is_some());
    let (y, m, d) = if let Some(i) = named {
        let m = month_from_name(pieces[i])?;
        let rest: Vec<(i64, usize)> = pieces
            .iter()
            .enumerate()
            .filter(|&(j, _)| j != i)
            .map(|(_, p)| num(p))
            .collect::<Option<_>>()?;
        // `Feb 12, 1969`, `12-Feb-69`, `1969 Feb 12`.
        let (d, y) = if rest[0].1 > 2 {
            (rest[1], rest[0])
        } else {
            (rest[0], rest[1])
        };
        if d.1 > 2 {
            return None;
        }
        (if y.1 <= 2 { window_year(y.0) } else { y.0 }, m, d.0)
    } else {
        let n: Vec<(i64, usize)> = pieces.iter().map(|p| num(p)).collect::<Option<_>>()?;
        if n[0].1 > 2 {
            (n[0].0, n[1].0, n[2].0)
        } else {
            if dotted && n[2].1 > 2 {
                return None;
            }
            let y = if n[2].1 <= 2 {
                window_year(n[2].0)
            } else {
                n[2].0
            };
            // `13/12/2020` is read day-first when month-first is impossible.
            if n[0].0 > 12 && n[1].0 <= 12 {
                (y, n[1].0, n[0].0)
            } else {
                (y, n[0].0, n[1].0)
            }
        }
    };
    if !(100..=9999).contains(&y) || !(1..=12).contains(&m) || d < 1 || d > days_in_month(y, m) {
        return None;
    }
    Some(serial(y, m, d) as i64)
}

/// Parses an en-US date and/or time: `m/d/yyyy`, `m-d-yy`, `m.d.yy`,
/// `yyyy-mm-dd`, `February 12, 1969`, `Feb 12, 1969`, `12-Feb-69`, each
/// optionally followed by `h:mm[:ss] [AM|PM]`, or a bare time.
pub fn parse_date(s: &str) -> Option<f64> {
    let mut tokens: Vec<&str> = s
        .split(|c: char| c.is_whitespace() || c == ',')
        .filter(|t| !t.is_empty())
        .collect();
    // Split the time (`4:35:17 PM`, `4:35PM`) off the end.
    let mut time = None;
    if let Some(i) = tokens.iter().position(|t| t.contains(':')) {
        let mut clock = tokens[i];
        let mut ampm = tokens.get(i + 1).copied();
        let upper = clock.to_ascii_uppercase();
        if upper.ends_with("AM") || upper.ends_with("PM") {
            if ampm.is_some() {
                return None;
            }
            ampm = Some(&clock[clock.len() - 2..]);
            clock = &clock[..clock.len() - 2];
        }
        if i + 1 + usize::from(tokens.get(i + 1).is_some()) != tokens.len() {
            return None;
        }
        time = Some(parse_time(clock, ampm)?);
        tokens.truncate(i);
    }
    let day = match tokens.as_slice() {
        [] if time.is_some() => 0,
        [one] => {
            let pieces: Vec<&str> = one.split(['/', '-', '.']).collect();
            parse_day(&pieces, one.contains('.'))?
        }
        many => parse_day(many, false)?,
    };
    Some(combine(day, time.unwrap_or(0.0)))
}

/// Formats like VBA with a US locale: `1/15/2024`, `1/15/2024 1:30:00 PM`, `1:30:00 PM`.
/// The date part of day 0 (12/30/1899) is never shown.
pub fn format_date(serial: f64) -> String {
    let (y, m, d) = ymd(serial);
    let (h, mi, s) = hms(serial);
    let has_time = h != 0 || mi != 0 || s != 0;
    let time = || {
        let (h12, ampm) = match h {
            0 => (12, "AM"),
            1..=11 => (h, "AM"),
            12 => (12, "PM"),
            _ => (h - 12, "PM"),
        };
        format!("{h12}:{mi:02}:{s:02} {ampm}")
    };
    if day_number(serial) == 0 {
        time()
    } else if has_time {
        format!("{m}/{d}/{y} {}", time())
    } else {
        format!("{m}/{d}/{y}")
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn round_trips() {
        assert_eq!(serial(2024, 1, 15), 45306.0);
        assert_eq!(ymd(45306.0), (2024, 1, 15));
        assert_eq!(
            format_date(parse_date("2024-01-15 13:30").unwrap()),
            "1/15/2024 1:30:00 PM"
        );
    }
}

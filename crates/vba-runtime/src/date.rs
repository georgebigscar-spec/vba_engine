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

pub fn ymd(serial: f64) -> (i64, i64, i64) {
    civil_from_days(serial.floor() as i64 + days_from_civil(1899, 12, 30))
}

pub fn hms(serial: f64) -> (i64, i64, i64) {
    let secs = ((serial - serial.floor()) * 86_400.0).round() as i64;
    (secs / 3600, secs / 60 % 60, secs % 60)
}

/// Accepts `m/d/yyyy`, `yyyy-mm-dd`, optionally followed by `hh:mm[:ss] [AM|PM]`,
/// or a bare time.
pub fn parse_date(s: &str) -> Option<f64> {
    let s = s.trim();
    let mut parts = s.splitn(2, ' ');
    let first = parts.next()?;
    let (date, time) = if first.contains(':') {
        (None, Some(s))
    } else {
        (Some(first), parts.next())
    };
    let mut out = 0.0;
    if let Some(d) = date {
        let nums: Vec<i64> = d
            .split(['/', '-', '.'])
            .map(|p| p.parse().ok())
            .collect::<Option<_>>()?;
        if nums.len() != 3 {
            return None;
        }
        let (y, m, dd) = if nums[0] > 31 {
            (nums[0], nums[1], nums[2])
        } else {
            (nums[2], nums[0], nums[1])
        };
        let y = if y < 30 {
            2000 + y
        } else if y < 100 {
            1900 + y
        } else {
            y
        };
        if !(1..=12).contains(&m) || !(1..=31).contains(&dd) {
            return None;
        }
        out = serial(y, m, dd);
    }
    if let Some(t) = time {
        let t = t.trim();
        let upper = t.to_ascii_uppercase();
        let (clock, pm, am) = (
            upper.trim_end_matches("PM").trim_end_matches("AM").trim(),
            upper.ends_with("PM"),
            upper.ends_with("AM"),
        );
        let nums: Vec<i64> = clock
            .split(':')
            .map(|p| p.parse().ok())
            .collect::<Option<_>>()?;
        let mut h = *nums.first()?;
        if pm && h < 12 {
            h += 12;
        }
        if am && h == 12 {
            h = 0;
        }
        let (mi, se) = (
            nums.get(1).copied().unwrap_or(0),
            nums.get(2).copied().unwrap_or(0),
        );
        out += (h * 3600 + mi * 60 + se) as f64 / 86_400.0;
    }
    Some(out)
}

/// Formats like VBA with a US locale: `1/15/2024`, `1/15/2024 1:30:00 PM`, `1:30:00 PM`.
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
    if serial.floor() == 0.0 && has_time {
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

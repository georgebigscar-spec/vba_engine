//! A1-style references.

pub const MAX_ROWS: u32 = 1_048_576;
pub const MAX_COLS: u32 = 16_384;

/// A rectangular block of cells, 1-based and inclusive.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Area {
    pub r1: u32,
    pub c1: u32,
    pub r2: u32,
    pub c2: u32,
}

impl Area {
    pub fn cell(r: u32, c: u32) -> Self {
        Area {
            r1: r,
            c1: c,
            r2: r,
            c2: c,
        }
    }

    pub fn rows(&self) -> u32 {
        self.r2 - self.r1 + 1
    }

    pub fn cols(&self) -> u32 {
        self.c2 - self.c1 + 1
    }

    pub fn union(&self, o: &Area) -> Area {
        Area {
            r1: self.r1.min(o.r1),
            c1: self.c1.min(o.c1),
            r2: self.r2.max(o.r2),
            c2: self.c2.max(o.c2),
        }
    }
}

pub fn col_name(mut c: u32) -> String {
    let mut s = Vec::new();
    while c > 0 {
        let d = (c - 1) % 26;
        s.push(b'A' + d as u8);
        c = (c - 1) / 26;
    }
    s.reverse();
    String::from_utf8(s).unwrap()
}

pub fn format_area(a: &Area) -> String {
    format_area_with(a, true, true)
}

/// `Range.Address(RowAbsolute, ColumnAbsolute)`.
pub fn format_area_with(a: &Area, row_abs: bool, col_abs: bool) -> String {
    let rd = if row_abs { "$" } else { "" };
    let cd = if col_abs { "$" } else { "" };
    let one = |r: u32, c: u32| format!("{cd}{}{rd}{r}", col_name(c));
    if a.r1 == a.r2 && a.c1 == a.c2 {
        one(a.r1, a.c1)
    } else if a.r1 == 1 && a.r2 == MAX_ROWS {
        format!("{cd}{}:{cd}{}", col_name(a.c1), col_name(a.c2))
    } else if a.c1 == 1 && a.c2 == MAX_COLS {
        format!("{rd}{}:{rd}{}", a.r1, a.r2)
    } else {
        format!("{}:{}", one(a.r1, a.c1), one(a.r2, a.c2))
    }
}

/// One side of a reference: `A1`, `A` (whole column) or `1` (whole row).
enum Part {
    Cell(u32, u32),
    Col(u32),
    Row(u32),
}

fn parse_part(s: &str) -> Option<Part> {
    let s: String = s.chars().filter(|&c| c != '$').collect();
    let letters: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    let digits = &s[letters.len()..];
    if !digits.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }
    let col = (!letters.is_empty()).then(|| {
        letters
            .to_ascii_uppercase()
            .bytes()
            .fold(0u32, |acc, b| acc * 26 + (b - b'A' + 1) as u32)
    });
    let row = digits.parse::<u32>().ok();
    match (col, row) {
        (Some(c), Some(r)) if (1..=MAX_COLS).contains(&c) && (1..=MAX_ROWS).contains(&r) => {
            Some(Part::Cell(r, c))
        }
        (Some(c), None) if (1..=MAX_COLS).contains(&c) => Some(Part::Col(c)),
        (None, Some(r)) if (1..=MAX_ROWS).contains(&r) => Some(Part::Row(r)),
        _ => None,
    }
}

/// Parses `A1`, `$B$2`, `A1:C3`, `A:C`, `2:5`. A sheet prefix (`Sheet1!A1`)
/// is returned separately.
pub fn parse(reference: &str) -> Option<(Option<String>, Area)> {
    let (sheet, addr) = match reference.rsplit_once('!') {
        Some((s, a)) => (Some(s.trim_matches('\'').to_string()), a),
        None => (None, reference),
    };
    let (a, b) = addr
        .trim()
        .split_once(':')
        .unwrap_or((addr.trim(), addr.trim()));
    let area = match (parse_part(a)?, parse_part(b)?) {
        (Part::Cell(r1, c1), Part::Cell(r2, c2)) => Area {
            r1: r1.min(r2),
            c1: c1.min(c2),
            r2: r1.max(r2),
            c2: c1.max(c2),
        },
        (Part::Col(c1), Part::Col(c2)) => Area {
            r1: 1,
            c1: c1.min(c2),
            r2: MAX_ROWS,
            c2: c1.max(c2),
        },
        (Part::Row(r1), Part::Row(r2)) => Area {
            r1: r1.min(r2),
            c1: 1,
            r2: r1.max(r2),
            c2: MAX_COLS,
        },
        _ => return None,
    };
    Some((sheet, area))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parse_and_format() {
        assert_eq!(parse("B3").unwrap().1, Area::cell(3, 2));
        assert_eq!(
            parse("$c$1:a2").unwrap().1,
            Area {
                r1: 1,
                c1: 1,
                r2: 2,
                c2: 3
            }
        );
        assert_eq!(format_area(&parse("A:B").unwrap().1), "$A:$B");
        assert_eq!(col_name(28), "AB");
        assert_eq!(
            parse("Data!AA10").unwrap(),
            (Some("Data".into()), Area::cell(10, 27))
        );
        assert!(parse("A0").is_none());
    }
}

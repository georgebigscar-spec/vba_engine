//! `Application.WorksheetFunction`: the commonly used worksheet functions.

use std::any::Any;
use std::cmp::Ordering;

use vba_runtime::value::{like, parse_number};
use vba_runtime::{Args, Invoke, Object, Value, VbArray, VbaError, VbaResult};

use crate::address::Area;
use crate::objects::{as_range, RangeObj};

pub struct WorksheetFunction;

/// The error Excel raises from VBA when a worksheet function fails.
fn fail(name: &str) -> VbaError {
    VbaError::new(
        1004,
        format!("Unable to get the {name} property of the WorksheetFunction class"),
    )
}

/// A rectangular source of values: a range or a VBA array (1-D is one row).
enum Grid<'a> {
    Range(&'a RangeObj),
    Array(&'a VbArray),
    Scalar(&'a Value),
}

impl<'a> Grid<'a> {
    fn of(v: &'a Value) -> Grid<'a> {
        match v {
            Value::Array(a) => Grid::Array(a),
            v => as_range(v).map_or(Grid::Scalar(v), Grid::Range),
        }
    }

    fn dims(&self) -> (u32, u32) {
        match self {
            Grid::Range(r) => (r.area.rows(), r.area.cols()),
            Grid::Array(a) => match a.bounds.as_slice() {
                [(l, u)] => (1, (u - l + 1).max(0) as u32),
                [(l1, u1), (l2, u2), ..] => {
                    ((u1 - l1 + 1).max(0) as u32, (u2 - l2 + 1).max(0) as u32)
                }
                [] => (0, 0),
            },
            Grid::Scalar(_) => (1, 1),
        }
    }

    /// Element at 1-based `(row, col)`.
    fn get(&self, r: u32, c: u32) -> Value {
        match self {
            Grid::Range(rg) => rg.sheet.get(rg.area.r1 + r - 1, rg.area.c1 + c - 1),
            Grid::Array(a) => {
                let idx: Vec<i64> = match a.bounds.as_slice() {
                    [(l, _)] => vec![l + c as i64 - 1],
                    [(l1, _), (l2, _)] => vec![l1 + r as i64 - 1, l2 + c as i64 - 1],
                    _ => return Value::Empty,
                };
                a.offset(&idx)
                    .map(|o| a.data[o].clone())
                    .unwrap_or(Value::Empty)
            }
            Grid::Scalar(v) => (*v).clone(),
        }
    }

    /// A single row or column as a list (for `Match`).
    fn vector(&self) -> Option<Vec<Value>> {
        let (rows, cols) = self.dims();
        if rows == 1 {
            Some((1..=cols).map(|c| self.get(1, c)).collect())
        } else if cols == 1 {
            Some((1..=rows).map(|r| self.get(r, 1)).collect())
        } else {
            None
        }
    }
}

/// Values referenced by a range (sparse: only stored cells, plus the count of
/// blank ones) or held by an array.
struct Values {
    values: Vec<Value>,
    blanks: u64,
    /// Values typed directly as arguments are coerced (`Sum("2")` = 2);
    /// values from ranges and arrays are not (`Sum(range of text)` = 0).
    direct: Vec<Value>,
}

fn collect(args: &Args) -> Values {
    let mut out = Values {
        values: Vec::new(),
        blanks: 0,
        direct: Vec::new(),
    };
    for v in args
        .positional
        .iter()
        .filter(|v| !matches!(v, Value::Missing))
    {
        if let Some(r) = as_range(v) {
            let a = r.area;
            let cells = r.sheet.cells.borrow();
            let mut filled = 0u64;
            for (&(_, c), d) in cells.range((a.r1, a.c1)..=(a.r2, a.c2)) {
                if c >= a.c1 && c <= a.c2 && !matches!(d.value, Value::Empty) {
                    out.values.push(d.value.clone());
                    filled += 1;
                }
            }
            out.blanks += a.rows() as u64 * a.cols() as u64 - filled;
        } else if let Value::Array(arr) = v {
            for x in &arr.data {
                match x {
                    Value::Empty => out.blanks += 1,
                    x => out.values.push(x.clone()),
                }
            }
        } else {
            out.direct.push(v.clone());
        }
    }
    out
}

fn is_number(v: &Value) -> bool {
    v.is_numeric_type() && !matches!(v, Value::Empty | Value::Boolean(_))
}

/// The numbers a statistical function (`Sum`, `Average`, `Max`) sees.
fn numbers(args: &Args, name: &str) -> VbaResult<Vec<f64>> {
    let vals = collect(args);
    let mut out: Vec<f64> = vals
        .values
        .iter()
        .filter(|v| is_number(v))
        .map(|v| v.to_f64().unwrap_or(0.0))
        .collect();
    for v in &vals.direct {
        if let Value::Error(_) = v {
            return Err(fail(name));
        }
        out.push(v.to_f64().map_err(|_| fail(name))?);
    }
    if vals.values.iter().any(|v| matches!(v, Value::Error(_))) {
        return Err(fail(name));
    }
    Ok(out)
}

/// `COUNTIF`-style criteria: `5`, `">5"`, `"<>x"`, `"ab*"`, `""`.
struct Criteria {
    op: &'static str,
    number: Option<f64>,
    text: String,
}

impl Criteria {
    fn parse(v: &Value) -> VbaResult<Criteria> {
        if let Some(r) = as_range(v) {
            return Criteria::parse(&r.value());
        }
        if is_number(v) || matches!(v, Value::Boolean(_)) {
            return Ok(Criteria {
                op: "=",
                number: Some(v.to_f64()?),
                text: v.to_string_value()?,
            });
        }
        let s = v.to_string_value()?;
        let (op, rest) = ["<>", ">=", "<=", "=", ">", "<"]
            .into_iter()
            .find_map(|op| s.strip_prefix(op).map(|r| (op, r)))
            .unwrap_or(("=", s.as_str()));
        Ok(Criteria {
            op,
            number: parse_number(rest),
            text: rest.to_string(),
        })
    }

    fn matches(&self, v: &Value) -> bool {
        let ord = match (self.number, v) {
            (Some(n), v) if is_number(v) => v.to_f64().ok().and_then(|x| x.partial_cmp(&n)),
            (Some(_), _) => None,
            (None, Value::String(s)) if matches!(self.op, "=" | "<>") => {
                let hit = like(
                    &s.to_lowercase(),
                    &wildcards(&self.text.to_lowercase()),
                    false,
                );
                return hit == (self.op == "=");
            }
            (None, Value::Empty) if self.text.is_empty() => {
                return self.op == "=";
            }
            (None, Value::String(s)) => Some(s.to_lowercase().cmp(&self.text.to_lowercase())),
            (None, _) => None,
        };
        match (self.op, ord) {
            ("=", o) => o == Some(Ordering::Equal),
            ("<>", o) => o != Some(Ordering::Equal),
            (">", o) => o == Some(Ordering::Greater),
            ("<", o) => o == Some(Ordering::Less),
            (">=", o) => matches!(o, Some(Ordering::Greater | Ordering::Equal)),
            ("<=", o) => matches!(o, Some(Ordering::Less | Ordering::Equal)),
            _ => false,
        }
    }
}

/// Excel wildcards (`*`, `?`, `~*`) as a `Like` pattern.
fn wildcards(p: &str) -> String {
    let mut out = String::new();
    let mut chars = p.chars();
    while let Some(c) = chars.next() {
        match c {
            '~' => match chars.next() {
                Some(n) => out.push_str(&format!("[{n}]")),
                None => out.push('~'),
            },
            '#' | '[' => out.push_str(&format!("[{c}]")),
            c => out.push(c),
        }
    }
    out
}

/// Cells of a criteria range clipped to the sheet's used area, as
/// `(row offset, col offset, value)`; arrays are listed in full.
fn cells_for_criteria(v: &Value) -> Vec<(u32, u32, Value)> {
    match Grid::of(v) {
        Grid::Range(r) => {
            let used = r.sheet.used_area();
            let a = r.area;
            let clip = Area {
                r1: a.r1.max(used.r1),
                c1: a.c1.max(used.c1),
                r2: a.r2.min(used.r2),
                c2: a.c2.min(used.c2),
            };
            let mut out = Vec::new();
            if clip.r1 <= clip.r2 && clip.c1 <= clip.c2 {
                for row in clip.r1..=clip.r2 {
                    for col in clip.c1..=clip.c2 {
                        out.push((row - a.r1, col - a.c1, r.sheet.get(row, col)));
                    }
                }
            }
            out
        }
        g => {
            let (rows, cols) = g.dims();
            (1..=rows)
                .flat_map(|r| (1..=cols).map(move |c| (r, c)))
                .map(|(r, c)| (r - 1, c - 1, g.get(r, c)))
                .collect()
        }
    }
}

/// Equality as lookup functions see it: numbers by value, text ignoring case.
fn lookup_cmp(a: &Value, b: &Value) -> Option<Ordering> {
    match (a, b) {
        (Value::String(x), Value::String(y)) => Some(x.to_lowercase().cmp(&y.to_lowercase())),
        (x, y) if is_number(x) && is_number(y) => x.to_f64().ok()?.partial_cmp(&y.to_f64().ok()?),
        (Value::Boolean(x), Value::Boolean(y)) => Some(x.cmp(y)),
        _ => None,
    }
}

fn lookup_eq(key: &Value, v: &Value) -> bool {
    match (key, v) {
        (Value::String(k), Value::String(s)) if k.contains(['*', '?']) => {
            like(&s.to_lowercase(), &wildcards(&k.to_lowercase()), false)
        }
        _ => lookup_cmp(key, v) == Some(Ordering::Equal),
    }
}

/// Position (1-based) of `key` in `list`. `mode` follows `MATCH`:
/// 0 exact, 1 largest ≤ key (ascending list), -1 smallest ≥ key (descending).
fn find(key: &Value, list: &[Value], mode: i64) -> Option<usize> {
    if mode == 0 {
        return list.iter().position(|v| lookup_eq(key, v)).map(|i| i + 1);
    }
    let want = if mode > 0 {
        Ordering::Greater
    } else {
        Ordering::Less
    };
    let mut best = None;
    for (i, v) in list.iter().enumerate() {
        match lookup_cmp(v, key) {
            Some(Ordering::Equal) => return Some(i + 1),
            Some(o) if o == want => break,
            Some(_) => best = Some(i + 1),
            None => {}
        }
    }
    best
}

fn round_half_away(x: f64, digits: i64) -> f64 {
    let f = 10f64.powi(digits as i32);
    let y = (x * f * (1.0 + f64::EPSILON)).abs().round() / f;
    y.copysign(x)
}

impl Object for WorksheetFunction {
    fn type_name(&self) -> &str {
        "WorksheetFunction"
    }

    fn invoke(&self, member: &str, _kind: Invoke, args: Args) -> VbaResult<Value> {
        let name = member;
        let m = member.to_ascii_lowercase();
        let num = |i: usize, n: &str| -> VbaResult<f64> {
            args.require(i, n)?.to_f64().map_err(|_| fail(name))
        };
        let v = match m.as_str() {
            "sum" => Value::Double(numbers(&args, name)?.iter().sum()),
            "product" => Value::Double(numbers(&args, name)?.iter().product()),
            "average" => {
                let n = numbers(&args, name)?;
                if n.is_empty() {
                    return Err(fail(name));
                }
                Value::Double(n.iter().sum::<f64>() / n.len() as f64)
            }
            "min" => Value::Double(
                numbers(&args, name)?
                    .into_iter()
                    .reduce(f64::min)
                    .unwrap_or(0.0),
            ),
            "max" => Value::Double(
                numbers(&args, name)?
                    .into_iter()
                    .reduce(f64::max)
                    .unwrap_or(0.0),
            ),
            "large" | "small" => {
                let mut n = numbers(&Args::new(vec![args.require(0, "Arg1")?.clone()]), name)?;
                let k = num(1, "Arg2")? as usize;
                n.sort_by(|a, b| a.total_cmp(b));
                if m == "large" {
                    n.reverse();
                }
                Value::Double(*n.get(k.wrapping_sub(1)).ok_or_else(|| fail(name))?)
            }
            "count" => {
                let vals = collect(&args);
                let direct = vals
                    .direct
                    .iter()
                    .filter(|v| v.to_f64().is_ok() && !matches!(v, Value::Error(_)))
                    .count();
                Value::Double((vals.values.iter().filter(|v| is_number(v)).count() + direct) as f64)
            }
            "counta" => {
                let vals = collect(&args);
                Value::Double((vals.values.len() + vals.direct.len()) as f64)
            }
            "countblank" => {
                let vals = collect(&args);
                let empty_text = vals
                    .values
                    .iter()
                    .filter(|v| matches!(v, Value::String(s) if s.is_empty()))
                    .count();
                Value::Double((vals.blanks + empty_text as u64) as f64)
            }
            "countif" => {
                let range = args.require(0, "Arg1")?;
                let crit = Criteria::parse(args.require(1, "Arg2")?)?;
                let hits = cells_for_criteria(range)
                    .iter()
                    .filter(|(_, _, v)| crit.matches(v))
                    .count();
                // Blank cells outside the used area match `""` too.
                let extra = match (as_range(range), crit.matches(&Value::Empty)) {
                    (Some(r), true) => {
                        let listed = cells_for_criteria(range).len() as u64;
                        r.area.rows() as u64 * r.area.cols() as u64 - listed
                    }
                    _ => 0,
                };
                Value::Double((hits as u64 + extra) as f64)
            }
            "sumif" | "averageif" => {
                let range = args.require(0, "Arg1")?;
                let crit = Criteria::parse(args.require(1, "Arg2")?)?;
                let sum_grid = Grid::of(args.get(2, "Arg3").unwrap_or(range));
                let (mut total, mut n) = (0.0, 0usize);
                for (dr, dc, v) in cells_for_criteria(range) {
                    if crit.matches(&v) {
                        let x = sum_grid.get(dr + 1, dc + 1);
                        if is_number(&x) {
                            total += x.to_f64()?;
                            n += 1;
                        }
                    }
                }
                if m == "averageif" {
                    if n == 0 {
                        return Err(fail(name));
                    }
                    total /= n as f64;
                }
                Value::Double(total)
            }
            "vlookup" | "hlookup" => {
                let key = args.require(0, "Arg1")?;
                let table = Grid::of(args.require(1, "Arg2")?);
                let index = num(2, "Arg3")? as u32;
                let approx = args
                    .get(3, "Arg4")
                    .map(|v| v.to_bool())
                    .transpose()?
                    .unwrap_or(true);
                let (rows, cols) = table.dims();
                let vertical = m == "vlookup";
                let (len, width) = if vertical { (rows, cols) } else { (cols, rows) };
                if index < 1 || index > width {
                    return Err(fail(name));
                }
                let at = |i: u32, j: u32| {
                    if vertical {
                        table.get(i, j)
                    } else {
                        table.get(j, i)
                    }
                };
                let keys: Vec<Value> = (1..=len).map(|i| at(i, 1)).collect();
                let pos = find(key, &keys, if approx { 1 } else { 0 }).ok_or_else(|| fail(name))?;
                at(pos as u32, index)
            }
            "match" => {
                let key = args.require(0, "Arg1")?;
                let list = Grid::of(args.require(1, "Arg2")?)
                    .vector()
                    .ok_or_else(|| fail(name))?;
                let mode = args
                    .get(2, "Arg3")
                    .map(|v| v.to_i64())
                    .transpose()?
                    .unwrap_or(1);
                Value::Double(find(key, &list, mode.signum()).ok_or_else(|| fail(name))? as f64)
            }
            "index" => {
                let grid = Grid::of(args.require(0, "Arg1")?);
                let (rows, cols) = grid.dims();
                let mut r = num(1, "Arg2")? as u32;
                let mut c = args
                    .get(2, "Arg3")
                    .map(|v| v.to_i64())
                    .transpose()?
                    .unwrap_or(0) as u32;
                // A single row/column takes one index.
                if c == 0 && rows == 1 {
                    (r, c) = (1, r);
                } else if c == 0 && cols == 1 {
                    c = 1;
                }
                if r < 1 || c < 1 || r > rows || c > cols {
                    return Err(fail(name));
                }
                match grid {
                    Grid::Range(rg) => {
                        return rg.invoke(
                            "Cells",
                            Invoke::Get,
                            Args::new(vec![Value::Long(r as i32), Value::Long(c as i32)]),
                        );
                    }
                    g => g.get(r, c),
                }
            }
            "round" | "roundup" | "rounddown" => {
                let x = num(0, "Arg1")?;
                let d = num(1, "Arg2")? as i64;
                let f = 10f64.powi(d as i32);
                Value::Double(match m.as_str() {
                    "round" => round_half_away(x, d),
                    "roundup" => (x * f).abs().ceil().copysign(x) / f,
                    _ => (x * f).abs().floor().copysign(x) / f,
                })
            }
            "trim" => {
                let s = args.require(0, "Arg1")?.to_string_value()?;
                Value::String(
                    s.split(' ')
                        .filter(|w| !w.is_empty())
                        .collect::<Vec<_>>()
                        .join(" "),
                )
            }
            "transpose" => {
                let grid = Grid::of(args.require(0, "Arg1")?);
                let (rows, cols) = grid.dims();
                let mut out = VbArray::new(
                    vec![(1, cols as i64), (1, rows as i64)],
                    vba_runtime::VarType::Variant,
                );
                for r in 1..=rows {
                    for c in 1..=cols {
                        out.data[((c - 1) * rows + (r - 1)) as usize] = grid.get(r, c);
                    }
                }
                Value::Array(Box::new(out))
            }
            _ => return Err(VbaError::not_supported(member)),
        };
        Ok(v)
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn criteria() {
        let c = Criteria::parse(&Value::str(">5")).unwrap();
        assert!(
            c.matches(&Value::Double(6.0))
                && !c.matches(&Value::Double(5.0))
                && !c.matches(&Value::str("x"))
        );
        let c = Criteria::parse(&Value::str("ap*")).unwrap();
        assert!(c.matches(&Value::str("Apple")) && !c.matches(&Value::str("pear")));
        let c = Criteria::parse(&Value::str("<>ap*")).unwrap();
        assert!(!c.matches(&Value::str("Apple")) && c.matches(&Value::str("pear")));
        let c = Criteria::parse(&Value::Long(3)).unwrap();
        assert!(c.matches(&Value::Double(3.0)));
        let c = Criteria::parse(&Value::str("")).unwrap();
        assert!(c.matches(&Value::Empty) && !c.matches(&Value::str("a")));
    }

    #[test]
    fn lookups() {
        let list: Vec<Value> = [1.0, 3.0, 5.0].iter().map(|&x| Value::Double(x)).collect();
        assert_eq!(find(&Value::Long(4), &list, 1), Some(2));
        assert_eq!(find(&Value::Long(4), &list, 0), None);
        assert_eq!(find(&Value::Long(5), &list, 0), Some(3));
        assert_eq!(round_half_away(2.5, 0), 3.0);
        assert_eq!(round_half_away(-2.5, 0), -3.0);
        assert_eq!(round_half_away(1.005, 2), 1.01);
    }
}

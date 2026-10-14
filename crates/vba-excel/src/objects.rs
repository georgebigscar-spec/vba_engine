//! The scriptable objects: `Application`, `Workbooks`, `Workbook`,
//! `Worksheets`, `Worksheet`, `Range`, `Font`, `Interior`.

use std::any::Any;
use std::rc::Rc;

use vba_runtime::builtins::format_value;
use vba_runtime::{Args, Invoke, Object, ObjectRef, Value, VarType, VbArray, VbaError, VbaResult};

use crate::address::{self, Area, MAX_COLS, MAX_ROWS};
use crate::functions::WorksheetFunction;
use crate::model::{AppData, BookData, CellData, SheetData, Style};

/// Above this many cells, whole-range writes (values, styles) are refused
/// instead of filling a million-cell column.
const MAX_WRITE_CELLS: u64 = 4_000_000;

fn lower(s: &str) -> String {
    s.to_ascii_lowercase()
}

fn arg_i64(args: &Args, i: usize, name: &str) -> VbaResult<i64> {
    args.require(i, name)?.to_i64()
}

fn opt_i64(args: &Args, i: usize, name: &str) -> VbaResult<Option<i64>> {
    args.get(i, name).map(|v| v.to_i64()).transpose()
}

pub(crate) fn app_error(msg: &str) -> VbaError {
    VbaError::new(1004, msg)
}

fn read_only(member: &str) -> VbaError {
    VbaError::new(1004, format!("Unable to set the {member} property"))
}

/// The `Range` behind a value, if it is one.
pub fn as_range(v: &Value) -> Option<&RangeObj> {
    match v {
        Value::Object(Some(o)) => o.as_any().downcast_ref::<RangeObj>(),
        _ => None,
    }
}

fn as_sheet(v: &Value) -> Option<&WorksheetObj> {
    match v {
        Value::Object(Some(o)) => o.as_any().downcast_ref::<WorksheetObj>(),
        _ => None,
    }
}

fn obj<T: Object + 'static>(o: T) -> Value {
    Value::Object(Some(Rc::new(o) as ObjectRef))
}

// ---- Range ------------------------------------------------------------------

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub enum RangeMode {
    Cells,
    Rows,
    Columns,
}

#[derive(Clone)]
pub struct RangeObj {
    pub app: Rc<AppData>,
    pub book: Rc<BookData>,
    pub sheet: Rc<SheetData>,
    pub area: Area,
    pub mode: RangeMode,
}

impl RangeObj {
    fn with_area(&self, area: Area) -> RangeObj {
        RangeObj {
            area,
            mode: RangeMode::Cells,
            ..self.clone()
        }
    }

    fn cells_count(&self) -> u64 {
        self.area.rows() as u64 * self.area.cols() as u64
    }

    fn check_write_size(&self) -> VbaResult<()> {
        if self.cells_count() > MAX_WRITE_CELLS {
            return Err(app_error("Range too large to write cell by cell"));
        }
        Ok(())
    }

    /// Cell `(r, c)` relative to this range's top-left corner (1-based; may
    /// fall outside the range, as in Excel).
    fn rel(&self, r: i64, c: i64) -> VbaResult<Area> {
        let r = self.area.r1 as i64 + r - 1;
        let c = self.area.c1 as i64 + c - 1;
        if r < 1 || c < 1 || r > MAX_ROWS as i64 || c > MAX_COLS as i64 {
            return Err(app_error("Application-defined or object-defined error"));
        }
        Ok(Area::cell(r as u32, c as u32))
    }

    pub fn value(&self) -> Value {
        let a = self.area;
        if a.r1 == a.r2 && a.c1 == a.c2 {
            return self.sheet.get(a.r1, a.c1);
        }
        let mut arr = VbArray::new(
            vec![(1, a.rows() as i64), (1, a.cols() as i64)],
            VarType::Variant,
        );
        for (&(r, c), d) in self.sheet.cells.borrow().range((a.r1, a.c1)..=(a.r2, a.c2)) {
            if c >= a.c1 && c <= a.c2 {
                let off = ((r - a.r1) * a.cols() + (c - a.c1)) as usize;
                arr.data[off] = d.value.clone();
            }
        }
        Value::Array(Box::new(arr))
    }

    /// Writes a scalar to every cell, or an array shaped like the range
    /// (a 1-D array fills each row).
    pub fn set_value(&self, v: Value) -> VbaResult<()> {
        self.check_write_size()?;
        let a = self.area;
        match v {
            Value::Array(arr) => {
                let get = |r: u32, c: u32| -> Value {
                    match arr.bounds.as_slice() {
                        [(l, u)] => {
                            let i = l + c as i64;
                            if i <= *u {
                                arr.data[(i - l) as usize].clone()
                            } else {
                                Value::Error(2042)
                            }
                        }
                        [(l1, u1), (l2, u2)] => {
                            let (i, j) = (l1 + r as i64, l2 + c as i64);
                            if i <= *u1 && j <= *u2 {
                                arr.data[((i - l1) * (u2 - l2 + 1) + (j - l2)) as usize].clone()
                            } else {
                                Value::Error(2042)
                            }
                        }
                        _ => Value::Error(2042),
                    }
                };
                for r in 0..a.rows() {
                    for c in 0..a.cols() {
                        self.sheet.set(a.r1 + r, a.c1 + c, get(r, c));
                    }
                }
            }
            Value::Object(Some(o)) => return self.set_value(o.default_value()?),
            v => {
                for r in a.r1..=a.r2 {
                    for c in a.c1..=a.c2 {
                        self.sheet.set(r, c, v.clone());
                    }
                }
            }
        }
        Ok(())
    }

    fn set_formula(&self, f: String) -> VbaResult<()> {
        self.check_write_size()?;
        let a = self.area;
        for r in a.r1..=a.r2 {
            for c in a.c1..=a.c2 {
                if f.starts_with('=') {
                    self.sheet.update(r, c, |d| {
                        d.formula = Some(f.clone());
                        d.value = Value::Empty;
                    });
                } else {
                    self.sheet.set(r, c, Value::String(f.clone()));
                }
            }
        }
        Ok(())
    }

    pub fn style(&self) -> Style {
        self.sheet.data(self.area.r1, self.area.c1).style
    }

    pub fn update_style(&self, f: impl Fn(&mut Style)) -> VbaResult<()> {
        let a = self.area;
        if a.r1 == 1 && a.r2 == MAX_ROWS || a.c1 == 1 && a.c2 == MAX_COLS {
            // Whole rows/columns: restyle the cells that exist instead of
            // materializing a million blank ones. Cells written later do not
            // inherit the format (Excel keeps row/column styles separately).
            let keys: Vec<(u32, u32)> = self
                .sheet
                .cells
                .borrow()
                .range((a.r1, a.c1)..=(a.r2, a.c2))
                .map(|(&k, _)| k)
                .filter(|&(_, c)| c >= a.c1 && c <= a.c2)
                .collect();
            for (r, c) in keys {
                self.sheet.update(r, c, |d| f(&mut d.style));
            }
            return Ok(());
        }
        self.check_write_size()?;
        for r in a.r1..=a.r2 {
            for c in a.c1..=a.c2 {
                self.sheet.update(r, c, |d| f(&mut d.style));
            }
        }
        Ok(())
    }

    fn text(&self) -> VbaResult<String> {
        let d = self.sheet.data(self.area.r1, self.area.c1);
        match (&d.style.number_format, &d.value) {
            (Some(fmt), v)
                if !fmt.eq_ignore_ascii_case("general") && !matches!(v, Value::Empty) =>
            {
                format_value(v, fmt)
            }
            (_, Value::Error(code)) => Ok(error_text(*code).to_string()),
            (_, v) => Ok(v.to_string()),
        }
    }

    fn sub_range(&self, args: &Args) -> VbaResult<RangeObj> {
        let first = args.require(0, "Cell1")?;
        let area = match first {
            Value::String(s) => {
                let (_, a) = address::parse(s)
                    .ok_or_else(|| app_error("Method 'Range' of object failed"))?;
                // Relative to this range's top-left corner.
                let rel = |a: Area| Area {
                    r1: a.r1 + self.area.r1 - 1,
                    c1: a.c1 + self.area.c1 - 1,
                    r2: a.r2 + self.area.r1 - 1,
                    c2: a.c2 + self.area.c1 - 1,
                };
                rel(a)
            }
            v => as_range(v)
                .map(|r| r.area)
                .ok_or_else(VbaError::type_mismatch)?,
        };
        let area = match args.get(1, "Cell2") {
            Some(v) => area.union(&self.area_of(v)?),
            None => area,
        };
        Ok(self.with_area(area))
    }

    fn area_of(&self, v: &Value) -> VbaResult<Area> {
        match v {
            Value::String(s) => address::parse(s)
                .map(|(_, a)| a)
                .ok_or_else(|| app_error("Method 'Range' of object failed")),
            v => as_range(v)
                .map(|r| r.area)
                .ok_or_else(VbaError::type_mismatch),
        }
    }

    fn cells(&self, args: &Args) -> VbaResult<RangeObj> {
        match (args.get(0, "RowIndex"), args.get(1, "ColumnIndex")) {
            (None, _) => Ok(self.with_area(self.area)),
            (Some(r), Some(c)) => {
                let c = match c {
                    Value::String(letters) => {
                        address::parse(&format!("{letters}1"))
                            .ok_or_else(VbaError::type_mismatch)?
                            .1
                            .c1 as i64
                    }
                    v => v.to_i64()?,
                };
                Ok(self.with_area(self.rel(r.to_i64()?, c)?))
            }
            (Some(i), None) => {
                let i = i.to_i64()? - 1;
                let cols = self.area.cols() as i64;
                Ok(self.with_area(self.rel(i / cols + 1, i % cols + 1)?))
            }
        }
    }

    fn worksheet(&self) -> WorksheetObj {
        WorksheetObj {
            app: self.app.clone(),
            book: self.book.clone(),
            sheet: self.sheet.clone(),
        }
    }

    fn items(&self) -> Vec<Value> {
        let a = self.area;
        match self.mode {
            RangeMode::Rows => (a.r1..=a.r2)
                .map(|r| obj(self.with_area(Area { r1: r, r2: r, ..a })))
                .collect(),
            RangeMode::Columns => (a.c1..=a.c2)
                .map(|c| obj(self.with_area(Area { c1: c, c2: c, ..a })))
                .collect(),
            RangeMode::Cells => {
                let mut v = Vec::new();
                for r in a.r1..=a.r2 {
                    for c in a.c1..=a.c2 {
                        v.push(obj(self.with_area(Area::cell(r, c))));
                    }
                }
                v
            }
        }
    }
}

pub fn error_text(code: i32) -> &'static str {
    match code {
        2000 => "#NULL!",
        2007 => "#DIV/0!",
        2015 => "#VALUE!",
        2023 => "#REF!",
        2029 => "#NAME?",
        2036 => "#NUM!",
        _ => "#N/A",
    }
}

impl Object for RangeObj {
    fn type_name(&self) -> &str {
        "Range"
    }

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        let a = self.area;
        match (lower(member).as_str(), kind) {
            ("" | "value" | "value2", Invoke::Let(v)) if args.is_empty() => self.set_value(v)?,
            ("" | "value" | "value2", _) if args.is_empty() => return Ok(self.value()),
            // `rng(1, 2)` — the default member with arguments is Item.
            ("" | "item", Invoke::Let(v)) => self.cells(&args)?.set_value(v)?,
            ("" | "item", _) => return Ok(obj(self.cells(&args)?)),
            ("formula", Invoke::Let(v)) => self.set_formula(v.to_string_value()?)?,
            ("formula", _) => {
                let d = self.sheet.data(a.r1, a.c1);
                return Ok(Value::String(
                    d.formula.unwrap_or_else(|| d.value.to_string()),
                ));
            }
            ("hasformula", _) => {
                return Ok(Value::Boolean(
                    self.sheet.data(a.r1, a.c1).formula.is_some(),
                ))
            }
            ("text", _) => return Ok(Value::String(self.text()?)),
            ("numberformat", Invoke::Let(v)) => {
                let f = v.to_string_value()?;
                self.update_style(|s| s.number_format = Some(f.clone()))?
            }
            ("numberformat", _) => {
                return Ok(Value::String(
                    self.style()
                        .number_format
                        .unwrap_or_else(|| "General".into()),
                ))
            }
            ("address", _) => {
                let flag = |i, n| {
                    args.get(i, n)
                        .map(|v| v.to_bool())
                        .transpose()
                        .map(|b| b.unwrap_or(true))
                };
                let (row_abs, col_abs) = (flag(0, "RowAbsolute")?, flag(1, "ColumnAbsolute")?);
                return Ok(Value::String(address::format_area_with(
                    &a, row_abs, col_abs,
                )));
            }
            ("row", _) => return Ok(Value::Long(a.r1 as i32)),
            ("column", _) => return Ok(Value::Long(a.c1 as i32)),
            ("count" | "countlarge", _) => {
                let n = match self.mode {
                    RangeMode::Cells => self.cells_count(),
                    RangeMode::Rows => a.rows() as u64,
                    RangeMode::Columns => a.cols() as u64,
                };
                return Ok(if n > i32::MAX as u64 {
                    Value::Double(n as f64)
                } else {
                    Value::Long(n as i32)
                });
            }
            ("cells", Invoke::Let(v)) => self.cells(&args)?.set_value(v)?,
            ("cells", _) => return Ok(obj(self.cells(&args)?)),
            ("range", Invoke::Let(v)) => self.sub_range(&args)?.set_value(v)?,
            ("range", _) => return Ok(obj(self.sub_range(&args)?)),
            ("rows", _) => {
                return Ok(match opt_i64(&args, 0, "Index")? {
                    Some(i) => obj(self.with_area(Area {
                        r1: self.rel(i, 1)?.r1,
                        r2: self.rel(i, 1)?.r1,
                        ..a
                    })),
                    None => obj(RangeObj {
                        mode: RangeMode::Rows,
                        ..self.clone()
                    }),
                })
            }
            ("columns", _) => {
                return Ok(match opt_i64(&args, 0, "Index")? {
                    Some(i) => obj(self.with_area(Area {
                        c1: self.rel(1, i)?.c1,
                        c2: self.rel(1, i)?.c1,
                        ..a
                    })),
                    None => obj(RangeObj {
                        mode: RangeMode::Columns,
                        ..self.clone()
                    }),
                })
            }
            ("entirerow", _) => {
                return Ok(obj(RangeObj {
                    mode: RangeMode::Rows,
                    ..self.with_area(Area {
                        c1: 1,
                        c2: MAX_COLS,
                        ..a
                    })
                }))
            }
            ("entirecolumn", _) => {
                return Ok(obj(RangeObj {
                    mode: RangeMode::Columns,
                    ..self.with_area(Area {
                        r1: 1,
                        r2: MAX_ROWS,
                        ..a
                    })
                }))
            }
            ("offset", _) => {
                let dr = opt_i64(&args, 0, "RowOffset")?.unwrap_or(0);
                let dc = opt_i64(&args, 1, "ColumnOffset")?.unwrap_or(0);
                let tl = self.rel(dr + 1, dc + 1)?;
                let br = self.with_area(Area::cell(a.r2, a.c2)).rel(dr + 1, dc + 1)?;
                return Ok(obj(self.with_area(Area {
                    r1: tl.r1,
                    c1: tl.c1,
                    r2: br.r1,
                    c2: br.c1,
                })));
            }
            ("resize", _) => {
                let rows = opt_i64(&args, 0, "RowSize")?.unwrap_or(a.rows() as i64);
                let cols = opt_i64(&args, 1, "ColumnSize")?.unwrap_or(a.cols() as i64);
                let br = self.rel(rows, cols)?;
                return Ok(obj(self.with_area(Area {
                    r2: br.r1,
                    c2: br.c1,
                    ..a
                })));
            }
            ("end", _) => {
                let (dr, dc) = match arg_i64(&args, 0, "Direction")? {
                    -4162 => (-1, 0),
                    -4121 => (1, 0),
                    -4159 => (0, -1),
                    -4161 => (0, 1),
                    _ => return Err(VbaError::invalid_call()),
                };
                let (r, c) = self.sheet.end(a.r1, a.c1, dr, dc);
                return Ok(obj(self.with_area(Area::cell(r, c))));
            }
            ("currentregion", _) => return Ok(obj(self.with_area(self.sheet.current_region(a)))),
            ("clearcontents", _) => {
                for r in a.r1..=a.r2 {
                    let keys: Vec<(u32, u32)> = self
                        .sheet
                        .cells
                        .borrow()
                        .range((r, a.c1)..=(r, a.c2))
                        .map(|(k, _)| *k)
                        .collect();
                    for (r, c) in keys {
                        self.sheet.update(r, c, |d| {
                            d.value = Value::Empty;
                            d.formula = None;
                        });
                    }
                }
            }
            ("clear", _) => {
                self.sheet
                    .cells
                    .borrow_mut()
                    .retain(|&(r, c), _| !(r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2));
            }
            ("copy", _) => {
                let dest = args
                    .get(0, "Destination")
                    .and_then(as_range)
                    .ok_or_else(|| app_error("Copy needs a Destination"))?;
                let cells: Vec<((u32, u32), CellData)> = self
                    .sheet
                    .cells
                    .borrow()
                    .range((a.r1, a.c1)..=(a.r2, a.c2))
                    .filter(|((_, c), _)| *c >= a.c1 && *c <= a.c2)
                    .map(|(k, d)| (*k, d.clone()))
                    .collect();
                let (dr, dc) = (
                    dest.area.r1 as i64 - a.r1 as i64,
                    dest.area.c1 as i64 - a.c1 as i64,
                );
                for r in 0..a.rows() {
                    for c in 0..a.cols() {
                        let (tr, tc) = ((dest.area.r1 + r), (dest.area.c1 + c));
                        dest.sheet.put(tr, tc, CellData::default());
                    }
                }
                for ((r, c), d) in cells {
                    dest.sheet
                        .put((r as i64 + dr) as u32, (c as i64 + dc) as u32, d);
                }
            }
            ("select" | "activate", _) => {
                self.app.activate_book(&self.book);
                if let Some(i) = self
                    .book
                    .sheets
                    .borrow()
                    .iter()
                    .position(|s| Rc::ptr_eq(s, &self.sheet))
                {
                    self.book.active.set(i);
                }
                self.app.active_cell.set((a.r1, a.c1));
            }
            ("font", _) => {
                return Ok(obj(StyleObj {
                    range: self.clone(),
                    kind: StyleKind::Font,
                }))
            }
            ("interior", _) => {
                return Ok(obj(StyleObj {
                    range: self.clone(),
                    kind: StyleKind::Interior,
                }))
            }
            ("worksheet" | "parent", _) => return Ok(obj(self.worksheet())),
            ("autofit" | "merge" | "unmerge" | "calculate", _) => {}
            (m, _) => return Err(VbaError::not_supported(&format!("Range.{m}"))),
        }
        Ok(Value::Empty)
    }

    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        if self.cells_count() > MAX_WRITE_CELLS && self.mode == RangeMode::Cells {
            return Err(app_error("Range too large to enumerate"));
        }
        Ok(self.items())
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        let mut v = vec![(
            "Address".to_string(),
            Value::String(address::format_area(&self.area)),
        )];
        if self.area.rows() == 1 && self.area.cols() == 1 {
            v.push(("Value".to_string(), self.value()));
        }
        v
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

// ---- Font / Interior --------------------------------------------------------

#[derive(Clone, Copy)]
enum StyleKind {
    Font,
    Interior,
}

struct StyleObj {
    range: RangeObj,
    kind: StyleKind,
}

impl Object for StyleObj {
    fn type_name(&self) -> &str {
        match self.kind {
            StyleKind::Font => "Font",
            StyleKind::Interior => "Interior",
        }
    }

    fn invoke(&self, member: &str, kind: Invoke, _args: Args) -> VbaResult<Value> {
        let st = self.range.style();
        let m = lower(member);
        match (self.kind, m.as_str(), kind) {
            (StyleKind::Font, "bold", Invoke::Let(v)) => {
                let b = v.to_bool()?;
                self.range.update_style(|s| s.bold = b)?
            }
            (StyleKind::Font, "bold", _) => return Ok(Value::Boolean(st.bold)),
            (StyleKind::Font, "italic", Invoke::Let(v)) => {
                let b = v.to_bool()?;
                self.range.update_style(|s| s.italic = b)?
            }
            (StyleKind::Font, "italic", _) => return Ok(Value::Boolean(st.italic)),
            (StyleKind::Font, "size", Invoke::Let(v)) => {
                let n = v.to_f64()?;
                self.range.update_style(|s| s.font_size = Some(n))?
            }
            (StyleKind::Font, "size", _) => return Ok(Value::Double(st.font_size.unwrap_or(11.0))),
            (StyleKind::Font, "color", Invoke::Let(v)) => {
                let c = v.to_i64()? as i32;
                self.range.update_style(|s| s.font_color = Some(c))?
            }
            (StyleKind::Font, "color", _) => {
                return Ok(Value::Double(st.font_color.unwrap_or(0) as f64))
            }
            (StyleKind::Interior, "color", Invoke::Let(v)) => {
                let c = v.to_i64()? as i32;
                self.range.update_style(|s| s.fill_color = Some(c))?
            }
            (StyleKind::Interior, "color", _) => {
                return Ok(Value::Double(st.fill_color.unwrap_or(0xFFFFFF) as f64))
            }
            (StyleKind::Interior, "colorindex" | "pattern", Invoke::Let(v)) => {
                // xlNone (-4142) clears the fill.
                if v.to_i64()? == -4142 {
                    self.range.update_style(|s| s.fill_color = None)?
                }
            }
            (_, m, _) => {
                return Err(VbaError::not_supported(&format!(
                    "{}.{m}",
                    self.type_name()
                )))
            }
        }
        Ok(Value::Empty)
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

// ---- Worksheet / Worksheets -------------------------------------------------

#[derive(Clone)]
pub struct WorksheetObj {
    pub app: Rc<AppData>,
    pub book: Rc<BookData>,
    pub sheet: Rc<SheetData>,
}

impl WorksheetObj {
    pub fn whole(&self) -> RangeObj {
        RangeObj {
            app: self.app.clone(),
            book: self.book.clone(),
            sheet: self.sheet.clone(),
            area: Area {
                r1: 1,
                c1: 1,
                r2: MAX_ROWS,
                c2: MAX_COLS,
            },
            mode: RangeMode::Cells,
        }
    }

    fn index(&self) -> usize {
        self.book
            .sheets
            .borrow()
            .iter()
            .position(|s| Rc::ptr_eq(s, &self.sheet))
            .unwrap_or(0)
    }

    /// `Range("A1")`, `Range("A1:B2")`, `Range(cell1, cell2)`, `Range("Sheet2!A1")`.
    pub fn range(&self, args: &Args) -> VbaResult<RangeObj> {
        if let Some(Value::String(s)) = args.get(0, "Cell1") {
            if let Some((Some(sheet), area)) = address::parse(s) {
                let sheet = self
                    .book
                    .sheet(&sheet)
                    .ok_or_else(VbaError::subscript_out_of_range)?;
                let ws = WorksheetObj {
                    sheet,
                    ..self.clone()
                };
                return Ok(ws.whole().with_area(area));
            }
        }
        self.whole().sub_range(args)
    }
}

impl Object for WorksheetObj {
    fn type_name(&self) -> &str {
        "Worksheet"
    }

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        match (lower(member).as_str(), kind) {
            ("name", Invoke::Let(v)) => {
                let name = v.to_string_value()?;
                if self
                    .book
                    .sheet(&name)
                    .is_some_and(|s| !Rc::ptr_eq(&s, &self.sheet))
                {
                    return Err(app_error("That name is already taken"));
                }
                *self.sheet.name.borrow_mut() = name;
            }
            ("name" | "codename", _) => return Ok(Value::String(self.sheet.name.borrow().clone())),
            ("index", _) => return Ok(Value::Long(self.index() as i32 + 1)),
            ("range", Invoke::Let(v)) => self.range(&args)?.set_value(v)?,
            ("range", _) => return Ok(obj(self.range(&args)?)),
            (m @ ("cells" | "rows" | "columns"), kind) => {
                return self.whole().invoke(m, kind, args)
            }
            ("usedrange", _) => return Ok(obj(self.whole().with_area(self.sheet.used_area()))),
            ("activate" | "select", _) => {
                self.app.activate_book(&self.book);
                self.book.active.set(self.index());
            }
            ("delete", _) => {
                let mut sheets = self.book.sheets.borrow_mut();
                if sheets.len() == 1 {
                    return Err(app_error(
                        "A workbook must contain at least one visible worksheet",
                    ));
                }
                sheets.retain(|s| !Rc::ptr_eq(s, &self.sheet));
                self.book.active.set(0);
            }
            ("parent", _) => {
                return Ok(obj(WorkbookObj {
                    app: self.app.clone(),
                    book: self.book.clone(),
                }))
            }
            ("calculate" | "protect" | "unprotect", _) => {}
            ("visible", Invoke::Let(_)) => {}
            ("visible", _) => return Ok(Value::Long(-1)),
            (m, _) => return Err(VbaError::not_supported(&format!("Worksheet.{m}"))),
        }
        Ok(Value::Empty)
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        vec![(
            "Name".into(),
            Value::String(self.sheet.name.borrow().clone()),
        )]
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

pub struct WorksheetsObj {
    pub app: Rc<AppData>,
    pub book: Rc<BookData>,
}

impl WorksheetsObj {
    fn sheet_obj(&self, sheet: Rc<SheetData>) -> Value {
        obj(WorksheetObj {
            app: self.app.clone(),
            book: self.book.clone(),
            sheet,
        })
    }

    fn item(&self, key: &Value) -> VbaResult<Value> {
        let sheet = match key {
            Value::String(name) => self.book.sheet(name),
            v => {
                let i = v.to_i64()?;
                self.book
                    .sheets
                    .borrow()
                    .get((i - 1).max(0) as usize)
                    .filter(|_| i >= 1)
                    .cloned()
            }
        };
        sheet
            .map(|s| self.sheet_obj(s))
            .ok_or_else(VbaError::subscript_out_of_range)
    }
}

impl Object for WorksheetsObj {
    fn type_name(&self) -> &str {
        "Sheets"
    }

    fn invoke(&self, member: &str, _kind: Invoke, args: Args) -> VbaResult<Value> {
        match lower(member).as_str() {
            "count" => Ok(Value::Long(self.book.sheets.borrow().len() as i32)),
            "" | "item" => self.item(args.require(0, "Index")?),
            "add" => {
                let count = opt_i64(&args, 2, "Count")?.unwrap_or(1).clamp(1, 255) as usize;
                let len = self.book.sheets.borrow().len();
                let at = match (
                    args.get(0, "Before").and_then(as_sheet),
                    args.get(1, "After").and_then(as_sheet),
                ) {
                    (Some(b), _) => b.index(),
                    (None, Some(a)) => a.index() + 1,
                    _ => self.book.active.get().min(len),
                };
                let mut last = None;
                for pos in at..at + count {
                    let sheet = SheetData::new(&self.book.next_sheet_name());
                    self.book.sheets.borrow_mut().insert(pos, sheet.clone());
                    self.book.active.set(pos);
                    last = Some(sheet);
                }
                Ok(self.sheet_obj(last.unwrap()))
            }
            m => Err(VbaError::not_supported(&format!("Sheets.{m}"))),
        }
    }

    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        Ok(self
            .book
            .sheets
            .borrow()
            .iter()
            .map(|s| self.sheet_obj(s.clone()))
            .collect())
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

// ---- Workbook / Workbooks ---------------------------------------------------

pub struct WorkbookObj {
    pub app: Rc<AppData>,
    pub book: Rc<BookData>,
}

impl Object for WorkbookObj {
    fn type_name(&self) -> &str {
        "Workbook"
    }

    fn invoke(&self, member: &str, _kind: Invoke, args: Args) -> VbaResult<Value> {
        let sheets = || {
            obj(WorksheetsObj {
                app: self.app.clone(),
                book: self.book.clone(),
            })
        };
        match lower(member).as_str() {
            "name" => Ok(Value::String(self.book.name.borrow().clone())),
            "fullname" => Ok(Value::String(
                self.book
                    .path
                    .borrow()
                    .as_ref()
                    .map(|p| p.display().to_string())
                    .unwrap_or_else(|| self.book.name.borrow().clone()),
            )),
            "path" => Ok(Value::String(
                self.book
                    .path
                    .borrow()
                    .as_ref()
                    .and_then(|p| p.parent())
                    .map(|p| p.display().to_string())
                    .unwrap_or_default(),
            )),
            "worksheets" | "sheets" => match args.get(0, "Index") {
                Some(key) => WorksheetsObj {
                    app: self.app.clone(),
                    book: self.book.clone(),
                }
                .item(key),
                None => Ok(sheets()),
            },
            "activesheet" => Ok(obj(WorksheetObj {
                app: self.app.clone(),
                book: self.book.clone(),
                sheet: self.book.active_sheet(),
            })),
            "activate" => {
                self.app.activate_book(&self.book);
                Ok(Value::Empty)
            }
            "save" => {
                let path = self
                    .book
                    .path
                    .borrow()
                    .clone()
                    .ok_or_else(|| app_error("Workbook has no file name; use SaveAs"))?;
                crate::save_book(&self.book, &path)?;
                Ok(Value::Empty)
            }
            "saveas" | "savecopyas" => {
                let path =
                    std::path::PathBuf::from(args.require(0, "Filename")?.to_string_value()?);
                crate::save_book(&self.book, &path)?;
                if lower(member) == "saveas" {
                    *self.book.name.borrow_mut() = path
                        .file_name()
                        .map(|n| n.to_string_lossy().into_owned())
                        .unwrap_or_default();
                    *self.book.path.borrow_mut() = Some(path);
                }
                Ok(Value::Empty)
            }
            "close" => {
                let mut books = self.app.books.borrow_mut();
                if books.len() > 1 {
                    books.retain(|b| !Rc::ptr_eq(b, &self.book));
                    self.app.active_book.set(0);
                }
                Ok(Value::Empty)
            }
            m => Err(VbaError::not_supported(&format!("Workbook.{m}"))),
        }
    }

    fn inspect(&self) -> Vec<(String, Value)> {
        vec![(
            "Name".into(),
            Value::String(self.book.name.borrow().clone()),
        )]
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

pub struct WorkbooksObj {
    pub app: Rc<AppData>,
}

impl WorkbooksObj {
    fn book_obj(&self, book: Rc<BookData>) -> Value {
        obj(WorkbookObj {
            app: self.app.clone(),
            book,
        })
    }
}

impl Object for WorkbooksObj {
    fn type_name(&self) -> &str {
        "Workbooks"
    }

    fn invoke(&self, member: &str, _kind: Invoke, args: Args) -> VbaResult<Value> {
        match lower(member).as_str() {
            "count" => Ok(Value::Long(self.app.books.borrow().len() as i32)),
            "" | "item" => {
                let books = self.app.books.borrow();
                let book = match args.require(0, "Index")? {
                    Value::String(name) => books
                        .iter()
                        .find(|b| b.name.borrow().eq_ignore_ascii_case(name))
                        .cloned(),
                    v => {
                        let i = v.to_i64()?;
                        books
                            .get((i - 1).max(0) as usize)
                            .filter(|_| i >= 1)
                            .cloned()
                    }
                };
                drop(books);
                book.map(|b| self.book_obj(b))
                    .ok_or_else(VbaError::subscript_out_of_range)
            }
            "add" => Ok(self.book_obj(self.app.add_book())),
            "open" => {
                let path =
                    std::path::PathBuf::from(args.require(0, "Filename")?.to_string_value()?);
                let book = crate::load_book(&path)?;
                self.app.books.borrow_mut().push(book.clone());
                self.app.activate_book(&book);
                Ok(self.book_obj(book))
            }
            m => Err(VbaError::not_supported(&format!("Workbooks.{m}"))),
        }
    }

    fn enumerate(&self) -> VbaResult<Vec<Value>> {
        Ok(self
            .app
            .books
            .borrow()
            .iter()
            .map(|b| self.book_obj(b.clone()))
            .collect())
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

// ---- Application ------------------------------------------------------------

pub struct ApplicationObj {
    pub app: Rc<AppData>,
}

impl ApplicationObj {
    pub fn active_sheet(&self) -> WorksheetObj {
        let book = self.app.active_book();
        WorksheetObj {
            app: self.app.clone(),
            sheet: book.active_sheet(),
            book,
        }
    }

    pub fn active_workbook(&self) -> Value {
        obj(WorkbookObj {
            app: self.app.clone(),
            book: self.app.active_book(),
        })
    }

    pub fn active_cell(&self) -> Value {
        let (r, c) = self.app.active_cell.get();
        obj(self.active_sheet().whole().with_area(Area::cell(r, c)))
    }
}

impl Object for ApplicationObj {
    fn type_name(&self) -> &str {
        "Application"
    }

    fn invoke(&self, member: &str, kind: Invoke, args: Args) -> VbaResult<Value> {
        let m = lower(member);
        match (m.as_str(), kind) {
            ("name", _) => Ok(Value::String("Microsoft Excel".into())),
            ("version", _) => Ok(Value::String("16.0".into())),
            ("application", _) => Ok(obj(ApplicationObj {
                app: self.app.clone(),
            })),
            ("workbooks", _) => {
                let wbs = WorkbooksObj {
                    app: self.app.clone(),
                };
                match args.get(0, "Index") {
                    Some(_) => wbs.invoke("item", Invoke::Get, args),
                    None => Ok(obj(wbs)),
                }
            }
            ("activeworkbook" | "thisworkbook", _) => Ok(self.active_workbook()),
            ("activesheet", _) => Ok(obj(self.active_sheet())),
            ("activecell", _) => Ok(self.active_cell()),
            ("selection", _) => Ok(self.active_cell()),
            ("worksheets" | "sheets", _) => WorkbookObj {
                app: self.app.clone(),
                book: self.app.active_book(),
            }
            .invoke(&m, Invoke::Get, args),
            ("range" | "cells" | "rows" | "columns", kind) => {
                self.active_sheet().invoke(&m, kind, args)
            }
            ("worksheetfunction", _) => Ok(obj(WorksheetFunction)),
            ("intersect", _) => {
                let a = args
                    .require(0, "Arg1")
                    .ok()
                    .and_then(as_range)
                    .ok_or_else(VbaError::type_mismatch)?;
                let b = args
                    .require(1, "Arg2")
                    .ok()
                    .and_then(as_range)
                    .ok_or_else(VbaError::type_mismatch)?;
                let i = Area {
                    r1: a.area.r1.max(b.area.r1),
                    c1: a.area.c1.max(b.area.c1),
                    r2: a.area.r2.min(b.area.r2),
                    c2: a.area.c2.min(b.area.c2),
                };
                if i.r1 > i.r2 || i.c1 > i.c2 || !Rc::ptr_eq(&a.sheet, &b.sheet) {
                    Ok(Value::Object(None))
                } else {
                    Ok(obj(a.with_area(i)))
                }
            }
            ("union", _) => {
                let mut area: Option<Area> = None;
                let mut first: Option<RangeObj> = None;
                for v in &args.positional {
                    let r = as_range(v).ok_or_else(VbaError::type_mismatch)?;
                    area = Some(area.map_or(r.area, |a| a.union(&r.area)));
                    first.get_or_insert_with(|| r.clone());
                }
                let first = first.ok_or_else(VbaError::wrong_arg_count)?;
                // Non-contiguous unions are approximated by their bounding box.
                Ok(obj(first.with_area(area.unwrap())))
            }
            ("calculate" | "wait" | "ontime" | "goto" | "quit", _) => Ok(Value::Empty),
            (
                "cutcopymode" | "screenupdating" | "displayalerts" | "enableevents" | "calculation"
                | "statusbar" | "displaystatusbar" | "username" | "caption",
                Invoke::Let(v),
            ) => {
                self.app.props.borrow_mut().insert(m, v);
                Ok(Value::Empty)
            }
            ("screenupdating" | "displayalerts" | "enableevents" | "displaystatusbar", _) => {
                Ok(self
                    .app
                    .props
                    .borrow()
                    .get(&m)
                    .cloned()
                    .unwrap_or(Value::Boolean(true)))
            }
            ("calculation", _) => Ok(self
                .app
                .props
                .borrow()
                .get(&m)
                .cloned()
                .unwrap_or(Value::Long(-4105))),
            ("cutcopymode" | "statusbar" | "username" | "caption", _) => Ok(self
                .app
                .props
                .borrow()
                .get(&m)
                .cloned()
                .unwrap_or(Value::Boolean(false))),
            (m, Invoke::Let(_)) => Err(read_only(m)),
            (m, _) => Err(VbaError::not_supported(m)),
        }
    }

    fn as_any(&self) -> &dyn Any {
        self
    }
}

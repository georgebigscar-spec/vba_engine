//! Workbook data: sheets of sparse cells.

use std::cell::{Cell as StdCell, RefCell};
use std::collections::{BTreeMap, HashMap};
use std::path::PathBuf;
use std::rc::Rc;

use vba_runtime::Value;

use crate::address::{Area, MAX_COLS, MAX_ROWS};

#[derive(Debug, Clone, Default, PartialEq)]
pub struct Style {
    pub bold: bool,
    pub italic: bool,
    /// VBA color (`RGB(r, g, b)` = r + g·256 + b·65536).
    pub font_color: Option<i32>,
    pub font_size: Option<f64>,
    pub fill_color: Option<i32>,
    pub number_format: Option<String>,
}

#[derive(Debug, Clone, Default)]
pub struct CellData {
    pub value: Value,
    /// Formula text including the leading `=`. Formulas are stored and
    /// saved, not calculated.
    pub formula: Option<String>,
    pub style: Style,
}

impl CellData {
    fn is_blank(&self) -> bool {
        matches!(self.value, Value::Empty)
            && self.formula.is_none()
            && self.style == Style::default()
    }
}

/// Normalizes a value for storage the way Excel does: every number is a Double.
pub fn to_cell_value(v: Value) -> Value {
    match v {
        Value::Byte(_)
        | Value::Integer(_)
        | Value::Long(_)
        | Value::LongLong(_)
        | Value::Single(_)
        | Value::Currency(_)
        | Value::Decimal(_) => Value::Double(v.to_f64().unwrap_or(0.0)),
        Value::Null | Value::Missing => Value::Empty,
        other => other,
    }
}

pub struct SheetData {
    pub name: RefCell<String>,
    /// Keyed by `(row, column)`, 1-based.
    pub cells: RefCell<BTreeMap<(u32, u32), CellData>>,
}

impl SheetData {
    pub fn new(name: &str) -> Rc<Self> {
        Rc::new(SheetData {
            name: RefCell::new(name.to_string()),
            cells: RefCell::default(),
        })
    }

    pub fn get(&self, r: u32, c: u32) -> Value {
        self.cells
            .borrow()
            .get(&(r, c))
            .map(|d| d.value.clone())
            .unwrap_or(Value::Empty)
    }

    pub fn data(&self, r: u32, c: u32) -> CellData {
        self.cells
            .borrow()
            .get(&(r, c))
            .cloned()
            .unwrap_or_default()
    }

    pub fn set(&self, r: u32, c: u32, v: Value) {
        self.update(r, c, |d| {
            d.value = to_cell_value(v);
            d.formula = None;
        });
    }

    pub fn update(&self, r: u32, c: u32, f: impl FnOnce(&mut CellData)) {
        let mut cells = self.cells.borrow_mut();
        let d = cells.entry((r, c)).or_default();
        f(d);
        if d.is_blank() {
            cells.remove(&(r, c));
        }
    }

    pub fn put(&self, r: u32, c: u32, d: CellData) {
        if d.is_blank() {
            self.cells.borrow_mut().remove(&(r, c));
        } else {
            self.cells.borrow_mut().insert((r, c), d);
        }
    }

    pub fn is_filled(&self, r: u32, c: u32) -> bool {
        self.cells
            .borrow()
            .get(&(r, c))
            .is_some_and(|d| !matches!(d.value, Value::Empty) || d.formula.is_some())
    }

    /// Bounding box of non-empty cells (`UsedRange`); `A1` for an empty sheet.
    pub fn used_area(&self) -> Area {
        let cells = self.cells.borrow();
        let Some(&(r, c)) = cells.keys().next() else {
            return Area::cell(1, 1);
        };
        cells
            .keys()
            .fold(Area::cell(r, c), |a, &(r, c)| a.union(&Area::cell(r, c)))
    }

    /// `Range.End`: Ctrl+arrow from `(r, c)`. `dr`/`dc` is the direction.
    pub fn end(&self, r: u32, c: u32, dr: i32, dc: i32) -> (u32, u32) {
        let vertical = dr != 0;
        let forward = dr > 0 || dc > 0;
        let (pos, max) = if vertical {
            (r, MAX_ROWS)
        } else {
            (c, MAX_COLS)
        };
        // Filled positions along this row/column, sorted.
        let filled: Vec<u32> = self
            .cells
            .borrow()
            .iter()
            .filter(|(&(cr, cc), d)| {
                (if vertical { cc == c } else { cr == r })
                    && (!matches!(d.value, Value::Empty) || d.formula.is_some())
            })
            .map(|(&(cr, cc), _)| if vertical { cr } else { cc })
            .collect();
        let has = |p: u32| filled.binary_search(&p).is_ok();
        let next = |p: u32| {
            if forward {
                (p < max).then(|| p + 1)
            } else {
                (p > 1).then(|| p - 1)
            }
        };
        let result = match next(pos) {
            None => pos,
            Some(n) if has(pos) && has(n) => {
                let mut p = n;
                while let Some(q) = next(p).filter(|&q| has(q)) {
                    p = q;
                }
                p
            }
            Some(_) => {
                let found = if forward {
                    filled.iter().copied().find(|&p| p > pos)
                } else {
                    filled.iter().rev().copied().find(|&p| p < pos)
                };
                found.unwrap_or(if forward { max } else { 1 })
            }
        };
        if vertical {
            (result, c)
        } else {
            (r, result)
        }
    }

    /// `Range.CurrentRegion`: grows the area while any neighbouring cell is filled.
    pub fn current_region(&self, start: Area) -> Area {
        let mut a = start;
        loop {
            let (r1, c1) = (a.r1.saturating_sub(1).max(1), a.c1.saturating_sub(1).max(1));
            let (r2, c2) = ((a.r2 + 1).min(MAX_ROWS), (a.c2 + 1).min(MAX_COLS));
            let cells = self.cells.borrow();
            let mut grown = a;
            for (&(r, c), d) in cells.range((r1, 0)..=(r2, u32::MAX)) {
                if c >= c1 && c <= c2 && (!matches!(d.value, Value::Empty) || d.formula.is_some()) {
                    grown = grown.union(&Area::cell(r, c));
                }
            }
            if grown == a {
                return a;
            }
            a = grown;
        }
    }
}

pub struct BookData {
    pub name: RefCell<String>,
    pub path: RefCell<Option<PathBuf>>,
    pub sheets: RefCell<Vec<Rc<SheetData>>>,
    pub active: StdCell<usize>,
}

impl BookData {
    pub fn new(name: &str, sheets: Vec<Rc<SheetData>>) -> Rc<Self> {
        Rc::new(BookData {
            name: RefCell::new(name.to_string()),
            path: RefCell::new(None),
            sheets: RefCell::new(sheets),
            active: StdCell::new(0),
        })
    }

    pub fn sheet(&self, name: &str) -> Option<Rc<SheetData>> {
        self.sheets
            .borrow()
            .iter()
            .find(|s| s.name.borrow().eq_ignore_ascii_case(name))
            .cloned()
    }

    pub fn active_sheet(&self) -> Rc<SheetData> {
        let sheets = self.sheets.borrow();
        sheets[self.active.get().min(sheets.len() - 1)].clone()
    }

    /// A sheet name not used yet: `Sheet2`, `Sheet3`, ...
    pub fn next_sheet_name(&self) -> String {
        (1..)
            .map(|i| format!("Sheet{i}"))
            .find(|n| self.sheet(n).is_none())
            .unwrap()
    }
}

pub struct AppData {
    pub books: RefCell<Vec<Rc<BookData>>>,
    pub active_book: StdCell<usize>,
    /// The active cell on the active sheet.
    pub active_cell: StdCell<(u32, u32)>,
    /// Settable application properties (`ScreenUpdating`, `StatusBar`, ...).
    pub props: RefCell<HashMap<String, Value>>,
    pub output: RefCell<Vec<String>>,
    pub echo: StdCell<bool>,
}

impl AppData {
    pub fn new() -> Rc<Self> {
        let app = Rc::new(AppData {
            books: RefCell::default(),
            active_book: StdCell::new(0),
            active_cell: StdCell::new((1, 1)),
            props: RefCell::default(),
            output: RefCell::default(),
            echo: StdCell::new(false),
        });
        app.add_book();
        app
    }

    pub fn add_book(&self) -> Rc<BookData> {
        let n = self.books.borrow().len() + 1;
        let book = BookData::new(&format!("Book{n}"), vec![SheetData::new("Sheet1")]);
        self.books.borrow_mut().push(book.clone());
        self.active_book.set(n - 1);
        book
    }

    pub fn active_book(&self) -> Rc<BookData> {
        let books = self.books.borrow();
        books[self.active_book.get().min(books.len() - 1)].clone()
    }

    pub fn activate_book(&self, book: &Rc<BookData>) {
        if let Some(i) = self.books.borrow().iter().position(|b| Rc::ptr_eq(b, book)) {
            self.active_book.set(i);
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn end_like_ctrl_arrow() {
        let s = SheetData::new("S");
        for r in [1, 2, 3, 7, 8] {
            s.set(r, 1, Value::Double(r as f64));
        }
        assert_eq!(s.end(MAX_ROWS, 1, -1, 0), (8, 1)); // Cells(Rows.Count, 1).End(xlUp)
        assert_eq!(s.end(1, 1, 1, 0), (3, 1)); // end of the first block
        assert_eq!(s.end(3, 1, 1, 0), (7, 1)); // jump to the next block
        assert_eq!(s.end(8, 1, 1, 0), (MAX_ROWS, 1));
        assert_eq!(s.end(1, 1, 0, 1), (1, MAX_COLS));
        assert_eq!(
            s.current_region(Area::cell(2, 1)),
            Area {
                r1: 1,
                c1: 1,
                r2: 3,
                c2: 1
            }
        );
    }
}

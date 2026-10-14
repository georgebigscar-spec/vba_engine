//! An in-memory Excel object model for `vba-runtime`: `Application`,
//! `Workbooks`, `Worksheets`, `Range`, `WorksheetFunction` and the `xl*`
//! constants, with `.xlsx` loading and saving (feature `xlsx`).
//!
//! ```
//! use vba_excel::ExcelHost;
//! use vba_runtime::Engine;
//!
//! let host = ExcelHost::new();
//! host.set_cell("Sheet1", "A1", 10.0.into());
//! let mut engine = Engine::new(host.clone());
//! engine.load_source("Sub Main()\n Range(\"B1\").Value = Range(\"A1\").Value * 2\nEnd Sub", "Module1").unwrap();
//! engine.run("Main", vec![]).unwrap();
//! assert_eq!(host.cell("Sheet1", "B1").unwrap().to_string(), "20");
//! ```

pub mod address;
pub mod functions;
pub mod model;
pub mod objects;

use std::path::Path;
use std::rc::Rc;

use vba_runtime::{Host, Invoke, Object, ObjectRef, Value, VbaError, VbaResult};

use model::{AppData, BookData};
use objects::{ApplicationObj, WorkbooksObj};

/// `xl*` constants available to macros.
const XL_CONSTANTS: &[(&str, i32)] = &[
    ("xlUp", -4162),
    ("xlDown", -4121),
    ("xlToLeft", -4159),
    ("xlToRight", -4161),
    ("xlCalculationAutomatic", -4105),
    ("xlCalculationManual", -4135),
    ("xlCalculationSemiautomatic", 2),
    ("xlValues", -4163),
    ("xlFormulas", -4123),
    ("xlWhole", 1),
    ("xlPart", 2),
    ("xlByRows", 1),
    ("xlByColumns", 2),
    ("xlNext", 1),
    ("xlPrevious", 2),
    ("xlNone", -4142),
    ("xlAutomatic", -4105),
    ("xlCenter", -4108),
    ("xlLeft", -4131),
    ("xlRight", -4152),
    ("xlTop", -4160),
    ("xlBottom", -4107),
    ("xlGeneral", 1),
    ("xlAscending", 1),
    ("xlDescending", 2),
    ("xlYes", 1),
    ("xlNo", 2),
    ("xlGuess", 0),
    ("xlSheetVisible", -1),
    ("xlSheetHidden", 0),
    ("xlSheetVeryHidden", 2),
    ("xlOpenXMLWorkbook", 51),
    ("xlOpenXMLWorkbookMacroEnabled", 52),
    ("xlWorkbookDefault", 51),
    ("xlCSV", 6),
    ("xlPasteAll", -4104),
    ("xlPasteValues", -4163),
    ("xlPasteFormats", -4122),
    ("xlCellTypeLastCell", 11),
    ("xlCellTypeConstants", 2),
    ("xlCellTypeBlanks", 4),
    ("xlContinuous", 1),
    ("xlThin", 2),
    ("xlMedium", -4138),
    ("xlThick", 4),
    ("xlEdgeTop", 8),
    ("xlEdgeBottom", 9),
    ("xlEdgeLeft", 7),
    ("xlEdgeRight", 10),
    ("xlErrDiv0", 2007),
    ("xlErrNA", 2042),
    ("xlErrValue", 2015),
    ("xlErrRef", 2023),
    ("xlErrName", 2029),
    ("xlErrNum", 2036),
    ("xlErrNull", 2000),
];

/// A [`Host`] backed by an in-memory Excel application. Cloning shares the
/// application, so tests and embedders keep a handle to inspect the
/// workbook after a macro runs.
#[derive(Clone)]
pub struct ExcelHost {
    pub app: Rc<AppData>,
}

impl Default for ExcelHost {
    fn default() -> Self {
        Self::new()
    }
}

impl ExcelHost {
    /// A fresh application with one workbook `Book1` holding `Sheet1`.
    pub fn new() -> Self {
        ExcelHost {
            app: AppData::new(),
        }
    }

    /// Opens a workbook from disk and makes it the only, active one.
    pub fn open(path: impl AsRef<Path>) -> VbaResult<Self> {
        let book = load_book(path.as_ref())?;
        let app = AppData::new();
        *app.books.borrow_mut() = vec![book];
        app.active_book.set(0);
        Ok(ExcelHost { app })
    }

    /// Echo `Debug.Print` / `MsgBox` to stdout as well as recording them.
    pub fn with_echo(self, echo: bool) -> Self {
        self.app.echo.set(echo);
        self
    }

    pub fn application(&self) -> ApplicationObj {
        ApplicationObj {
            app: self.app.clone(),
        }
    }

    pub fn active_workbook(&self) -> Rc<BookData> {
        self.app.active_book()
    }

    /// Value of a cell in the active workbook, `None` if the sheet or the
    /// address does not exist.
    pub fn cell(&self, sheet: &str, addr: &str) -> Option<Value> {
        let s = self.active_workbook().sheet(sheet)?;
        let (_, area) = address::parse(addr)?;
        Some(s.get(area.r1, area.c1))
    }

    /// Sets a cell in the active workbook; creates the sheet if needed.
    pub fn set_cell(&self, sheet: &str, addr: &str, value: Value) {
        let book = self.active_workbook();
        let s = book.sheet(sheet).unwrap_or_else(|| {
            let s = model::SheetData::new(sheet);
            book.sheets.borrow_mut().push(s.clone());
            s
        });
        let (_, area) = address::parse(addr).expect("valid cell address");
        s.set(area.r1, area.c1, value);
    }

    /// Lines written with `Debug.Print` and `MsgBox`.
    pub fn lines(&self) -> Vec<String> {
        self.app.output.borrow().clone()
    }

    /// Saves the active workbook to `path` (`.xlsx`).
    pub fn save_as(&self, path: impl AsRef<Path>) -> VbaResult<()> {
        let book = self.active_workbook();
        save_book(&book, path.as_ref())?;
        *book.path.borrow_mut() = Some(path.as_ref().to_path_buf());
        Ok(())
    }
}

impl Host for ExcelHost {
    fn global(&self, name: &str) -> Option<Value> {
        let app = self.application();
        let member = match name.to_ascii_lowercase().as_str() {
            "application" => return Some(Value::Object(Some(Rc::new(app) as ObjectRef))),
            "workbooks" => {
                return Some(Value::Object(Some(Rc::new(WorkbooksObj {
                    app: self.app.clone(),
                }) as ObjectRef)))
            }
            m @ ("activeworkbook" | "thisworkbook" | "activesheet" | "activecell" | "selection"
            | "worksheets" | "sheets" | "worksheetfunction") => m.to_string(),
            _ => {
                return XL_CONSTANTS
                    .iter()
                    .find(|(n, _)| n.eq_ignore_ascii_case(name))
                    .map(|&(_, v)| Value::Long(v));
            }
        };
        app.invoke(&member, Invoke::Get, Default::default()).ok()
    }

    fn global_object(&self) -> Option<ObjectRef> {
        Some(Rc::new(self.application()))
    }

    fn debug_print(&mut self, text: &str) {
        if self.app.echo.get() {
            println!("{text}");
        }
        self.app.output.borrow_mut().push(text.to_string());
    }
}

// ---- .xlsx I/O ----------------------------------------------------------------

#[cfg(not(feature = "xlsx"))]
pub fn load_book(path: &Path) -> VbaResult<Rc<BookData>> {
    Err(VbaError::new(
        1004,
        format!("{}: built without the `xlsx` feature", path.display()),
    ))
}

#[cfg(not(feature = "xlsx"))]
pub fn save_book(_book: &Rc<BookData>, path: &Path) -> VbaResult<()> {
    Err(VbaError::new(
        1004,
        format!("{}: built without the `xlsx` feature", path.display()),
    ))
}

#[cfg(feature = "xlsx")]
pub use xlsx::{load_book, save_book};

#[cfg(feature = "xlsx")]
mod xlsx {
    use super::*;
    use calamine::{open_workbook_auto, Data, Reader};
    use rust_xlsxwriter::{Color, Format, Workbook};

    use crate::model::{CellData, SheetData, Style};

    fn io_error(path: &Path, e: impl std::fmt::Display) -> VbaError {
        VbaError::new(1004, format!("{}: {e}", path.display()))
    }

    fn error_code(e: &calamine::CellErrorType) -> i32 {
        use calamine::CellErrorType::*;
        match e {
            Div0 => 2007,
            NA => 2042,
            Name => 2029,
            Null => 2000,
            Num => 2036,
            Ref => 2023,
            Value => 2015,
            _ => 2042,
        }
    }

    fn to_value(d: &Data) -> Value {
        match d {
            Data::Empty => Value::Empty,
            Data::Int(i) => Value::Double(*i as f64),
            Data::Float(f) => Value::Double(*f),
            Data::String(s) => Value::String(s.clone()),
            Data::Bool(b) => Value::Boolean(*b),
            Data::DateTime(dt) => Value::Date(dt.as_f64()),
            Data::DateTimeIso(s) | Data::DurationIso(s) => Value::String(s.clone()),
            Data::Error(e) => Value::Error(error_code(e)),
        }
    }

    pub fn load_book(path: &Path) -> VbaResult<Rc<BookData>> {
        let mut wb = open_workbook_auto(path).map_err(|e| io_error(path, e))?;
        let mut sheets = Vec::new();
        for name in wb.sheet_names() {
            let sheet = SheetData::new(&name);
            let range = wb.worksheet_range(&name).map_err(|e| io_error(path, e))?;
            if let Some((r0, c0)) = range.start() {
                for (r, c, d) in range.used_cells() {
                    sheet.set(r0 + r as u32 + 1, c0 + c as u32 + 1, to_value(d));
                }
            }
            if let Ok(formulas) = wb.worksheet_formula(&name) {
                if let Some((r0, c0)) = formulas.start() {
                    for (r, c, f) in formulas.used_cells() {
                        if !f.is_empty() {
                            sheet.update(r0 + r as u32 + 1, c0 + c as u32 + 1, |d| {
                                d.formula = Some(format!("={f}"))
                            });
                        }
                    }
                }
            }
            sheets.push(sheet);
        }
        if sheets.is_empty() {
            sheets.push(SheetData::new("Sheet1"));
        }
        let name = path.file_name().and_then(|n| n.to_str()).unwrap_or("Book1");
        let book = BookData::new(name, sheets);
        *book.path.borrow_mut() = Some(path.to_path_buf());
        Ok(book)
    }

    /// VBA colors are `&HBBGGRR`; xlsx wants `0xRRGGBB`.
    fn color(bgr: i32) -> Color {
        let (r, g, b) = (bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
        Color::RGB(((r << 16) | (g << 8) | b) as u32)
    }

    fn format(style: &Style, is_date: bool) -> Option<Format> {
        let date_fmt = is_date && style.number_format.is_none();
        if *style == Style::default() && !date_fmt {
            return None;
        }
        let mut f = Format::new();
        if style.bold {
            f = f.set_bold();
        }
        if style.italic {
            f = f.set_italic();
        }
        if let Some(c) = style.font_color {
            f = f.set_font_color(color(c));
        }
        if let Some(s) = style.font_size {
            f = f.set_font_size(s);
        }
        if let Some(c) = style.fill_color {
            f = f.set_background_color(color(c));
        }
        if let Some(n) = &style.number_format {
            f = f.set_num_format(n);
        } else if date_fmt {
            f = f.set_num_format("yyyy-mm-dd hh:mm:ss");
        }
        Some(f)
    }

    fn write_cell(
        ws: &mut rust_xlsxwriter::Worksheet,
        r: u32,
        c: u16,
        d: &CellData,
    ) -> Result<(), rust_xlsxwriter::XlsxError> {
        let fmt = format(&d.style, matches!(d.value, Value::Date(_)));
        if let Some(formula) = &d.formula {
            match &fmt {
                Some(f) => ws.write_formula_with_format(r, c, formula.as_str(), f)?,
                None => ws.write_formula(r, c, formula.as_str())?,
            };
            return Ok(());
        }
        match (&d.value, &fmt) {
            (Value::String(s), Some(f)) => ws.write_string_with_format(r, c, s, f)?,
            (Value::String(s), None) => ws.write_string(r, c, s)?,
            (Value::Boolean(b), Some(f)) => ws.write_boolean_with_format(r, c, *b, f)?,
            (Value::Boolean(b), None) => ws.write_boolean(r, c, *b)?,
            (Value::Empty, Some(f)) => ws.write_blank(r, c, f)?,
            (Value::Empty, None) => ws,
            (Value::Error(code), _) => ws.write_string(r, c, crate::objects::error_text(*code))?,
            (v, f) => {
                let x = v.to_f64().unwrap_or(0.0);
                match f {
                    Some(f) => ws.write_number_with_format(r, c, x, f)?,
                    None => ws.write_number(r, c, x)?,
                }
            }
        };
        Ok(())
    }

    pub fn save_book(book: &Rc<BookData>, path: &Path) -> VbaResult<()> {
        let mut wb = Workbook::new();
        for sheet in book.sheets.borrow().iter() {
            let ws = wb.add_worksheet();
            ws.set_name(sheet.name.borrow().as_str())
                .map_err(|e| io_error(path, e))?;
            for (&(r, c), d) in sheet.cells.borrow().iter() {
                write_cell(ws, r - 1, (c - 1) as u16, d).map_err(|e| io_error(path, e))?;
            }
        }
        wb.save(path).map_err(|e| io_error(path, e))
    }
}

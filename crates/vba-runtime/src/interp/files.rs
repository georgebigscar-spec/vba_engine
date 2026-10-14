//! File I/O (`Open`, `Print #`, `Line Input #`, ...) and the string
//! statements `LSet`, `RSet`, `Mid(...) =` and `Erase`.
//!
//! Files go through the [`Host`](crate::Host): input files are read whole
//! when opened, output is buffered and written when the file is closed (or
//! when the macro ends).

use vba_syntax::ast::{Expr, OpenMode, PrintItem};

use super::Interpreter;
use crate::date;
use crate::error::{VbaError, VbaResult};
use crate::object::Args;
use crate::value::{self, Value, VbArray};

pub(super) struct OpenFile {
    path: String,
    mode: OpenMode,
    data: Vec<u8>,
    /// Read position for `Input` files.
    pos: usize,
    /// Current output column, for `,` print zones and `Tab`.
    pub(super) col: usize,
}

fn bad_file_number() -> VbaError {
    VbaError::new(52, "Bad file name or number")
}

fn bad_file_mode() -> VbaError {
    VbaError::new(54, "Bad file mode")
}

fn past_end() -> VbaError {
    VbaError::new(62, "Input past end of file")
}

/// `Write #` formatting: strings quoted, `#TRUE#`, `#NULL#`, dates as `#yyyy-mm-dd#`.
fn write_text(v: &Value) -> String {
    match v {
        Value::String(s) => format!("\"{s}\""),
        Value::Boolean(b) => if *b { "#TRUE#" } else { "#FALSE#" }.to_string(),
        Value::Null => "#NULL#".to_string(),
        Value::Empty => String::new(),
        Value::Error(n) => format!("#ERROR {n}#"),
        Value::Date(d) => {
            let (y, m, day) = date::ymd(*d);
            let (h, mi, s) = date::hms(*d);
            match (d.trunc() != 0.0, d.fract() != 0.0) {
                (_, false) => format!("#{y:04}-{m:02}-{day:02}#"),
                (false, true) => format!("#{h:02}:{mi:02}:{s:02}#"),
                (true, true) => format!("#{y:04}-{m:02}-{day:02} {h:02}:{mi:02}:{s:02}#"),
            }
        }
        v => v.to_string(),
    }
}

/// A field read by `Input #`, typed the way VBA reads it back.
fn input_value(raw: &str, quoted: bool) -> Value {
    if quoted {
        return Value::String(raw.to_string());
    }
    let t = raw.trim();
    match t.to_ascii_uppercase().as_str() {
        "" => return Value::Empty,
        "#TRUE#" => return Value::Boolean(true),
        "#FALSE#" => return Value::Boolean(false),
        "#NULL#" => return Value::Null,
        _ => {}
    }
    if let Some(d) = t.strip_prefix('#').and_then(|r| r.strip_suffix('#')) {
        if let Some(serial) = date::parse_date(d) {
            return Value::Date(serial);
        }
    }
    match value::parse_number(t) {
        Some(n) => Value::Double(n),
        None => Value::String(t.to_string()),
    }
}

/// A `Dir` / `Kill` wildcard pattern as a `Like` pattern.
fn wildcard(pattern: &str) -> String {
    pattern
        .chars()
        .map(|c| match c {
            '#' | '[' => format!("[{c}]"),
            c => c.to_string(),
        })
        .collect()
}

fn split_path(path: &str) -> (&str, &str) {
    path.rsplit_once(['/', '\\']).unwrap_or(("", path))
}

fn join_path(dir: &str, name: &str) -> String {
    if dir.is_empty() {
        name.to_string()
    } else {
        let sep = if dir.contains('\\') && !dir.contains('/') {
            '\\'
        } else {
            '/'
        };
        format!("{dir}{sep}{name}")
    }
}

impl Interpreter {
    fn file_number(&mut self, e: &Expr) -> VbaResult<i64> {
        let n = self.eval(e)?;
        let n = self.scalar(n)?.to_i64()?;
        if (1..=511).contains(&n) {
            Ok(n)
        } else {
            Err(bad_file_number())
        }
    }

    fn open_file(&mut self, n: i64) -> VbaResult<&mut OpenFile> {
        self.files.get_mut(&n).ok_or_else(bad_file_number)
    }

    fn output_file(&mut self, n: i64) -> VbaResult<&mut OpenFile> {
        let f = self.open_file(n)?;
        match f.mode {
            OpenMode::Output | OpenMode::Append => Ok(f),
            _ => Err(bad_file_mode()),
        }
    }

    fn input_file(&mut self, n: i64) -> VbaResult<&mut OpenFile> {
        let f = self.open_file(n)?;
        match f.mode {
            OpenMode::Input => Ok(f),
            _ => Err(bad_file_mode()),
        }
    }

    pub(super) fn exec_open(&mut self, path: &Expr, mode: OpenMode, file: &Expr) -> VbaResult<()> {
        let path = self.eval(path)?;
        let path = self.scalar(path)?.to_string_value()?;
        let n = self.file_number(file)?;
        if self.files.contains_key(&n) {
            return Err(VbaError::new(55, "File already open"));
        }
        let data = match mode {
            OpenMode::Input => self.host.read_file(&path)?,
            OpenMode::Output => {
                // Opening for output creates (and empties) the file at once.
                self.host.write_file(&path, &[])?;
                Vec::new()
            }
            OpenMode::Append => match self.host.read_file(&path) {
                Ok(d) => d,
                Err(e) if e.number == 53 => Vec::new(),
                Err(e) => return Err(e),
            },
            OpenMode::Binary | OpenMode::Random => {
                return Err(VbaError::new(
                    -1,
                    "Binary and Random file modes are not supported",
                ))
            }
        };
        let col = data.iter().rev().take_while(|&&b| b != b'\n').count();
        self.files.insert(
            n,
            OpenFile {
                path,
                mode,
                data,
                pos: 0,
                col,
            },
        );
        Ok(())
    }

    pub(super) fn exec_close(&mut self, files: &[Expr]) -> VbaResult<()> {
        let numbers = if files.is_empty() {
            let mut all: Vec<i64> = self.files.keys().copied().collect();
            all.sort();
            all
        } else {
            files
                .iter()
                .map(|f| self.file_number(f))
                .collect::<VbaResult<_>>()?
        };
        for n in numbers {
            self.close_file(n)?;
        }
        Ok(())
    }

    fn close_file(&mut self, n: i64) -> VbaResult<()> {
        if let Some(f) = self.files.remove(&n) {
            if matches!(f.mode, OpenMode::Output | OpenMode::Append) {
                self.host.write_file(&f.path, &f.data)?;
            }
        }
        Ok(())
    }

    /// Flushes and closes every file; called when a macro ends.
    pub(super) fn close_all_files(&mut self) -> VbaResult<()> {
        self.exec_close(&[])
    }

    pub(super) fn exec_print_file(&mut self, file: &Expr, items: &[PrintItem]) -> VbaResult<()> {
        let n = self.file_number(file)?;
        let col = self.output_file(n)?.col;
        let (line, open) = self.print_items(items, String::new(), col)?;
        let f = self.output_file(n)?;
        f.data.extend_from_slice(line.as_bytes());
        if open {
            f.col += line.chars().count();
        } else {
            f.data.extend_from_slice(b"\r\n");
            f.col = 0;
        }
        Ok(())
    }

    pub(super) fn exec_write_file(&mut self, file: &Expr, items: &[PrintItem]) -> VbaResult<()> {
        let n = self.file_number(file)?;
        self.output_file(n)?;
        let mut line = String::new();
        let mut open = false;
        for item in items {
            match item {
                PrintItem::Expr(e) => {
                    let v = self.eval(e)?;
                    let v = match v {
                        Value::Null => v,
                        v => self.scalar(v)?,
                    };
                    line.push_str(&write_text(&v));
                    open = false;
                }
                PrintItem::Comma | PrintItem::Semicolon => {
                    line.push(',');
                    open = true;
                }
            }
        }
        let f = self.output_file(n)?;
        f.data.extend_from_slice(line.as_bytes());
        if open {
            f.col += line.chars().count();
        } else {
            f.data.extend_from_slice(b"\r\n");
            f.col = 0;
        }
        Ok(())
    }

    pub(super) fn exec_line_input(&mut self, file: &Expr, target: &Expr) -> VbaResult<()> {
        let n = self.file_number(file)?;
        let f = self.input_file(n)?;
        if f.pos >= f.data.len() {
            return Err(past_end());
        }
        let rest = &f.data[f.pos..];
        let end = rest
            .iter()
            .position(|&b| b == b'\n' || b == b'\r')
            .unwrap_or(rest.len());
        let line = String::from_utf8_lossy(&rest[..end]).into_owned();
        let mut next = f.pos + end;
        if f.data.get(next) == Some(&b'\r') {
            next += 1;
        }
        if f.data.get(next) == Some(&b'\n') {
            next += 1;
        }
        f.pos = next;
        self.assign(target, Value::String(line), false)
    }

    pub(super) fn exec_input_file(&mut self, file: &Expr, targets: &[Expr]) -> VbaResult<()> {
        let n = self.file_number(file)?;
        for target in targets {
            let f = self.input_file(n)?;
            let data = &f.data;
            let mut i = f.pos;
            while i < data.len() && matches!(data[i], b' ' | b'\t') {
                i += 1;
            }
            if i >= data.len() {
                return Err(past_end());
            }
            let (raw, quoted) = if data[i] == b'"' {
                let start = i + 1;
                let end = data[start..]
                    .iter()
                    .position(|&b| b == b'"')
                    .map_or(data.len(), |p| start + p);
                i = (end + 1).min(data.len());
                (
                    String::from_utf8_lossy(&data[start..end]).into_owned(),
                    true,
                )
            } else {
                let start = i;
                while i < data.len() && !matches!(data[i], b',' | b'\r' | b'\n') {
                    i += 1;
                }
                (String::from_utf8_lossy(&data[start..i]).into_owned(), false)
            };
            // Skip to after the delimiter: a comma or the end of the line.
            while i < data.len() && !matches!(data[i], b',' | b'\r' | b'\n') {
                i += 1;
            }
            if data.get(i) == Some(&b',') {
                i += 1;
            } else {
                if data.get(i) == Some(&b'\r') {
                    i += 1;
                }
                if data.get(i) == Some(&b'\n') {
                    i += 1;
                }
            }
            f.pos = i;
            self.assign(target, input_value(&raw, quoted), false)?;
        }
        Ok(())
    }

    /// `EOF`, `LOF`, `FreeFile`, `Dir`, `Kill`, `FileLen`, `FileCopy`,
    /// `MkDir`: built-ins that need the open-file table or the host's files.
    pub(super) fn file_builtin(&mut self, name: &str, args: &Args) -> Option<VbaResult<Value>> {
        let lname = name.to_ascii_lowercase();
        if !matches!(
            lname.as_str(),
            "eof" | "lof" | "freefile" | "dir" | "kill" | "filelen" | "filecopy" | "mkdir"
        ) {
            return None;
        }
        Some(self.file_builtin_inner(&lname, args))
    }

    fn file_builtin_inner(&mut self, name: &str, args: &Args) -> VbaResult<Value> {
        let text =
            |i: usize, n: &str| -> VbaResult<String> { args.require(i, n)?.to_string_value() };
        Ok(match name {
            "eof" => {
                let n = args.require(0, "FileNumber")?.to_i64()?;
                let f = self.open_file(n)?;
                Value::Boolean(f.mode != OpenMode::Input || f.pos >= f.data.len())
            }
            "lof" => {
                let n = args.require(0, "FileNumber")?.to_i64()?;
                Value::Long(self.open_file(n)?.data.len() as i32)
            }
            "freefile" => {
                let n = (1..=511)
                    .find(|n| !self.files.contains_key(n))
                    .ok_or_else(|| VbaError::new(67, "Too many files"))?;
                Value::Integer(n as i16)
            }
            "dir" => match args.get(0, "PathName") {
                Some(p) => {
                    let path = p.to_string_value()?;
                    let (dir, pattern) = split_path(&path);
                    let pattern = if pattern.is_empty() { "*" } else { pattern };
                    let names = match self.host.list_dir(dir) {
                        Ok(n) => n,
                        Err(e) if e.number == 53 => Vec::new(),
                        Err(e) => return Err(e),
                    };
                    let like = wildcard(pattern);
                    let mut found: Vec<String> = names
                        .into_iter()
                        .filter(|n| value::like(n, &like, true))
                        .collect();
                    found.reverse();
                    let first = found.pop().unwrap_or_default();
                    self.dir_pending = Some(found);
                    Value::String(first)
                }
                None => match &mut self.dir_pending {
                    Some(rest) => Value::String(rest.pop().unwrap_or_default()),
                    None => return Err(VbaError::invalid_call()),
                },
            },
            "kill" => {
                let path = text(0, "PathName")?;
                let (dir, pattern) = split_path(&path);
                if pattern.contains(['*', '?']) {
                    let like = wildcard(pattern);
                    let names = self.host.list_dir(dir)?;
                    let matching: Vec<String> = names
                        .into_iter()
                        .filter(|n| value::like(n, &like, true))
                        .collect();
                    if matching.is_empty() {
                        return Err(VbaError::new(53, format!("File not found: {path}")));
                    }
                    for n in matching {
                        self.host.delete_file(&join_path(dir, &n))?;
                    }
                } else {
                    self.host.delete_file(&path)?;
                }
                Value::Empty
            }
            "filelen" => {
                let path = text(0, "PathName")?;
                Value::Long(self.host.read_file(&path)?.len() as i32)
            }
            "filecopy" => {
                let (src, dst) = (text(0, "Source")?, text(1, "Destination")?);
                let data = self.host.read_file(&src)?;
                self.host.write_file(&dst, &data)?;
                Value::Empty
            }
            "mkdir" => {
                let path = text(0, "Path")?;
                self.host.create_dir(&path)?;
                Value::Empty
            }
            _ => unreachable!(),
        })
    }

    /// `LSet s = v` / `RSet s = v`: `v` left- or right-aligned in `s`'s length.
    pub(super) fn exec_align_set(
        &mut self,
        right: bool,
        target: &Expr,
        value: &Expr,
    ) -> VbaResult<()> {
        let current = self.eval(target)?;
        let width = self.scalar(current)?.to_string_value()?.chars().count();
        let v = self.eval(value)?;
        let v: Vec<char> = self
            .scalar(v)?
            .to_string_value()?
            .chars()
            .take(width)
            .collect();
        let pad = " ".repeat(width - v.len());
        let v: String = v.into_iter().collect();
        let out = if right { pad + &v } else { v + &pad };
        self.assign(target, Value::String(out), false)
    }

    /// `Mid(s, start[, length]) = v`: overwrites characters, keeping `s`'s length.
    pub(super) fn exec_mid_assign(
        &mut self,
        target: &Expr,
        start: &Expr,
        length: Option<&Expr>,
        value: &Expr,
    ) -> VbaResult<()> {
        let s = self.eval(target)?;
        let mut chars: Vec<char> = self.scalar(s)?.to_string_value()?.chars().collect();
        let start = self.eval(start)?.to_i64()?;
        if start < 1 || start as usize > chars.len() {
            return Err(VbaError::invalid_call());
        }
        let length = match length {
            Some(e) => Some(self.eval(e)?.to_i64()?.max(0) as usize),
            None => None,
        };
        let v = self.eval(value)?;
        let repl: Vec<char> = self.scalar(v)?.to_string_value()?.chars().collect();
        let at = start as usize - 1;
        let count = repl
            .len()
            .min(length.unwrap_or(usize::MAX))
            .min(chars.len() - at);
        chars[at..at + count].copy_from_slice(&repl[..count]);
        self.assign(target, Value::String(chars.into_iter().collect()), false)
    }

    /// `Erase a`: a fixed-size array keeps its bounds with every element
    /// reset; a dynamic one is deallocated.
    pub(super) fn exec_erase(&mut self, targets: &[Expr]) -> VbaResult<()> {
        for target in targets {
            let Expr::Ident(name) = target else {
                return Err(VbaError::type_mismatch());
            };
            let slot = self
                .find_slot(name)
                .ok_or_else(|| VbaError::variable_not_defined(name))?;
            let (arr, fixed) = {
                let var = slot.borrow();
                match &var.value {
                    Value::Array(a) => ((**a).clone(), var.fixed_array),
                    _ => return Err(VbaError::type_mismatch()),
                }
            };
            let erased = if fixed {
                let mut arr = arr;
                for v in arr.data.iter_mut() {
                    *v = match v {
                        Value::Udt(u) => {
                            let def = self.types.get(&u.type_name.to_ascii_lowercase()).cloned();
                            match def {
                                Some(def) => self.new_udt(&def)?,
                                None => arr.elem_type.default_value(),
                            }
                        }
                        _ => arr.elem_type.default_value(),
                    };
                }
                arr
            } else {
                VbArray::new(Vec::new(), arr.elem_type)
            };
            slot.borrow_mut().value = Value::Array(Box::new(erased));
        }
        Ok(())
    }
}

//! The host interface: how an Office-like application plugs into the runtime.

use std::cell::RefCell;
use std::collections::BTreeMap;
use std::rc::Rc;

use crate::error::{VbaError, VbaResult};
use crate::object::ObjectRef;
use crate::value::Value;

/// Implemented by the embedding application (Excel-, Word-like hosts, tests).
pub trait Host {
    /// Resolves a global name not defined in VBA code: `Application`,
    /// `ActiveSheet`, `ThisWorkbook`, host constants like `xlUp`.
    fn global(&self, _name: &str) -> Option<Value> {
        None
    }

    /// The application object whose members are in scope unqualified, so that
    /// `Range("A1")` means `Application.Range("A1")`.
    fn global_object(&self) -> Option<ObjectRef> {
        None
    }

    /// Output of `Debug.Print` (the Immediate window).
    fn debug_print(&mut self, text: &str);

    /// `MsgBox`; returns the clicked button (`vbOK` = 1).
    fn msgbox(&mut self, prompt: &str, _buttons: i32, _title: &str) -> VbaResult<i32> {
        self.debug_print(&format!("[MsgBox] {prompt}"));
        Ok(1)
    }

    /// `InputBox`; `None` means Cancel.
    fn inputbox(
        &mut self,
        _prompt: &str,
        _title: &str,
        default: &str,
    ) -> VbaResult<Option<String>> {
        Ok(Some(default.to_string()))
    }

    /// `CreateObject("Scripting.Dictionary")` and friends.
    fn create_object(&mut self, prog_id: &str) -> VbaResult<Value> {
        Err(VbaError::new(
            429,
            format!("ActiveX component can't create object: {prog_id}"),
        ))
    }

    /// `New ClassName` for host-provided classes.
    fn new_object(&mut self, class: &str) -> VbaResult<Value> {
        Err(VbaError::new(429, format!("Unknown class: {class}")))
    }

    /// Current local time as an OLE date (`Now`).
    fn now(&self) -> f64 {
        let secs = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|d| d.as_secs_f64())
            .unwrap_or(0.0);
        25_569.0 + secs / 86_400.0
    }

    /// Called periodically during execution — the place to honour `DoEvents`
    /// or a host-side cancel button. Returning `false` interrupts the macro.
    fn poll(&mut self) -> bool {
        true
    }

    // ---- files: `Open`, `Kill`, `Dir`, ... ---------------------------------
    //
    // Denied by default so an embedded engine cannot touch the disk unless
    // the host opts in; [`StdHost`] uses the real file system through
    // [`fs`], [`CaptureHost`] an in-memory one.

    /// Whole contents of a file (`Open ... For Input`, `FileLen`).
    fn read_file(&mut self, path: &str) -> VbaResult<Vec<u8>> {
        Err(fs::denied(path))
    }

    /// Replaces a file's contents (`Close` after `Output` / `Append`).
    fn write_file(&mut self, path: &str, _data: &[u8]) -> VbaResult<()> {
        Err(fs::denied(path))
    }

    /// `Kill path`.
    fn delete_file(&mut self, path: &str) -> VbaResult<()> {
        Err(fs::denied(path))
    }

    /// `MkDir path`.
    fn create_dir(&mut self, path: &str) -> VbaResult<()> {
        Err(fs::denied(path))
    }

    /// Names of the entries in a directory (`Dir`), without the directory.
    fn list_dir(&mut self, dir: &str) -> VbaResult<Vec<String>> {
        Err(fs::denied(dir))
    }
}

/// File access on the real file system, for hosts that allow it.
pub mod fs {
    use crate::error::{VbaError, VbaResult};

    pub fn denied(path: &str) -> VbaError {
        VbaError::new(75, format!("Path/File access error: {path}"))
    }

    fn io_error(path: &str, e: std::io::Error) -> VbaError {
        match e.kind() {
            std::io::ErrorKind::NotFound => VbaError::new(53, format!("File not found: {path}")),
            _ => VbaError::new(75, format!("Path/File access error: {path}: {e}")),
        }
    }

    pub fn read_file(path: &str) -> VbaResult<Vec<u8>> {
        std::fs::read(path).map_err(|e| io_error(path, e))
    }

    pub fn write_file(path: &str, data: &[u8]) -> VbaResult<()> {
        std::fs::write(path, data).map_err(|e| io_error(path, e))
    }

    pub fn delete_file(path: &str) -> VbaResult<()> {
        std::fs::remove_file(path).map_err(|e| io_error(path, e))
    }

    pub fn create_dir(path: &str) -> VbaResult<()> {
        std::fs::create_dir(path).map_err(|e| io_error(path, e))
    }

    pub fn list_dir(dir: &str) -> VbaResult<Vec<String>> {
        let dir = if dir.is_empty() { "." } else { dir };
        let mut names: Vec<String> = std::fs::read_dir(dir)
            .map_err(|e| io_error(dir, e))?
            .filter_map(|e| e.ok()?.file_name().into_string().ok())
            .collect();
        names.sort();
        Ok(names)
    }
}

/// Writes `Debug.Print` / `MsgBox` to stdout.
#[derive(Default)]
pub struct StdHost;

impl Host for StdHost {
    fn debug_print(&mut self, text: &str) {
        println!("{text}");
    }

    fn read_file(&mut self, path: &str) -> VbaResult<Vec<u8>> {
        fs::read_file(path)
    }

    fn write_file(&mut self, path: &str, data: &[u8]) -> VbaResult<()> {
        fs::write_file(path, data)
    }

    fn delete_file(&mut self, path: &str) -> VbaResult<()> {
        fs::delete_file(path)
    }

    fn create_dir(&mut self, path: &str) -> VbaResult<()> {
        fs::create_dir(path)
    }

    fn list_dir(&mut self, dir: &str) -> VbaResult<Vec<String>> {
        fs::list_dir(dir)
    }
}

/// Collects output in memory; handy for tests.
#[derive(Default, Clone)]
pub struct CaptureHost {
    pub output: Rc<RefCell<Vec<String>>>,
    pub globals: Vec<(String, Value)>,
    /// An in-memory file system: path → contents. Directories are implied
    /// by the paths (`C:\data\a.txt` lists under `C:\data`).
    pub files: Rc<RefCell<BTreeMap<String, Vec<u8>>>>,
}

impl CaptureHost {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn lines(&self) -> Vec<String> {
        self.output.borrow().clone()
    }

    pub fn file_text(&self, path: &str) -> Option<String> {
        self.files
            .borrow()
            .get(path)
            .map(|d| String::from_utf8_lossy(d).into_owned())
    }
}

impl Host for CaptureHost {
    fn global(&self, name: &str) -> Option<Value> {
        self.globals
            .iter()
            .find(|(n, _)| n.eq_ignore_ascii_case(name))
            .map(|(_, v)| v.clone())
    }

    fn debug_print(&mut self, text: &str) {
        self.output.borrow_mut().push(text.to_string());
    }

    fn read_file(&mut self, path: &str) -> VbaResult<Vec<u8>> {
        self.files
            .borrow()
            .get(path)
            .cloned()
            .ok_or_else(|| VbaError::new(53, format!("File not found: {path}")))
    }

    fn write_file(&mut self, path: &str, data: &[u8]) -> VbaResult<()> {
        self.files
            .borrow_mut()
            .insert(path.to_string(), data.to_vec());
        Ok(())
    }

    fn delete_file(&mut self, path: &str) -> VbaResult<()> {
        self.files
            .borrow_mut()
            .remove(path)
            .map(|_| ())
            .ok_or_else(|| VbaError::new(53, format!("File not found: {path}")))
    }

    fn create_dir(&mut self, _path: &str) -> VbaResult<()> {
        Ok(())
    }

    fn list_dir(&mut self, dir: &str) -> VbaResult<Vec<String>> {
        let dir = dir.trim_end_matches(['/', '\\']);
        Ok(self
            .files
            .borrow()
            .keys()
            .filter_map(|p| {
                let (d, name) = p.rsplit_once(['/', '\\']).unwrap_or(("", p));
                d.eq_ignore_ascii_case(dir).then(|| name.to_string())
            })
            .collect())
    }
}

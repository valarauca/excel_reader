
use super::calamine::{Excel, DataType,Range};
use std::path::PathBuf;
use std::fmt::Debug;
use std::io::Write;
use std::fs::File;
use super::cell::to_csv;

/*
 * All Errors are fatal
 */
fn exit_on_err<T: Debug>(msg: String, e: T) -> ! {
    println!("\n\nFatal Error\n{}\nError: {:?}\nBye\n\n",msg,e);
    ::std::process::exit(1);
}
macro_rules! err {
    ($e: expr, $fmt_str: expr => $($arg:expr),* ) => {
        match $e {
            Ok(x) => x,
            Err(errr) => exit_on_err(format!($fmt_str, $($arg),*), errr)
        }
    }
}

/// What we're doing
#[derive(Debug,Clone)]
pub enum Action {
    ListNames(PathBuf),
    Convert(PathBuf),
    HasVBA(PathBuf),
    VBARef(PathBuf),
    DumpVBA(PathBuf),
    VBAModuleNames(PathBuf)
}

/*
 * These are the actual actions
 */
fn list_names(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    let names = err!(excel.sheet_names(), "On reading names\nFile: {:?}" => p);
    for name in names {
        println!("{}", name);
    }
}
fn has_vba(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    if excel.has_vba() {
        println!("YES VBA");
    } else {
        println!("NO VBA");
    }
}
fn vba_ref(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    let mut vbaproj = err!(excel.vba_project(), "On reading vba proj\nFile: {:?}" => p);
    for r in vbaproj.get_references() {
        println!("{:#?}", r);
    }
}
fn vba_mod_names(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    let mut vbaproj = err!(excel.vba_project(), "On reading vba proj\nFile: {:?}" => p);
    for m in vbaproj.get_module_names() {
        println!("{:#?}", m);
    }
}
fn dump_vba(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    let mut vbaproj = err!(excel.vba_project(), "On reading vba proj\nFile: {:?}" => p);
    for m in vbaproj.get_module_names() {
        let mut path = p.parent().unwrap().to_path_buf();
        path.push(format!("{}.vba", m));
        let text = err!(vbaproj.get_module(m), "On reading vba module\nFile: {:?}\nModule: {:?}" => p, m);
        let mut f = err!(File::create(&path), "On creating vba module file\nFile Input: {:?}\nFile Output (failure): {:?}\nModule: {:?}" => p, &path, &m);
        err!(f.write_all(text.as_bytes()), "On writing file\nFile Output (failure): {:?}\nModule: {:?}" => &path, &m);
    }
}
fn dump_csv(p: &PathBuf) {
    let mut excel = err!(Excel::open(p),"On opening\nFile: {:?}" => p);
    let names = err!(excel.sheet_names(), "On reading names\nFile: {:?}" => p);
    for name in names {
        let mut path = p.parent().unwrap().to_path_buf();
        path.push(format!("{}.csv", name));
        let cells = err!(excel.worksheet_range(name.as_str()), "On reading worksheet data\n File: {:?}\nSheet: {:?}" => p, name);
        let sheet = to_csv(&cells);
        let mut f = err!(File::create(&path), "On creating csv file\nFile Input: {:?}\nFile Output (failure): {:?}\nSheet: {:?}" => p, &path, name);
        err!(f.write_all(sheet.as_bytes()), "On writing file\nFile Output (failure): {:?}\nSheet: {:?}" => &path, name);
    }
}

impl Action {

    /// Do what CLI told us
    pub fn exec(self) {
        match self {
            Action::ListNames(ref p) => list_names(p),
            Action::HasVBA(ref p) => has_vba(p),
            Action::VBARef(ref p) => vba_ref(p),
            Action::VBAModuleNames(ref p) => vba_mod_names(p),
            Action::DumpVBA(ref p) => dump_vba(p),
            Action::Convert(ref p) => dump_csv(p)
        }
    }
}

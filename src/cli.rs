use super::clap::{App,Arg,ArgMatches};
use super::action::Action;
use std::path::PathBuf;


const ABOUT: &'static str = "
Reads Excel Files

";

pub fn fetch<'a>() -> ArgMatches<'a> {
    App::new("excel reader")
        .author("Cody Laeder")
        .version("1.0.0")
        .about(ABOUT)
        .arg(Arg::with_name("input")
             .short("i")
             .long("input")
             .value_name("PATH")
             .takes_value(true)
             .help("Path to input xls, xlsx, xlsm, xlsb"))
        .arg(Arg::with_name("sheets")
             .long("sheet_names")
             .takes_value(false)
             .help("Displays sheets"))
        .arg(Arg::with_name("csv")
            .long("convert_csv")
            .takes_value(false)
            .help("Writes each sheet of the CSV to the filesystem."))
        .arg(Arg::with_name("has_vba")
            .long("has_vba")
            .takes_value(false)
            .help("States if the excel file has VBA"))
        .arg(Arg::with_name("vba_ref")
            .long("vba_ref")
            .takes_value(false)
            .help("Lists references of internal VBA files."))
        .arg(Arg::with_name("dump_vba")
            .long("dump_vba")
            .takes_value(false)
            .help("Writes VBA files to the filesystem."))
        .arg(Arg::with_name("vba_module_names")
            .long("vba_module_names")
            .takes_value(false)
            .help("lists names of vba modules"))
        .get_matches()
}

pub fn to_exec<'a>(x: &ArgMatches<'a> ) {
    let p = match x.value_of("input") {
        Option::None => {
            println!("\n\nNo input file Supplied\nPlease use -h or --help to see help\n\n");
            ::std::process::exit(0);
        },
        Option::Some(x) => x
    };
    if x.is_present("csv") {
        Action::Convert(PathBuf::from(p)).exec();
    }
    if x.is_present("sheets") {
        Action::ListNames(PathBuf::from(p)).exec();
    }
    if x.is_present("has_vba") {
        Action::HasVBA(PathBuf::from(p)).exec();
    }
    if x.is_present("vba_ref") {
        Action::VBARef(PathBuf::from(p)).exec();
    }
    if x.is_present("vba_module_names") {
        Action::VBAModuleNames(PathBuf::from(p)).exec();
    }
    if x.is_present("dump_vba") {
        Action::DumpVBA(PathBuf::from(p)).exec();
    }
    ::std::process::exit(0);
}

excel_reader
---

CLI tool for interacting with xls, xlsx, xlsm, xlsb files

####Manual
```
excel reader 1.0.0

Reads Excel Files



USAGE:
    excel_reader.exe [FLAGS] [OPTIONS]

FLAGS:
        --convert_csv         Writes each sheet of the xl* file as a CSV to the filesystem.
        --dump_vba            Writes VBA files to the filesystem.
        --has_vba             States if the excel file has VBA
    -h, --help                Prints help information
        --sheet_names         Displays sheets
        --vba_module_names    lists names of vba modules
        --vba_ref             Lists references of internal VBA files.
    -V, --version             Prints version information

OPTIONS:
    -i, --input <PATH>    Path to input xls, xlsx, xlsm, xlsb
```

####Dependent Library/Thanks:

clap v2.20.0 thanks to Kevin K. [Repository](https://github.com/kbknapp/clap-rs) MIT-Licensed

calamine v0.3.2 thanks to Johann Tuffle [Repository](https://github.com/tafia/calamine) MIT-Licensed

####Building MacOS/Linux/_ BSD

On POSIX you will need `gcc` (or a compariable C compiler), `rustc` (stable), and `cargo`

1. within the `excel_reader/` directory run `cargo build --release`
2. The finished binary will be located at `excel_reader/target/release/excel_reader.exe`

####Building Windows MSVC (Visual Studio Toolchain)

1. Ensure you have [Visual Studio C++ Build Tools](http://landinghub.visualstudio.com/visual-cpp-build-tools)
2. Install [Rustup](https://www.rust-lang.org/en-US/install.html) which will install cargo+rustc (use default options)
3. within the `excel_reader/` directory run `cargo build --release`
4. The finished binary will be located at `excel_reader/target/release/excel_reader.exe`

####Building: Windows GNU (mingw)

I couldn't get this to work

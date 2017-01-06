extern crate clap;
extern crate calamine;


mod cli;
mod action;
mod cell;

fn main() {

    let args = cli::fetch();
    cli::to_exec(&args);

}

extern crate calamine;
extern crate csv;
extern crate getopts;

use calamine::{open_workbook_auto, Reader};
use csv::Writer;
use getopts::Options;
use std::env;
use std::process;
use std::path::PathBuf;

fn main() {
    let args: Vec<String> = env::args().collect();
    let program = args[0].clone();

    let mut opts = Options::new();
    opts.optopt("o", "output", "set output file name", "OUTPUT");
    opts.optopt("s", "sheet", "set sheet name", "SHEET");
    opts.optflag("h", "help", "print this help menu");
    let matches = match opts.parse(&args[1..]) {
        Ok(m) => m,
        Err(err) => {
            println!("{}\n", err);
            print_usage(&program, opts);
            process::exit(1);
        },
    };

    if matches.opt_present("h") {
        print_usage(&program, opts);
        return;
    }

    let input_path = if !matches.free.is_empty() {
        matches.free[0].clone()
    } else {
        print_usage(&program, opts);
        process::exit(1);
    };

    let input_path = PathBuf::from(input_path);

    let output_path = match matches.opt_str("o") {
        Some(path) => PathBuf::from(path),
        None => input_path.with_extension("csv"),
    };

    let mut xl = match open_workbook_auto(&input_path) {
        Ok(xl) => xl,
        Err(err) => {
            println!("Failed to read input file: {}", err);
            process::exit(1);
        },
    };

    let mut csv = match Writer::from_path(output_path) {
        Ok(csv) => csv,
        Err(err) => {
            println!("Failed to open output file: {}", err);
            process::exit(1);
        },
    };

    if xl.sheet_names().is_empty() {
        println!("Workbook has no sheets");
        process::exit(1);
    }

    let sheet_name = matches.opt_str("s").unwrap_or_else(|| xl.sheet_names()[0].clone());
    let range = match xl.worksheet_range(&sheet_name) {
        Some(Ok(range)) => range,
        Some(Err(err)) => {
            println!("Error reading worksheet: {}", err);
            process::exit(1);
        },
        None => {
            println!("Worksheet not found: {}", sheet_name);
            process::exit(1);
        },
    };

    for r in range.rows() {
        if let Err(err) = csv.write_record(r.iter().map(|c| c.to_string())) {
            println!("Failed to write CSV row: {}", err);
            process::exit(1);
        }
    }
}

fn print_usage(program: &str, opts: Options) {
    let brief = format!("Usage: {} FILE [options]", program);
    print!("{}", opts.usage(&brief));
}

extern crate quicli;
// use quicli::prelude::*;
use structopt::{self, StructOpt};
extern crate clap_log_flag;
extern crate clap_verbosity_flag;

#[allow(unused_imports)]
use calamine::{open_workbook_auto, DataType, Range, Reader};
use csv;
use serde_json::{self, Map, Value};

use std::error::Error;
#[allow(unused_imports)]
use std::io::{BufRead, BufReader, BufWriter, Cursor, Read, Seek, SeekFrom, Write};
use std::path::PathBuf;

#[derive(Debug, StructOpt)]
#[structopt(
    about = "Convert excel to csv",
    raw(setting = "structopt::clap::AppSettings::ColoredHelp")
)]
pub struct Cli {
    #[structopt(parse(from_os_str), help = "Which excel to convert")]
    pub path: PathBuf,
    #[structopt(long = "sheet", help = "Which sheet to convert")]
    pub sheet: Option<String>,
    #[structopt(
        long = "json",
        short = "J",
        help = "output in JSON instead <experimental>"
    )]
    pub json: bool,
    #[structopt(
        long = "first",
        short = "f",
        help = "process the 1st if multiple sheets were found. Or use '--sheet' to clarify"
    )]
    pub first: bool,
    #[structopt(flatten)]
    pub verbose: clap_verbosity_flag::Verbosity,
    #[structopt(flatten)]
    pub log: clap_log_flag::Log,
}

fn main() -> Result<(), Box<Error>> {
    let args = Cli::from_args();
    args.log.log_all(Some(args.verbose.log_level()))?;

    let path = &args.path;

    let mut dest = Cursor::new(Vec::new());
    // let dest = ::std::io::stdout();
    let mut xl = open_workbook_auto(&path).unwrap();

    let sheets = xl.sheet_names().to_owned();
    let sheet = match &args.sheet {
        Some(sheet) => sheet,
        None => {
            if sheets.len() > 1 {
                if args.first {
                    eprintln!(
                        "The file has {} sheets but we choose the FIRST one only: {}",
                        sheets.len(),
                        &sheets[0],
                    );
                } else {
                    panic!("Which sheet to convert? Candidates are: {:?}", sheets);
                }
            }
            &sheets[0]
        }
    };

    let range = xl.worksheet_range(&sheet).unwrap().unwrap();
    {
        let mut wtr = csv::Writer::from_writer(&mut dest);
        for r in range.rows() {
            let row = r
                .into_iter()
                .map(|c| match c {
                    DataType::Empty => "".to_string(),
                    DataType::String(ref s) => format!("{}", s),
                    DataType::Float(ref f) => format!("{}", f),
                    DataType::Int(ref i) => format!("{}", i),
                    DataType::Error(ref e) => format!("{:?}", e),
                    DataType::Bool(ref b) => format!("{}", b),
                })
                .collect::<Vec<String>>();
            wtr.write_record(&row)?;
            // println!("{:?}", row);
        }
    };
    dest.seek(SeekFrom::Start(0))?;

    if args.json {
        type Row = Map<String, Value>;
        let buff = BufReader::new(dest);
        let mut rdr = csv::Reader::from_reader(buff);
        let array: Vec<Row> = rdr
            .deserialize()
            .into_iter()
            .map(|result| {
                let record: Row = result.unwrap();
                record
            })
            .collect();
        println!("{}", serde_json::to_string_pretty(&array)?);
    } else {
        let s = String::from_utf8(dest.into_inner())?;
        println!("{}", s);
    }

    Ok(())
}

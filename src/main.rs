use std::io;
use clap::Parser;

use rs_xsheet2jsonl::xpath2sheet2rows2stdout;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None, arg_required_else_help = true)]
struct Args {
    /// Path to the XLSX file
    #[arg(short, long)]
    xlsx_path: String,

    /// Name of the sheet to read
    #[arg(short, long, default_value = "Sheet1")]
    sheet_name: String,
}

fn main() -> Result<(), io::Error> {
    let args = Args::parse();
    xpath2sheet2rows2stdout(&args.xlsx_path, &args.sheet_name)?;
    Ok(())
}

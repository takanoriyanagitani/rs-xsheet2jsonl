use std::io;
use std::path::Path;

use io::BufWriter;
use io::Write;

use serde_json::Map;
use serde_json::Value;
use serde_json::json;

use chrono::NaiveDateTime;

use calamine::CellType;
use calamine::Data;
use calamine::ExcelDateTime;
use calamine::Range;
use calamine::Reader;
use calamine::Rows;
use calamine::Xlsx;

pub struct XlRange<T>(pub Range<T>);

impl<T> XlRange<T>
where
    T: CellType,
{
    pub fn cell_pos_end(&self) -> Option<(u32, u32)> {
        self.0.end()
    }

    pub fn cell_pos_start(&self) -> Option<(u32, u32)> {
        self.0.start()
    }
}

impl<T> XlRange<T>
where
    T: CellType,
{
    pub fn rows(&self) -> Rows<'_, T> {
        self.0.rows()
    }
}

impl XlRange<Data> {
    pub fn rows2writer<W>(&self, wtr: &mut W) -> Result<(), io::Error>
    where
        W: FnMut(usize, &[Data]) -> Result<(), io::Error>,
    {
        let irow = self.rows();
        for (rno, row) in irow.enumerate() {
            wtr(rno, row)?;
        }
        Ok(())
    }
}

pub fn date2val(dt: ExcelDateTime) -> Value {
    let ondt: Option<NaiveDateTime> = dt.as_datetime();
    match ondt {
        Some(ndt) => json!(ndt.to_string()),
        None => json!(dt.to_string()),
    }
}

pub fn dat2val(dat: &Data) -> Value {
    match dat {
        Data::Int(v) => json!(v),
        Data::Float(v) => json!(v),
        Data::String(v) => json!(v),
        Data::Bool(v) => json!(v),
        Data::DateTime(v) => date2val(*v),
        Data::DateTimeIso(v) => json!(v),
        Data::DurationIso(v) => json!(v),
        Data::Error(v) => json!(v.to_string()),
        Data::Empty => json!(null),
    }
}

impl XlRange<Data> {
    pub fn rows2io_writer<W>(&self, mut wtr: W) -> Result<(), io::Error>
    where
        W: Write,
    {
        let mut row: Map<String, Value> = Map::default();
        self.rows2writer(&mut move |rowno: usize, cols: &[Data]| {
            row.clear();
            for (cno, col) in cols.iter().enumerate() {
                row.insert(format!("{cno}"), dat2val(col));
            }
            row.insert("row_number".into(), json!(rowno));
            serde_json::to_writer(&mut wtr, &row)?;
            writeln!(&mut wtr)?;
            Ok(())
        })?;
        Ok(())
    }
}

pub fn xpath2sheet2rows2writer<P, W>(xpath: P, sheet: &str, mut wtr: W) -> Result<(), io::Error>
where
    W: Write,
    P: AsRef<Path>,
{
    let mut wkbk: Xlsx<_> = calamine::open_workbook(xpath).map_err(io::Error::other)?;
    let rng: Range<Data> = wkbk.worksheet_range(sheet).map_err(io::Error::other)?;
    XlRange(rng).rows2io_writer(&mut wtr)?;
    wtr.flush()
}

pub fn xpath2sheet2rows2stdout<P>(xpath: P, sheet: &str) -> Result<(), io::Error>
where
    P: AsRef<Path>,
{
    let o = io::stdout();
    let mut ol = o.lock();
    xpath2sheet2rows2writer(xpath, sheet, BufWriter::new(&mut ol))?;
    ol.flush()
}

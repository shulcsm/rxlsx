use crate::shared::Strings;
use crate::types::PYTHON_TYPES;
use pyo3::ffi::{PyDateTime_GET_DAY, PyDateTime_GET_MONTH, PyDateTime_GET_YEAR, PyObject};
use pyo3::prelude::*;
use pyo3::types::PyAny;

#[derive(Debug)]
pub enum CellValue {
    Bool(bool),
    SharedString(usize),
    InlineString(String),
    Number(f64),
    // Dates required Format
    // Date(f64),
    // DateTime(f64),     // figure out tz
    // Currency
    Formula(String),
    // Error?
}

unsafe fn epcoh_from_date(ob: *mut PyObject) -> f64 {
    let y = PyDateTime_GET_YEAR(ob);
    let m = PyDateTime_GET_MONTH(ob) as u32;
    let d = PyDateTime_GET_DAY(ob) as u32;
    chrono::NaiveDate::from_ymd(y, m, d)
        .and_hms(0, 0, 0)
        .timestamp() as f64
}

pub fn cell_value(ob: &PyAny, strings: &mut Strings) -> PyResult<Option<CellValue>> {
    Ok(unsafe {
        let ob_type = ob.get_type_ptr();

        if ob_type == PYTHON_TYPES.int {
            Some(CellValue::Number(ob.extract()?))
        } else if ob_type == PYTHON_TYPES.float {
            Some(CellValue::Number(ob.extract()?))
        } else if ob_type == PYTHON_TYPES.decimal {
            Some(CellValue::Number(ob.extract()?))
        } else if ob_type == PYTHON_TYPES.str {
            // Could be empty, when do we want inline

            let string: String = ob.extract()?;

            if string.starts_with("=") {
                Some(CellValue::Formula(string[1..].into()))
            } else {
                let idx = strings.insert(string.into());
                Some(CellValue::SharedString(idx))
            }
        } else if ob_type == PYTHON_TYPES.bool {
            Some(CellValue::Bool(ob.extract()?))
        } else if ob_type == PYTHON_TYPES.none {
            None
        // } else if ob_type == PYTHON_TYPES.date {
        //     Some(CellValue::Date(epcoh_from_date(ob.into_ptr())))
        } else {
            unimplemented!("UNHANDLED TYPE {}", ob);
        }
    })
}

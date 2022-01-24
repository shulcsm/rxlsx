use crate::shared::Strings;
use crate::types::PYTHON_TYPES;
use chrono::Datelike;
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
    Date(f64),
    // DateTime(f64),     // figure out tz
    // Currency
    Formula(String),
    // Error?
}

unsafe fn serial_from_pydate(ob: *mut PyObject) -> f64 {
    // @TODO different modes, test
    // https://support.microsoft.com/en-us/office/date-systems-in-excel-e7fe7167-48a9-4b96-bb53-5612a800b487?ui=en-us&rs=en-us&ad=us
    // https://support.microsoft.com/en-us/office/date-function-e36c0c8c-4104-49da-ab83-82328b832349?ui=en-us&rs=en-us&ad=us
    // https://docs.microsoft.com/en-US/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
    let y = PyDateTime_GET_YEAR(ob);
    let m = PyDateTime_GET_MONTH(ob) as u32;
    let d = PyDateTime_GET_DAY(ob) as u32;
    let chrono = chrono::NaiveDate::from_ymd(y, m, d).num_days_from_ce();
    (chrono - 719_163 + 25569) as f64
}
pub fn cell_value(ob: &PyAny, strings: &mut Strings) -> PyResult<Option<CellValue>> {
    // @TODO
    // ob.is_none();

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
        } else if ob_type == PYTHON_TYPES.date {
            Some(CellValue::Date(serial_from_pydate(ob.into_ptr())))
        } else {
            unimplemented!("UNHANDLED TYPE {}", ob);
        }
    })
}

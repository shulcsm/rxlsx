use crate::shared::Strings;
use crate::types::PYTHON_TYPES;
use pyo3::types::PyAny;

use pyo3::prelude::*;

#[derive(Debug)]
pub enum CellValue {
    Bool(bool),
    SharedString(usize),
    InlineString(String),
    Number(f64),
    // Date
    // Currency
    Formula(String),
    // Error?
}

pub fn cell_value(ob: &PyAny, strings: &mut Strings) -> PyResult<CellValue> {
    unsafe {
        let ob_type = ob.get_type_ptr();

        if ob_type == PYTHON_TYPES.int {
            return Ok(CellValue::Number(ob.extract()?));
        } else if ob_type == PYTHON_TYPES.float {
            return Ok(CellValue::Number(ob.extract()?));
        } else if ob_type == PYTHON_TYPES.decimal {
            return Ok(CellValue::Number(ob.extract()?));
        } else if ob_type == PYTHON_TYPES.str {
            // Could be empty, when do we want inline

            let string: String = ob.extract()?;

            if string.starts_with("=") {
                return Ok(CellValue::Formula(string[1..].into()));
            } else {
                let idx = strings.insert(string.into());
                return Ok(CellValue::SharedString(idx));
            }
        } else if ob_type == PYTHON_TYPES.bool {
            return Ok(CellValue::Bool(ob.extract()?));
        }
    };
    unimplemented!("UNHANDLED TYPE {}", ob);
}

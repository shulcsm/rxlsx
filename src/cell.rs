use crate::types::PYTHON_TYPES;
use pyo3::conversion::FromPyObject;
use pyo3::types::PyAny;

use pyo3::prelude::*;

#[derive(Debug)]
pub enum CellValue {
    Bool(bool),
    String(String), // InlineString | SharedString
    Number(f64),
    // Date
    // Currency
    Formula(String),
    // Error?
}

impl<'source> FromPyObject<'source> for CellValue {
    fn extract(ob: &'source PyAny) -> PyResult<Self> {
        unsafe {
            let ob_type = ob.get_type_ptr();

            if ob_type == PYTHON_TYPES.int {
                return Ok(CellValue::Number(ob.extract()?));
            } else if ob_type == PYTHON_TYPES.float {
                return Ok(CellValue::Number(ob.extract()?));
            } else if ob_type == PYTHON_TYPES.decimal {
                return Ok(CellValue::Number(ob.extract()?));
            } else if ob_type == PYTHON_TYPES.str {
                let string: &str = ob.extract()?;
                if string.starts_with("=") {
                    return Ok(CellValue::Formula(string[1..].into()));
                } else {
                    return Ok(CellValue::String(string.into()));
                }
            } else if ob_type == PYTHON_TYPES.bool {
                return Ok(CellValue::Bool(ob.extract()?));
            }
        };
        unimplemented!("UNHANDLED TYPE {}", ob);
    }
}

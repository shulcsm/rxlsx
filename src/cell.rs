use crate::types::PYTHON_TYPES;
use pyo3::conversion::FromPyObject;
use pyo3::types::PyAny;

use pyo3::prelude::*;

#[derive(Debug)]
pub enum CellValue {
    Bool(bool),
    InlineString(String),
    Number(f64),
    // Date
    // Currency
    // Formula
    // SharedString
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
                return Ok(CellValue::InlineString(ob.extract()?));
            } else if ob_type == PYTHON_TYPES.bool {
                return Ok(CellValue::Bool(ob.extract()?));
            }
        };
        unimplemented!("UNHANDLED TYPE {}", ob);
    }
}

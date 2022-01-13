use pyo3::prelude::*;
mod util;
mod workbook;
mod worksheet;
mod writer;

#[pymodule]
fn rxlsx(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_class::<workbook::Workbook>()?;
    m.add_class::<worksheet::Worksheet>()?;
    Ok(())
}
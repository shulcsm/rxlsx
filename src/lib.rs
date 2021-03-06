use pyo3::prelude::*;
mod cell;
mod shared;
mod types;
mod util;
mod workbook;
mod worksheet;
mod writer;

#[pymodule]
fn rxlsx(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_class::<workbook::Workbook>()?;
    m.add_class::<worksheet::Worksheet>()?;
    m.add_class::<worksheet::Column>()?;
    Ok(())
}

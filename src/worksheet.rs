use crate::cell::CellValue;
use crate::workbook::Workbook;
use pyo3::prelude::*;
use pyo3::types::PyList;
use std::collections::HashMap;

pub type Cells = HashMap<(usize, usize), CellValue>;

#[pyclass]
pub struct Worksheet {
    pub parent: Py<Workbook>,
    #[pyo3(get)]
    pub title: String, // @TODO name
    #[pyo3(get)]
    pub max_row_idx: usize,
    #[pyo3(get)]
    pub max_col_idx: usize,
    pub cells: Cells,
    // "_dense_cells",
    // "_sparse_cells",
    // "_styles",
    // "_row_styles",
    // "_col_styles",
    // "_parent",
    // "_merges",
    // "_attributes",
    // "_panes",
    // "_show_grid_lines",
    // "auto_filter"
}

#[pymethods]
impl Worksheet {
    // should be any iterable but TypeError: argument 'values': 'list' object cannot be converted to 'Iterator'
    fn append(&mut self, values: &PyList) {
        let next_row = self.max_row_idx + 1;

        let cols = values.len();
        for (idx, v) in values.iter().enumerate() {
            let c: CellValue = v.extract().unwrap();

            self.cells.insert((next_row, idx + 1), c);
        }
        self.max_row_idx = next_row;
        if cols > self.max_col_idx {
            self.max_col_idx = cols;
        }
    }
}

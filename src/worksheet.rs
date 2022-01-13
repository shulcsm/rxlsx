use crate::workbook::Workbook;
// use derive_more::{Add, Display, From, Into};
use pyo3::prelude::*;
use pyo3::types::PyList;
use std::collections::HashMap;
/*
class DataTypes(object):
    BOOLEAN = 0
    DATE = 1
    ERROR = 2
    INLINE_STRING = 3
    NUMBER = 4
    SHARED_STRING = 5
    STRING = 6
    FORMULA = 7
    EXCEL_BASE_DATE = datetime(1900, 1, 1, 0, 0, 0)
*/

// // @TODO ref to doc and 1 - 1,048,576
// #[derive(Eq, PartialEq, From, Add, Debug, Clone, Copy, Hash)]
// pub struct Row(pub usize);
// // @TODO ref to doc and 1 - 16,384
// #[derive(Eq, PartialEq, From, Add, Debug, Clone, Copy, Hash)]
// pub struct Col(pub usize);

// #[derive(Eq, PartialEq, Hash)]
// pub struct Coords(pub Row, pub Col);

#[derive(Debug, Clone, PartialEq, FromPyObject)]
pub enum CellValue {
    Bool(bool),
    Str(String),
    Number(f64),
    // Date
    // Currency
}

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

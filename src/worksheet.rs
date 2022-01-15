use crate::cell::{cell_value, CellValue};
use crate::shared::SharedRef;
use pyo3::prelude::*;
use pyo3::types::PyList;
use std::collections::HashMap;

pub type Cells = HashMap<(usize, usize), CellValue>;

#[pyclass]
pub struct Worksheet {
    #[pyo3(get)]
    pub name: String,
    #[pyo3(get)]
    pub max_row_idx: usize,
    #[pyo3(get)]
    pub max_col_idx: usize,
    // @TODO private
    pub cells: Cells,
    shared: SharedRef,
}

impl Worksheet {
    pub fn new(name: String, shared: SharedRef) -> Self {
        Worksheet {
            name,
            max_row_idx: 0,
            max_col_idx: 0,
            cells: HashMap::new(),
            shared,
        }
    }
}

#[pymethods]
impl Worksheet {
    fn insert(&mut self, row_index: usize, column_index: usize, value: &PyAny) {
        let mut shared = self.shared.write().unwrap();
        let strings = &mut shared.strings;

        strings.insert("a".into());

        let c: CellValue = cell_value(value, strings).unwrap();
        self.cells.insert((row_index, column_index), c);
    }

    // should be any iterable but TypeError: argument 'values': 'list' object cannot be converted to 'Iterator'
    fn append(&mut self, values: &PyList) {
        let next_row = self.max_row_idx + 1;

        let cols = values.len();
        for (idx, v) in values.iter().enumerate() {
            self.insert(next_row, idx + 1, v)
        }
        self.max_row_idx = next_row;
        if cols > self.max_col_idx {
            self.max_col_idx = cols;
        }
    }
}

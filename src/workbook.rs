use crate::util::zip_from_path_or_file;
use crate::worksheet::Worksheet;
use crate::writer::WorkbookWriter;
use pyo3::prelude::*;
use std::collections::HashMap;

#[pyclass]
pub struct Workbook {
    pub worksheets: Vec<Py<Worksheet>>,
    //  shared_strings
    //     )  # "_styles", "_items", "_has_macros", "_encoding", "_writer")
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Workbook {
        Workbook {
            worksheets: Vec::new(),
        }
    }

    fn _refcount(self_: PyRefMut<Self>, py: Python) -> PyResult<isize> {
        Ok(Py::from(self_).get_refcnt(py) - 1)
    }

    fn create_sheet(
        self_: Py<Self>,
        py: Python,
        title: Option<String>,
        index: Option<usize>,
    ) -> Py<Worksheet> {
        let ws = Py::new(
            py,
            Worksheet {
                parent: self_.clone(),
                title: title.unwrap_or("Sheet".to_owned()),
                max_row_idx: 0,
                max_col_idx: 0,
                cells: HashMap::new(),
            },
        )
        .unwrap();
        let mut s: PyRefMut<Self> = self_.extract(py).unwrap();
        if let Some(i) = index {
            // @TODO ensure len or error
            s.worksheets.insert(i, ws.clone());
        } else {
            s.worksheets.push(ws.clone());
        }
        ws
    }

    fn remove_sheet(&mut self, worksheet: Py<Worksheet>) -> PyResult<()> {
        self.worksheets.remove(self.index(worksheet));
        Ok(())
    }

    fn index(&self, worksheet: Py<Worksheet>) -> usize {
        // @TODO IndexError
        self.worksheets
            .iter()
            .position(|x| *x == worksheet)
            .expect("sheet not found")
    }

    fn save(&self, _py: Python, path_or_file: PyObject) -> PyResult<()> {
        let mut zip = zip_from_path_or_file(path_or_file)?;
        let writer = WorkbookWriter::new(self, &mut zip);
        writer.save()
    }
}

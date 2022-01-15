use crate::shared::{Shared, SharedRef};
use crate::util::zip_from_path_or_file;
use crate::worksheet::Worksheet;
use crate::writer::WorkbookWriter;
use pyo3::prelude::*;
use std::sync::{Arc, RwLock};

#[pyclass]
pub struct Workbook {
    pub worksheets: Vec<Py<Worksheet>>,
    pub shared: SharedRef,
}

#[pymethods]
impl Workbook {
    #[new]
    fn new() -> Workbook {
        Workbook {
            worksheets: Vec::new(),
            shared: Arc::new(RwLock::new(Shared::new())),
        }
    }

    fn _refcount(self_: PyRefMut<Self>, py: Python) -> PyResult<isize> {
        Ok(Py::from(self_).get_refcnt(py) - 1)
    }

    fn create_sheet(
        &mut self,
        py: Python,
        name: Option<String>,
        index: Option<usize>,
    ) -> Py<Worksheet> {
        let ws = Py::new(
            py,
            Worksheet::new(name.unwrap_or("Sheet".to_owned()), self.shared.clone()),
        )
        .unwrap();

        if let Some(i) = index {
            // @TODO ensure len or error
            self.worksheets.insert(i, ws.clone());
        } else {
            self.worksheets.push(ws.clone());
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

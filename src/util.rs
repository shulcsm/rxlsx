use pyo3::exceptions::PyTypeError;
use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyString};
use std::fs::File;
use std::io;
use std::io::{Seek, SeekFrom, Write};

// https://github.com/omerbenamram/pyo3-file/blob/master/src/lib.rs

#[derive(Debug)]
pub struct PyFile {
    inner: PyObject,
}

fn pyerr_to_io_err(e: PyErr) -> io::Error {
    let gil = Python::acquire_gil();
    let py = gil.python();
    let e_as_object: PyObject = e.into_py(py);

    match e_as_object.call_method(py, "__str__", (), None) {
        Ok(repr) => match repr.extract::<String>(py) {
            Ok(s) => io::Error::new(io::ErrorKind::Other, s),
            Err(_e) => io::Error::new(io::ErrorKind::Other, "An unknown error has occurred"),
        },
        Err(_) => io::Error::new(io::ErrorKind::Other, "Err doesn't have __str__"),
    }
}

pub enum Writeable {
    NativeFile(File),
    PythonFile(PyFile),
}

impl Write for Writeable {
    fn write(&mut self, buf: &[u8]) -> Result<usize, io::Error> {
        match self {
            Writeable::NativeFile(ref mut f) => f.write(buf),
            Writeable::PythonFile(ref mut f) => {
                let gil = Python::acquire_gil();
                let py = gil.python();
                let pybytes = PyBytes::new(py, buf);
                let number_bytes_written = f
                    .inner
                    .call_method(py, "write", (pybytes,), None)
                    .map_err(pyerr_to_io_err)?;
                number_bytes_written.extract(py).map_err(pyerr_to_io_err)
            }
        }
    }

    fn flush(&mut self) -> Result<(), io::Error> {
        match self {
            Writeable::NativeFile(ref mut f) => f.flush(),
            Writeable::PythonFile(ref mut f) => {
                let gil = Python::acquire_gil();
                let py = gil.python();
                f.inner
                    .call_method(py, "flush", (), None)
                    .map_err(pyerr_to_io_err)?;
                Ok(())
            }
        }
    }
}

impl Seek for Writeable {
    fn seek(&mut self, pos: SeekFrom) -> Result<u64, io::Error> {
        match self {
            Writeable::NativeFile(ref mut f) => f.seek(pos),
            Writeable::PythonFile(ref mut f) => {
                let gil = Python::acquire_gil();
                let py = gil.python();
                let (whence, offset) = match pos {
                    SeekFrom::Start(i) => (0, i as i64),
                    SeekFrom::Current(i) => (1, i as i64),
                    SeekFrom::End(i) => (2, i as i64),
                };
                let new_position = f
                    .inner
                    .call_method(py, "seek", (offset, whence), None)
                    .map_err(pyerr_to_io_err)?;
                new_position.extract(py).map_err(pyerr_to_io_err)
            }
        }
    }
}

impl Writeable {
    fn new(path_or_file: PyObject) -> PyResult<Writeable> {
        // @TODO handle pathlib.Path
        let gil = Python::acquire_gil();
        let py = gil.python();

        if let Ok(string_ref) = path_or_file.cast_as::<PyString>(py) {
            let string_path = string_ref.to_string_lossy().to_string();
            let file = File::create(string_path)?;
            return Ok(Writeable::NativeFile(file));
        }

        if path_or_file.getattr(py, "write").is_err() {
            return Err(PyErr::new::<PyTypeError, _>(
                "Object does not have a .write() method.",
            ));
        }
        if path_or_file.getattr(py, "seek").is_err() {
            return Err(PyErr::new::<PyTypeError, _>(
                "Object does not have a .seek() method.",
            ));
        }
        dbg!("writable");
        Ok(Writeable::PythonFile(PyFile {
            inner: path_or_file,
        }))
    }
}

pub type Zip = zip::ZipWriter<Writeable>;

pub fn zip_from_path_or_file(path_or_file: PyObject) -> PyResult<Zip> {
    let writeable = Writeable::new(path_or_file)?;

    Ok(zip::ZipWriter::new(writeable))
}

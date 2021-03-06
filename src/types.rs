use once_cell::sync::Lazy;
use pyo3::ffi::{
    PyBool_Type, PyDateTimeAPI, PyFloat_Type, PyLong_Type, PyTypeObject, PyUnicode_Type, Py_None,
};
use pyo3::prelude::*;
use pyo3::AsPyPointer;

pub struct PythonTypes {
    pub int: *mut PyTypeObject,
    pub decimal: *mut PyTypeObject,
    pub float: *mut PyTypeObject,
    pub str: *mut PyTypeObject,
    pub bool: *mut PyTypeObject,
    pub none: *mut PyTypeObject,
    pub date: *mut PyTypeObject,
    pub date_time: *mut PyTypeObject,
}

pub static mut PYTHON_TYPES: Lazy<PythonTypes> = Lazy::new(|| unsafe {
    let gil = Python::acquire_gil();
    let py = gil.python();
    let sys = py.import("decimal").unwrap();
    let decimal = sys.getattr("Decimal").unwrap();

    PythonTypes {
        int: &mut PyLong_Type as *mut PyTypeObject,
        decimal: decimal.as_ptr() as *mut PyTypeObject,
        float: &mut PyFloat_Type as *mut PyTypeObject,
        str: &mut PyUnicode_Type as *mut PyTypeObject,
        bool: &mut PyBool_Type as *mut PyTypeObject,
        none: (*Py_None()).ob_type,
        date: PyDateTimeAPI.DateType,
        date_time: PyDateTimeAPI.DateTimeType,
    }
});

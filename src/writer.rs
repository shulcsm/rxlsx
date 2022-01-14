use crate::cell::CellValue;
use crate::util::Zip;
use crate::workbook::Workbook;
use crate::worksheet::Worksheet;
use pyo3::prelude::*;
use std::io::Write;

pub struct WorkbookWriter<'a> {
    inner: &'a Workbook,
    writer: &'a mut Zip,
}

impl<'a> WorkbookWriter<'a> {
    pub fn new(workbook: &'a Workbook, writer: &'a mut Zip) -> Self {
        WorkbookWriter {
            inner: workbook,
            writer,
        }
    }

    fn file(&mut self, name: &str) {
        let options: zip::write::FileOptions = zip::write::FileOptions::default()
            .compression_method(zip::CompressionMethod::Deflated)
            .unix_permissions(0o600);

        self.writer.start_file(name, options).unwrap();
    }

    fn write_docprops_app(&mut self) {
        self.file("docProps/app.xml");

        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>Microsoft Excel</Application>
    <AppVersion>3.0</AppVersion>
</Properties>"#;
        self.writer.write(xml).unwrap();
    }

    fn write_docprops_core(&mut self) {
        self.file("docProps/core.xml");

        // @TODO: creator, created, modified
        //  Utc::now().format("%Y-%m-%dT%H:%M:00Z").to_string(),
        // <cp:corePropertie
        //     <dc:creator>rxlsx</dc:creator>
        //     <dcterms:created xsi:type="dcterms:W3CDTF">{{ date }}</dcterms:created>
        //     <dcterms:modified xsi:type="dcterms:W3CDTF">{{ date }}</dcterms:modified>
        // </cp:coreProperties>

        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
    xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
    xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>"#;
        self.writer.write(xml).unwrap();
    }

    fn write_content_type(&mut self) {
        self.file("[Content_Types].xml");

        // @TODO: macros, themes, shared strings

        let head = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"#;
        self.writer.write(head).unwrap();

        for idx in 1..=self.inner.worksheets.len() {
            let wb = format!(
                "<Override PartName=\"/xl/worksheets/sheet{}.xml\"
                    ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
                idx
            );
            self.writer.write(wb.as_bytes()).unwrap();
        }

        let tail = br#"</Types>"#;
        self.writer.write(tail).unwrap();
    }

    fn write_rels_rels(&mut self) {
        self.file("_rels/.rels");

        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#;

        self.writer.write(xml).unwrap();
    }

    fn write_xl_styles(&mut self) {
        self.file("xl/styles.xml");

        // @TODO actual styles
        let xml = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"></styleSheet>"#;

        self.writer.write(xml).unwrap();
    }

    fn write_xl_workbook(&mut self) {
        self.file("xl/workbook.xml");

        // @TODO workbookPr, workbookProtection, calcPr(calc on load?), bookViews?, definedNames?

        let head = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr/>
    <sheets>"#;
        self.writer.write(head).unwrap();

        let gil = Python::acquire_gil();
        let py = gil.python();

        // @TODO should rId not collide with rels?
        for (idx, sheet) in self.inner.worksheets.iter().enumerate() {
            let wb = format!(
                "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
                sheet.borrow(py).title,
                idx + 1,
                idx + 3 // 1 theme, 2 styles
            );
            self.writer.write(wb.as_bytes()).unwrap();
        }

        let tail = br#"
    </sheets>
</workbook>"#;
        self.writer.write(tail).unwrap();
    }

    fn write_xl_workbook_rels(&mut self) {
        self.file("xl/_rels/workbook.xml.rels");

        // @TODO <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
        // @TODO shared strings
        let head = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#;

        self.writer.write(head).unwrap();

        for idx in 1..=self.inner.worksheets.len() {
            let wb = format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>", idx + 2, idx);
            self.writer.write(wb.as_bytes()).unwrap();
        }

        let tail = br#"
</Relationships>"#;

        self.writer.write(tail).unwrap();
    }

    pub fn save(mut self) -> PyResult<()> {
        // @TODO ZipResult into PyResult
        self.write_docprops_app();
        self.write_docprops_core();
        self.write_content_type();
        self.write_rels_rels();
        self.write_xl_styles();
        self.write_xl_workbook();
        self.write_xl_workbook_rels();

        let gil = Python::acquire_gil();
        let py = gil.python();

        for (idx, pyws) in self.inner.worksheets.iter().enumerate() {
            let id = idx + 1;
            let file_name = format!("xl/worksheets/sheet{}.xml", id);
            self.file(&file_name);

            let ws = pyws.borrow(py);
            let sheet_writer = WorksheetWriter::new(ws, &mut self.writer);
            sheet_writer.save()?;
        }
        Ok(())
    }
}

pub struct WorksheetWriter<'a> {
    inner: PyRef<'a, Worksheet>,
    writer: &'a mut Zip,
}

pub fn column_to_letter(index: usize) -> String {
    let alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    let mut result = String::from("");
    let mut work = index - 1;
    loop {
        result.push(alphabet.chars().nth(work % 26).unwrap());
        if work < 26 {
            break;
        }
        work = (work / 26) - 1;
    }
    result.chars().rev().collect()
}

fn escape_str_value(s: &str) -> String {
    s.replace("&", "&amp;").replace("<", "&lt;")
}

pub fn index_to_coord(column_index: usize, row_index: usize) -> String {
    format!("{}{}", column_to_letter(column_index), row_index)
}

impl<'a> WorksheetWriter<'a> {
    pub fn new(worksheet: PyRef<'a, Worksheet>, writer: &'a mut Zip) -> Self {
        WorksheetWriter {
            inner: worksheet,
            writer,
        }
    }

    fn write_cell(&self, buff: &mut Vec<u8>, row_index: usize, column_index: usize) {
        let coord = index_to_coord(column_index, row_index);

        if let Some(value) = self.inner.cells.get(&(row_index, column_index)) {
            match value {
                CellValue::Number(value) => {
                    let r = format!("<c r=\"{}\"><v>{}</v></c>", coord, value);
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::Bool(value) => {
                    let v = if *value { 1 } else { 0 };
                    let r = format!("<c r=\"{}\" t=\"b\"><v>{}</v></c>", coord, v);
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::InlineString(value) => {
                    let r = format!(
                        "<c r=\"{}\" t=\"str\"><v>{}</v></c>",
                        coord,
                        escape_str_value(&value)
                    );
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::Formula(value) => {
                    let r = format!(
                        "<c r=\"{}\" t=\"str\"><f>{}</f></c>",
                        coord,
                        escape_str_value(&value)
                    );
                    buff.write(r.as_bytes()).unwrap();
                }
            };
        }
    }

    fn write_row(&self, buff: &mut Vec<u8>, row_index: usize) {
        let head = format!("<row r=\"{}\">", row_index);

        buff.write(head.as_bytes()).unwrap();

        for column_index in 1..=self.inner.max_col_idx {
            self.write_cell(buff, row_index, column_index)
        }

        let tail = br#"</row>"#;

        buff.write(tail).unwrap();
    }

    fn write_sheet_data(&mut self) {
        let head = br#"<sheetData>"#;
        self.writer.write_all(head).unwrap();

        let mut buff: Vec<u8> = Vec::new();
        for row_index in 1..=self.inner.max_row_idx {
            self.write_row(&mut buff, row_index);
            self.writer.write_all(&buff).unwrap();
            buff.clear();
        }

        let tail = br#"</sheetData>"#;
        self.writer.write_all(tail).unwrap();
    }

    pub fn save(mut self) -> PyResult<()> {
        let head = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#;
        self.writer.write_all(head).unwrap();

        self.write_sheet_data();

        let tail = br#"</worksheet>"#;
        self.writer.write_all(tail).unwrap();
        Ok(())
    }
}

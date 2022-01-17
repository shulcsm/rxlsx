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

        // @TODO: macros, themes

        let head = br#"<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
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

        for (idx, sheet) in self.inner.worksheets.iter().enumerate() {
            let wb = format!(
                "<sheet name=\"{}\" sheetId=\"{}\" r:id=\"rId{}\"/>",
                sheet.borrow(py).name,
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
        let head = br#"<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>"#;

        self.writer.write(head).unwrap();

        let total = self.inner.worksheets.len();
        for idx in 1..=total {
            let wb = format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet{}.xml\"/>", idx + 2, idx);
            self.writer.write(wb.as_bytes()).unwrap();
        }

        let shared_idx = total + 3;
        let shared = format!("<Relationship Id=\"rId{}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>", shared_idx);
        self.writer.write(shared.as_bytes()).unwrap();

        let tail = br#"
</Relationships>"#;

        self.writer.write(tail).unwrap();
    }

    fn write_xls_shared_strings(&mut self) {
        let strings = &self.inner.shared.read().unwrap().strings;

        self.file("xl/sharedStrings.xml");

        let head = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#;
        self.writer.write(head).unwrap();

        let wb = format!("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"{}\" uniqueCount=\"{}\">", strings.size, strings.total);
        self.writer.write(wb.as_bytes()).unwrap();

        for val in strings.index.keys() {
            let tag = format!(
                "<si><t xml:space=\"preserve\">{}</t></si>",
                escape_str_value(val)
            );
            self.writer.write(tag.as_bytes()).unwrap();
        }

        let tail = br#"
</sst>"#;
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
        self.write_xls_shared_strings();
        Ok(())
    }
}

pub struct WorksheetWriter<'a> {
    inner: PyRef<'a, Worksheet>,
    writer: &'a mut Zip,
}

pub fn column_to_letter(index: usize) -> String {
    // 0 indexed
    let mut col = index - 1;
    if col < 26 {
        ((b'A' + col as u8) as char).to_string()
    } else {
        let mut rev = String::new();
        while col >= 26 {
            let c = col % 26;
            rev.push((b'A' + c as u8) as char);
            col -= c;
            col /= 26;
        }
        rev.chars().rev().collect()
    }
}

fn escape_str_value(s: &str) -> String {
    s.replace("&", "&amp;").replace("<", "&lt;")
}

pub fn index_to_coord(column_index: usize, row_index: usize) -> String {
    column_to_letter(column_index) + row_index.to_string().as_str()
}

impl<'a> WorksheetWriter<'a> {
    pub fn new(worksheet: PyRef<'a, Worksheet>, writer: &'a mut Zip) -> Self {
        WorksheetWriter {
            inner: worksheet,
            writer,
        }
    }

    fn write_cell(&mut self, buff: &mut Vec<u8>, row_index: usize, column_index: usize) {
        let coord = index_to_coord(column_index, row_index);

        if let Some(value) = self.inner.cells.get(&(row_index, column_index)) {
            match value {
                CellValue::Number(ref value) => {
                    let r = format!("<c r=\"{}\"><v>{}</v></c>", coord, value);
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::Bool(ref value) => {
                    let v = if *value { 1 } else { 0 };
                    let r = format!("<c r=\"{}\" t=\"b\"><v>{}</v></c>", coord, v);
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::SharedString(value) => {
                    let r = format!("<c r=\"{}\" t=\"s\"><v>{}</v></c>", coord, value);
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::InlineString(ref value) => {
                    let r = format!(
                        "<c r=\"{}\" t=\"str\"><v>{}</v></c>",
                        coord,
                        escape_str_value(value)
                    );
                    buff.write(r.as_bytes()).unwrap();
                }
                CellValue::Formula(ref value) => {
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

    fn write_row(&mut self, buff: &mut Vec<u8>, row_index: usize) {
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

        let gil = Python::acquire_gil();
        let py = gil.python();

        if !self.inner.columns.is_empty() {
            self.writer.write_all(br#"<cols>"#).unwrap();
            for column_index in 1..=self.inner.max_col_idx {
                if let Some(col) = self.inner.columns.get(&column_index) {
                    let c = col.borrow(py);

                    self.writer.write_all(
                        format!(
                            "<col min=\"{}\" max=\"{}\" width=\"{}\" customWidth=\"1\"/>\n",
                            &column_index, &column_index, c.width
                        )
                        .as_bytes(),
                    )?;
                } else {
                    // is this necessary?

                    self.writer.write_all(
                        format!(
                            "<col min=\"{}\" max=\"{}\" />\n",
                            &column_index, &column_index
                        )
                        .as_bytes(),
                    )?;
                }
            }
            self.writer.write_all(br#"</cols>"#).unwrap();
        }

        self.write_sheet_data();

        let tail = br#"</worksheet>"#;
        self.writer.write_all(tail).unwrap();
        Ok(())
    }
}

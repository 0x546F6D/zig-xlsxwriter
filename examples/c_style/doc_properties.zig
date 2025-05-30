//
// Example of setting document properties such as Author, Title, etc., for an
// Excel spreadsheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-doc_properties.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Create a properties structure and set some of the fields.
    var properties = lxw.lxw_doc_properties{
        .title = "This is an example spreadsheet",
        .subject = "With document properties",
        .author = "John McNamara",
        .manager = "Dr. Heinz Doofenshmirtz",
        .company = "of Wolves",
        .category = "Example spreadsheets",
        .keywords = "Sample, Example, Properties",
        .comments = "Created with libxlsxwriter",
        .status = "Quo",
    };

    // Set the properties in the workbook.
    _ = lxw.workbook_set_properties(workbook, &properties);

    // Add some text to the file.
    _ = lxw.worksheet_set_column(worksheet, 0, 0, 50, null);
    _ = lxw.worksheet_write_string(
        worksheet,
        0,
        0,
        "Select 'Workbook Properties' to see properties.",
        null,
    );

    _ = lxw.workbook_close(workbook);
}

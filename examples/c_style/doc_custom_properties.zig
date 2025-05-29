//
// Example of setting custom document properties for an Excel spreadsheet
// using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-doc_custom_properties.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);
    var datetime = lxw.lxw_datetime{
        .year = 2016,
        .month = 12,
        .day = 12,
        .hour = 0,
        .min = 0,
        .sec = 0.0,
    };

    // Set some custom document properties in the workbook.
    _ = lxw.workbook_set_custom_property_string(
        workbook,
        "Checked by",
        "Eve",
    );
    _ = lxw.workbook_set_custom_property_datetime(
        workbook,
        "Date completed",
        &datetime,
    );
    _ = lxw.workbook_set_custom_property_number(
        workbook,
        "Document number",
        12345,
    );
    _ = lxw.workbook_set_custom_property_number(
        workbook,
        "Reference number",
        1.2345,
    );
    _ = lxw.workbook_set_custom_property_boolean(
        workbook,
        "Has Review",
        1,
    );
    _ = lxw.workbook_set_custom_property_boolean(
        workbook,
        "Signed off",
        0,
    );

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

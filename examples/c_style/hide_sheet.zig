//
// Example of how to hide a worksheet using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-hide_sheet.xlsx");
    const worksheet1 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet2 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet3 = lxw.workbook_add_worksheet(workbook, null);

    // Hide Sheet2. It won't be visible until it is unhidden in Excel.
    _ = lxw.worksheet_hide(worksheet2);

    _ = lxw.worksheet_write_string(worksheet1, 0, 0, "Sheet2 is hidden", null);
    _ = lxw.worksheet_write_string(worksheet2, 0, 0, "Now it's my turn to find you!", null);
    _ = lxw.worksheet_write_string(worksheet3, 0, 0, "Sheet2 is hidden", null);

    // Make the first column wider to make the text clearer.
    _ = lxw.worksheet_set_column(worksheet1, 0, 0, 30, null);
    _ = lxw.worksheet_set_column(worksheet2, 0, 0, 30, null);
    _ = lxw.worksheet_set_column(worksheet3, 0, 0, 30, null);

    _ = lxw.workbook_close(workbook);
}

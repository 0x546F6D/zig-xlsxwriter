//
// An example of how to hide rows and columns using the libxlsxwriter
// library.
//
// In order to hide rows without setting each one, (of approximately 1 million
// rows), Excel uses an optimization to hide all rows that don't have data. In
// Libxlsxwriter we replicate that using the worksheet_set_default_row()
// function.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    // Create a new workbook and add a worksheet
    const workbook = lxw.workbook_new("zig-hide_row_col.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write some data
    _ = lxw.worksheet_write_string(worksheet, 0, 3, "Some hidden columns.", null);
    _ = lxw.worksheet_write_string(worksheet, 7, 0, "Some hidden rows.", null);

    // Hide all rows without data
    _ = lxw.worksheet_set_default_row(worksheet, 15, lxw.LXW_TRUE);

    // Set the height of empty rows that we want to display even if it is
    // the default height
    var row: lxw.lxw_row_t = 1;
    while (row <= 6) : (row += 1) {
        _ = lxw.worksheet_set_row(worksheet, row, 15, null);
    }

    // Columns can be hidden explicitly. This doesn't increase the file size
    var options = lxw.lxw_row_col_options{
        .hidden = 1,
        .level = 0,
        .collapsed = 0,
    };

    // Use COLS macro equivalent for "G:XFD" range
    _ = lxw.worksheet_set_column_opt(worksheet, 6, 16383, 8.43, null, &options);

    _ = lxw.workbook_close(workbook);
}

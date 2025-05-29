//
// Example of using libxlsxwriter for writing large files in constant memory
// mode.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const row_max: u32 = 1000;
    const col_max: u16 = 50;

    // Set the workbook options
    var options = lxw.lxw_workbook_options{
        .constant_memory = lxw.LXW_TRUE,
        .tmpdir = null,
        .use_zip64 = lxw.LXW_FALSE,
        .output_buffer = null,
        .output_buffer_size = null,
    };

    // Create a new workbook with options
    const workbook = lxw.workbook_new_opt("zig-constant_memory.xlsx", &options);
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    var row: u32 = 0;
    while (row < row_max) : (row += 1) {
        var col: u16 = 0;
        while (col < col_max) : (col += 1) {
            _ = lxw.worksheet_write_number(worksheet, row, col, 123.45, null);
        }
    }

    _ = lxw.workbook_close(workbook);
}

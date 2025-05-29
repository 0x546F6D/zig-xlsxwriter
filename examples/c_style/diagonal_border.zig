//
// A simple formatting example that demonstrates how to add diagonal
// cell borders using the libxlsxwriter library.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-diagonal_border.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Add some diagonal border formats.
    const format1 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_diag_type(format1, lxw.LXW_DIAGONAL_BORDER_UP);

    const format2 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_diag_type(format2, lxw.LXW_DIAGONAL_BORDER_DOWN);

    const format3 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_diag_type(format3, lxw.LXW_DIAGONAL_BORDER_UP_DOWN);

    const format4 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_diag_type(format4, lxw.LXW_DIAGONAL_BORDER_UP_DOWN);
    _ = lxw.format_set_diag_border(format4, lxw.LXW_BORDER_HAIR);
    _ = lxw.format_set_diag_color(format4, lxw.LXW_COLOR_RED);

    _ = lxw.worksheet_write_string(worksheet, 2, 1, "Text", format1);
    _ = lxw.worksheet_write_string(worksheet, 5, 1, "Text", format2);
    _ = lxw.worksheet_write_string(worksheet, 8, 1, "Text", format3);
    _ = lxw.worksheet_write_string(worksheet, 11, 1, "Text", format4);

    _ = lxw.workbook_close(workbook);
}

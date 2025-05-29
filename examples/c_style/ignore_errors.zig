//
// An example of turning off worksheet cells errors/warnings using
// libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-ignore_errors.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write strings that looks like numbers. This will cause an Excel warning.
    _ = lxw.worksheet_write_string(worksheet, 1, 2, "123", null);
    _ = lxw.worksheet_write_string(worksheet, 2, 2, "123", null);

    // Write a divide by zero formula. This will also cause an Excel warning.
    _ = lxw.worksheet_write_formula(worksheet, 4, 2, "=1/0", null);
    _ = lxw.worksheet_write_formula(worksheet, 5, 2, "=1/0", null);

    // Turn off some of the warnings:
    _ = lxw.worksheet_ignore_errors(worksheet, lxw.LXW_IGNORE_NUMBER_STORED_AS_TEXT, "C3");
    _ = lxw.worksheet_ignore_errors(worksheet, lxw.LXW_IGNORE_EVAL_ERROR, "C6");

    // Write some descriptions for the cells and make the column wider for clarity.
    _ = lxw.worksheet_set_column(worksheet, 1, 1, 16, null);
    _ = lxw.worksheet_write_string(worksheet, 1, 1, "Warning:", null);
    _ = lxw.worksheet_write_string(worksheet, 2, 1, "Warning turned off:", null);
    _ = lxw.worksheet_write_string(worksheet, 4, 1, "Warning:", null);
    _ = lxw.worksheet_write_string(worksheet, 5, 1, "Warning turned off:", null);

    _ = lxw.workbook_close(workbook);
}

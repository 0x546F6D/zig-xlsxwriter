// An a simple example of how to add conditional formatting to an
// libxlsxwriter file.
//
// See conditional_format.c for a more comprehensive example.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-conditional_format1.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Write some sample data.
    const row_b1 = lxw.lxw_name_to_row("B1");
    const col_b1 = lxw.lxw_name_to_col("B1");
    _ = lxw.worksheet_write_number(worksheet, row_b1, col_b1, 34, null);

    const row_b2 = lxw.lxw_name_to_row("B2");
    const col_b2 = lxw.lxw_name_to_col("B2");
    _ = lxw.worksheet_write_number(worksheet, row_b2, col_b2, 32, null);

    const row_b3 = lxw.lxw_name_to_row("B3");
    const col_b3 = lxw.lxw_name_to_col("B3");
    _ = lxw.worksheet_write_number(worksheet, row_b3, col_b3, 31, null);

    const row_b4 = lxw.lxw_name_to_row("B4");
    const col_b4 = lxw.lxw_name_to_col("B4");
    _ = lxw.worksheet_write_number(worksheet, row_b4, col_b4, 35, null);

    const row_b5 = lxw.lxw_name_to_row("B5");
    const col_b5 = lxw.lxw_name_to_col("B5");
    _ = lxw.worksheet_write_number(worksheet, row_b5, col_b5, 36, null);

    const row_b6 = lxw.lxw_name_to_row("B6");
    const col_b6 = lxw.lxw_name_to_col("B6");
    _ = lxw.worksheet_write_number(worksheet, row_b6, col_b6, 30, null);

    const row_b7 = lxw.lxw_name_to_row("B7");
    const col_b7 = lxw.lxw_name_to_col("B7");
    _ = lxw.worksheet_write_number(worksheet, row_b7, col_b7, 38, null);

    const row_b8 = lxw.lxw_name_to_row("B8");
    const col_b8 = lxw.lxw_name_to_col("B8");
    _ = lxw.worksheet_write_number(worksheet, row_b8, col_b8, 38, null);

    const row_b9 = lxw.lxw_name_to_row("B9");
    const col_b9 = lxw.lxw_name_to_col("B9");
    _ = lxw.worksheet_write_number(worksheet, row_b9, col_b9, 32, null);

    // Add a format with red text.
    const custom_format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_font_color(custom_format, lxw.LXW_COLOR_RED);

    // Create a conditional format object.
    var conditional_format: lxw.lxw_conditional_format = std.mem.zeroes(lxw.lxw_conditional_format);

    // Set the format type: a cell conditional:
    conditional_format.type = lxw.LXW_CONDITIONAL_TYPE_CELL;

    // Set the criteria to use:
    conditional_format.criteria = lxw.LXW_CONDITIONAL_CRITERIA_LESS_THAN;

    // Set the value to which the criteria will be applied:
    conditional_format.value = 33;

    // Set the format to use if the criteria/value applies:
    conditional_format.format = custom_format;

    // Now apply the format to data range.
    // RANGE("B1:B9") expands to first_row, first_col, last_row, last_col
    const first_row = lxw.lxw_name_to_row("B1");
    const first_col = lxw.lxw_name_to_col("B1");
    const last_row = lxw.lxw_name_to_row("B9");
    const last_col = lxw.lxw_name_to_col("B9");
    _ = lxw.worksheet_conditional_format_range(worksheet, first_row, first_col, last_row, last_col, &conditional_format);

    _ = lxw.workbook_close(workbook);
}

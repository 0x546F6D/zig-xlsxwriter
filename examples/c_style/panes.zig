const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    var row: u32 = 0;
    var col: u16 = 0;

    // Create a new workbook and add some worksheets.
    const workbook = lxw.workbook_new("zig-panes.xlsx");

    const worksheet1 = lxw.workbook_add_worksheet(workbook, "Panes 1");
    const worksheet2 = lxw.workbook_add_worksheet(workbook, "Panes 2");
    const worksheet3 = lxw.workbook_add_worksheet(workbook, "Panes 3");
    const worksheet4 = lxw.workbook_add_worksheet(workbook, "Panes 4");

    // Set up some formatting and text to highlight the panes.
    const header = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_align(header, lxw.LXW_ALIGN_CENTER);
    _ = lxw.format_set_align(header, lxw.LXW_ALIGN_VERTICAL_CENTER);
    _ = lxw.format_set_fg_color(header, 0xD7E4BC);
    _ = lxw.format_set_bold(header);
    _ = lxw.format_set_border(header, lxw.LXW_BORDER_THIN);

    const center = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_align(center, lxw.LXW_ALIGN_CENTER);

    //
    // Example 1. Freeze pane on the top row.
    //
    _ = lxw.worksheet_freeze_panes(worksheet1, 1, 0);

    // Some sheet formatting.
    _ = lxw.worksheet_set_column(worksheet1, 0, 8, 16, null);
    _ = lxw.worksheet_set_row(worksheet1, 0, 20, null);
    _ = lxw.worksheet_set_selection(worksheet1, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 0;
    while (col < 9) : (col += 1) {
        _ = lxw.worksheet_write_string(worksheet1, 0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 100) : (row += 1) {
        col = 0;
        while (col < 9) : (col += 1) {
            _ = lxw.worksheet_write_number(worksheet1, row, col, @floatFromInt(row + 1), center);
        }
    }

    //
    // Example 2. Freeze pane on the left column.
    //
    _ = lxw.worksheet_freeze_panes(worksheet2, 0, 1);

    // Some sheet formatting.
    _ = lxw.worksheet_set_column(worksheet2, 0, 0, 16, null);
    _ = lxw.worksheet_set_selection(worksheet2, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    row = 0;
    while (row < 50) : (row += 1) {
        _ = lxw.worksheet_write_string(worksheet2, row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = lxw.worksheet_write_number(worksheet2, row, col, @floatFromInt(col), center);
        }
    }

    //
    // Example 3. Freeze pane on the top row and left column.
    //
    _ = lxw.worksheet_freeze_panes(worksheet3, 1, 1);

    // Some sheet formatting.
    _ = lxw.worksheet_set_column(worksheet3, 0, 25, 16, null);
    _ = lxw.worksheet_set_row(worksheet3, 0, 20, null);
    _ = lxw.worksheet_write_string(worksheet3, 0, 0, "", header);
    _ = lxw.worksheet_set_selection(worksheet3, 4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        _ = lxw.worksheet_write_string(worksheet3, 0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        _ = lxw.worksheet_write_string(worksheet3, row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = lxw.worksheet_write_number(worksheet3, row, col, @floatFromInt(col), center);
        }
    }

    //
    // Example 4. Split pane on the top row and left column.
    //
    // The divisions must be specified in terms of row and column dimensions.
    // The default row height is 15 and the default column width is 8.43
    //
    _ = lxw.worksheet_split_panes(worksheet4, 15, 8.43);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        _ = lxw.worksheet_write_string(worksheet4, 0, col, "Scroll", center);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        _ = lxw.worksheet_write_string(worksheet4, row, 0, "Scroll", center);

        col = 1;
        while (col < 26) : (col += 1) {
            _ = lxw.worksheet_write_number(worksheet4, row, col, @floatFromInt(col), center);
        }
    }

    _ = lxw.workbook_close(workbook);
}

//
// Example of how to set Excel worksheet tab colors using libxlsxwriter.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-tab_colors.xlsx");

    // Set up some worksheets.
    const worksheet1 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet2 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet3 = lxw.workbook_add_worksheet(workbook, null);
    const worksheet4 = lxw.workbook_add_worksheet(workbook, null);

    // Set the tab colors.
    _ = lxw.worksheet_set_tab_color(worksheet1, lxw.LXW_COLOR_RED);
    _ = lxw.worksheet_set_tab_color(worksheet2, lxw.LXW_COLOR_GREEN);
    _ = lxw.worksheet_set_tab_color(worksheet3, 0xFF9900); // Orange.

    // worksheet4 will have the default color.
    _ = lxw.worksheet_write_string(worksheet4, 0, 0, "Hello", null);

    _ = lxw.workbook_close(workbook);
}

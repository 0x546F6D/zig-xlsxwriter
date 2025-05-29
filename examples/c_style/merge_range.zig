const std = @import("std");
const lxw = @import("lxw");

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-merge_range.xlsx");
    const worksheet = lxw.workbook_add_worksheet(workbook, null);
    const merge_format = lxw.workbook_add_format(workbook);

    // Configure a format for the merged range.
    _ = lxw.format_set_align(merge_format, lxw.LXW_ALIGN_CENTER);
    _ = lxw.format_set_align(merge_format, lxw.LXW_ALIGN_VERTICAL_CENTER);
    _ = lxw.format_set_bold(merge_format);
    _ = lxw.format_set_bg_color(merge_format, lxw.LXW_COLOR_YELLOW);
    _ = lxw.format_set_border(merge_format, lxw.LXW_BORDER_THIN);

    // Increase the cell size of the merged cells to highlight the formatting.
    _ = lxw.worksheet_set_column(worksheet, 1, 3, 12, null);
    _ = lxw.worksheet_set_row(worksheet, 3, 30, null);
    _ = lxw.worksheet_set_row(worksheet, 6, 30, null);
    _ = lxw.worksheet_set_row(worksheet, 7, 30, null);

    // Merge 3 cells.
    _ = lxw.worksheet_merge_range(worksheet, 3, 1, 3, 3, "Merged Range", merge_format);

    // Merge 3 cells over two rows.
    _ = lxw.worksheet_merge_range(worksheet, 6, 1, 7, 3, "Merged Range", merge_format);

    _ = lxw.workbook_close(workbook);
}

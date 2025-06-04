pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    const merge_format = try workbook.addFormat();

    // Configure a format for the merged range.
    merge_format.setAlign(.center);
    merge_format.setAlign(.vertical_center);
    merge_format.setBold();
    merge_format.setBgColor(.yellow);
    merge_format.setBorder(.thin);

    // Increase the cell size of the merged cells to highlight the formatting.
    try worksheet.setColumn(.{ .first = 1, .last = 3 }, 12, .default, null);
    try worksheet.setRow(3, 30, .default, null);
    try worksheet.setRow(6, 30, .default, null);
    try worksheet.setRow(7, 30, .default, null);

    // Merge 3 cells.
    try worksheet.mergeRange(
        .{ .first_row = 3, .first_col = 1, .last_row = 3, .last_col = 3 },

        "Merged Range",
        merge_format,
    );

    // Merge 3 cells over two rows.
    try worksheet.mergeRange(
        .{ .first_row = 6, .first_col = 1, .last_row = 7, .last_col = 3 },
        "Merged Range",
        merge_format,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
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
    try worksheet.setColumn(1, 3, 12, .none);
    try worksheet.setRow(3, 30, .none);
    try worksheet.setRow(6, 30, .none);
    try worksheet.setRow(7, 30, .none);

    // Merge 3 cells.
    try worksheet.mergeRange(3, 1, 3, 3, "Merged Range", merge_format);

    // Merge 3 cells over two rows.
    try worksheet.mergeRange(6, 1, 7, 3, "Merged Range", merge_format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // get image path
    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Change some of the column widths for clarity
    try worksheet.setColumn(xwz.cols("A:B"), 30, .default, null);

    // Embed an image
    try worksheet.writeString(xwz.cell("A2"), "Embed an image in a cell:", .default);
    try worksheet.embedImage(xwz.cell("B2"), asset_path, null);

    // Make a row bigger and embed the image
    try worksheet.setRow(3, 72, .default, null);

    try worksheet.writeString(xwz.cell("A4"), "Embed an image in a cell:", .default);
    try worksheet.embedImage(xwz.cell("B4"), asset_path, null);

    // Make a row bigger and embed the image
    try worksheet.setRow(5, 150, .default, null);
    try worksheet.writeString(xwz.cell("A6"), "Embed an image in a cell:", .default);
    try worksheet.embedImage(xwz.cell("B6"), asset_path, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

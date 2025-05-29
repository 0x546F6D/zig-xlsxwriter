pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // get image path
    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Change some of the column widths for clarity
    const cols = xwz.cols("A:B");
    try worksheet.setColumn(cols.first_col, cols.last_col, 30, .none);

    // Embed an image
    var cell = xwz.cell("A2");
    try worksheet.writeString(cell.row, cell.col, "Embed an image in a cell:", .none);

    cell = xwz.cell("B2");
    try worksheet.embedImage(cell.row, cell.col, asset_path);

    // Make a row bigger and embed the image
    try worksheet.setRow(3, 72, .none);

    cell = xwz.cell("A4");
    try worksheet.writeString(cell.row, cell.col, "Embed an image in a cell:", .none);

    cell = xwz.cell("B4");
    try worksheet.embedImage(cell.row, cell.col, asset_path);

    // Make a row bigger and embed the image
    try worksheet.setRow(5, 150, .none);

    cell = xwz.cell("A6");
    try worksheet.writeString(cell.row, cell.col, "Embed an image in a cell:", .none);

    cell = xwz.cell("B6");
    try worksheet.embedImage(cell.row, cell.col, asset_path);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

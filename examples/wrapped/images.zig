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

    // Change some of the column widths for clarity.
    try worksheet.setColumn(xwz.cols("A:A"), 30, .default, null);

    // Insert an image.
    try worksheet.writeString(xwz.cell("A2"), "Insert an image in a cell:", .default);

    try worksheet.insertImage(xwz.cell("B2"), asset_path, null);

    // Insert an image offset in the cell.
    try worksheet.writeString(xwz.cell("A12"), "Insert an offset image:", .default);

    const options1 = xwz.ImageOptions{
        .x_offset = 15,
        .y_offset = 10,
    };

    try worksheet.insertImage(xwz.cell("B12"), asset_path, options1);

    // Insert an image with scaling.
    try worksheet.writeString(xwz.cell("A22"), "Insert a scaled image:", .default);

    const options2 = xwz.ImageOptions{
        .x_scale = 0.5,
        .y_scale = 0.5,
    };

    try worksheet.insertImage(xwz.cell("B22"), asset_path, options2);

    // Insert an image with a hyperlink.
    try worksheet.writeString(xwz.cell("A32"), "Insert an image with a hyperlink:", .default);

    const options3 = xwz.ImageOptions{ .url = "https://github.com/jmcnamara" };

    try worksheet.insertImage(xwz.cell("B32"), asset_path, options3);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

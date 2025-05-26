pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // get image path
    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Change some of the column widths for clarity.
    const cols = xlsxwriter.cols("A:A");
    try worksheet.setColumn(cols.first_col, cols.last_col, 30, .none);

    // Insert an image.
    var cell = xlsxwriter.cell("A2");
    try worksheet.writeString(cell.row, cell.col, "Insert an image in a cell:", .none);

    cell = xlsxwriter.cell("B2");
    try worksheet.insertImage(cell.row, cell.col, asset_path);

    // Insert an image offset in the cell.
    cell = xlsxwriter.cell("A12");
    try worksheet.writeString(cell.row, cell.col, "Insert an offset image:", .none);

    const options1 = xlsxwriter.ImageOptions{
        .x_offset = 15,
        .y_offset = 10,
    };

    cell = xlsxwriter.cell("B12");
    try worksheet.insertImageOpt(cell.row, cell.col, asset_path, options1);

    // Insert an image with scaling.
    cell = xlsxwriter.cell("A22");
    try worksheet.writeString(cell.row, cell.col, "Insert a scaled image:", .none);

    const options2 = xlsxwriter.ImageOptions{
        .x_scale = 0.5,
        .y_scale = 0.5,
    };

    cell = xlsxwriter.cell("B22");
    try worksheet.insertImageOpt(cell.row, cell.col, asset_path, options2);

    // Insert an image with a hyperlink.
    cell = xlsxwriter.cell("A32");
    try worksheet.writeString(cell.row, cell.col, "Insert an image with a hyperlink:", .none);

    const options3 = xlsxwriter.ImageOptions{ .url = "https://github.com/jmcnamara" };

    cell = xlsxwriter.cell("B32");
    try worksheet.insertImageOpt(cell.row, cell.col, asset_path, options3);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

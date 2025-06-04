pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Write some data
    try worksheet.writeString(.{ .col = 3 }, "Some hidden columns.", .default);
    try worksheet.writeString(.{ .row = 7 }, "Some hidden rows.", .default);

    // Hide all rows without data
    worksheet.setDefaultRow(15, true);

    // Set the height of empty rows that we want to display even if it is
    // the default height
    var row: u32 = 1;
    while (row <= 6) : (row += 1) try worksheet.setRow(row, xwz.def_row_height, .default, null);

    // Columns can be hidden explicitly. This doesn't increase the file size
    const options = xwz.RowColOptions{ .hidden = true };

    // Use COLS macro equivalent for "G:XFD" range
    try worksheet.setColumn(xwz.cols("G:XFD"), 8.43, .default, options);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet(null);
    const worksheet2 = try workbook.addWorkSheet(null);
    const worksheet3 = try workbook.addWorkSheet(null);

    // Hide Sheet2. It won't be visible until it is unhidden in Excel.
    worksheet2.hide();

    try worksheet1.writeString(.{}, "Sheet2 is hidden", .default);
    try worksheet2.writeString(.{}, "Now it's my turn to find you!", .default);
    try worksheet3.writeString(.{}, "Sheet2 is hidden", .default);

    // Make the first column wider to make the text clearer.
    try worksheet1.setColumn(.{}, 30, .default, null);
    try worksheet2.setColumn(.{}, 30, .default, null);
    try worksheet3.setColumn(.{}, 30, .default, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

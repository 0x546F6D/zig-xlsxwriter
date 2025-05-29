pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet(null);
    const worksheet2 = try workbook.addWorkSheet(null);
    const worksheet3 = try workbook.addWorkSheet(null);

    // Hide Sheet2. It won't be visible until it is unhidden in Excel.
    worksheet2.hide();

    try worksheet1.writeString(0, 0, "Sheet2 is hidden", .none);
    try worksheet2.writeString(0, 0, "Now it's my turn to find you!", .none);
    try worksheet3.writeString(0, 0, "Sheet2 is hidden", .none);

    // Make the first column wider to make the text clearer.
    try worksheet1.setColumn(0, 0, 30, .none);
    try worksheet2.setColumn(0, 0, 30, .none);
    try worksheet3.setColumn(0, 0, 30, .none);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

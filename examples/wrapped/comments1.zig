pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    try worksheet.writeString(0, 0, "Hello", .none);

    try worksheet.writeComment(0, 0, "This is a comment");
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

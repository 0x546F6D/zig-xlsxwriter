pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a format with date formatting.
    const format = try workbook.addFormat();
    format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 20, .none);

    // Write some Unix datetimes with formatting.

    // 1970-01-01. The Unix epoch.
    try worksheet.writeUnixTime(0, 0, 0, format);

    // 2000-01-01.
    try worksheet.writeUnixTime(1, 0, 1577836800, format);

    // 1900-01-01.
    try worksheet.writeUnixTime(2, 0, -2208988800, format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

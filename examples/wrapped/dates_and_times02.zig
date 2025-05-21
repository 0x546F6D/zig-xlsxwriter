// A datetime to display.
var datetime: xlsxwriter.DateTime = .{
    .year = 2013,
    .month = 2,
    .day = 28,
    .hour = 12,
    .min = 0,
    .sec = 0.0,
};

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a format with date formatting.
    const format = try workbook.addFormat();
    format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 20, .none);

    // Write the datetime without formatting.
    try worksheet.writeDateTime(0, 0, datetime, .none);

    // Write the datetime with formatting.
    try worksheet.writeDateTime(1, 0, datetime, format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

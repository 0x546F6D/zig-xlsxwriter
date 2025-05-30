// A number to display as a date.
const number = 41333.5;

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch |err| {
        std.debug.print("{}: {s}\n", .{ err, xwz.strError(err) });
    };
    const worksheet = try workbook.addWorkSheet(null);

    // Add a format with date formatting.
    const format = try workbook.addFormat();
    format.setNumFormat("mmm d yyyy hh:mm AM/PM");

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 20, .default);

    // Write the number without formatting.
    try worksheet.writeNumber(0, 0, number, .default);

    // Write the number with formatting. Note: the worksheet_write_datetime()
    // or worksheet_write_unixtime() functions are preferable for writing
    // dates and times. This is for demonstration purposes only.
    try worksheet.writeNumber(1, 0, number, format);
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

// A datetime to display.
var datetime: xlsxwriter.DateTime = .{
    .year = 2013,
    .month = 1,
    .day = 23,
    .hour = 12,
    .min = 30,
    .sec = 5.123,
};

// Examples date and time formats. In the output file compare how changing
// the format strings changes the appearance of the date.
const date_formats = [_][]const u8{
    "dd/mm/yy",
    "mm/dd/yy",
    "dd m yy",
    "d mm yy",
    "d mmm yy",
    "d mmmm yy",
    "d mmmm yyy",
    "d mmmm yyyy",
    "dd/mm/yy hh:mm",
    "dd/mm/yy hh:mm:ss",
    "dd/mm/yy hh:mm:ss.000",
    "hh:mm",
    "hh:mm:ss",
    "hh:mm:ss.000",
};

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format.
    const bold = try workbook.addFormat();
    bold.setBold();

    // Write the column headers.
    try worksheet.writeString(0, 0, "Formatted date", bold);
    try worksheet.writeString(0, 1, "Format", bold);

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 20, .none);

    // Write the same date and time using each of the above formats.
    for (date_formats, 1..) |date_format, row| {

        // Create a format for the date or time.
        const format = try workbook.addFormat();
        format.setNumFormat(date_format.ptr);
        format.setAlign(.left);

        // Write the datetime with each format.
        try worksheet.writeDateTime(@intCast(row), 0, datetime, format);

        // Also write the format string for comparison.
        try worksheet.writeString(@intCast(row), 1, date_format.ptr, .none);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

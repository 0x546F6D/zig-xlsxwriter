pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(.{}, 30, .default, null);

    // Add some formats.
    const format01 = try workbook.addFormat();
    const format02 = try workbook.addFormat();
    const format03 = try workbook.addFormat();
    const format04 = try workbook.addFormat();
    const format05 = try workbook.addFormat();
    const format06 = try workbook.addFormat();
    const format07 = try workbook.addFormat();
    const format08 = try workbook.addFormat();
    const format09 = try workbook.addFormat();
    const format10 = try workbook.addFormat();
    const format11 = try workbook.addFormat();

    // Set some example number formats.
    format01.setNumFormat("0.000");
    format02.setNumFormat("#,##0");
    format03.setNumFormat("#,##0.00");
    format04.setNumFormat("0.00");
    format05.setNumFormat("mm/dd/yy");
    format06.setNumFormat("mmm d yyyy");
    format07.setNumFormat("d mmmm yyyy");
    format08.setNumFormat("dd/mm/yyyy hh:mm AM/PM");
    format09.setNumFormat("0 \"dollar and\" .00 \"cents\"");

    // Write data using the formats.
    try worksheet.writeNumber(.{}, 3.1415926, .default); // 3.1415926
    try worksheet.writeNumber(.{ .row = 1 }, 3.1415926, format01); // 3.142
    try worksheet.writeNumber(.{ .row = 2 }, 1234.56, format02); // 1,235
    try worksheet.writeNumber(.{ .row = 3 }, 1234.56, format03); // 1,234.56
    try worksheet.writeNumber(.{ .row = 4 }, 49.99, format04); // 49.99
    try worksheet.writeNumber(.{ .row = 5 }, 36892.521, format05); // 01/01/01
    try worksheet.writeNumber(.{ .row = 6 }, 36892.521, format06); // Jan 1 2001
    try worksheet.writeNumber(.{ .row = 7 }, 36892.521, format07); // 1 January 2001
    try worksheet.writeNumber(.{ .row = 8 }, 36892.521, format08); // 01/01/2001 12:30 AM
    try worksheet.writeNumber(.{ .row = 9 }, 1.87, format09); // 1 dollar and .87 cents

    // Show limited conditional number formats.
    format10.setNumFormat("[Green]General;[Red]-General;General");
    try worksheet.writeNumber(.{ .row = 10 }, 123, format10); // > 0 Green
    try worksheet.writeNumber(.{ .row = 11 }, -45, format10); // < 0 Red
    try worksheet.writeNumber(.{ .row = 12 }, 0, format10); // = 0 Default color

    // Format a Zip code.
    format11.setNumFormat("00000");
    try worksheet.writeNumber(.{ .row = 13 }, 1209, format11);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

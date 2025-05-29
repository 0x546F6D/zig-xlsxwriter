pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 30, .none);

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
    try worksheet.writeNumber(0, 0, 3.1415926, .none); // 3.1415926
    try worksheet.writeNumber(1, 0, 3.1415926, format01); // 3.142
    try worksheet.writeNumber(2, 0, 1234.56, format02); // 1,235
    try worksheet.writeNumber(3, 0, 1234.56, format03); // 1,234.56
    try worksheet.writeNumber(4, 0, 49.99, format04); // 49.99
    try worksheet.writeNumber(5, 0, 36892.521, format05); // 01/01/01
    try worksheet.writeNumber(6, 0, 36892.521, format06); // Jan 1 2001
    try worksheet.writeNumber(7, 0, 36892.521, format07); // 1 January 2001
    try worksheet.writeNumber(8, 0, 36892.521, format08); // 01/01/2001 12:30 AM
    try worksheet.writeNumber(9, 0, 1.87, format09); // 1 dollar and .87 cents

    // Show limited conditional number formats.
    format10.setNumFormat("[Green]General;[Red]-General;General");
    try worksheet.writeNumber(10, 0, 123, format10); // > 0 Green
    try worksheet.writeNumber(11, 0, -45, format10); // < 0 Red
    try worksheet.writeNumber(12, 0, 0, format10); // = 0 Default color

    // Format a Zip code.
    format11.setNumFormat("00000");
    try worksheet.writeNumber(13, 0, 1209, format11);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

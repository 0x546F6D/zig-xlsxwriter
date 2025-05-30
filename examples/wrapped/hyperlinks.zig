pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Get the default url format (used in the overwriting examples below)
    const url_format = try workbook.getDefaultUrlFormat();

    // Create a user defined link format
    const red_format = try workbook.addFormat();
    red_format.setUnderline(.single);
    red_format.setFontColor(.red);
    // Or use an undefined color:
    red_format.setFontColor(@enumFromInt(0xFF0000));

    // Widen the first column to make the text clearer
    try worksheet.setColumn(0, 0, 30, .default);

    // Write a hyperlink. A default blue underline will be used if the format is NULL
    try worksheet.writeUrl(0, 0, "http://libxlsxwriter.github.io", .default);

    // Write a hyperlink but overwrite the displayed string. Note, we need to
    // specify the format for the string to match the default hyperlink
    try worksheet.writeUrl(2, 0, "http://libxlsxwriter.github.io", .default);
    try worksheet.writeString(2, 0, "Read the documentation.", url_format);

    // Write a hyperlink with a different format
    try worksheet.writeUrl(4, 0, "http://libxlsxwriter.github.io", red_format);

    // Write a mail hyperlink
    try worksheet.writeUrl(6, 0, "mailto:jmcnamara@cpan.org", .default);

    // Write a mail hyperlink and overwrite the displayed string. We again
    // specify the format for the string to match the default hyperlink
    try worksheet.writeUrl(8, 0, "mailto:jmcnamara@cpan.org", .default);
    try worksheet.writeString(8, 0, "Drop me a line.", url_format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

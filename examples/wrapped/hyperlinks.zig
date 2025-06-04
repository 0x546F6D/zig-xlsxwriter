pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
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
    try worksheet.setColumn(.{}, 30, .default, null);

    // Write a hyperlink. A default blue underline will be used if the format is NULL
    try worksheet.writeUrl(.{}, "http://libxlsxwriter.github.io", .default, null);

    // Write a hyperlink but overwrite the displayed string. Note, we need to
    // specify the format for the string to match the default hyperlink
    try worksheet.writeUrl(.{ .row = 2 }, "http://libxlsxwriter.github.io", .default, null);
    try worksheet.writeString(.{ .row = 2 }, "Read the documentation.", url_format);

    // Write a hyperlink with a different format
    try worksheet.writeUrl(.{ .row = 4 }, "http://libxlsxwriter.github.io", red_format, null);

    // Write a mail hyperlink
    try worksheet.writeUrl(.{ .row = 6 }, "mailto:jmcnamara@cpan.org", .default, null);

    // Write a mail hyperlink and overwrite the displayed string. We again
    // specify the format for the string to match the default hyperlink
    try worksheet.writeUrl(.{ .row = 8 }, "mailto:jmcnamara@cpan.org", .default, null);
    try worksheet.writeString(.{ .row = 8 }, "Drop me a line.", url_format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

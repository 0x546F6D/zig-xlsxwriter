pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Set up some formats to use.
    const bold = try workbook.addFormat();
    bold.setBold();

    const italic = try workbook.addFormat();
    italic.setItalic();

    const red = try workbook.addFormat();
    red.setFontColor(.red);

    const blue = try workbook.addFormat();
    blue.setFontColor(.blue);

    const center = try workbook.addFormat();
    center.setAlign(.center);

    const superscript = try workbook.addFormat();
    superscript.setFontScript(.superscript);

    // Make the first column wider for clarity.
    try worksheet.setColumn(0, 0, 30, .none);

    // Example 1: Bold and italic text
    // Write individual cells with appropriate formatting
    const fragment11: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = "This is ",
    };
    const fragment12: xlsxwriter.RichStringTuple = .{
        .format_c = bold.format_c,
        .string = "bold",
    };
    const fragment13: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = " and this is ",
    };
    const fragment14: xlsxwriter.RichStringTuple = .{
        .format_c = italic.format_c,
        .string = "italic",
    };
    const rich_string1: xlsxwriter.RichStringType = &.{
        &fragment11,
        &fragment12,
        &fragment13,
        &fragment14,
    };
    try worksheet.writeRichString(0, 0, rich_string1, .none);

    // Example 2: Red and blue text

    const fragment21: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = "This is ",
    };
    const fragment22: xlsxwriter.RichStringTuple = .{
        .format_c = red.format_c,
        .string = "red",
    };
    const fragment23: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = " and this is ",
    };
    const fragment24: xlsxwriter.RichStringTuple = .{
        .format_c = blue.format_c,
        .string = "blue",
    };
    const rich_string2: xlsxwriter.RichStringType = &.{
        &fragment21,
        &fragment22,
        &fragment23,
        &fragment24,
    };
    try worksheet.writeRichString(2, 0, rich_string2, .none);

    // Example 3. A rich string plus cell formatting.
    const fragment31: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = "Some ",
    };
    const fragment32: xlsxwriter.RichStringTuple = .{
        .format_c = bold.format_c,
        .string = "bold text",
    };
    const fragment33: xlsxwriter.RichStringTuple = .{
        .format_c = null,
        .string = " centered",
    };
    const rich_string3: xlsxwriter.RichStringType = &.{
        &fragment31,
        &fragment32,
        &fragment33,
    };
    try worksheet.writeRichString(4, 0, rich_string3, center);

    // Example 4: Math example with superscript
    const fragment41: xlsxwriter.RichStringTuple = .{
        .format_c = italic.format_c,
        .string = "j =k",
    };
    const fragment42: xlsxwriter.RichStringTuple = .{
        .format_c = superscript.format_c,
        .string = "(n-1)",
    };
    const rich_string4: xlsxwriter.RichStringType = &.{
        &fragment41,
        &fragment42,
    };
    try worksheet.writeRichString(6, 0, rich_string4, center);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    // pass an 'Allocator' to initWorkBook() because we use the writeRichString() function
    // Use writeRichStringNoAlloc() if you do not want to use allocation
    const workbook = try xwz.initWorkBook(alloc, xlsx_path.ptr);
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
    try worksheet.setColumn(0, 0, 30, .default);

    // Example 1: Bold and italic text
    // Write individual cells with appropriate formatting
    const fragment11: xwz.RichStringTuple = .{
        .string = "This is ",
    };
    const fragment12: xwz.RichStringTuple = .{
        .format = bold,
        .string = "bold",
    };
    const fragment13: xwz.RichStringTuple = .{
        .string = " and this is ",
    };
    const fragment14: xwz.RichStringTuple = .{
        .format = italic,
        .string = "italic",
    };
    const rich_string1 = &.{
        fragment11,
        fragment12,
        fragment13,
        fragment14,
    };
    worksheet.writeRichString(0, 0, rich_string1, .default) catch |err| {
        std.debug.print("writeSrichString error: {s}\n", .{xwz.strError(err)});
        return err;
    };

    // Example 2: Red and blue text

    const fragment21: xwz.RichStringTuple = .{
        .string = "This is ",
    };
    const fragment22: xwz.RichStringTuple = .{
        .format = red,
        .string = "red",
    };
    const fragment23: xwz.RichStringTuple = .{
        .string = " and this is ",
    };
    const fragment24: xwz.RichStringTuple = .{
        .format = blue,
        .string = "blue",
    };
    const rich_string2 = &.{
        fragment21,
        fragment22,
        fragment23,
        fragment24,
    };
    try worksheet.writeRichString(2, 0, rich_string2, .default);

    // Example 3. A rich string plus cell formatting.
    const fragment31: xwz.RichStringTuple = .{
        .string = "Some ",
    };
    const fragment32: xwz.RichStringTuple = .{
        .format = bold,
        .string = "bold text",
    };
    const fragment33: xwz.RichStringTuple = .{
        .string = " centered",
    };
    const rich_string3 = &.{
        fragment31,
        fragment32,
        fragment33,
    };
    try worksheet.writeRichString(4, 0, rich_string3, center);

    // Example 4: Math example with superscript
    const fragment41: xwz.RichStringTuple = .{
        .format = italic,
        .string = "j =k",
    };
    const fragment42: xwz.RichStringTuple = .{
        .format = superscript,
        .string = "(n-1)",
    };
    const rich_string4 = &.{
        fragment41,
        fragment42,
    };
    try worksheet.writeRichString(6, 0, rich_string4, center);
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

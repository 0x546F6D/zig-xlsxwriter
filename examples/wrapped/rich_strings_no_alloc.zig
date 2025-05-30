pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    // No need to pass an 'Allocator' to initWorkBook()
    // because we use the writeRichStringNoAlloc() function
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
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
    const fragment11: xwz.RichStringTupleNoAlloc = .{
        .string = "This is ",
    };
    const fragment12: xwz.RichStringTupleNoAlloc = .{
        .format_c = bold.format_c,
        .string = "bold",
    };
    const fragment13: xwz.RichStringTupleNoAlloc = .{
        .string = " and this is ",
    };
    const fragment14: xwz.RichStringTupleNoAlloc = .{
        .format_c = italic.format_c,
        .string = "italic",
    };
    const rich_string1: xwz.RichStringNoAllocArray = &.{
        &fragment11,
        &fragment12,
        &fragment13,
        &fragment14,
    };
    worksheet.writeRichStringNoAlloc(0, 0, rich_string1, .default) catch |err| {
        std.debug.print("writeSrichString error: {s}\n", .{xwz.strError(err)});
        return err;
    };

    // Example 2: Red and blue text

    const fragment21: xwz.RichStringTupleNoAlloc = .{
        .string = "This is ",
    };
    const fragment22: xwz.RichStringTupleNoAlloc = .{
        .format_c = red.format_c,
        .string = "red",
    };
    const fragment23: xwz.RichStringTupleNoAlloc = .{
        .string = " and this is ",
    };
    const fragment24: xwz.RichStringTupleNoAlloc = .{
        .format_c = blue.format_c,
        .string = "blue",
    };
    const rich_string2: xwz.RichStringNoAllocArray = &.{
        &fragment21,
        &fragment22,
        &fragment23,
        &fragment24,
    };
    try worksheet.writeRichStringNoAlloc(2, 0, rich_string2, .default);

    // Example 3. A rich string plus cell formatting.
    const fragment31: xwz.RichStringTupleNoAlloc = .{
        .string = "Some ",
    };
    const fragment32: xwz.RichStringTupleNoAlloc = .{
        .format_c = bold.format_c,
        .string = "bold text",
    };
    const fragment33: xwz.RichStringTupleNoAlloc = .{
        .string = " centered",
    };
    const rich_string3: xwz.RichStringNoAllocArray = &.{
        &fragment31,
        &fragment32,
        &fragment33,
    };
    try worksheet.writeRichStringNoAlloc(4, 0, rich_string3, center);

    // Example 4: Math example with superscript
    const fragment41: xwz.RichStringTupleNoAlloc = .{
        .format_c = italic.format_c,
        .string = "j =k",
    };
    const fragment42: xwz.RichStringTupleNoAlloc = .{
        .format_c = superscript.format_c,
        .string = "(n-1)",
    };
    const rich_string4: xwz.RichStringNoAllocArray = &.{
        &fragment41,
        &fragment42,
    };
    try worksheet.writeRichStringNoAlloc(6, 0, rich_string4, center);
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(.{}, 20, .default, null);

    // Add some formats.
    const format1 = try workbook.addFormat();
    const format2 = try workbook.addFormat();
    const format3 = try workbook.addFormat();

    // Set the bold property for format 1.
    format1.setBold();

    // Set the italic property for format 2.
    format2.setItalic();

    // Set the bold and italic properties for format 3.
    format3.setBold();
    format3.setItalic();

    // Write some formatted strings.
    try worksheet.writeString(.{}, "This is bold", format1);
    try worksheet.writeString(.{ .row = 1 }, "This is italic", format2);
    try worksheet.writeString(.{ .row = 2 }, "Bold and italic", format3);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Widen the first column to make the text clearer.
    try worksheet.setColumn(0, 0, 20, .none);

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
    try worksheet.writeString(0, 0, "This is bold", format1);
    try worksheet.writeString(1, 0, "This is italic", format2);
    try worksheet.writeString(2, 0, "Bold and italic", format3);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

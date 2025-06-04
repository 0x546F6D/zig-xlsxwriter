pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Add some diagonal border formats.
    const format1 = try workbook.addFormat();
    format1.setDiagType(.up);

    const format2 = try workbook.addFormat();
    format2.setDiagType(.down);

    const format3 = try workbook.addFormat();
    format3.setDiagType(.up_down);

    const format4 = try workbook.addFormat();
    format4.setDiagType(.up_down);
    format4.setDiagBorder(.hair);
    format4.setDiagColor(.red);

    try worksheet.writeString(.{ .row = 2, .col = 1 }, "Text", format1);
    try worksheet.writeString(.{ .row = 5, .col = 1 }, "Text", format2);
    try worksheet.writeString(.{ .row = 8, .col = 1 }, "Text", format3);
    try worksheet.writeString(.{ .row = 11, .col = 1 }, "Text", format4);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

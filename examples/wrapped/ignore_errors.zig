pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Write strings that looks like numbers. This will cause an Excel warning.
    try worksheet.writeString(.{ .row = 1, .col = 2 }, "123", .default);
    try worksheet.writeString(.{ .row = 2, .col = 2 }, "123", .default);

    // Write a divide by zero formula. This will also cause an Excel warning.
    try worksheet.writeFormula(.{ .row = 4, .col = 2 }, "=1/0", .default);
    try worksheet.writeFormula(.{ .row = 5, .col = 2 }, "=1/0", .default);

    // Turn off some of the warnings:
    try worksheet.ignoreErrors(.number_stored_as_text, "C3");
    try worksheet.ignoreErrors(.eval_error, "C6");

    // Write some descriptions for the cells and make the column wider for clarity.
    try worksheet.setColumn(.{ .first = 1, .last = 1 }, 16, .default, null);
    try worksheet.writeString(.{ .row = 1, .col = 1 }, "Warning:", .default);
    try worksheet.writeString(.{ .row = 2, .col = 1 }, "Warning turned off:", .default);
    try worksheet.writeString(.{ .row = 4, .col = 1 }, "Warning:", .default);
    try worksheet.writeString(.{ .row = 5, .col = 1 }, "Warning turned off:", .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

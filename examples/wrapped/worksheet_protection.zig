pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Turn worksheet protection on without a password
    worksheet.protect(null, .default);

    const unlocked = try workbook.addFormat();
    unlocked.setUnlocked();

    const hidden = try workbook.addFormat();
    hidden.setHidden();

    // Widen the first column to make the text clearer
    try worksheet.setColumn(0, 0, 40, .default);

    // Write a locked, unlocked and hidden cell
    try worksheet.writeString(0, 0, "B1 is locked. It cannot be edited.", .default);
    try worksheet.writeString(1, 0, "B2 is unlocked. It can be edited.", .default);
    try worksheet.writeString(2, 0, "B3 is hidden. The formula isn't visible.", .default);

    try worksheet.writeFormula(0, 1, "=1+2", .default); // Locked by default
    try worksheet.writeFormula(1, 1, "=1+2", unlocked);
    try worksheet.writeFormula(2, 1, "=1+2", hidden);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

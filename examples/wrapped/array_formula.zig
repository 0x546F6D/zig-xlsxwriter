pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the formulas.
    try worksheet.writeNumber(0, 1, 500, .none);
    try worksheet.writeNumber(1, 1, 10, .none);
    try worksheet.writeNumber(4, 1, 1, .none);
    try worksheet.writeNumber(5, 1, 2, .none);
    try worksheet.writeNumber(6, 1, 3, .none);

    try worksheet.writeNumber(0, 2, 300, .none);
    try worksheet.writeNumber(1, 2, 15, .none);
    try worksheet.writeNumber(4, 2, 20234, .none);
    try worksheet.writeNumber(5, 2, 21003, .none);
    try worksheet.writeNumber(6, 2, 10000, .none);

    // Write an array formula that returns a single value.
    try worksheet.writeArrayFormula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", .none);

    // Similar to above but using the RANGE macro.
    try worksheet.writeArrayFormula(1, xlsxwriter.range("A2:A2"), 0, 0, "{=SUM(B1:C1*B2:C2)}", .none);

    // Write an array formula that returns a range of values.
    try worksheet.writeArrayFormula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", .none);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the formulas.
    try worksheet.writeNumber(0, 1, 500, .default);
    try worksheet.writeNumber(1, 1, 10, .default);
    try worksheet.writeNumber(4, 1, 1, .default);
    try worksheet.writeNumber(5, 1, 2, .default);
    try worksheet.writeNumber(6, 1, 3, .default);

    try worksheet.writeNumber(0, 2, 300, .default);
    try worksheet.writeNumber(1, 2, 15, .default);
    try worksheet.writeNumber(4, 2, 20234, .default);
    try worksheet.writeNumber(5, 2, 21003, .default);
    try worksheet.writeNumber(6, 2, 10000, .default);

    // Write an array formula that returns a single value.
    try worksheet.writeArrayFormula(0, 0, 0, 0, "{=SUM(B1:C1*B2:C2)}", .default);

    // Similar to above but using the RANGE macro.
    const range = xwz.range("A2:A2");
    try worksheet.writeArrayFormula(
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        "{=SUM(B1:C1*B2:C2)}",
        .default,
    );

    // Write an array formula that returns a range of values.
    try worksheet.writeArrayFormula(4, 0, 6, 0, "{=TREND(C5:C7,B5:B7)}", .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Write some data for the formulas.
    try worksheet.writeNumber(.{ .row = 0, .col = 1 }, 500, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 10, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 1 }, 1, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 1 }, 2, .default);
    try worksheet.writeNumber(.{ .row = 6, .col = 1 }, 3, .default);

    try worksheet.writeNumber(.{ .row = 0, .col = 2 }, 300, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 2 }, 15, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 2 }, 20234, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 2 }, 21003, .default);
    try worksheet.writeNumber(.{ .row = 6, .col = 2 }, 10000, .default);

    // Write an array formula that returns a single value.
    try worksheet.writeArrayFormula(.{}, "{=SUM(B1:C1*B2:C2)}", .default);

    // Similar to above but using the RANGE macro.
    try worksheet.writeArrayFormula(
        xwz.range("A2:A2"),
        "{=SUM(B1:C1*B2:C2)}",
        .default,
    );

    // Write an array formula that returns a range of values.
    try worksheet.writeArrayFormula(
        .{
            .first_row = 4,
            .last_row = 6,
        },
        "{=TREND(C5:C7,B5:B7)}",
        .default,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

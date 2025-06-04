pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Write some sample data.
    try worksheet.writeNumber(xwz.cell("B1"), 34, .default);
    try worksheet.writeNumber(xwz.cell("B2"), 32, .default);
    try worksheet.writeNumber(xwz.cell("B3"), 31, .default);
    try worksheet.writeNumber(xwz.cell("B4"), 35, .default);
    try worksheet.writeNumber(xwz.cell("B5"), 36, .default);
    try worksheet.writeNumber(xwz.cell("B6"), 30, .default);
    try worksheet.writeNumber(xwz.cell("B7"), 38, .default);
    try worksheet.writeNumber(xwz.cell("B8"), 38, .default);
    try worksheet.writeNumber(xwz.cell("B9"), 32, .default);

    // Add a format with red text.
    const custom_format = try workbook.addFormat();
    custom_format.setFontColor(.red);

    // Create a conditional format object.
    const conditional_format: xwz.ConditionalFormat = .{
        // Set the format type: a cell conditional:
        .type = .cell,
        // Set the criteria to use:
        .criteria = .less_than,
        // Set the value to which the criteria will be applied:
        .value = 33,
        // Set the format to use if the criteria/value applies:
        .format = custom_format,
    };

    // Now apply the format to data range.
    // RANGE("B1:B9") expands to first_row, first_col, last_row, last_col
    try worksheet.conditionalFormatRange(
        xwz.range("B1:B9"),
        conditional_format,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

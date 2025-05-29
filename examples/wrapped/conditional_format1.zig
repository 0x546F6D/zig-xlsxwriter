pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    var workbook = try xwz.initWorkBook(alloc, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Write some sample data.
    const row_b1 = xwz.nameToRow("B1");
    const col_b1 = xwz.nameToCol("B1");
    try worksheet.writeNumber(row_b1, col_b1, 34, .none);

    const row_b2 = xwz.nameToRow("B2");
    const col_b2 = xwz.nameToCol("B2");
    try worksheet.writeNumber(row_b2, col_b2, 32, .none);

    const row_b3 = xwz.nameToRow("B3");
    const col_b3 = xwz.nameToCol("B3");
    try worksheet.writeNumber(row_b3, col_b3, 31, .none);

    const row_b4 = xwz.nameToRow("B4");
    const col_b4 = xwz.nameToCol("B4");
    try worksheet.writeNumber(row_b4, col_b4, 35, .none);

    const row_b5 = xwz.nameToRow("B5");
    const col_b5 = xwz.nameToCol("B5");
    try worksheet.writeNumber(row_b5, col_b5, 36, .none);

    const row_b6 = xwz.nameToRow("B6");
    const col_b6 = xwz.nameToCol("B6");
    try worksheet.writeNumber(row_b6, col_b6, 30, .none);

    const row_b7 = xwz.nameToRow("B7");
    const col_b7 = xwz.nameToCol("B7");
    try worksheet.writeNumber(row_b7, col_b7, 38, .none);

    const row_b8 = xwz.nameToRow("B8");
    const col_b8 = xwz.nameToCol("B8");
    try worksheet.writeNumber(row_b8, col_b8, 38, .none);

    const row_b9 = xwz.nameToRow("B9");
    const col_b9 = xwz.nameToCol("B9");
    try worksheet.writeNumber(row_b9, col_b9, 32, .none);

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
    const range = xwz.range("B1:B9");
    try worksheet.conditionalFormatRange(
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        conditional_format,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

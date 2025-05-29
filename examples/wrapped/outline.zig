pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    var workbook = try xwz.initWorkBook(alloc, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet("Outlined Rows");
    const worksheet2 = try workbook.addWorkSheet("Collapsed Rows");
    const worksheet3 = try workbook.addWorkSheet("Outline Columns");
    const worksheet4 = try workbook.addWorkSheet("Outline levels");
    var worksheet: xwz.WorkSheet = undefined;

    const bold = try workbook.addFormat();
    bold.setBold();

    // Example 1: Create a worksheet with outlined rows. It also includes
    // SUBTOTAL() functions so that it looks like the type of automatic
    // outlines that are generated when you use the 'Sub Totals' option.
    //
    // For outlines the important parameters are 'hidden' and 'level'. Rows
    // with the same 'level' are grouped together. The group will be collapsed
    // if 'hidden' is non-zero.
    worksheet = worksheet1;

    // The option structs with the outline level set.
    const options1 = xwz.RowColOptions{ .level = 2 };
    const options2 = xwz.RowColOptions{ .level = 1 };

    // Set the column width for clarity.
    try worksheet.setColumn(0, 0, 20, .none);

    // Set the row options with the outline level.
    try worksheet.setRowOpt(1, def_row_height, .none, options1);
    try worksheet.setRowOpt(2, def_row_height, .none, options1);
    try worksheet.setRowOpt(3, def_row_height, .none, options1);
    try worksheet.setRowOpt(4, def_row_height, .none, options1);
    try worksheet.setRowOpt(5, def_row_height, .none, options2);
    try worksheet.setRowOpt(6, def_row_height, .none, options1);
    try worksheet.setRowOpt(7, def_row_height, .none, options1);
    try worksheet.setRowOpt(8, def_row_height, .none, options1);
    try worksheet.setRowOpt(9, def_row_height, .none, options1);
    try worksheet.setRowOpt(10, def_row_height, .none, options2);

    // Add data and formulas to the worksheet.
    try worksheet.writeString(0, 0, "Region", bold);
    try worksheet.writeString(1, 0, "North", .none);
    try worksheet.writeString(2, 0, "North", .none);
    try worksheet.writeString(3, 0, "North", .none);
    try worksheet.writeString(4, 0, "North", .none);
    try worksheet.writeString(5, 0, "North Total", bold);

    try worksheet.writeString(0, 1, "Sales", bold);
    try worksheet.writeNumber(1, 1, 1000, .none);
    try worksheet.writeNumber(2, 1, 1200, .none);
    try worksheet.writeNumber(3, 1, 900, .none);
    try worksheet.writeNumber(4, 1, 1200, .none);
    try worksheet.writeFormula(5, 1, "=SUBTOTAL(9,B2:B5)", bold);

    try worksheet.writeString(6, 0, "South", .none);
    try worksheet.writeString(7, 0, "South", .none);
    try worksheet.writeString(8, 0, "South", .none);
    try worksheet.writeString(9, 0, "South", .none);
    try worksheet.writeString(10, 0, "South Total", bold);

    try worksheet.writeNumber(6, 1, 400, .none);
    try worksheet.writeNumber(7, 1, 600, .none);
    try worksheet.writeNumber(8, 1, 500, .none);
    try worksheet.writeNumber(9, 1, 600, .none);
    try worksheet.writeFormula(10, 1, "=SUBTOTAL(9,B7:B10)", bold);

    try worksheet.writeString(11, 0, "Grand Total", bold);
    try worksheet.writeFormula(11, 1, "=SUBTOTAL(9,B2:B10)", bold);

    // Example 2: Create a worksheet with outlined rows. This is the same as
    // the previous example except that the rows are collapsed. Note: We need
    // to indicate the row that contains the collapsed symbol '+' with the
    // optional parameter, 'collapsed'.
    worksheet = worksheet2;

    // The option structs with the outline level and collapsed property set.
    const options3 = xwz.RowColOptions{ .hidden = true, .level = 2 };
    const options4 = xwz.RowColOptions{ .hidden = true, .level = 1 };
    const options5 = xwz.RowColOptions{ .collapsed = true };

    // Set the column width for clarity.
    try worksheet.setColumn(0, 0, 20, .none);

    // Set the row options with the outline level.
    try worksheet.setRowOpt(1, def_row_height, .none, options3);
    try worksheet.setRowOpt(2, def_row_height, .none, options3);
    try worksheet.setRowOpt(3, def_row_height, .none, options3);
    try worksheet.setRowOpt(4, def_row_height, .none, options3);
    try worksheet.setRowOpt(5, def_row_height, .none, options4);

    try worksheet.setRowOpt(6, def_row_height, .none, options3);
    try worksheet.setRowOpt(7, def_row_height, .none, options3);
    try worksheet.setRowOpt(8, def_row_height, .none, options3);
    try worksheet.setRowOpt(9, def_row_height, .none, options3);
    try worksheet.setRowOpt(10, def_row_height, .none, options4);
    try worksheet.setRowOpt(11, def_row_height, .none, options5);

    // Add data and formulas to the worksheet.
    try worksheet.writeString(0, 0, "Region", bold);
    try worksheet.writeString(1, 0, "North", .none);
    try worksheet.writeString(2, 0, "North", .none);
    try worksheet.writeString(3, 0, "North", .none);
    try worksheet.writeString(4, 0, "North", .none);
    try worksheet.writeString(5, 0, "North Total", bold);

    try worksheet.writeString(0, 1, "Sales", bold);
    try worksheet.writeNumber(1, 1, 1000, .none);
    try worksheet.writeNumber(2, 1, 1200, .none);
    try worksheet.writeNumber(3, 1, 900, .none);
    try worksheet.writeNumber(4, 1, 1200, .none);
    try worksheet.writeFormula(5, 1, "=SUBTOTAL(9,B2:B5)", bold);

    try worksheet.writeString(6, 0, "South", .none);
    try worksheet.writeString(7, 0, "South", .none);
    try worksheet.writeString(8, 0, "South", .none);
    try worksheet.writeString(9, 0, "South", .none);
    try worksheet.writeString(10, 0, "South Total", bold);

    try worksheet.writeNumber(6, 1, 400, .none);
    try worksheet.writeNumber(7, 1, 600, .none);
    try worksheet.writeNumber(8, 1, 500, .none);
    try worksheet.writeNumber(9, 1, 600, .none);
    try worksheet.writeFormula(10, 1, "=SUBTOTAL(9,B7:B10)", bold);

    try worksheet.writeString(11, 0, "Grand Total", bold);
    try worksheet.writeFormula(11, 1, "=SUBTOTAL(9,B2:B10)", bold);

    // Example 3: Create a worksheet with outlined columns.
    worksheet = worksheet3;
    const options6 = xwz.RowColOptions{ .level = 1 };

    // Add data and formulas to the worksheet.
    try worksheet.writeString(0, 0, "Month", .none);
    try worksheet.writeString(0, 1, "Jan", .none);
    try worksheet.writeString(0, 2, "Feb", .none);
    try worksheet.writeString(0, 3, "Mar", .none);
    try worksheet.writeString(0, 4, "Apr", .none);
    try worksheet.writeString(0, 5, "May", .none);
    try worksheet.writeString(0, 6, "Jun", .none);
    try worksheet.writeString(0, 7, "Total", .none);

    try worksheet.writeString(1, 0, "North", .none);
    try worksheet.writeNumber(1, 1, 50, .none);
    try worksheet.writeNumber(1, 2, 20, .none);
    try worksheet.writeNumber(1, 3, 15, .none);
    try worksheet.writeNumber(1, 4, 25, .none);
    try worksheet.writeNumber(1, 5, 65, .none);
    try worksheet.writeNumber(1, 6, 80, .none);
    try worksheet.writeFormula(1, 7, "=SUM(B2:G2)", .none);

    try worksheet.writeString(2, 0, "South", .none);
    try worksheet.writeNumber(2, 1, 10, .none);
    try worksheet.writeNumber(2, 2, 20, .none);
    try worksheet.writeNumber(2, 3, 30, .none);
    try worksheet.writeNumber(2, 4, 50, .none);
    try worksheet.writeNumber(2, 5, 50, .none);
    try worksheet.writeNumber(2, 6, 50, .none);
    try worksheet.writeFormula(2, 7, "=SUM(B3:G3)", .none);

    try worksheet.writeString(3, 0, "East", .none);
    try worksheet.writeNumber(3, 1, 45, .none);
    try worksheet.writeNumber(3, 2, 75, .none);
    try worksheet.writeNumber(3, 3, 50, .none);
    try worksheet.writeNumber(3, 4, 15, .none);
    try worksheet.writeNumber(3, 5, 75, .none);
    try worksheet.writeNumber(3, 6, 100, .none);
    try worksheet.writeFormula(3, 7, "=SUM(B4:G4)", .none);

    try worksheet.writeString(4, 0, "West", .none);
    try worksheet.writeNumber(4, 1, 15, .none);
    try worksheet.writeNumber(4, 2, 15, .none);
    try worksheet.writeNumber(4, 3, 55, .none);
    try worksheet.writeNumber(4, 4, 35, .none);
    try worksheet.writeNumber(4, 5, 20, .none);
    try worksheet.writeNumber(4, 6, 50, .none);
    try worksheet.writeFormula(4, 7, "=SUM(B5:G5)", .none);

    try worksheet.writeFormula(5, 7, "=SUM(H2:H5)", bold);

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold);

    // Set column formatting and the outline level.
    try worksheet.setColumn(0, 0, 10, bold);
    try worksheet.setColumnOpt(1, 6, 5, .none, options6);
    try worksheet.setColumn(7, 7, 10, .none);

    // Example 4: Show all possible outline levels.
    worksheet = worksheet4;
    const level1 = xwz.RowColOptions{ .level = 1 };
    const level2 = xwz.RowColOptions{ .level = 2 };
    const level3 = xwz.RowColOptions{ .level = 3 };
    const level4 = xwz.RowColOptions{ .level = 4 };
    const level5 = xwz.RowColOptions{ .level = 5 };
    const level6 = xwz.RowColOptions{ .level = 6 };
    const level7 = xwz.RowColOptions{ .level = 7 };

    try worksheet.writeString(0, 0, "Level 1", .none);
    try worksheet.writeString(1, 0, "Level 2", .none);
    try worksheet.writeString(2, 0, "Level 3", .none);
    try worksheet.writeString(3, 0, "Level 4", .none);
    try worksheet.writeString(4, 0, "Level 5", .none);
    try worksheet.writeString(5, 0, "Level 6", .none);
    try worksheet.writeString(6, 0, "Level 7", .none);
    try worksheet.writeString(7, 0, "Level 6", .none);
    try worksheet.writeString(8, 0, "Level 5", .none);
    try worksheet.writeString(9, 0, "Level 4", .none);
    try worksheet.writeString(10, 0, "Level 3", .none);
    try worksheet.writeString(11, 0, "Level 2", .none);
    try worksheet.writeString(12, 0, "Level 1", .none);

    try worksheet.setRowOpt(0, def_row_height, .none, level1);
    try worksheet.setRowOpt(1, def_row_height, .none, level2);
    try worksheet.setRowOpt(2, def_row_height, .none, level3);
    try worksheet.setRowOpt(3, def_row_height, .none, level4);
    try worksheet.setRowOpt(4, def_row_height, .none, level5);
    try worksheet.setRowOpt(5, def_row_height, .none, level6);
    try worksheet.setRowOpt(6, def_row_height, .none, level7);
    try worksheet.setRowOpt(7, def_row_height, .none, level6);
    try worksheet.setRowOpt(8, def_row_height, .none, level5);
    try worksheet.setRowOpt(9, def_row_height, .none, level4);
    try worksheet.setRowOpt(10, def_row_height, .none, level3);
    try worksheet.setRowOpt(11, def_row_height, .none, level2);
    try worksheet.setRowOpt(12, def_row_height, .none, level1);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const def_row_height = xwz.def_row_height;

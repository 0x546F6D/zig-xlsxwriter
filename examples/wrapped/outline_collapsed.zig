pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet("Outlined Rows");
    const worksheet2 = try workbook.addWorkSheet("Collapsed Rows 1");
    const worksheet3 = try workbook.addWorkSheet("Collapsed Rows 2");
    const worksheet4 = try workbook.addWorkSheet("Collapsed Rows 3");
    const worksheet5 = try workbook.addWorkSheet("Outline Columns");
    const worksheet6 = try workbook.addWorkSheet("Collapsed Columns");
    var worksheet: xlsxwriter.WorkSheet = undefined;

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
    const options1 = xlsxwriter.RowColOptions{ .level = 2 };
    const options2 = xlsxwriter.RowColOptions{ .level = 1 };

    // Set the row outline properties.
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

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 2: Create a worksheet with collapsed outlined rows.
    worksheet = worksheet2;
    // This is the same as the example 1 except that all rows are collapsed.
    const options3 = xlsxwriter.RowColOptions{ .hidden = xw_true, .level = 2 };
    const options4 = xlsxwriter.RowColOptions{ .hidden = xw_true, .level = 1 };
    const options5 = xlsxwriter.RowColOptions{ .collapsed = xw_true };

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

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 3: Create a worksheet with collapsed outlined rows. Same as the
    // example 1 except that the two sub-totals are collapsed.
    worksheet = worksheet3;
    const options6 = xlsxwriter.RowColOptions{ .hidden = xw_true, .level = 2 };
    const options7 = xlsxwriter.RowColOptions{ .level = 1, .collapsed = xw_true };

    // Set the row options with the outline level.
    try worksheet.setRowOpt(1, def_row_height, .none, options6);
    try worksheet.setRowOpt(2, def_row_height, .none, options6);
    try worksheet.setRowOpt(3, def_row_height, .none, options6);
    try worksheet.setRowOpt(4, def_row_height, .none, options6);
    try worksheet.setRowOpt(5, def_row_height, .none, options7);

    try worksheet.setRowOpt(6, def_row_height, .none, options6);
    try worksheet.setRowOpt(7, def_row_height, .none, options6);
    try worksheet.setRowOpt(8, def_row_height, .none, options6);
    try worksheet.setRowOpt(9, def_row_height, .none, options6);
    try worksheet.setRowOpt(10, def_row_height, .none, options7);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 4: Create a worksheet with outlined rows. Same as the example 1
    // except that the two sub-totals are collapsed.
    worksheet = worksheet4;
    const options8 = xlsxwriter.RowColOptions{ .hidden = xw_true, .level = 2 };
    const options9 = xlsxwriter.RowColOptions{ .hidden = xw_true, .level = 1, .collapsed = xw_true };
    const options10 = xlsxwriter.RowColOptions{ .collapsed = xw_true };

    // Set the row options with the outline level.
    try worksheet.setRowOpt(1, def_row_height, .none, options8);
    try worksheet.setRowOpt(2, def_row_height, .none, options8);
    try worksheet.setRowOpt(3, def_row_height, .none, options8);
    try worksheet.setRowOpt(4, def_row_height, .none, options8);
    try worksheet.setRowOpt(5, def_row_height, .none, options9);

    try worksheet.setRowOpt(6, def_row_height, .none, options8);
    try worksheet.setRowOpt(7, def_row_height, .none, options8);
    try worksheet.setRowOpt(8, def_row_height, .none, options8);
    try worksheet.setRowOpt(9, def_row_height, .none, options8);
    try worksheet.setRowOpt(10, def_row_height, .none, options9);
    try worksheet.setRowOpt(11, def_row_height, .none, options10);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 5: Create a worksheet with outlined columns.
    worksheet = worksheet5;
    const options11 = xlsxwriter.RowColOptions{ .level = 1 };

    // Write the sub-total data that is common to the column examples.
    try createColExampleData(worksheet5, bold);

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold);

    // Set column formatting and the outline level.
    try worksheet.setColumn(0, 0, 10, bold);
    try worksheet.setColumnOpt(1, 6, 5, .none, options11);
    try worksheet.setColumn(7, 7, 10, .none);

    // Example 6: Create a worksheet with outlined columns.
    worksheet = worksheet6;
    const options12 = xlsxwriter.RowColOptions{
        .hidden = 1,
        .level = 1,
        .collapsed = 0,
    };
    const options13 = xlsxwriter.RowColOptions{
        .hidden = 0,
        .level = 0,
        .collapsed = 1,
    };

    // Write the sub-total data that is common to the column examples.
    try createColExampleData(worksheet6, bold);

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold);

    // Set column formatting and the outline level.
    try worksheet.setColumn(0, 0, 10, bold);
    try worksheet.setColumnOpt(1, 6, 5, .none, options12);
    try worksheet.setColumnOpt(7, 7, 10, .none, options13);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 1-4.
fn createRowExampleData(worksheet: WorkSheet, bold: Format) !void {
    // Set the column width for clarity.
    try worksheet.setColumn(0, 0, 20, .none);

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
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 5-6.
fn createColExampleData(worksheet: WorkSheet, bold: Format) !void {
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
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = xlsxwriter.WorkSheet;
const Format = xlsxwriter.Format;
const xw_true = xlsxwriter.xw_true;
const xw_false = xlsxwriter.xw_false;
const def_row_height = xlsxwriter.def_row_height;

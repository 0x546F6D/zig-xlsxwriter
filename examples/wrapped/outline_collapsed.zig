pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet("Outlined Rows");
    const worksheet2 = try workbook.addWorkSheet("Collapsed Rows 1");
    const worksheet3 = try workbook.addWorkSheet("Collapsed Rows 2");
    const worksheet4 = try workbook.addWorkSheet("Collapsed Rows 3");
    const worksheet5 = try workbook.addWorkSheet("Outline Columns");
    const worksheet6 = try workbook.addWorkSheet("Collapsed Columns");
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

    // Set the row outline properties.
    try worksheet.setRow(1, def_row_height, .default, options1);
    try worksheet.setRow(2, def_row_height, .default, options1);
    try worksheet.setRow(3, def_row_height, .default, options1);
    try worksheet.setRow(4, def_row_height, .default, options1);
    try worksheet.setRow(5, def_row_height, .default, options2);

    try worksheet.setRow(6, def_row_height, .default, options1);
    try worksheet.setRow(7, def_row_height, .default, options1);
    try worksheet.setRow(8, def_row_height, .default, options1);
    try worksheet.setRow(9, def_row_height, .default, options1);
    try worksheet.setRow(10, def_row_height, .default, options2);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 2: Create a worksheet with collapsed outlined rows.
    worksheet = worksheet2;
    // This is the same as the example 1 except that all rows are collapsed.
    const options3 = xwz.RowColOptions{ .hidden = true, .level = 2 };
    const options4 = xwz.RowColOptions{ .hidden = true, .level = 1 };
    const options5 = xwz.RowColOptions{ .collapsed = true };

    // Set the row options with the outline level.
    try worksheet.setRow(1, def_row_height, .default, options3);
    try worksheet.setRow(2, def_row_height, .default, options3);
    try worksheet.setRow(3, def_row_height, .default, options3);
    try worksheet.setRow(4, def_row_height, .default, options3);
    try worksheet.setRow(5, def_row_height, .default, options4);

    try worksheet.setRow(6, def_row_height, .default, options3);
    try worksheet.setRow(7, def_row_height, .default, options3);
    try worksheet.setRow(8, def_row_height, .default, options3);
    try worksheet.setRow(9, def_row_height, .default, options3);
    try worksheet.setRow(10, def_row_height, .default, options4);
    try worksheet.setRow(11, def_row_height, .default, options5);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 3: Create a worksheet with collapsed outlined rows. Same as the
    // example 1 except that the two sub-totals are collapsed.
    worksheet = worksheet3;
    const options6 = xwz.RowColOptions{ .hidden = true, .level = 2 };
    const options7 = xwz.RowColOptions{ .level = 1, .collapsed = true };

    // Set the row options with the outline level.
    try worksheet.setRow(1, def_row_height, .default, options6);
    try worksheet.setRow(2, def_row_height, .default, options6);
    try worksheet.setRow(3, def_row_height, .default, options6);
    try worksheet.setRow(4, def_row_height, .default, options6);
    try worksheet.setRow(5, def_row_height, .default, options7);

    try worksheet.setRow(6, def_row_height, .default, options6);
    try worksheet.setRow(7, def_row_height, .default, options6);
    try worksheet.setRow(8, def_row_height, .default, options6);
    try worksheet.setRow(9, def_row_height, .default, options6);
    try worksheet.setRow(10, def_row_height, .default, options7);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 4: Create a worksheet with outlined rows. Same as the example 1
    // except that the two sub-totals are collapsed.
    worksheet = worksheet4;
    const options8 = xwz.RowColOptions{ .hidden = true, .level = 2 };
    const options9 = xwz.RowColOptions{ .hidden = true, .level = 1, .collapsed = true };
    const options10 = xwz.RowColOptions{ .collapsed = true };

    // Set the row options with the outline level.
    try worksheet.setRow(1, def_row_height, .default, options8);
    try worksheet.setRow(2, def_row_height, .default, options8);
    try worksheet.setRow(3, def_row_height, .default, options8);
    try worksheet.setRow(4, def_row_height, .default, options8);
    try worksheet.setRow(5, def_row_height, .default, options9);

    try worksheet.setRow(6, def_row_height, .default, options8);
    try worksheet.setRow(7, def_row_height, .default, options8);
    try worksheet.setRow(8, def_row_height, .default, options8);
    try worksheet.setRow(9, def_row_height, .default, options8);
    try worksheet.setRow(10, def_row_height, .default, options9);
    try worksheet.setRow(11, def_row_height, .default, options10);

    // Write the sub-total data that is common to the row examples.
    try createRowExampleData(worksheet, bold);

    // Example 5: Create a worksheet with outlined columns.
    worksheet = worksheet5;
    const options11 = xwz.RowColOptions{ .level = 1 };

    // Write the sub-total data that is common to the column examples.
    try createColExampleData(worksheet5, bold);

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold, null);

    // Set column formatting and the outline level.
    try worksheet.setColumn(.{}, 10, bold, null);
    try worksheet.setColumn(.{ .first = 1, .last = 6 }, 5, .default, options11);
    try worksheet.setColumn(.{ .first = 7, .last = 7 }, 10, .default, null);

    // Example 6: Create a worksheet with outlined columns.
    worksheet = worksheet6;
    const options12 = xwz.RowColOptions{
        .hidden = true,
        .level = 1,
    };
    const options13 = xwz.RowColOptions{
        .level = 0,
        .collapsed = true,
    };

    // Write the sub-total data that is common to the column examples.
    try createColExampleData(worksheet6, bold);

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold, null);

    // Set column formatting and the outline level.
    try worksheet.setColumn(.{}, 10, bold, null);
    try worksheet.setColumn(.{ .first = 1, .last = 6 }, 5, .default, options12);
    try worksheet.setColumn(.{ .first = 7, .last = 7 }, 10, .default, options13);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 1-4.
fn createRowExampleData(worksheet: WorkSheet, bold: Format) !void {
    // Set the column width for clarity.
    try worksheet.setColumn(.{}, 20, .default, null);

    // Add data and formulas to the worksheet.
    try worksheet.writeString(.{}, "Region", bold);
    try worksheet.writeString(.{ .row = 1 }, "North", .default);
    try worksheet.writeString(.{ .row = 2 }, "North", .default);
    try worksheet.writeString(.{ .row = 3 }, "North", .default);
    try worksheet.writeString(.{ .row = 4 }, "North", .default);
    try worksheet.writeString(.{ .row = 5 }, "North Total", bold);

    try worksheet.writeString(.{ .col = 1 }, "Sales", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 1000, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 1200, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 900, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 1 }, 1200, .default);
    try worksheet.writeFormula(.{ .row = 5, .col = 1 }, "=SUBTOTAL(9,B2:B5)", bold);

    try worksheet.writeString(.{ .row = 6 }, "South", .default);
    try worksheet.writeString(.{ .row = 7 }, "South", .default);
    try worksheet.writeString(.{ .row = 8 }, "South", .default);
    try worksheet.writeString(.{ .row = 9 }, "South", .default);
    try worksheet.writeString(.{ .row = 10 }, "South Total", bold);

    try worksheet.writeNumber(.{ .row = 6, .col = 1 }, 400, .default);
    try worksheet.writeNumber(.{ .row = 7, .col = 1 }, 600, .default);
    try worksheet.writeNumber(.{ .row = 8, .col = 1 }, 500, .default);
    try worksheet.writeNumber(.{ .row = 9, .col = 1 }, 600, .default);
    try worksheet.writeFormula(.{ .row = 10, .col = 1 }, "=SUBTOTAL(9,B7:B10)", bold);

    try worksheet.writeString(.{ .row = 11 }, "Grand Total", bold);
    try worksheet.writeFormula(.{ .row = 11, .col = 1 }, "=SUBTOTAL(9,B2:B10)", bold);
}

// This function will generate the same data and sub-totals on each worksheet.
// Used in the examples 5-6.
fn createColExampleData(worksheet: WorkSheet, bold: Format) !void {
    // Add data and formulas to the worksheet.
    try worksheet.writeString(.{}, "Month", .default);
    try worksheet.writeString(.{ .col = 1 }, "Jan", .default);
    try worksheet.writeString(.{ .col = 2 }, "Feb", .default);
    try worksheet.writeString(.{ .col = 3 }, "Mar", .default);
    try worksheet.writeString(.{ .col = 4 }, "Apr", .default);
    try worksheet.writeString(.{ .col = 5 }, "May", .default);
    try worksheet.writeString(.{ .col = 6 }, "Jun", .default);
    try worksheet.writeString(.{ .col = 7 }, "Total", .default);

    try worksheet.writeString(.{ .row = 1 }, "North", .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 2 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 3 }, 15, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 4 }, 25, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 5 }, 65, .default);
    try worksheet.writeNumber(.{ .row = 1, .col = 6 }, 80, .default);
    try worksheet.writeFormula(.{ .row = 1, .col = 7 }, "=SUM(B2:G2)", .default);

    try worksheet.writeString(.{ .row = 2 }, "South", .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 10, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 2 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 3 }, 30, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 4 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 5 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 6 }, 50, .default);
    try worksheet.writeFormula(.{ .row = 2, .col = 7 }, "=SUM(B3:G3)", .default);

    try worksheet.writeString(.{ .row = 3 }, "East", .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 45, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 2 }, 75, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 3 }, 50, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 4 }, 15, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 5 }, 75, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 6 }, 100, .default);
    try worksheet.writeFormula(.{ .row = 3, .col = 7 }, "=SUM(B4:G4)", .default);

    try worksheet.writeString(.{ .row = 4 }, "West", .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 1 }, 15, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 2 }, 15, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 3 }, 55, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 4 }, 35, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 5 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 6 }, 50, .default);
    try worksheet.writeFormula(.{ .row = 4, .col = 7 }, "=SUM(B5:G5)", .default);

    try worksheet.writeFormula(.{ .row = 5, .col = 7 }, "=SUM(H2:H5)", bold);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;
const def_row_height = xwz.def_row_height;

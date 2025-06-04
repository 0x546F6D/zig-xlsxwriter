pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
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
    try worksheet.setColumn(.{}, 20, .default, null);

    // Set the row options with the outline level.
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
    try worksheet.setColumn(.{}, 20, .default, null);

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

    // Example 3: Create a worksheet with outlined columns.
    worksheet = worksheet3;
    const options6 = xwz.RowColOptions{ .level = 1 };

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

    // Add bold format to the first row.
    try worksheet.setRow(0, def_row_height, bold, null);

    // Set column formatting and the outline level.
    try worksheet.setColumn(.{}, 10, bold, null);
    try worksheet.setColumn(.{ .first = 1, .last = 6 }, 5, .default, options6);
    try worksheet.setColumn(.{ .first = 7, .last = 7 }, 10, .default, null);

    // Example 4: Show all possible outline levels.
    worksheet = worksheet4;
    const level1 = xwz.RowColOptions{ .level = 1 };
    const level2 = xwz.RowColOptions{ .level = 2 };
    const level3 = xwz.RowColOptions{ .level = 3 };
    const level4 = xwz.RowColOptions{ .level = 4 };
    const level5 = xwz.RowColOptions{ .level = 5 };
    const level6 = xwz.RowColOptions{ .level = 6 };
    const level7 = xwz.RowColOptions{ .level = 7 };

    try worksheet.writeString(.{}, "Level 1", .default);
    try worksheet.writeString(.{ .row = 1 }, "Level 2", .default);
    try worksheet.writeString(.{ .row = 2 }, "Level 3", .default);
    try worksheet.writeString(.{ .row = 3 }, "Level 4", .default);
    try worksheet.writeString(.{ .row = 4 }, "Level 5", .default);
    try worksheet.writeString(.{ .row = 5 }, "Level 6", .default);
    try worksheet.writeString(.{ .row = 6 }, "Level 7", .default);
    try worksheet.writeString(.{ .row = 7 }, "Level 6", .default);
    try worksheet.writeString(.{ .row = 8 }, "Level 5", .default);
    try worksheet.writeString(.{ .row = 9 }, "Level 4", .default);
    try worksheet.writeString(.{ .row = 10 }, "Level 3", .default);
    try worksheet.writeString(.{ .row = 11 }, "Level 2", .default);
    try worksheet.writeString(.{ .row = 12 }, "Level 1", .default);

    try worksheet.setRow(0, def_row_height, .default, level1);
    try worksheet.setRow(1, def_row_height, .default, level2);
    try worksheet.setRow(2, def_row_height, .default, level3);
    try worksheet.setRow(3, def_row_height, .default, level4);
    try worksheet.setRow(4, def_row_height, .default, level5);
    try worksheet.setRow(5, def_row_height, .default, level6);
    try worksheet.setRow(6, def_row_height, .default, level7);
    try worksheet.setRow(7, def_row_height, .default, level6);
    try worksheet.setRow(8, def_row_height, .default, level5);
    try worksheet.setRow(9, def_row_height, .default, level4);
    try worksheet.setRow(10, def_row_height, .default, level3);
    try worksheet.setRow(11, def_row_height, .default, level2);
    try worksheet.setRow(12, def_row_height, .default, level1);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const def_row_height = xwz.def_row_height;

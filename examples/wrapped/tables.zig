pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet(null);

    const worksheet2 = try workbook.addWorkSheet(null);
    const worksheet3 = try workbook.addWorkSheet(null);
    const worksheet4 = try workbook.addWorkSheet(null);
    const worksheet5 = try workbook.addWorkSheet(null);
    const worksheet6 = try workbook.addWorkSheet(null);
    const worksheet7 = try workbook.addWorkSheet(null);
    const worksheet8 = try workbook.addWorkSheet(null);
    const worksheet9 = try workbook.addWorkSheet(null);
    const worksheet10 = try workbook.addWorkSheet(null);
    const worksheet11 = try workbook.addWorkSheet(null);
    const worksheet12 = try workbook.addWorkSheet(null);
    const worksheet13 = try workbook.addWorkSheet(null);

    const currency_format = try workbook.addFormat();
    currency_format.setNumFormat("$#,##0");

    // Example 1. Default table with no data
    try write_worksheet(
        worksheet1,
        "Default table with no data.",
    );
    try worksheet1.addTable(2, 1, 6, 5, .empty);

    // Example 2: Default table with data
    try write_worksheet(
        worksheet2,
        "Default table with data.",
    );
    try write_worksheet_data(worksheet2, .none);
    try worksheet2.addTable(2, 1, 6, 5, .empty);

    // Example 3: Table without autofilter
    try write_worksheet(
        worksheet3,
        "Table without autofilter.",
    );
    try write_worksheet_data(worksheet3, .none);
    const options3 = xlsxwriter.TableOptions{
        .no_autofilter = xlsxwriter.xw_true,
    };
    try worksheet3.addTable(2, 1, 6, 5, options3);

    // Example 4: Table without default header row
    try write_worksheet(
        worksheet4,
        "Table without default header row.",
    );
    try write_worksheet_data(worksheet4, .none);
    const options4 = xlsxwriter.TableOptions{
        .no_header_row = xlsxwriter.xw_true,
    };
    try worksheet4.addTable(3, 1, 6, 5, options4);

    // Example 5: Default table with "First Column" and "Last Column" options
    try write_worksheet(
        worksheet5,
        "Default table with \"First Column\" and \"Last Column\" options.",
    );
    const options5 = xlsxwriter.TableOptions{
        .first_column = xlsxwriter.xw_true,
        .last_column = xlsxwriter.xw_true,
    };
    try write_worksheet_data(worksheet5, .none);
    try worksheet5.addTable(2, 1, 6, 5, options5);

    // Example 6: Table with banded columns but without default banded rows
    try write_worksheet(
        worksheet6,
        "Table with banded columns but without default banded rows.",
    );
    try write_worksheet_data(worksheet6, .none);
    const options6 = xlsxwriter.TableOptions{
        .no_banded_rows = xlsxwriter.xw_true,
        .banded_columns = xlsxwriter.xw_true,
    };
    try worksheet6.addTable(2, 1, 6, 5, options6);

    // Example 7: Table with user defined column headers
    try write_worksheet(
        worksheet7,
        "Table with user defined column headers.",
    );
    try write_worksheet_data(worksheet7, .none);
    var col7_1 = xlsxwriter.TableColumn{ .header = "Product" };
    var col7_2 = xlsxwriter.TableColumn{ .header = "Quarter 1" };
    var col7_3 = xlsxwriter.TableColumn{ .header = "Quarter 2" };
    var col7_4 = xlsxwriter.TableColumn{ .header = "Quarter 3" };
    var col7_5 = xlsxwriter.TableColumn{ .header = "Quarter 4" };
    const columns7: xlsxwriter.TableColumnArray = &.{
        &col7_1,
        &col7_2,
        &col7_3,
        &col7_4,
        &col7_5,
    };
    const options7 = xlsxwriter.TableOptions{
        .columns = columns7,
    };
    try worksheet7.addTable(2, 1, 6, 5, options7);

    // Example 8: Table with user defined column headers and formula
    try write_worksheet(
        worksheet8,
        "Table with user defined column headers.",
    );
    try write_worksheet_data(worksheet8, .none);
    var col8_1 = xlsxwriter.TableColumn{ .header = "Product" };
    var col8_2 = xlsxwriter.TableColumn{ .header = "Quarter 1" };
    var col8_3 = xlsxwriter.TableColumn{ .header = "Quarter 2" };
    var col8_4 = xlsxwriter.TableColumn{ .header = "Quarter 3" };
    var col8_5 = xlsxwriter.TableColumn{ .header = "Quarter 4" };
    var col8_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table8[@[Quarter 1]:[Quarter 4]])",
    };
    const columns8: xlsxwriter.TableColumnArray = &.{
        &col8_1,
        &col8_2,
        &col8_3,
        &col8_4,
        &col8_5,
        &col8_6,
    };
    const options8 = xlsxwriter.TableOptions{
        .columns = columns8,
    };
    try worksheet8.addTable(2, 1, 6, 6, options8);

    // Example 9: Table with totals row (but no caption or totals)
    try write_worksheet(
        worksheet9,
        "Table with totals row (but no caption or totals).",
    );
    try write_worksheet_data(worksheet9, .none);
    var col9_1 = xlsxwriter.TableColumn{ .header = "Product" };
    var col9_2 = xlsxwriter.TableColumn{ .header = "Quarter 1" };
    var col9_3 = xlsxwriter.TableColumn{ .header = "Quarter 2" };
    var col9_4 = xlsxwriter.TableColumn{ .header = "Quarter 3" };
    var col9_5 = xlsxwriter.TableColumn{ .header = "Quarter 4" };
    var col9_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table9[@[Quarter 1]:[Quarter 4]])",
    };
    const columns9: xlsxwriter.TableColumnArray = &.{
        &col9_1,
        &col9_2,
        &col9_3,
        &col9_4,
        &col9_5,
        &col9_6,
    };
    const options9 = xlsxwriter.TableOptions{
        .total_row = xlsxwriter.xw_true,
        .columns = columns9,
    };
    try worksheet9.addTable(2, 1, 7, 6, options9);

    // Example 10: Table with totals row with user captions and functions
    try write_worksheet(
        worksheet10,
        "Table with totals row with user captions and functions.",
    );
    try write_worksheet_data(worksheet10, .none);
    var col10_1 = xlsxwriter.TableColumn{
        .header = "Product",
        .total_string = "Totals",
    };
    var col10_2 = xlsxwriter.TableColumn{
        .header = "Quarter 1",
        .total_function = .sum,
    };
    var col10_3 = xlsxwriter.TableColumn{
        .header = "Quarter 2",
        .total_function = .sum,
    };
    var col10_4 = xlsxwriter.TableColumn{
        .header = "Quarter 3",
        .total_function = .sum,
    };
    var col10_5 = xlsxwriter.TableColumn{
        .header = "Quarter 4",
        .total_function = .sum,
    };
    var col10_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table10[@[Quarter 1]:[Quarter 4]])",
        .total_function = .sum,
    };
    const columns10: xlsxwriter.TableColumnArray = &.{
        &col10_1,
        &col10_2,
        &col10_3,
        &col10_4,
        &col10_5,
        &col10_6,
    };
    const options10 = xlsxwriter.TableOptions{
        .total_row = xlsxwriter.xw_true,
        .columns = columns10,
    };
    try worksheet10.addTable(2, 1, 7, 6, options10);

    // Example 11: Table with alternative Excel style
    try write_worksheet(
        worksheet11,
        "Table with alternative Excel style.",
    );
    try write_worksheet_data(worksheet11, .none);
    var col11_1 = xlsxwriter.TableColumn{
        .header = "Product",
        .total_string = "Totals",
    };
    var col11_2 = xlsxwriter.TableColumn{
        .header = "Quarter 1",
        .total_function = .sum,
    };
    var col11_3 = xlsxwriter.TableColumn{
        .header = "Quarter 2",
        .total_function = .sum,
    };
    var col11_4 = xlsxwriter.TableColumn{
        .header = "Quarter 3",
        .total_function = .sum,
    };
    var col11_5 = xlsxwriter.TableColumn{
        .header = "Quarter 4",
        .total_function = .sum,
    };
    var col11_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table11[@[Quarter 1]:[Quarter 4]])",
        .total_function = .sum,
    };
    const columns11: xlsxwriter.TableColumnArray = &.{
        &col11_1,
        &col11_2,
        &col11_3,
        &col11_4,
        &col11_5,
        &col11_6,
    };
    const options11 = xlsxwriter.TableOptions{
        .style_type = .light,
        .style_type_number = 11,
        .total_row = xlsxwriter.xw_true,
        .columns = columns11,
    };
    try worksheet11.addTable(2, 1, 7, 6, options11);

    // Example 12: Table with Excel style removed
    try write_worksheet(
        worksheet12,
        "Table with Excel style removed.",
    );
    try write_worksheet_data(worksheet12, .none);
    var col12_1 = xlsxwriter.TableColumn{
        .header = "Product",
        .total_string = "Totals",
    };
    var col12_2 = xlsxwriter.TableColumn{
        .header = "Quarter 1",
        .total_function = .sum,
    };
    var col12_3 = xlsxwriter.TableColumn{
        .header = "Quarter 2",
        .total_function = .sum,
    };
    var col12_4 = xlsxwriter.TableColumn{
        .header = "Quarter 3",
        .total_function = .sum,
    };
    var col12_5 = xlsxwriter.TableColumn{
        .header = "Quarter 4",
        .total_function = .sum,
    };
    var col12_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table12[@[Quarter 1]:[Quarter 4]])",
        .total_function = .sum,
    };
    const columns12: xlsxwriter.TableColumnArray = &.{
        &col12_1,
        &col12_2,
        &col12_3,
        &col12_4,
        &col12_5,
        &col12_6,
    };
    const options12 = xlsxwriter.TableOptions{
        .style_type = .light,
        .total_row = xlsxwriter.xw_true,
        .columns = columns12,
    };
    try worksheet12.addTable(2, 1, 7, 6, options12);

    // Example 13: Table with column formats
    try write_worksheet(
        worksheet13,
        "Table with column formats.",
    );
    try write_worksheet_data(worksheet13, currency_format);
    var col13_1 = xlsxwriter.TableColumn{
        .header = "Product",
        .total_string = "Totals",
    };
    var col13_2 = xlsxwriter.TableColumn{
        .header = "Quarter 1",
        .format_c = currency_format.format_c,
        .total_function = .sum,
    };
    var col13_3 = xlsxwriter.TableColumn{
        .header = "Quarter 2",
        .format_c = currency_format.format_c,
        .total_function = .sum,
    };
    var col13_4 = xlsxwriter.TableColumn{
        .header = "Quarter 3",
        .format_c = currency_format.format_c,
        .total_function = .sum,
    };
    var col13_5 = xlsxwriter.TableColumn{
        .header = "Quarter 4",
        .format_c = currency_format.format_c,
        .total_function = .sum,
    };
    var col13_6 = xlsxwriter.TableColumn{
        .header = "Year",
        .formula = "=SUM(Table13[@[Quarter 1]:[Quarter 4]])",
        .format_c = currency_format.format_c,
        .total_function = .sum,
    };
    const columns13: xlsxwriter.TableColumnArray = &.{
        &col13_1,
        &col13_2,
        &col13_3,
        &col13_4,
        &col13_5,
        &col13_6,
    };
    const options13 = xlsxwriter.TableOptions{
        .total_row = xlsxwriter.xw_true,
        .columns = columns13,
    };
    try worksheet13.addTable(2, 1, 7, 6, options13);
}

fn write_worksheet(
    worksheet: WorkSheet,
    title: [*c]const u8,
) !void {
    // Example 1: Default table with no data
    try worksheet.setColumn(1, 6, 12, .none);
    try worksheet.writeString(0, 1, title, .none);
}

fn write_worksheet_data(worksheet: WorkSheet, format: Format) !void {
    // array of strings "Apples", "Pears", "Bananas", "Oranges"
    const rowDescriptions =
        [_][:0]const u8{ "Apples", "Pears", "Bananas", "Oranges" };

    // data is a 4x4 array of f64
    const data = [_][4]f64{
        [4]f64{ 10000, 5000, 8000, 6000 },
        [4]f64{ 2000, 3000, 4000, 5000 },
        [4]f64{ 6000, 6000, 6500, 6000 },
        [4]f64{ 500, 300, 200, 700 },
    };
    const start_row: usize = 3;

    for (rowDescriptions, start_row..) |str, i| {
        // Write the first row strings
        try worksheet.writeString(@intCast(i), 1, str, format);
    }

    for (data, start_row..) |row, i| {
        for (row, 2..) |value, j| {
            try worksheet.writeNumber(@intCast(i), @intCast(j), value, format);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;

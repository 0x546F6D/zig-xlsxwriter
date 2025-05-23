// A simple function and data structure to populate some of the worksheets.
const WorksheetData = struct {
    col1: [:0]const u8,
    col2: [:0]const u8,
    col3: [:0]const u8,
    col4: i32,
};

const data = [_]WorksheetData{
    .{ .col1 = "East", .col2 = "Tom", .col3 = "Apple", .col4 = 6380 },
    .{ .col1 = "West", .col2 = "Fred", .col3 = "Grape", .col4 = 5619 },
    .{ .col1 = "North", .col2 = "Amy", .col3 = "Pear", .col4 = 4565 },
    .{ .col1 = "South", .col2 = "Sal", .col3 = "Banana", .col4 = 5323 },
    .{ .col1 = "East", .col2 = "Fritz", .col3 = "Apple", .col4 = 4394 },
    .{ .col1 = "West", .col2 = "Sravan", .col3 = "Grape", .col4 = 7195 },
    .{ .col1 = "North", .col2 = "Xi", .col3 = "Pear", .col4 = 5231 },
    .{ .col1 = "South", .col2 = "Hector", .col3 = "Banana", .col4 = 2427 },
    .{ .col1 = "East", .col2 = "Tom", .col3 = "Banana", .col4 = 4213 },
    .{ .col1 = "West", .col2 = "Fred", .col3 = "Pear", .col4 = 3239 },
    .{ .col1 = "North", .col2 = "Amy", .col3 = "Grape", .col4 = 6520 },
    .{ .col1 = "South", .col2 = "Sal", .col3 = "Apple", .col4 = 1310 },
    .{ .col1 = "East", .col2 = "Fritz", .col3 = "Banana", .col4 = 6274 },
    .{ .col1 = "West", .col2 = "Sravan", .col3 = "Pear", .col4 = 4894 },
    .{ .col1 = "North", .col2 = "Xi", .col3 = "Grape", .col4 = 7580 },
    .{ .col1 = "South", .col2 = "Hector", .col3 = "Apple", .col4 = 9814 },
};

fn writeWorksheetData(worksheet: WorkSheet, header: Format) void {
    worksheet.writeString(0, 0, "Region", header);
    worksheet.writeString(0, 1, "Sales Rep", header);
    worksheet.writeString(0, 2, "Product", header);
    worksheet.writeString(0, 3, "Units", header);

    for (data, 1..) |item, row| {
        worksheet.writeString(@intCast(row), 0, item.col1.ptr, .none);
        worksheet.writeString(@intCast(row), 1, item.col2.ptr, .none);
        worksheet.writeString(@intCast(row), 2, item.col3.ptr, .none);
        worksheet.writeNumber(@intCast(row), 3, @floatFromInt(item.col4), .none);
    }
}

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet("Filter");
    const worksheet2 = try workbook.addWorkSheet("Unique");
    const worksheet3 = try workbook.addWorkSheet("Sort");
    const worksheet4 = try workbook.addWorkSheet("Sortby");
    const worksheet5 = try workbook.addWorkSheet("Xlookup");
    const worksheet6 = try workbook.addWorkSheet("Xmatch");
    const worksheet7 = try workbook.addWorkSheet("Randarray");
    const worksheet8 = try workbook.addWorkSheet("Sequence");
    const worksheet9 = try workbook.addWorkSheet("Spill ranges");
    const worksheet10 = try workbook.addWorkSheet("Older functions");

    const header1 = try workbook.addFormat();
    header1.setBgColor(@enumFromInt(0x74AC4C));
    header1.setFontColor(@enumFromInt(0xFFFFFF));

    const header2 = try workbook.addFormat();
    header2.setBgColor(@enumFromInt(0x528FD3));
    header2.setFontColor(@enumFromInt(0xFFFFFF));

    // WorkSheet1: Example of using the FILTER() function.
    try worksheet1.writeDynamicFormula(
        1,
        5,
        "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet1.writeString(0, 10, "Product", header2);
    try worksheet1.writeString(1, 10, "Apple", .none);
    try worksheet1.writeString(0, 5, "Region", header2);
    try worksheet1.writeString(0, 6, "Sales Rep", header2);
    try worksheet1.writeString(0, 7, "Product", header2);
    try worksheet1.writeString(0, 8, "Units", header2);

    writeWorksheetData(worksheet1, header1);
    try worksheet1.setColumnPixels(4, 4, 20, .none);
    try worksheet1.setColumnPixels(9, 9, 20, .none);

    // WorkSheet2: Example of using the UNIQUE() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet2,
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        null,
    );

    // A more complex example combining SORT and UNIQUE.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet2,
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 5, "Sales Rep", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet2, 0, 7, "Sales Rep", header2);

    writeWorksheetData(worksheet2, header1);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet2, 4, 4, 20, null);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet2, 6, 6, 20, null);

    // Example of using the SORT() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet3,
        1,
        5,
        "=_xlfn._xlws.SORT(B2:B17)",
        null,
    );

    // WorkSheet3: A more complex example combining SORT and FILTER.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet3,
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 5, "Sales Rep", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 7, "Product", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet3, 0, 8, "Units", header2);

    writeWorksheetData(worksheet3, header1);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet3, 4, 4, 20, null);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet3, 6, 6, 20, null);

    // WorkSheet4: Example of using the SORTBY() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet4,
        1,
        3,
        "=_xlfn.SORTBY(A2:B9,B2:B9)",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 0, "Name", header1);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 1, "Age", header1);

    _ = xlsxwriter.worksheet_write_string(worksheet4, 1, 0, "Tom", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 2, 0, "Fred", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 3, 0, "Amy", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 4, 0, "Sal", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 5, 0, "Fritz", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 6, 0, "Srivan", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 7, 0, "Xi", null);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 8, 0, "Hector", null);

    _ = xlsxwriter.worksheet_write_number(worksheet4, 1, 1, 52, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 2, 1, 65, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 3, 1, 22, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 4, 1, 73, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 5, 1, 19, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 6, 1, 39, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 7, 1, 19, null);
    _ = xlsxwriter.worksheet_write_number(worksheet4, 8, 1, 66, null);

    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 3, "Name", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet4, 0, 4, "Age", header2);

    _ = xlsxwriter.worksheet_set_column_pixels(worksheet4, 2, 2, 20, null);

    // WorkSheet5: Example of using the XLOOKUP() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet5,
        0,
        5,
        "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet5, 0, 0, "Country", header1);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 0, 1, "Abr", header1);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 0, 2, "Prefix", header1);

    _ = xlsxwriter.worksheet_write_string(worksheet5, 1, 0, "China", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 2, 0, "India", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 3, 0, "United States", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 4, 0, "Indonesia", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 5, 0, "Brazil", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 6, 0, "Pakistan", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 7, 0, "Nigeria", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 8, 0, "Bangladesh", null);

    _ = xlsxwriter.worksheet_write_string(worksheet5, 1, 1, "CN", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 2, 1, "IN", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 3, 1, "US", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 4, 1, "ID", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 5, 1, "BR", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 6, 1, "PK", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 7, 1, "NG", null);
    _ = xlsxwriter.worksheet_write_string(worksheet5, 8, 1, "BD", null);

    _ = xlsxwriter.worksheet_write_number(worksheet5, 1, 2, 86, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 2, 2, 91, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 3, 2, 1, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 4, 2, 62, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 5, 2, 55, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 6, 2, 92, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 7, 2, 234, null);
    _ = xlsxwriter.worksheet_write_number(worksheet5, 8, 2, 880, null);

    _ = xlsxwriter.worksheet_write_string(worksheet5, 0, 4, "Brazil", header2);

    _ = xlsxwriter.worksheet_set_column_pixels(worksheet5, 0, 0, 100, null);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet5, 3, 3, 20, null);

    // Example of using the XMATCH() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet6,
        1,
        3,
        "=_xlfn.XMATCH(C2,A2:A6)",
        null,
    );

    // WorkSheet6: Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet6, 0, 0, "Product", header1);

    _ = xlsxwriter.worksheet_write_string(worksheet6, 1, 0, "Apple", null);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 2, 0, "Grape", null);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 3, 0, "Pear", null);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 4, 0, "Banana", null);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 5, 0, "Cherry", null);

    _ = xlsxwriter.worksheet_write_string(worksheet6, 0, 2, "Product", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 0, 3, "Position", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet6, 1, 2, "Grape", null);

    _ = xlsxwriter.worksheet_set_column_pixels(worksheet6, 1, 1, 20, null);

    // WorkSheet7: Example of using the RANDARRAY() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet7,
        0,
        0,
        "=_xlfn.RANDARRAY(5,3,1,100, TRUE)",
        null,
    );

    // WorkSheet8: Example of using the SEQUENCE() function.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet8,
        0,
        0,
        "=_xlfn.SEQUENCE(4,5)",
        null,
    );

    // WorkSheet9: Example of using the Spill range operator.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        7,
        "=_xlfn.ANCHORARRAY(F2)",
        null,
    );

    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        9,
        "=COUNTA(_xlfn.ANCHORARRAY(F2))",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        null,
    );

    _ = xlsxwriter.worksheet_write_string(worksheet9, 0, 5, "Unique", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 0, 7, "Spill", header2);
    _ = xlsxwriter.worksheet_write_string(worksheet9, 0, 9, "Spill", header2);

    writeWorksheetData(worksheet9, header1);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet9, 4, 4, 20, null);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet9, 6, 6, 20, null);
    _ = xlsxwriter.worksheet_set_column_pixels(worksheet9, 8, 8, 20, null);

    // WorkSheet10: Example of using dynamic ranges with older Excel functions.
    _ = xlsxwriter.worksheet_write_dynamic_array_formula(
        worksheet10,
        0,
        1,
        2,
        1,
        "=LEN(A1:A3)",
        null,
    );

    // Write the data the function will work on.
    _ = xlsxwriter.worksheet_write_string(worksheet10, 0, 0, "Foo", null);
    _ = xlsxwriter.worksheet_write_string(worksheet10, 1, 0, "Food", null);
    _ = xlsxwriter.worksheet_write_string(worksheet10, 2, 0, "Frood", null);

    _ = xlsxwriter.workbook_close(workbook);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;

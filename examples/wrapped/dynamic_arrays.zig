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

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
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

    try writeWorksheetData(worksheet1, header1);
    try worksheet1.setColumnPixels(4, 4, 20, .none);
    try worksheet1.setColumnPixels(9, 9, 20, .none);

    // WorkSheet2: Example of using the UNIQUE() function.
    try worksheet2.writeDynamicFormula(
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        .none,
    );

    // A more complex example combining SORT and UNIQUE.
    try worksheet2.writeDynamicFormula(
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))",
        .none,
    );

    // Write the data the function will work on.
    try worksheet2.writeString(0, 5, "Sales Rep", header2);
    try worksheet2.writeString(0, 7, "Sales Rep", header2);

    try writeWorksheetData(worksheet2, header1);
    try worksheet2.setColumnPixels(4, 4, 20, .none);
    try worksheet2.setColumnPixels(6, 6, 20, .none);

    // WorkSheet3: Example of using the SORT() function.
    try worksheet3.writeDynamicFormula(
        1,
        5,
        "=_xlfn._xlws.SORT(B2:B17)",
        .none,
    );

    // A more complex example combining SORT and FILTER.
    try worksheet3.writeDynamicFormula(
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet3.writeString(0, 5, "Sales Rep", header2);
    try worksheet3.writeString(0, 7, "Product", header2);
    try worksheet3.writeString(0, 8, "Units", header2);

    try writeWorksheetData(worksheet3, header1);
    try worksheet3.setColumnPixels(4, 4, 20, .none);
    try worksheet3.setColumnPixels(6, 6, 20, .none);

    // WorkSheet4: Example of using the SORTBY() function.
    try worksheet4.writeDynamicFormula(
        1,
        3,
        "=_xlfn.SORTBY(A2:B9,B2:B9)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet4.writeString(0, 0, "Name", header1);
    try worksheet4.writeString(0, 1, "Age", header1);

    try worksheet4.writeString(1, 0, "Tom", .none);
    try worksheet4.writeString(2, 0, "Fred", .none);
    try worksheet4.writeString(3, 0, "Amy", .none);
    try worksheet4.writeString(4, 0, "Sal", .none);
    try worksheet4.writeString(5, 0, "Fritz", .none);
    try worksheet4.writeString(6, 0, "Srivan", .none);
    try worksheet4.writeString(7, 0, "Xi", .none);
    try worksheet4.writeString(8, 0, "Hector", .none);

    try worksheet4.writeNumber(1, 1, 52, .none);
    try worksheet4.writeNumber(2, 1, 65, .none);
    try worksheet4.writeNumber(3, 1, 22, .none);
    try worksheet4.writeNumber(4, 1, 73, .none);
    try worksheet4.writeNumber(5, 1, 19, .none);
    try worksheet4.writeNumber(6, 1, 39, .none);
    try worksheet4.writeNumber(7, 1, 19, .none);
    try worksheet4.writeNumber(8, 1, 66, .none);

    try worksheet4.writeString(0, 3, "Name", header2);
    try worksheet4.writeString(0, 4, "Age", header2);

    try worksheet4.setColumnPixels(2, 2, 20, .none);

    // WorkSheet5: Example of using the XLOOKUP() function.
    try worksheet5.writeDynamicFormula(
        0,
        5,
        "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet5.writeString(0, 0, "Country", header1);
    try worksheet5.writeString(0, 1, "Abr", header1);
    try worksheet5.writeString(0, 2, "Prefix", header1);

    try worksheet5.writeString(1, 0, "China", .none);
    try worksheet5.writeString(2, 0, "India", .none);
    try worksheet5.writeString(3, 0, "United States", .none);
    try worksheet5.writeString(4, 0, "Indonesia", .none);
    try worksheet5.writeString(5, 0, "Brazil", .none);
    try worksheet5.writeString(6, 0, "Pakistan", .none);
    try worksheet5.writeString(7, 0, "Nigeria", .none);
    try worksheet5.writeString(8, 0, "Bangladesh", .none);

    try worksheet5.writeString(1, 1, "CN", .none);
    try worksheet5.writeString(2, 1, "IN", .none);
    try worksheet5.writeString(3, 1, "US", .none);
    try worksheet5.writeString(4, 1, "ID", .none);
    try worksheet5.writeString(5, 1, "BR", .none);
    try worksheet5.writeString(6, 1, "PK", .none);
    try worksheet5.writeString(7, 1, "NG", .none);
    try worksheet5.writeString(8, 1, "BD", .none);

    try worksheet5.writeNumber(1, 2, 86, .none);
    try worksheet5.writeNumber(2, 2, 91, .none);
    try worksheet5.writeNumber(3, 2, 1, .none);
    try worksheet5.writeNumber(4, 2, 62, .none);
    try worksheet5.writeNumber(5, 2, 55, .none);
    try worksheet5.writeNumber(6, 2, 92, .none);
    try worksheet5.writeNumber(7, 2, 234, .none);
    try worksheet5.writeNumber(8, 2, 880, .none);

    try worksheet5.writeString(0, 4, "Brazil", header2);

    try worksheet5.setColumnPixels(0, 0, 100, .none);
    try worksheet5.setColumnPixels(3, 3, 20, .none);

    // WorkSheet6: Example of using the XMATCH() function.
    try worksheet6.writeDynamicFormula(
        1,
        3,
        "=_xlfn.XMATCH(C2,A2:A6)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet6.writeString(0, 0, "Product", header1);

    try worksheet6.writeString(1, 0, "Apple", .none);
    try worksheet6.writeString(2, 0, "Grape", .none);
    try worksheet6.writeString(3, 0, "Pear", .none);
    try worksheet6.writeString(4, 0, "Banana", .none);
    try worksheet6.writeString(5, 0, "Cherry", .none);

    try worksheet6.writeString(0, 2, "Product", header2);
    try worksheet6.writeString(0, 3, "Position", header2);
    try worksheet6.writeString(1, 2, "Grape", .none);

    try worksheet6.setColumnPixels(1, 1, 20, .none);

    // WorkSheet7: Example of using the RANDARRAY() function.
    try worksheet7.writeDynamicFormula(
        0,
        0,
        "=_xlfn.RANDARRAY(5,3,1,100, TRUE)",
        .none,
    );

    // WorkSheet8: Example of using the SEQUENCE() function.
    try worksheet8.writeDynamicFormula(
        0,
        0,
        "=_xlfn.SEQUENCE(4,5)",
        .none,
    );

    // WorkSheet9: Example of using the Spill range operator.
    try worksheet9.writeDynamicFormula(
        1,
        7,
        "=_xlfn.ANCHORARRAY(F2)",
        .none,
    );

    try worksheet9.writeDynamicFormula(
        1,
        9,
        "=COUNTA(_xlfn.ANCHORARRAY(F2))",
        .none,
    );

    // Write the data the function will work on.
    try worksheet9.writeDynamicFormula(
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        .none,
    );

    try worksheet9.writeString(0, 5, "Unique", header2);
    try worksheet9.writeString(0, 7, "Spill", header2);
    try worksheet9.writeString(0, 9, "Spill", header2);

    try writeWorksheetData(worksheet9, header1);
    try worksheet9.setColumnPixels(4, 4, 20, .none);
    try worksheet9.setColumnPixels(6, 6, 20, .none);
    try worksheet9.setColumnPixels(8, 8, 20, .none);

    // WorkSheet10: Example of using dynamic ranges with older Excel functions.
    try worksheet10.writeDynamicArrayFormula(
        0,
        1,
        2,
        1,
        "=LEN(A1:A3)",
        .none,
    );

    // Write the data the function will work on.
    try worksheet10.writeString(0, 0, "Foo", .none);
    try worksheet10.writeString(1, 0, "Food", .none);
    try worksheet10.writeString(2, 0, "Frood", .none);
}

fn writeWorksheetData(worksheet: WorkSheet, header: Format) !void {
    try worksheet.writeString(0, 0, "Region", header);
    try worksheet.writeString(0, 1, "Sales Rep", header);
    try worksheet.writeString(0, 2, "Product", header);
    try worksheet.writeString(0, 3, "Units", header);

    for (data, 1..) |item, row| {
        try worksheet.writeString(@intCast(row), 0, item.col1.ptr, .none);
        try worksheet.writeString(@intCast(row), 1, item.col2.ptr, .none);
        try worksheet.writeString(@intCast(row), 2, item.col3.ptr, .none);
        try worksheet.writeNumber(@intCast(row), 3, @floatFromInt(item.col4), .none);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;

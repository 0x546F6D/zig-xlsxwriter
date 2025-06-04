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
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
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
        .{ .row = 1, .col = 5 },
        "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet1.writeString(.{ .col = 10 }, "Product", header2);
    try worksheet1.writeString(.{ .row = 1, .col = 10 }, "Apple", .default);
    try worksheet1.writeString(.{ .col = 5 }, "Region", header2);
    try worksheet1.writeString(.{ .col = 6 }, "Sales Rep", header2);
    try worksheet1.writeString(.{ .col = 7 }, "Product", header2);
    try worksheet1.writeString(.{ .col = 8 }, "Units", header2);

    try writeWorksheetData(worksheet1, header1);
    try worksheet1.setColumnPixels(.{ .first = 4, .last = 4 }, 20, .default, null);
    try worksheet1.setColumnPixels(.{ .first = 9, .last = 9 }, 20, .default, null);

    // WorkSheet2: Example of using the UNIQUE() function.
    try worksheet2.writeDynamicFormula(
        .{ .row = 1, .col = 5 },
        "=_xlfn.UNIQUE(B2:B17)",
        .default,
    );

    // A more complex example combining SORT and UNIQUE.
    try worksheet2.writeDynamicFormula(
        .{ .row = 1, .col = 7 },
        "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))",
        .default,
    );

    // Write the data the function will work on.
    try worksheet2.writeString(.{ .col = 5 }, "Sales Rep", header2);
    try worksheet2.writeString(.{ .col = 7 }, "Sales Rep", header2);

    try writeWorksheetData(worksheet2, header1);
    try worksheet2.setColumnPixels(.{ .first = 4, .last = 4 }, 20, .default, null);
    try worksheet2.setColumnPixels(.{ .first = 6, .last = 6 }, 20, .default, null);

    // WorkSheet3: Example of using the SORT() function.
    try worksheet3.writeDynamicFormula(
        .{ .row = 1, .col = 5 },
        "=_xlfn._xlws.SORT(B2:B17)",
        .default,
    );

    // A more complex example combining SORT and FILTER.
    try worksheet3.writeDynamicFormula(
        .{ .row = 1, .col = 7 },
        "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet3.writeString(.{ .col = 5 }, "Sales Rep", header2);
    try worksheet3.writeString(.{ .col = 7 }, "Product", header2);
    try worksheet3.writeString(.{ .col = 8 }, "Units", header2);

    try writeWorksheetData(worksheet3, header1);
    try worksheet3.setColumnPixels(.{ .first = 4, .last = 4 }, 20, .default, null);
    try worksheet3.setColumnPixels(.{ .first = 6, .last = 6 }, 20, .default, null);

    // WorkSheet4: Example of using the SORTBY() function.
    try worksheet4.writeDynamicFormula(
        .{ .row = 1, .col = 3 },
        "=_xlfn.SORTBY(A2:B9,B2:B9)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet4.writeString(.{}, "Name", header1);
    try worksheet4.writeString(.{ .col = 1 }, "Age", header1);

    try worksheet4.writeString(.{ .row = 1 }, "Tom", .default);
    try worksheet4.writeString(.{ .row = 2 }, "Fred", .default);
    try worksheet4.writeString(.{ .row = 3 }, "Amy", .default);
    try worksheet4.writeString(.{ .row = 4 }, "Sal", .default);
    try worksheet4.writeString(.{ .row = 5 }, "Fritz", .default);
    try worksheet4.writeString(.{ .row = 6 }, "Srivan", .default);
    try worksheet4.writeString(.{ .row = 7 }, "Xi", .default);
    try worksheet4.writeString(.{ .row = 8 }, "Hector", .default);

    try worksheet4.writeNumber(.{ .row = 1, .col = 1 }, 52, .default);
    try worksheet4.writeNumber(.{ .row = 2, .col = 1 }, 65, .default);
    try worksheet4.writeNumber(.{ .row = 3, .col = 1 }, 22, .default);
    try worksheet4.writeNumber(.{ .row = 4, .col = 1 }, 73, .default);
    try worksheet4.writeNumber(.{ .row = 5, .col = 1 }, 19, .default);
    try worksheet4.writeNumber(.{ .row = 6, .col = 1 }, 39, .default);
    try worksheet4.writeNumber(.{ .row = 7, .col = 1 }, 19, .default);
    try worksheet4.writeNumber(.{ .row = 8, .col = 1 }, 66, .default);

    try worksheet4.writeString(.{ .col = 3 }, "Name", header2);
    try worksheet4.writeString(.{ .col = 4 }, "Age", header2);

    try worksheet4.setColumnPixels(.{ .first = 2, .last = 2 }, 20, .default, null);

    // WorkSheet5: Example of using the XLOOKUP() function.
    try worksheet5.writeDynamicFormula(
        .{ .col = 5 },
        "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet5.writeString(.{}, "Country", header1);
    try worksheet5.writeString(.{ .col = 1 }, "Abr", header1);
    try worksheet5.writeString(.{ .col = 2 }, "Prefix", header1);

    try worksheet5.writeString(.{ .row = 1 }, "China", .default);
    try worksheet5.writeString(.{ .row = 2 }, "India", .default);
    try worksheet5.writeString(.{ .row = 3 }, "United States", .default);
    try worksheet5.writeString(.{ .row = 4 }, "Indonesia", .default);
    try worksheet5.writeString(.{ .row = 5 }, "Brazil", .default);
    try worksheet5.writeString(.{ .row = 6 }, "Pakistan", .default);
    try worksheet5.writeString(.{ .row = 7 }, "Nigeria", .default);
    try worksheet5.writeString(.{ .row = 8 }, "Bangladesh", .default);

    try worksheet5.writeString(.{ .row = 1, .col = 1 }, "CN", .default);
    try worksheet5.writeString(.{ .row = 2, .col = 1 }, "IN", .default);
    try worksheet5.writeString(.{ .row = 3, .col = 1 }, "US", .default);
    try worksheet5.writeString(.{ .row = 4, .col = 1 }, "ID", .default);
    try worksheet5.writeString(.{ .row = 5, .col = 1 }, "BR", .default);
    try worksheet5.writeString(.{ .row = 6, .col = 1 }, "PK", .default);
    try worksheet5.writeString(.{ .row = 7, .col = 1 }, "NG", .default);
    try worksheet5.writeString(.{ .row = 8, .col = 1 }, "BD", .default);

    try worksheet5.writeNumber(.{ .row = 1, .col = 2 }, 86, .default);
    try worksheet5.writeNumber(.{ .row = 2, .col = 2 }, 91, .default);
    try worksheet5.writeNumber(.{ .row = 3, .col = 2 }, 1, .default);
    try worksheet5.writeNumber(.{ .row = 4, .col = 2 }, 62, .default);
    try worksheet5.writeNumber(.{ .row = 5, .col = 2 }, 55, .default);
    try worksheet5.writeNumber(.{ .row = 6, .col = 2 }, 92, .default);
    try worksheet5.writeNumber(.{ .row = 7, .col = 2 }, 234, .default);
    try worksheet5.writeNumber(.{ .row = 8, .col = 2 }, 880, .default);

    try worksheet5.writeString(.{ .col = 4 }, "Brazil", header2);

    try worksheet5.setColumnPixels(.{}, 100, .default, null);
    try worksheet5.setColumnPixels(.{ .first = 3, .last = 3 }, 20, .default, null);

    // WorkSheet6: Example of using the XMATCH() function.
    try worksheet6.writeDynamicFormula(
        .{ .row = 1, .col = 3 },
        "=_xlfn.XMATCH(C2,A2:A6)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet6.writeString(.{}, "Product", header1);

    try worksheet6.writeString(.{ .row = 1 }, "Apple", .default);
    try worksheet6.writeString(.{ .row = 2 }, "Grape", .default);
    try worksheet6.writeString(.{ .row = 3 }, "Pear", .default);
    try worksheet6.writeString(.{ .row = 4 }, "Banana", .default);
    try worksheet6.writeString(.{ .row = 5 }, "Cherry", .default);

    try worksheet6.writeString(.{ .col = 2 }, "Product", header2);
    try worksheet6.writeString(.{ .col = 3 }, "Position", header2);
    try worksheet6.writeString(.{ .row = 1, .col = 2 }, "Grape", .default);

    try worksheet6.setColumnPixels(.{ .first = 1, .last = 1 }, 20, .default, null);

    // WorkSheet7: Example of using the RANDARRAY() function.
    try worksheet7.writeDynamicFormula(
        .{},
        "=_xlfn.RANDARRAY(5,3,1,100, TRUE)",
        .default,
    );

    // WorkSheet8: Example of using the SEQUENCE() function.
    try worksheet8.writeDynamicFormula(
        .{},
        "=_xlfn.SEQUENCE(4,5)",
        .default,
    );

    // WorkSheet9: Example of using the Spill range operator.
    try worksheet9.writeDynamicFormula(
        .{ .row = 1, .col = 7 },
        "=_xlfn.ANCHORARRAY(F2)",
        .default,
    );

    try worksheet9.writeDynamicFormula(
        .{ .row = 1, .col = 9 },
        "=COUNTA(_xlfn.ANCHORARRAY(F2))",
        .default,
    );

    // Write the data the function will work on.
    try worksheet9.writeDynamicFormula(
        .{ .row = 1, .col = 5 },
        "=_xlfn.UNIQUE(B2:B17)",
        .default,
    );

    try worksheet9.writeString(.{ .col = 5 }, "Unique", header2);
    try worksheet9.writeString(.{ .col = 7 }, "Spill", header2);
    try worksheet9.writeString(.{ .col = 9 }, "Spill", header2);

    try writeWorksheetData(worksheet9, header1);
    try worksheet9.setColumnPixels(.{ .first = 4, .last = 4 }, 20, .default, null);
    try worksheet9.setColumnPixels(.{ .first = 6, .last = 6 }, 20, .default, null);
    try worksheet9.setColumnPixels(.{ .first = 8, .last = 8 }, 20, .default, null);

    // WorkSheet10: Example of using dynamic ranges with older Excel functions.
    try worksheet10.writeDynamicArrayFormula(
        .{ .first_col = 1, .last_row = 2, .last_col = 1 },
        "=LEN(A1:A3)",
        .default,
    );

    // Write the data the function will work on.
    try worksheet10.writeString(.{}, "Foo", .default);
    try worksheet10.writeString(.{ .row = 1 }, "Food", .default);
    try worksheet10.writeString(.{ .row = 2 }, "Frood", .default);
}

fn writeWorksheetData(worksheet: WorkSheet, header: Format) !void {
    try worksheet.writeString(.{}, "Region", header);
    try worksheet.writeString(.{ .col = 1 }, "Sales Rep", header);
    try worksheet.writeString(.{ .col = 2 }, "Product", header);
    try worksheet.writeString(.{ .col = 3 }, "Units", header);

    for (data, 1..) |item, row| {
        try worksheet.writeString(.{ .row = @intCast(row) }, item.col1, .default);
        try worksheet.writeString(.{ .row = @intCast(row), .col = 1 }, item.col2, .default);
        try worksheet.writeString(.{ .row = @intCast(row), .col = 2 }, item.col3, .default);
        try worksheet.writeNumber(.{ .row = @intCast(row), .col = 3 }, @floatFromInt(item.col4), .default);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;

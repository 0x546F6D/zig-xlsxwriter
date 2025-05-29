//
// An example of how to use libxlsxwriter to write functions that create
// dynamic arrays. These functions are new to Excel 365. The examples mirror
// the examples in the Excel documentation on these functions.
//
// Copyright 2014-2025, John McNamara, jmcnamara@cpan.org
//
//

const std = @import("std");
const lxw = @import("lxw");

// A simple function and data structure to populate some of the worksheets.
const WorksheetData = struct {
    col1: [:0]const u8,
    col2: [:0]const u8,
    col3: [:0]const u8,
    col4: i32,
};

fn writeWorksheetData(worksheet: *lxw.lxw_worksheet, header: *lxw.lxw_format) void {
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

    _ = lxw.worksheet_write_string(worksheet, 0, 0, "Region", header);
    _ = lxw.worksheet_write_string(worksheet, 0, 1, "Sales Rep", header);
    _ = lxw.worksheet_write_string(worksheet, 0, 2, "Product", header);
    _ = lxw.worksheet_write_string(worksheet, 0, 3, "Units", header);

    for (data, 0..) |item, row| {
        _ = lxw.worksheet_write_string(
            worksheet,
            @intCast(row + 1),
            0,
            item.col1.ptr,
            null,
        );
        _ = lxw.worksheet_write_string(
            worksheet,
            @intCast(row + 1),
            1,
            item.col2.ptr,
            null,
        );
        _ = lxw.worksheet_write_string(
            worksheet,
            @intCast(row + 1),
            2,
            item.col3.ptr,
            null,
        );
        _ = lxw.worksheet_write_number(
            worksheet,
            @intCast(row + 1),
            3,
            @floatFromInt(item.col4),
            null,
        );
    }
}

pub fn main() !void {
    const workbook = lxw.workbook_new("zig-dynamic_arrays.xlsx");
    const worksheet1 = lxw.workbook_add_worksheet(workbook, "Filter");
    const worksheet2 = lxw.workbook_add_worksheet(workbook, "Unique");
    const worksheet3 = lxw.workbook_add_worksheet(workbook, "Sort");
    const worksheet4 = lxw.workbook_add_worksheet(workbook, "Sortby");
    const worksheet5 = lxw.workbook_add_worksheet(workbook, "Xlookup");
    const worksheet6 = lxw.workbook_add_worksheet(workbook, "Xmatch");
    const worksheet7 = lxw.workbook_add_worksheet(workbook, "Randarray");
    const worksheet8 = lxw.workbook_add_worksheet(workbook, "Sequence");
    const worksheet9 = lxw.workbook_add_worksheet(workbook, "Spill ranges");
    const worksheet10 = lxw.workbook_add_worksheet(workbook, "Older functions");

    const header1 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bg_color(header1, 0x74AC4C);
    _ = lxw.format_set_font_color(header1, 0xFFFFFF);

    const header2 = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bg_color(header2, 0x528FD3);
    _ = lxw.format_set_font_color(header2, 0xFFFFFF);

    // Example of using the FILTER() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet1,
        1,
        5,
        "=_xlfn._xlws.FILTER(A1:D17,C1:C17=K2)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet1, 0, 10, "Product", header2);
    _ = lxw.worksheet_write_string(worksheet1, 1, 10, "Apple", null);
    _ = lxw.worksheet_write_string(worksheet1, 0, 5, "Region", header2);
    _ = lxw.worksheet_write_string(worksheet1, 0, 6, "Sales Rep", header2);
    _ = lxw.worksheet_write_string(worksheet1, 0, 7, "Product", header2);
    _ = lxw.worksheet_write_string(worksheet1, 0, 8, "Units", header2);

    writeWorksheetData(worksheet1, header1);
    _ = lxw.worksheet_set_column_pixels(worksheet1, 4, 4, 20, null);
    _ = lxw.worksheet_set_column_pixels(worksheet1, 9, 9, 20, null);

    // Example of using the UNIQUE() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet2,
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        null,
    );

    // A more complex example combining SORT and UNIQUE.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet2,
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn.UNIQUE(B2:B17))",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet2, 0, 5, "Sales Rep", header2);
    _ = lxw.worksheet_write_string(worksheet2, 0, 7, "Sales Rep", header2);

    writeWorksheetData(worksheet2, header1);
    _ = lxw.worksheet_set_column_pixels(worksheet2, 4, 4, 20, null);
    _ = lxw.worksheet_set_column_pixels(worksheet2, 6, 6, 20, null);

    // Example of using the SORT() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet3,
        1,
        5,
        "=_xlfn._xlws.SORT(B2:B17)",
        null,
    );

    // A more complex example combining SORT and FILTER.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet3,
        1,
        7,
        "=_xlfn._xlws.SORT(_xlfn._xlws.FILTER(C2:D17,D2:D17>5000,\"\"),2,1)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet3, 0, 5, "Sales Rep", header2);
    _ = lxw.worksheet_write_string(worksheet3, 0, 7, "Product", header2);
    _ = lxw.worksheet_write_string(worksheet3, 0, 8, "Units", header2);

    writeWorksheetData(worksheet3, header1);
    _ = lxw.worksheet_set_column_pixels(worksheet3, 4, 4, 20, null);
    _ = lxw.worksheet_set_column_pixels(worksheet3, 6, 6, 20, null);

    // Example of using the SORTBY() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet4,
        1,
        3,
        "=_xlfn.SORTBY(A2:B9,B2:B9)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet4, 0, 0, "Name", header1);
    _ = lxw.worksheet_write_string(worksheet4, 0, 1, "Age", header1);

    _ = lxw.worksheet_write_string(worksheet4, 1, 0, "Tom", null);
    _ = lxw.worksheet_write_string(worksheet4, 2, 0, "Fred", null);
    _ = lxw.worksheet_write_string(worksheet4, 3, 0, "Amy", null);
    _ = lxw.worksheet_write_string(worksheet4, 4, 0, "Sal", null);
    _ = lxw.worksheet_write_string(worksheet4, 5, 0, "Fritz", null);
    _ = lxw.worksheet_write_string(worksheet4, 6, 0, "Srivan", null);
    _ = lxw.worksheet_write_string(worksheet4, 7, 0, "Xi", null);
    _ = lxw.worksheet_write_string(worksheet4, 8, 0, "Hector", null);

    _ = lxw.worksheet_write_number(worksheet4, 1, 1, 52, null);
    _ = lxw.worksheet_write_number(worksheet4, 2, 1, 65, null);
    _ = lxw.worksheet_write_number(worksheet4, 3, 1, 22, null);
    _ = lxw.worksheet_write_number(worksheet4, 4, 1, 73, null);
    _ = lxw.worksheet_write_number(worksheet4, 5, 1, 19, null);
    _ = lxw.worksheet_write_number(worksheet4, 6, 1, 39, null);
    _ = lxw.worksheet_write_number(worksheet4, 7, 1, 19, null);
    _ = lxw.worksheet_write_number(worksheet4, 8, 1, 66, null);

    _ = lxw.worksheet_write_string(worksheet4, 0, 3, "Name", header2);
    _ = lxw.worksheet_write_string(worksheet4, 0, 4, "Age", header2);

    _ = lxw.worksheet_set_column_pixels(worksheet4, 2, 2, 20, null);

    // Example of using the XLOOKUP() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet5,
        0,
        5,
        "=_xlfn.XLOOKUP(E1,A2:A9,C2:C9)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet5, 0, 0, "Country", header1);
    _ = lxw.worksheet_write_string(worksheet5, 0, 1, "Abr", header1);
    _ = lxw.worksheet_write_string(worksheet5, 0, 2, "Prefix", header1);

    _ = lxw.worksheet_write_string(worksheet5, 1, 0, "China", null);
    _ = lxw.worksheet_write_string(worksheet5, 2, 0, "India", null);
    _ = lxw.worksheet_write_string(worksheet5, 3, 0, "United States", null);
    _ = lxw.worksheet_write_string(worksheet5, 4, 0, "Indonesia", null);
    _ = lxw.worksheet_write_string(worksheet5, 5, 0, "Brazil", null);
    _ = lxw.worksheet_write_string(worksheet5, 6, 0, "Pakistan", null);
    _ = lxw.worksheet_write_string(worksheet5, 7, 0, "Nigeria", null);
    _ = lxw.worksheet_write_string(worksheet5, 8, 0, "Bangladesh", null);

    _ = lxw.worksheet_write_string(worksheet5, 1, 1, "CN", null);
    _ = lxw.worksheet_write_string(worksheet5, 2, 1, "IN", null);
    _ = lxw.worksheet_write_string(worksheet5, 3, 1, "US", null);
    _ = lxw.worksheet_write_string(worksheet5, 4, 1, "ID", null);
    _ = lxw.worksheet_write_string(worksheet5, 5, 1, "BR", null);
    _ = lxw.worksheet_write_string(worksheet5, 6, 1, "PK", null);
    _ = lxw.worksheet_write_string(worksheet5, 7, 1, "NG", null);
    _ = lxw.worksheet_write_string(worksheet5, 8, 1, "BD", null);

    _ = lxw.worksheet_write_number(worksheet5, 1, 2, 86, null);
    _ = lxw.worksheet_write_number(worksheet5, 2, 2, 91, null);
    _ = lxw.worksheet_write_number(worksheet5, 3, 2, 1, null);
    _ = lxw.worksheet_write_number(worksheet5, 4, 2, 62, null);
    _ = lxw.worksheet_write_number(worksheet5, 5, 2, 55, null);
    _ = lxw.worksheet_write_number(worksheet5, 6, 2, 92, null);
    _ = lxw.worksheet_write_number(worksheet5, 7, 2, 234, null);
    _ = lxw.worksheet_write_number(worksheet5, 8, 2, 880, null);

    _ = lxw.worksheet_write_string(worksheet5, 0, 4, "Brazil", header2);

    _ = lxw.worksheet_set_column_pixels(worksheet5, 0, 0, 100, null);
    _ = lxw.worksheet_set_column_pixels(worksheet5, 3, 3, 20, null);

    // Example of using the XMATCH() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet6,
        1,
        3,
        "=_xlfn.XMATCH(C2,A2:A6)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet6, 0, 0, "Product", header1);

    _ = lxw.worksheet_write_string(worksheet6, 1, 0, "Apple", null);
    _ = lxw.worksheet_write_string(worksheet6, 2, 0, "Grape", null);
    _ = lxw.worksheet_write_string(worksheet6, 3, 0, "Pear", null);
    _ = lxw.worksheet_write_string(worksheet6, 4, 0, "Banana", null);
    _ = lxw.worksheet_write_string(worksheet6, 5, 0, "Cherry", null);

    _ = lxw.worksheet_write_string(worksheet6, 0, 2, "Product", header2);
    _ = lxw.worksheet_write_string(worksheet6, 0, 3, "Position", header2);
    _ = lxw.worksheet_write_string(worksheet6, 1, 2, "Grape", null);

    _ = lxw.worksheet_set_column_pixels(worksheet6, 1, 1, 20, null);

    // Example of using the RANDARRAY() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet7,
        0,
        0,
        "=_xlfn.RANDARRAY(5,3,1,100, TRUE)",
        null,
    );

    // Example of using the SEQUENCE() function.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet8,
        0,
        0,
        "=_xlfn.SEQUENCE(4,5)",
        null,
    );

    // Example of using the Spill range operator.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        7,
        "=_xlfn.ANCHORARRAY(F2)",
        null,
    );

    _ = lxw.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        9,
        "=COUNTA(_xlfn.ANCHORARRAY(F2))",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_dynamic_formula(
        worksheet9,
        1,
        5,
        "=_xlfn.UNIQUE(B2:B17)",
        null,
    );

    _ = lxw.worksheet_write_string(worksheet9, 0, 5, "Unique", header2);
    _ = lxw.worksheet_write_string(worksheet9, 0, 7, "Spill", header2);
    _ = lxw.worksheet_write_string(worksheet9, 0, 9, "Spill", header2);

    writeWorksheetData(worksheet9, header1);
    _ = lxw.worksheet_set_column_pixels(worksheet9, 4, 4, 20, null);
    _ = lxw.worksheet_set_column_pixels(worksheet9, 6, 6, 20, null);
    _ = lxw.worksheet_set_column_pixels(worksheet9, 8, 8, 20, null);

    // Example of using dynamic ranges with older Excel functions.
    _ = lxw.worksheet_write_dynamic_array_formula(
        worksheet10,
        0,
        1,
        2,
        1,
        "=LEN(A1:A3)",
        null,
    );

    // Write the data the function will work on.
    _ = lxw.worksheet_write_string(worksheet10, 0, 0, "Foo", null);
    _ = lxw.worksheet_write_string(worksheet10, 1, 0, "Food", null);
    _ = lxw.worksheet_write_string(worksheet10, 2, 0, "Frood", null);

    _ = lxw.workbook_close(workbook);
}

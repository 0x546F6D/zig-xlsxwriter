const Row = struct {
    region: [:0]const u8,
    item: [:0]const u8,
    volume: i32,
    month: [:0]const u8,
};

var data = [_]Row{
    .{ .region = "East", .item = "Apple", .volume = 9000, .month = "July" },
    .{ .region = "East", .item = "Apple", .volume = 5000, .month = "July" },
    .{ .region = "South", .item = "Orange", .volume = 9000, .month = "September" },
    .{ .region = "North", .item = "Apple", .volume = 2000, .month = "November" },
    .{ .region = "West", .item = "Apple", .volume = 9000, .month = "November" },
    .{ .region = "South", .item = "Pear", .volume = 7000, .month = "October" },
    .{ .region = "North", .item = "Pear", .volume = 9000, .month = "August" },
    .{ .region = "West", .item = "Orange", .volume = 1000, .month = "December" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "November" },
    .{ .region = "South", .item = "Pear", .volume = 10000, .month = "April" },
    .{ .region = "West", .item = "Grape", .volume = 6000, .month = "January" },
    .{ .region = "South", .item = "Orange", .volume = 3000, .month = "May" },
    .{ .region = "North", .item = "Apple", .volume = 3000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 7000, .month = "February" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "December" },
    .{ .region = "East", .item = "Grape", .volume = 8000, .month = "February" },
    .{ .region = "South", .item = "Grape", .volume = 10000, .month = "June" },
    .{ .region = "West", .item = "Pear", .volume = 7000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 2000, .month = "October" },
    .{ .region = "East", .item = "Grape", .volume = 7000, .month = "December" },
    .{ .region = "North", .item = "Grape", .volume = 6000, .month = "April" },
    .{ .region = "East", .item = "Pear", .volume = 8000, .month = "February" },
    .{ .region = "North", .item = "Apple", .volume = 7000, .month = "August" },
    .{ .region = "North", .item = "Orange", .volume = 7000, .month = "July" },
    .{ .region = "North", .item = "Apple", .volume = 6000, .month = "June" },
    .{ .region = "South", .item = "Grape", .volume = 8000, .month = "September" },
    .{ .region = "West", .item = "Apple", .volume = 3000, .month = "October" },
    .{ .region = "South", .item = "Orange", .volume = 10000, .month = "November" },
    .{ .region = "West", .item = "Grape", .volume = 4000, .month = "July" },
    .{ .region = "North", .item = "Orange", .volume = 5000, .month = "August" },
    .{ .region = "East", .item = "Orange", .volume = 1000, .month = "November" },
    .{ .region = "East", .item = "Orange", .volume = 4000, .month = "October" },
    .{ .region = "North", .item = "Grape", .volume = 5000, .month = "August" },
    .{ .region = "East", .item = "Apple", .volume = 1000, .month = "December" },
    .{ .region = "South", .item = "Apple", .volume = 10000, .month = "March" },
    .{ .region = "East", .item = "Grape", .volume = 7000, .month = "October" },
    .{ .region = "West", .item = "Grape", .volume = 1000, .month = "September" },
    .{ .region = "East", .item = "Grape", .volume = 10000, .month = "October" },
    .{ .region = "South", .item = "Orange", .volume = 8000, .month = "March" },
    .{ .region = "North", .item = "Apple", .volume = 4000, .month = "July" },
    .{ .region = "South", .item = "Orange", .volume = 5000, .month = "July" },
    .{ .region = "West", .item = "Apple", .volume = 4000, .month = "June" },
    .{ .region = "East", .item = "Apple", .volume = 5000, .month = "April" },
    .{ .region = "North", .item = "Pear", .volume = 3000, .month = "August" },
    .{ .region = "East", .item = "Grape", .volume = 9000, .month = "November" },
    .{ .region = "North", .item = "Orange", .volume = 8000, .month = "October" },
    .{ .region = "East", .item = "Apple", .volume = 10000, .month = "June" },
    .{ .region = "South", .item = "Pear", .volume = 1000, .month = "December" },
    .{ .region = "North", .item = "Grape", .volume = 10000, .month = "July" },
    .{ .region = "East", .item = "Grape", .volume = 6000, .month = "February" },
};

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    // set header format
    const header = try workbook.addFormat();
    header.setBold();

    // Add worksheets
    const worksheet1 = try workbook.addWorkSheet(null);
    const worksheet2 = try workbook.addWorkSheet(null);
    const worksheet3 = try workbook.addWorkSheet(null);
    const worksheet4 = try workbook.addWorkSheet(null);
    const worksheet5 = try workbook.addWorkSheet(null);
    const worksheet6 = try workbook.addWorkSheet(null);
    const worksheet7 = try workbook.addWorkSheet(null);

    const hidden: xwz.RowColOptions = .{ .hidden = true };

    // Example 1. Autofilter without conditions.
    try prepWorkSheet(worksheet1, header);

    // Example 2. Autofilter with a filter condition in the first column.
    try prepWorkSheet(worksheet2, header);

    // Add the filter criteria
    const filter_rule2: xwz.FilterRule = .{
        .criteria = .equal_to,
        .value_string = "East",
    };
    try worksheet2.filterColumn(0, filter_rule2);

    // It isn't sufficient to just apply the filter condition below.
    // We must also hide the rows that don't match the criteria
    // since Excel doesn't do that automatically.
    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East")) {
            try worksheet2.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }

    // Example 3. Autofilter with a dual filter condition in one of the columns.
    try prepWorkSheet(worksheet3, header);

    // Add the filter criterias
    const filter_rule3a: xwz.FilterRule = .{
        .criteria = .equal_to,
        .value_string = "East",
    };
    const filter_rule3b: xwz.FilterRule = .{
        .criteria = .equal_to,
        .value_string = "South",
    };
    try worksheet3.filterColumn2(0, filter_rule3a, filter_rule3b, xwz.filter_or);

    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East") and
            !std.mem.eql(u8, row.region, "South"))
        {
            try worksheet3.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }

    // Example 4. Autofilter with filter conditions in two columns.
    try prepWorkSheet(worksheet4, header);

    // Add the filter criteria
    const filter_rule4a: xwz.FilterRule = .{
        .criteria = .equal_to,
        .value_string = "East",
    };
    const filter_rule4b: xwz.FilterRule = .{
        .criteria = .greater_than,
        .value = 3000,
    };
    const filter_rule4c: xwz.FilterRule = .{
        .criteria = .less_than,
        .value = 8000,
    };
    try worksheet4.filterColumn(0, filter_rule4a);
    try worksheet4.filterColumn2(2, filter_rule4b, filter_rule4c, xwz.filter_and);

    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East") or
            row.volume < 3000 or row.volume > 8000)
        {
            try worksheet4.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }

    // Example 5. Autofilter with a list filter condition.
    try prepWorkSheet(worksheet5, header);

    // Add the filter list
    const list_rule5: xwz.CStringArray = &.{ "East", "North", "South" };
    try worksheet5.filterList(0, list_rule5);

    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (!std.mem.eql(u8, row.region, "East") and
            !std.mem.eql(u8, row.region, "North") and
            !std.mem.eql(u8, row.region, "South"))
        {
            try worksheet5.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }

    // Example 6. Autofilter with filter for blanks.
    // - modify data for examples 6 & 7 with blank filter criteria
    data[5].region = "";
    try prepWorkSheet(worksheet6, header);

    // Add the filter criteria
    const filter_rule6: xwz.FilterRule = .{ .criteria = .blanks };

    try worksheet6.filterColumn(0, filter_rule6);

    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (!std.mem.eql(u8, row.region, "")) {
            try worksheet6.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }

    // Example 7. Autofilter with filter for non-blanks.
    try prepWorkSheet(worksheet7, header);

    // Add the filter criteria
    const filter_rule7: xwz.FilterRule = .{ .criteria = .non_blanks };
    try worksheet7.filterColumn(0, filter_rule7);

    // Hide rows that don't match the filter
    for (data, 1..) |row, i| {
        if (std.mem.eql(u8, row.region, "")) {
            try worksheet7.setRowOpt(@intCast(i), xwz.def_row_height, .default, hidden);
        }
    }
}

fn prepWorkSheet(worksheet: WorkSheet, header: Format) !void {
    // write header
    try writeWorksheetHeader(worksheet, header);
    // write the data table
    try writeData(worksheet);
    // Add the autofilter range
    try worksheet.autoFilter(0, 0, 50, 3);
}

fn writeWorksheetHeader(worksheet: WorkSheet, header: Format) !void {
    // Make the columns wider for clarity
    try worksheet.setColumn(0, 3, 12, .default);

    // Write the column headers
    try worksheet.setRow(0, 20, header);
    try worksheet.writeString(0, 0, "Region", .default);
    try worksheet.writeString(0, 1, "Item", .default);
    try worksheet.writeString(0, 2, "Volume", .default);
    try worksheet.writeString(0, 3, "Month", .default);
}

fn writeData(worksheet: WorkSheet) !void {
    for (data, 1..) |row, i| {
        try worksheet.writeString(@intCast(i), 0, row.region, .default);
        try worksheet.writeString(@intCast(i), 1, row.item, .default);
        try worksheet.writeNumber(@intCast(i), 2, @floatFromInt(row.volume), .default);
        try worksheet.writeString(@intCast(i), 3, row.month, .default);
    }
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;

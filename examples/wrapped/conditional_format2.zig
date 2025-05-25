// Write some data to the worksheet.
fn writeWorksheetData(worksheet: WorkSheet) !void {
    const data = [_][10]u8{
        .{ 34, 72, 38, 30, 75, 48, 75, 66, 84, 86 },
        .{ 6, 24, 1, 84, 54, 62, 60, 3, 26, 59 },
        .{ 28, 79, 97, 13, 85, 93, 93, 22, 5, 14 },
        .{ 27, 71, 40, 17, 18, 79, 90, 93, 29, 47 },
        .{ 88, 25, 33, 23, 67, 1, 59, 79, 47, 36 },
        .{ 24, 100, 20, 88, 29, 33, 38, 54, 54, 88 },
        .{ 6, 57, 88, 28, 10, 26, 37, 7, 41, 48 },
        .{ 52, 78, 1, 96, 26, 45, 47, 33, 96, 36 },
        .{ 60, 54, 81, 66, 81, 90, 80, 93, 12, 55 },
        .{ 70, 5, 46, 14, 71, 19, 66, 36, 41, 21 },
    };

    for (data, 0..) |array, row| {
        for (array, 0..) |val, col| {
            try worksheet.writeNumber(@intCast(row + 2), @intCast(col + 1), @floatFromInt(val), .none);
        }
    }
}

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

    // Add a format. Light red fill with dark red text.
    const format1 = try workbook.addFormat();
    format1.setBgColor(@enumFromInt(0xFFC7CE));
    format1.setFontColor(@enumFromInt(0x9C0006));

    // Add a format. Green fill with dark green text.
    const format2 = try workbook.addFormat();
    format2.setBgColor(@enumFromInt(0xC6EFCE));
    format2.setFontColor(@enumFromInt(0x006100));

    // Create a single conditional format object to reuse in the examples.
    var conditional_format: xlsxwriter.ConditionalFormat = .empty;

    // Example 1. Conditional formatting based on simple cell based criteria.
    try writeWorksheetData(worksheet1);

    try worksheet1.writeString(
        0,
        0,
        "Cells with values >= 50 are in light red. Values < 50 are in light green.",
        .none,
    );

    conditional_format.type = .cell;
    conditional_format.criteria = .greater_than_or_equal_to;
    conditional_format.value = 50;
    conditional_format.format_c = format1.format_c;
    try worksheet1.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.criteria = .less_than;
    conditional_format.format_c = format2.format_c;
    try worksheet1.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    // Example 2. Conditional formatting based on max and min values.
    try writeWorksheetData(worksheet2);

    try worksheet2.writeString(
        0,
        0,
        "Values between 30 and 70 are in light red. Values outside that range are in light green.",
        .none,
    );

    conditional_format.criteria = .between;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format_c = format1.format_c;
    try worksheet2.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.criteria = .not_between;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format_c = format2.format_c;
    try worksheet2.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    // Example 3. Conditional formatting with duplicate and unique values.
    try writeWorksheetData(worksheet3);

    try worksheet3.writeString(
        0,
        0,
        "Duplicate values are in light red. Unique values are in light green.",
        .none,
    );

    conditional_format.type = .duplicate;
    conditional_format.format_c = format1.format_c;
    try worksheet3.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.type = .unique;
    conditional_format.format_c = format2.format_c;
    try worksheet3.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    // Example 4. Conditional formatting with above and below average values.
    try writeWorksheetData(worksheet4);

    try worksheet4.writeString(
        0,
        0,
        "Above average values are in light red. Below average values are in light green.",
        .none,
    );

    conditional_format.type = .average;
    conditional_format.criteria = .average_above;
    conditional_format.format_c = format1.format_c;
    try worksheet4.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.criteria = .average_below;
    conditional_format.format_c = format2.format_c;
    try worksheet4.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    // Example 5. Conditional formatting with top and bottom values.
    try writeWorksheetData(worksheet5);
    conditional_format = .empty;

    try worksheet5.writeString(
        0,
        0,
        "Top 10 values are in light red. Bottom 10 values are in light green.",
        .none,
    );

    conditional_format.type = .top;
    conditional_format.value = 10;
    conditional_format.format_c = format1.format_c;
    try worksheet5.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.type = .bottom;
    conditional_format.value = 10;
    conditional_format.format_c = format2.format_c;
    try worksheet5.conditionalFormatRange(2, 5, 11, 10, conditional_format);

    // Example 6. Conditional formatting with multiple ranges.
    try writeWorksheetData(worksheet6);

    try worksheet6.writeString(
        0,
        0,
        "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.",
        .none,
    );

    conditional_format.type = .cell;
    conditional_format.criteria = .greater_than_or_equal_to;
    conditional_format.value = 50;
    conditional_format.format_c = format1.format_c;
    conditional_format.multi_range = "B3:K6 B9:K12";
    try worksheet6.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    conditional_format.criteria = .less_than;
    conditional_format.value = 50;
    conditional_format.format_c = format2.format_c;
    conditional_format.multi_range = "B3:K6 B9:K12";
    try worksheet6.conditionalFormatRange(2, 1, 11, 10, conditional_format);

    // Reset the options before the next example.
    conditional_format = .empty;

    // Example 7. Conditional formatting with 2 color scales.
    // Write the worksheet data.
    for (1..13) |i| {
        try worksheet7.writeNumber(@intCast(i + 1), 1, @floatFromInt(i), .none);
        try worksheet7.writeNumber(@intCast(i + 1), 3, @floatFromInt(i), .none);
        try worksheet7.writeNumber(@intCast(i + 1), 6, @floatFromInt(i), .none);
        try worksheet7.writeNumber(@intCast(i + 1), 8, @floatFromInt(i), .none);
    }

    try worksheet7.writeString(
        0,
        0,
        "Examples of color scales with default and user colors.",
        .none,
    );

    try worksheet7.writeString(1, 1, "2 Color Scale", .none);
    try worksheet7.writeString(1, 3, "2 Color Scale + user colors", .none);
    try worksheet7.writeString(1, 6, "3 Color Scale", .none);
    try worksheet7.writeString(1, 8, "3 Color Scale + user colors", .none);

    // 2 color scale with standard colors.
    conditional_format.type = .color_2_scale;
    try worksheet7.conditionalFormatRange(2, 1, 13, 1, conditional_format);

    // 2 color scale with user defined colors.
    conditional_format.type = .color_2_scale;
    conditional_format.min_color = @enumFromInt(0xFF0000);
    conditional_format.max_color = @enumFromInt(0x00FF00);
    try worksheet7.conditionalFormatRange(2, 3, 13, 3, conditional_format);

    // Reset the colors before the next example.
    conditional_format = .empty;

    // 3 color scale with standard colors.
    conditional_format.type = .color_3_scale;
    try worksheet7.conditionalFormatRange(2, 6, 13, 6, conditional_format);

    // 3 color scale with user defined colors.
    conditional_format.type = .color_3_scale;
    conditional_format.min_color = @enumFromInt(0xC5D9F1);
    conditional_format.mid_color = @enumFromInt(0x8DB4E3);
    conditional_format.max_color = @enumFromInt(0x538ED5);
    try worksheet7.conditionalFormatRange(2, 8, 13, 8, conditional_format);

    // Example 8. Conditional formatting with data bars.
    // First data bar example.
    try worksheet8.writeString(
        0,
        0,
        "Examples of data bars.",
        .none,
    );

    // Write the worksheet data.
    for (1..13) |i| {
        try worksheet8.writeNumber(@intCast(i + 1), 1, @floatFromInt(i), .none);
        try worksheet8.writeNumber(@intCast(i + 1), 3, @floatFromInt(i), .none);
        try worksheet8.writeNumber(@intCast(i + 1), 5, @floatFromInt(i), .none);
        try worksheet8.writeNumber(@intCast(i + 1), 7, @floatFromInt(i), .none);
        try worksheet8.writeNumber(@intCast(i + 1), 9, @floatFromInt(i), .none);
    }
    const data = [_]f64{ -1, -2, -3, -2, -1, 0, 1, 2, 3, 2, 1, 0 };
    for (data, 1..) |val, i| {
        try worksheet8.writeNumber(@intCast(i + 1), 11, val, .none);
        try worksheet8.writeNumber(@intCast(i + 1), 13, val, .none);
    }

    try worksheet8.writeString(1, 1, "Default data bars", .none);
    try worksheet8.writeString(1, 3, "Bars only", .none);
    try worksheet8.writeString(1, 5, "With user color", .none);
    try worksheet8.writeString(1, 7, "Solid bars", .none);
    try worksheet8.writeString(1, 9, "Right to left", .none);
    try worksheet8.writeString(1, 11, "Excel 2010 style", .none);
    try worksheet8.writeString(1, 13, "Negative same as positive", .none);

    // Default data bars.
    conditional_format.type = .data_bar;
    try worksheet8.conditionalFormatRange(2, 1, 13, 1, conditional_format);

    // Data bars with border.
    conditional_format.bar_only = xlsxwriter.xw_true;
    // conditional_format.bar_border_color = .black;
    try worksheet8.conditionalFormatRange(2, 3, 13, 3, conditional_format);

    // User defined color.
    conditional_format = .empty;
    conditional_format.type = .data_bar;
    conditional_format.bar_color = @enumFromInt(0x63C384);
    try worksheet8.conditionalFormatRange(2, 5, 13, 5, conditional_format);

    // Same color for negative values.
    conditional_format = .empty;
    conditional_format.type = .data_bar;
    conditional_format.bar_solid = xlsxwriter.xw_true;
    // conditional_format.bar_negative_color_same = 1;
    try worksheet8.conditionalFormatRange(2, 7, 13, 7, conditional_format);

    // Right to left.
    conditional_format = .empty;
    conditional_format.type = .data_bar;
    conditional_format.bar_direction = .right_to_left;
    try worksheet8.conditionalFormatRange(2, 9, 13, 9, conditional_format);

    // Excel 2010 style.
    conditional_format = .empty;
    conditional_format.type = .data_bar;
    conditional_format.data_bar_2010 = 1;
    try worksheet8.conditionalFormatRange(2, 11, 13, 11, conditional_format);

    // Zero axis.
    conditional_format = .empty;
    conditional_format.type = .data_bar;
    conditional_format.bar_negative_color_same = xlsxwriter.xw_true;
    conditional_format.bar_negative_border_color_same = xlsxwriter.xw_true;
    try worksheet8.conditionalFormatRange(2, 13, 13, 13, conditional_format);

    // Example 9. Conditional formatting with icon sets.
    try worksheet9.writeString(
        0,
        0,
        "Examples of conditional formats with icon sets.",
        .none,
    );

    // Write the worksheet data.
    for (1..4) |i| {
        try worksheet9.writeNumber(2, @intCast(i), @floatFromInt(i), .none);
        try worksheet9.writeNumber(3, @intCast(i), @floatFromInt(i), .none);
        try worksheet9.writeNumber(4, @intCast(i), @floatFromInt(i), .none);
        try worksheet9.writeNumber(5, @intCast(i), @floatFromInt(i), .none);
    }

    for (1..5) |i| {
        try worksheet9.writeNumber(6, @intCast(i), @floatFromInt(i), .none);
        try worksheet9.writeNumber(7, @intCast(i), @floatFromInt(i), .none);
        try worksheet9.writeNumber(8, @intCast(i), @floatFromInt(i), .none);
    }

    try worksheet9.writeNumber(7, 5, 5, .none);
    try worksheet9.writeNumber(8, 5, 5, .none);

    // Reset the conditional format.
    conditional_format = .empty;

    // Three traffic lights (default style).
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    try worksheet9.conditionalFormatRange(2, 1, 2, 3, conditional_format);

    // Three traffic lights (unrimmed style).
    conditional_format.reverse_icons = xlsxwriter.xw_true;
    try worksheet9.conditionalFormatRange(3, 1, 3, 3, conditional_format);

    // Three traffic lights (unrimmed style).
    conditional_format = .empty;
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    conditional_format.icons_only = xlsxwriter.xw_true;
    // conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    try worksheet9.conditionalFormatRange(4, 1, 4, 3, conditional_format);

    // Three arrows.
    conditional_format = .empty;
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_arrows_colored;
    try worksheet9.conditionalFormatRange(5, 1, 5, 3, conditional_format);

    // Four arrows.
    conditional_format.icon_style = .icons_4_arrows_colored;
    try worksheet9.conditionalFormatRange(6, 1, 6, 4, conditional_format);

    // Five arrows.
    conditional_format.icon_style = .icons_5_arrows_colored;
    try worksheet9.conditionalFormatRange(7, 1, 7, 5, conditional_format);

    // Five ratings.
    conditional_format.icon_style = .icons_5_ratings;
    try worksheet9.conditionalFormatRange(8, 1, 8, 5, conditional_format);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;

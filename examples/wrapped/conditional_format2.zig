pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
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
    var worksheet: WorkSheet = undefined;

    // Add a format. Light red fill with dark red text.
    const format1 = try workbook.addFormat();
    format1.setBgColor(@enumFromInt(0xFFC7CE));
    format1.setFontColor(@enumFromInt(0x9C0006));

    // Add a format. Green fill with dark green text.
    const format2 = try workbook.addFormat();
    format2.setBgColor(@enumFromInt(0xC6EFCE));
    format2.setFontColor(@enumFromInt(0x006100));

    // Create a single conditional format object to reuse in the examples.
    var conditional_format: xwz.ConditionalFormat = .default;

    // Example 1. Conditional formatting based on simple cell based criteria.
    worksheet = worksheet1;
    try writeWorksheetData(worksheet);

    try worksheet.writeString(
        .{},
        "Cells with values >= 50 are in light red. Values < 50 are in light green.",
        .default,
    );

    conditional_format.type = .cell;
    conditional_format.criteria = .greater_than_or_equal_to;
    conditional_format.value = 50;
    conditional_format.format = format1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.criteria = .less_than;
    conditional_format.format = format2;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Example 2. Conditional formatting based on max and min values.
    worksheet = worksheet2;
    try writeWorksheetData(worksheet);

    try worksheet.writeString(
        .{},
        "Values between 30 and 70 are in light red. Values outside that range are in light green.",
        .default,
    );

    conditional_format.criteria = .between;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format = format1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.criteria = .not_between;
    conditional_format.min_value = 30;
    conditional_format.max_value = 70;
    conditional_format.format = format2;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Example 3. Conditional formatting with duplicate and unique values.
    worksheet = worksheet3;
    try writeWorksheetData(worksheet);

    try worksheet.writeString(
        .{},
        "Duplicate values are in light red. Unique values are in light green.",
        .default,
    );

    conditional_format.type = .duplicate;
    conditional_format.format = format1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.type = .unique;
    conditional_format.format = format2;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Example 4. Conditional formatting with above and below average values.
    worksheet = worksheet4;
    try writeWorksheetData(worksheet);

    try worksheet.writeString(
        .{},
        "Above average values are in light red. Below average values are in light green.",
        .default,
    );

    conditional_format.type = .average;
    conditional_format.criteria = .average_above;
    conditional_format.format = format1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.criteria = .average_below;
    conditional_format.format = format2;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Example 5. Conditional formatting with top and bottom values.
    worksheet = worksheet5;
    try writeWorksheetData(worksheet);
    conditional_format = .default;

    try worksheet.writeString(
        .{},
        "Top 10 values are in light red. Bottom 10 values are in light green.",
        .default,
    );

    conditional_format.type = .top;
    conditional_format.value = 10;
    conditional_format.format = format1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.type = .bottom;
    conditional_format.value = 10;
    conditional_format.format = format2;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 5, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Example 6. Conditional formatting with multiple ranges.
    worksheet = worksheet6;
    try writeWorksheetData(worksheet);

    try worksheet.writeString(
        .{},
        "Cells with values >= 50 are in light red. Values < 50 are in light green. Non-contiguous ranges.",
        .default,
    );

    conditional_format.type = .cell;
    conditional_format.criteria = .greater_than_or_equal_to;
    conditional_format.value = 50;
    conditional_format.format = format1;
    conditional_format.multi_range = "B3:K6 B9:K12";
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    conditional_format.criteria = .less_than;
    conditional_format.value = 50;
    conditional_format.format = format2;
    conditional_format.multi_range = "B3:K6 B9:K12";
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 11, .last_col = 10 },
        conditional_format,
    );

    // Reset the options before the next example.
    conditional_format = .default;

    // Example 7. Conditional formatting with 2 color scales.
    worksheet = worksheet7;
    // Write the worksheet data.
    for (1..13) |i| {
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 1 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 3 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 6 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 8 }, @floatFromInt(i), .default);
    }

    try worksheet.writeString(
        .{},
        "Examples of color scales with default and user colors.",
        .default,
    );

    try worksheet.writeString(.{ .row = 1, .col = 1 }, "2 Color Scale", .default);
    try worksheet.writeString(.{ .row = 1, .col = 3 }, "2 Color Scale + user colors", .default);
    try worksheet.writeString(.{ .row = 1, .col = 6 }, "3 Color Scale", .default);
    try worksheet.writeString(.{ .row = 1, .col = 8 }, "3 Color Scale + user colors", .default);

    // 2 color scale with standard colors.
    conditional_format.type = .color_2_scale;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 13, .last_col = 1 },
        conditional_format,
    );

    // 2 color scale with user defined colors.
    conditional_format.type = .color_2_scale;
    conditional_format.min_color = @enumFromInt(0xFF0000);
    conditional_format.max_color = @enumFromInt(0x00FF00);
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 3, .last_row = 13, .last_col = 3 },
        conditional_format,
    );

    // Reset the colors before the next example.
    conditional_format = .default;

    // 3 color scale with standard colors.
    conditional_format.type = .color_3_scale;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 6, .last_row = 13, .last_col = 6 },
        conditional_format,
    );

    // 3 color scale with user defined colors.
    conditional_format.type = .color_3_scale;
    conditional_format.min_color = @enumFromInt(0xC5D9F1);
    conditional_format.mid_color = @enumFromInt(0x8DB4E3);
    conditional_format.max_color = @enumFromInt(0x538ED5);
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 8, .last_row = 13, .last_col = 8 },
        conditional_format,
    );

    // Example 8. Conditional formatting with data bars.
    worksheet = worksheet8;
    // First data bar example.
    try worksheet.writeString(
        .{},
        "Examples of data bars.",
        .default,
    );

    // Write the worksheet data.
    for (1..13) |i| {
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 1 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 3 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 5 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 7 }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 9 }, @floatFromInt(i), .default);
    }
    const data = [_]f64{ -1, -2, -3, -2, -1, 0, 1, 2, 3, 2, 1, 0 };
    for (data, 1..) |val, i| {
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 11 }, val, .default);
        try worksheet.writeNumber(.{ .row = @intCast(i + 1), .col = 13 }, val, .default);
    }

    try worksheet.writeString(.{ .row = 1, .col = 1 }, "Default data bars", .default);
    try worksheet.writeString(.{ .row = 1, .col = 3 }, "Bars only", .default);
    try worksheet.writeString(.{ .row = 1, .col = 5 }, "With user color", .default);
    try worksheet.writeString(.{ .row = 1, .col = 7 }, "Solid bars", .default);
    try worksheet.writeString(.{ .row = 1, .col = 9 }, "Right to left", .default);
    try worksheet.writeString(.{ .row = 1, .col = 11 }, "Excel 2010 style", .default);
    try worksheet.writeString(.{ .row = 1, .col = 13 }, "Negative same as positive", .default);

    // Default data bars.
    conditional_format.type = .data_bar;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 13, .last_col = 1 },
        conditional_format,
    );

    // Data bars with border.
    conditional_format.bar_only = true;
    // conditional_format.bar_border_color = .black;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 3, .last_row = 13, .last_col = 3 },
        conditional_format,
    );

    // User defined color.
    conditional_format = .default;
    conditional_format.type = .data_bar;
    conditional_format.bar_color = @enumFromInt(0x63C384);
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 5, .last_row = 13, .last_col = 5 },
        conditional_format,
    );

    // Same color for negative values.
    conditional_format = .default;
    conditional_format.type = .data_bar;
    conditional_format.bar_solid = true;
    // conditional_format.bar_negative_color_same = 1;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 7, .last_row = 13, .last_col = 7 },
        conditional_format,
    );

    // Right to left.
    conditional_format = .default;
    conditional_format.type = .data_bar;
    conditional_format.bar_direction = .right_to_left;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 9, .last_row = 13, .last_col = 9 },
        conditional_format,
    );

    // Excel 2010 style.
    conditional_format = .default;
    conditional_format.type = .data_bar;
    conditional_format.data_bar_2010 = true;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 11, .last_row = 13, .last_col = 11 },
        conditional_format,
    );

    // Zero axis.
    conditional_format = .default;
    conditional_format.type = .data_bar;
    conditional_format.bar_negative_color_same = true;
    conditional_format.bar_negative_border_color_same = true;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 13, .last_row = 13, .last_col = 13 },
        conditional_format,
    );

    // Example 9. Conditional formatting with icon sets.
    worksheet = worksheet9;
    try worksheet.writeString(
        .{},
        "Examples of conditional formats with icon sets.",
        .default,
    );

    // Write the worksheet data.
    for (1..4) |i| {
        try worksheet.writeNumber(.{ .row = 2, .col = @intCast(i) }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = 3, .col = @intCast(i) }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = 4, .col = @intCast(i) }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = 5, .col = @intCast(i) }, @floatFromInt(i), .default);
    }

    for (1..5) |i| {
        try worksheet.writeNumber(.{ .row = 6, .col = @intCast(i) }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = 7, .col = @intCast(i) }, @floatFromInt(i), .default);
        try worksheet.writeNumber(.{ .row = 8, .col = @intCast(i) }, @floatFromInt(i), .default);
    }

    try worksheet.writeNumber(.{ .row = 7, .col = 5 }, 5, .default);
    try worksheet.writeNumber(.{ .row = 8, .col = 5 }, 5, .default);

    // Reset the conditional format.
    conditional_format = .default;

    // Three traffic lights (default style).
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 2, .first_col = 1, .last_row = 2, .last_col = 3 },
        conditional_format,
    );

    // Three traffic lights (unrimmed style).
    conditional_format.reverse_icons = true;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 3, .first_col = 1, .last_row = 3, .last_col = 3 },
        conditional_format,
    );

    // Three traffic lights (unrimmed style).
    conditional_format = .default;
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    conditional_format.icons_only = true;
    // conditional_format.icon_style = .icons_3_traffic_lights_unrimmed;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 4, .first_col = 1, .last_row = 4, .last_col = 3 },
        conditional_format,
    );

    // Three arrows.
    conditional_format = .default;
    conditional_format.type = .icon_sets;
    conditional_format.icon_style = .icons_3_arrows_colored;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 5, .first_col = 1, .last_row = 5, .last_col = 3 },
        conditional_format,
    );

    // Four arrows.
    conditional_format.icon_style = .icons_4_arrows_colored;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 6, .first_col = 1, .last_row = 6, .last_col = 4 },
        conditional_format,
    );

    // Five arrows.
    conditional_format.icon_style = .icons_5_arrows_colored;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 7, .first_col = 1, .last_row = 7, .last_col = 5 },
        conditional_format,
    );

    // Five ratings.
    conditional_format.icon_style = .icons_5_ratings;
    try worksheet.conditionalFormatRange(
        .{ .first_row = 8, .first_col = 1, .last_row = 8, .last_col = 5 },
        conditional_format,
    );
}

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
            try worksheet.writeNumber(.{ .row = @intCast(row + 2), .col = @intCast(col + 1) }, @floatFromInt(val), .default);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;

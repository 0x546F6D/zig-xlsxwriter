pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet1 = try workbook.addWorkSheet("Panes 1");
    const worksheet2 = try workbook.addWorkSheet("Panes 2");
    const worksheet3 = try workbook.addWorkSheet("Panes 3");
    const worksheet4 = try workbook.addWorkSheet("Panes 4");
    var worksheet: WorkSheet = undefined;

    var row: u32 = 0;
    var col: u16 = 0;

    // Set up some formatting and text to highlight the panes.
    const header = try workbook.addFormat();
    header.setAlign(.center);
    header.setAlign(.vertical_center);
    header.setFgColor(@enumFromInt(0xD7E4BC));
    header.setBold();
    header.setBorder(.thin);

    const center = try workbook.addFormat();
    center.setAlign(.center);

    // Example 1. Freeze pane on the top row.
    worksheet = worksheet1;
    worksheet.freezePanes(1, 0);

    // Some sheet formatting.
    try worksheet.setColumn(0, 8, 16, .default);
    try worksheet.setRow(0, 20, .default);
    try worksheet.setSelection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 0;
    while (col < 9) : (col += 1) {
        try worksheet.writeString(0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 100) : (row += 1) {
        col = 0;
        while (col < 9) : (col += 1) {
            try worksheet.writeNumber(row, col, @floatFromInt(row + 1), center);
        }
    }

    // Example 2. Freeze pane on the left column.
    worksheet = worksheet2;
    worksheet.freezePanes(0, 1);

    // Some sheet formatting.
    try worksheet.setColumn(0, 0, 16, .default);
    try worksheet.setSelection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    row = 0;
    while (row < 50) : (row += 1) {
        try worksheet.writeString(row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            try worksheet.writeNumber(row, col, @floatFromInt(col), center);
        }
    }

    // Example 3. Freeze pane on the top row and left column.
    worksheet = worksheet3;
    worksheet.freezePanes(1, 1);

    // Some sheet formatting.
    try worksheet.setColumn(0, 25, 16, .default);
    try worksheet.setRow(0, 20, .default);
    try worksheet.writeString(0, 0, "", header);
    try worksheet.setSelection(4, 3, 4, 3);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        try worksheet.writeString(0, col, "Scroll down", header);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        try worksheet.writeString(row, 0, "Scroll right", header);

        col = 1;
        while (col < 26) : (col += 1) {
            try worksheet.writeNumber(row, col, @floatFromInt(col), center);
        }
    }

    // Example 4. Split pane on the top row and left column.
    worksheet = worksheet4;
    // The divisions must be specified in terms of row and column dimensions.
    // The default row height is 15 and the default column width is 8.43
    //
    worksheet.splitPanes(15, 8.43);

    // Some worksheet text to demonstrate scrolling.
    col = 1;
    while (col < 26) : (col += 1) {
        try worksheet.writeString(0, col, "Scroll", center);
    }

    row = 1;
    while (row < 50) : (row += 1) {
        try worksheet.writeString(row, 0, "Scroll", center);

        col = 1;
        while (col < 26) : (col += 1) {
            try worksheet.writeNumber(row, col, @floatFromInt(col), center);
        }
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;

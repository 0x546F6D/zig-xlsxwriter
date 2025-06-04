pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a format.
    const format = try workbook.addFormat();

    // Set the bold property for the format
    format.setBold();

    // Change the column width for clarity.
    try worksheet.setColumn(.{}, 20, .default, null);

    // Write some simple text.
    try worksheet.writeString(.{}, "Hello", .default);

    // Text with formatting.
    try worksheet.writeString(.{ .row = 1 }, "World", format);

    // Write some numbers.
    try worksheet.writeNumber(.{ .row = 2 }, 123, .default);
    try worksheet.writeNumber(.{ .row = 3 }, 123.456, .default);

    // Insert an image.
    const use_buffer = false;
    if (use_buffer) {
        // using embed + insertImageBuffer
        const logo_data = @import("assets").logo_data;
        try worksheet.insertImageBuffer(.{ .row = 1, .col = 2 }, logo_data, null);
    } else {
        // using image path + insertImage
        const asset_path = try h.getAssetPath(alloc, "logo.png");
        defer alloc.free(asset_path);
        try worksheet.insertImage(.{ .row = 1, .col = 2 }, asset_path, null);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

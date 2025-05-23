pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a format.
    const format = try workbook.addFormat();

    // Set the bold property for the format
    format.setBold();

    // Change the column width for clarity.
    try worksheet.setColumn(0, 0, 20, .none);

    // Write some simple text.
    try worksheet.writeString(0, 0, "Hello", .none);

    // Text with formatting.
    try worksheet.writeString(1, 0, "World", format);

    // Write some numbers.
    try worksheet.writeNumber(2, 0, 123, .none);
    try worksheet.writeNumber(3, 0, 123.456, .none);

    // Insert an image.
    const use_buffer = true;
    if (use_buffer) {
        // using embed + insertImageBuffer
        const logo_data = @import("assets").logo_data;
        try worksheet.insertImageBuffer(1, 2, logo_data, logo_data.len);
    } else {
        // using image path + insertImage
        const asset_path = try h.getAssetPath(alloc, "logo.png");
        defer alloc.free(asset_path);
        try worksheet.insertImage(1, 2, asset_path);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    const asset_path = try h.getAssetPath(alloc, "logo.png");
    defer alloc.free(asset_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    const datetime = xwz.DateTime{
        .year = 2016,
        .month = 12,
        .day = 12,
        .hour = 0,
        .min = 0,
        .sec = 0.0,
    };

    // Set some custom document properties in the workbook.
    try workbook.setCustomPropertyString("Checked by", "Eve");
    try workbook.setCustomPropertyDateTime("Date completed", datetime);
    try workbook.setCustomPropertyNumber("Document number", 12345);
    try workbook.setCustomPropertyNumber("Reference number", 1.2345);
    try workbook.setCustomPropertyBoolean("Has Review", true);
    try workbook.setCustomPropertyBoolean("Signed off", false);

    // Add some text to the file.
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, "Select 'Workbook Properties' to see properties.", .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

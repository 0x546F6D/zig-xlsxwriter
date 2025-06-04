pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Create a properties structure and set some of the fields.
    const properties = xwz.DocProperties{
        .title = "This is an example spreadsheet",
        .subject = "With document properties",
        .author = "John McNamara",
        .manager = "Dr. Heinz Doofenshmirtz",
        .company = "of Wolves",
        .category = "Example spreadsheets",
        .keywords = "Sample, Example, Properties",
        .comments = "Created with libxlswriter",
        .status = "Quo",
    };

    // Set the properties in the workbook.
    try workbook.setProperties(properties);

    // Add some text to the file.
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(
        .{},
        "Select 'Workbook Properties' to see properties.",
        .default,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

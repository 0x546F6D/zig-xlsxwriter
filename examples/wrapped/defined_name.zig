pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    // pass an 'Allocator' to initWorkBook() because we use the getWorkSheets() function
    // use 'var workbook' instead of 'const' because we use the getWorkSheets() function
    var workbook = try xwz.initWorkBook(alloc, xlsx_path, null);
    defer workbook.deinit() catch {};
    // We don't use the returned worksheets in this example and use a generic
    // loop below instead.
    _ = try workbook.addWorkSheet(null);
    _ = try workbook.addWorkSheet(null);

    // Define some global/workbook names
    try workbook.defineName("Exchange_rate", "=0.96");
    try workbook.defineName("Sales", "=Sheet1!$G$1:$H$10");

    // Define a local/worksheet name. This overrides the global "Sales" name
    // with a local defined name.
    try workbook.defineName("Sheet2!Sales", "=Sheet2!$G$1:$G$10");

    // Write some text to the worksheets and one of the defined names in a formula
    const worksheets_o = workbook.getWorkSheets() catch |err| {
        std.debug.print("GetWorkSheets error: {s}\n", .{xwz.strError(err)});
        return err;
    };
    const worksheets = if (worksheets_o) |worksheets| worksheets else return;

    for (worksheets) |worksheet| {
        try worksheet.setColumn(.{}, 45, .default, null);
        try worksheet.writeString(.{}, "This worksheet contains some defined names.", .default);
        try worksheet.writeString(.{ .row = 1 }, "See Formulas -> Name Manager above.", .default);
        try worksheet.writeString(.{ .row = 2 }, "Example formula in cell B3 ->", .default);
        try worksheet.writeFormula(.{ .row = 2, .col = 1 }, "=Exchange_rate", .default);
    }
}

const std = @import("std");
var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

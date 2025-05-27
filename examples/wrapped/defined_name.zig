pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
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
    const worksheets_o = try workbook.getWorkSheets(alloc);
    const worksheets = if (worksheets_o) |worksheets| worksheets else return;
    defer alloc.free(worksheets);

    for (worksheets) |worksheet| {
        try worksheet.setColumn(0, 0, 45, .none);
        try worksheet.writeString(0, 0, "This worksheet contains some defined names.", .none);
        try worksheet.writeString(1, 0, "See Formulas -> Name Manager above.", .none);
        try worksheet.writeString(2, 0, "Example formula in cell B3 ->", .none);
        try worksheet.writeFormula(2, 1, "=Exchange_rate", .none);
    }
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");

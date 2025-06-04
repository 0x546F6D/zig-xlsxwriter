pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Note that the formula name is prefixed with "_xlfn." and that the
    // lambda function parameters are prefixed with "_xlpm.". These prefixes
    // won't show up in Excel.
    try worksheet.writeDynamicFormula(
        xwz.cell("A1"),
        "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))(32)",
        .default,
    );

    // Create the lambda function as a defined name and write it as a dynamic formula
    try workbook.defineName(
        "ToCelsius",
        "=_xlfn.LAMBDA(_xlpm.temp, (5/9) * (_xlpm.temp-32))",
    );

    try worksheet.writeDynamicFormula(
        xwz.cell("A2"),
        "=ToCelsius(212)",
        .default,
    );
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

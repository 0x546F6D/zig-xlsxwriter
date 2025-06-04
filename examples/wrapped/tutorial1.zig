const Expense = struct {
    item: [:0]const u8,
    cost: f64,
};

var expenses = [_]Expense{
    .{ .item = "Rent", .cost = 1000.0 },
    .{ .item = "Gas", .cost = 100.0 },
    .{ .item = "Food", .cost = 300.0 },
    .{ .item = "Gym", .cost = 50.0 },
};

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Iterate over the data and write it out element by element.
    for (expenses) |expense| {
        try worksheet.writeString(.{ .row = row, .col = col }, expense.item, .default);
        try worksheet.writeNumber(.{ .row = row, .col = col + 1 }, expense.cost, .default);
        row += 1;
    }

    try worksheet.writeString(.{ .row = row, .col = col }, "Total", .default);
    try worksheet.writeFormula(.{ .row = row, .col = col + 1 }, "=SUM(B1:B4)", .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

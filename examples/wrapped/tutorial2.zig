const Expense = struct {
    item: [*:0]const u8,
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
    const workbook = try xwz.initWorkBook(null, xlsx_path.ptr);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format to use to highlight cells.
    const bold = try workbook.addFormat();
    bold.setBold();

    // Add a number format for cells with money.
    const money = try workbook.addFormat();
    money.setNumFormat("$#,##0");

    // Write some data header.
    try worksheet.writeString(0, 0, "Item", bold);
    try worksheet.writeString(0, 1, "Cost", bold);

    // Iterate over the data and write it out element by element.
    for (expenses, 1..) |expense, row| {
        try worksheet.writeString(@intCast(row), 0, expense.item, .default);
        try worksheet.writeNumber(@intCast(row), 1, expense.cost, money);
    }

    // Write a total using a formula.
    try worksheet.writeString(expenses.len + 1, 0, "Total", bold);
    try worksheet.writeFormula(expenses.len + 1, 1, "=SUM(B2:B5)", money);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

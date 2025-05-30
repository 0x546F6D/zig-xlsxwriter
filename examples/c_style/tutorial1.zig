const lxw = @import("lxw");

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

pub fn main() void {

    // Create a workbook and add a worksheet.
    const workbook: *lxw.lxw_workbook = lxw.workbook_new("zig-tutorial1.xlsx");
    const worksheet: *lxw.lxw_worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Iterate over the data and write it out element by element.
    while (row < 4) {
        _ = lxw.worksheet_write_string(worksheet, row, col, expenses[row].item, null);
        _ = lxw.worksheet_write_number(worksheet, row, col + 1, expenses[row].cost, null);
        defer row += 1;
    }

    // Write a total using a formula.
    _ = lxw.worksheet_write_string(worksheet, row, col, "Total", 0);
    _ = lxw.worksheet_write_formula(worksheet, row, col + 1, "=SUM(B1:B4)", null);

    // Save the workbook and free any allocated memory.
    _ = lxw.workbook_close(workbook);
}

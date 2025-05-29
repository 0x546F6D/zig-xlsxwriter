const lxw = @import("lxw");

const Expense = struct {
    item: [*:0]const u8,
    cost: i32,
};

var expenses = [_]Expense{
    .{ .item = "Rent", .cost = 1000 },
    .{ .item = "Gas", .cost = 100 },
    .{ .item = "Food", .cost = 300 },
    .{ .item = "Gym", .cost = 50 },
};

pub fn main() void {

    // Create a workbook and add a worksheet.
    const workbook: ?*lxw.lxw_workbook = lxw.workbook_new("zig-tutorial2.xlsx");
    const worksheet: ?*lxw.lxw_worksheet = lxw.workbook_add_worksheet(workbook, null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Add a bold format to use to highlight cells.
    const bold: ?*lxw.lxw_format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_bold(bold);

    // Add a number format for cells with money.
    const money: ?*lxw.lxw_format = lxw.workbook_add_format(workbook);
    _ = lxw.format_set_num_format(money, "$#,##0");

    // Write some data header.
    _ = lxw.worksheet_write_string(worksheet, row, col, "Item", bold);
    _ = lxw.worksheet_write_string(worksheet, row, col + 1, "Cost", bold);

    // Iterate over the data and write it out element by element.
    for (expenses, 0..) |expense, i| {
        // Write from the first cell below the headers.
        row = @intCast(i + 1);
        _ = lxw.worksheet_write_string(worksheet, row, col, expense.item, null);
        _ = lxw.worksheet_write_number(worksheet, row, col + 1, @floatFromInt(expense.cost), money);
    }

    // Write a total using a formula.
    _ = lxw.worksheet_write_string(worksheet, row + 1, col, "Total", bold);
    _ = lxw.worksheet_write_formula(worksheet, row + 1, col + 1, "=SUM(B2:B5)", money);

    // Save the workbook and free any allocated memory.
    _ = lxw.workbook_close(workbook);
}

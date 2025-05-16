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

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.WorkBook.init("out/tutorial1.xlsx");
    const worksheet = try workbook.addWorkSheet(null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Iterate over the data and write it out element by element.
    while (row < 4) {
        defer row += 1;

        try worksheet.writeString(
            row,
            col,
            expenses[row].item,
            .none,
        );

        try worksheet.writeNumber(
            row,
            col + 1,
            expenses[row].cost,
            .none,
        );
    }

    try worksheet.writeString(
        row,
        col,
        "Total",
        .none,
    );

    try worksheet.writeFormula(
        row,
        col + 1,
        "=SUM(B1:B4)",
        .none,
    );

    // Save the workbook and free any allocated memory.
    try workbook.deinit();
}

const xlsxwriter = @import("xlsxwriter");

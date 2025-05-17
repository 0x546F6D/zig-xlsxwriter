# zig-xlsxwriter

The `zig-xlsxwriter` is a wrapper of [`libxlsxwriter`](https://github.com/jmcnamara/libxlsxwriter) that can be used to write text, numbers,
dates and formulas to multiple worksheets in a new Excel 2007+ xlsx file. It
has a focus on performance and on fidelity with the file format created by
Excel. It cannot be used to modify an existing file.

## Requirements

- [zig v0.14.0](https://ziglang.org/download)
- [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)

## Example

Sample code from tutorial2.zig to generate excel file.

```zig
const Expense = struct {
    item: [*:0]const u8,
    cost: f64,
    datetime: xlsxwriter.datetime,
};

var expenses = [_]Expense{
    .{
        .item = "Rent",
        .cost = 1000.0,
        .datetime = .{
            .year = 2013,
            .month = 1,
            .day = 13,
            .hour = 8,
            .min = 34,
            .sec = 65.45,
        },
    },
    .{
        .item = "Gas",
        .cost = 100.0,
        .datetime = .{
            .year = 2013,
            .month = 1,
            .day = 14,
            .hour = 12,
            .min = 17,
            .sec = 23.34,
        },
    },
    .{
        .item = "Food",
        .cost = 300.0,
        .datetime = .{
            .year = 2013,
            .month = 1,
            .day = 16,
            .hour = 4,
            .min = 29,
            .sec = 54.05,
        },
    },
    .{
        .item = "Gym",
        .cost = 50.0,
        .datetime = .{
            .year = 2013,
            .month = 1,
            .day = 20,
            .hour = 6,
            .min = 55,
            .sec = 32.16,
        },
    },
};

pub fn main() !void {

    // Create a workbook and add a worksheet.
    const workbook = try xlsxwriter.WorkBook.init("out/tutorial2.xlsx");
    const worksheet = try workbook.addWorkSheet(null);

    // Start from the first cell. Rows and columns are zero indexed.
    var row: u32 = 0;
    const col: u16 = 0;

    // Add a bold format to use to highlight cells.
    const bold = try workbook.addFormat();
    bold.setBold();

    // Add a number format for cells with money.
    const money = try workbook.addFormat();
    money.setNumFormat("$#,##0");

    // Add an Excel date format.
    const date_format = try workbook.addFormat();
    date_format.setNumFormat("mmmm d yyyy");

    // Adjust the column width.
    try worksheet.setColumn(
        0,
        0,
        15,
        .none,
    );

    // Write some data header.
    try worksheet.writeString(
        row,
        col,
        "Item",
        bold,
    );
    try worksheet.writeString(
        row,
        col + 2,
        "Cost",
        bold,
    );

    // Iterate over the data and write it out element by element.
    for (&expenses) |*expense| {
        row += 1;
        try worksheet.writeString(
            row,
            col,
            expense.item,
            .none,
        );
        try worksheet.writeDateTime(
            row,
            col + 1,
            &expense.datetime,
            date_format,
        );
        try worksheet.writeNumber(
            row,
            col + 2,
            expense.cost,
            money,
        );
    }
    row += 1;

    // Write a total using a formula.
    try worksheet.writeString(
        row,
        col,
        "Total",
        .none,
    );
    try worksheet.writeFormula(
        row,
        col + 2,
        "=SUM(C2:C5)",
        .none,
    );

    // Save the workbook and free any allocated memory.
    try workbook.deinit();
}

const xlsxwriter = @import("xlsxwriter");
```


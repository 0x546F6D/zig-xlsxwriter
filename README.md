# zig-xlsxwriter

Branch to add the unit and functional tests as well as re-implement the python .xlsx files comparison code in zig.

The `zig-xlsxwriter` is a wrapper of [`libxlsxwriter`](https://github.com/jmcnamara/libxlsxwriter) that can be used to write text, numbers,
dates and formulas to multiple worksheets in a new Excel 2007+ xlsx file. It
has a focus on performance and on fidelity with the file format created by
Excel. It cannot be used to modify an existing file.

In order to keep the API consistent, some functions require an allocator:

- workbook.getWorkSheets()
- workbook.getChartSheets()
- worksheet.addTable()
- worksheet.writeRichString()
- series.setPoints()
- series.setLabelsCustom()

This allocator is passed while initialising the workbook and the allocations are freed either at the end of function call, or in the case of .getWorkSheets()/.getChartSheets, during the workbook.deinit():

```zig
const worbook = xwz.initWorkBook(allocator, filename, workbook_options);
defer workbook.deinit() catch {};
```

Except for .getWorkSheets() and getChartSheets(), it is possible to use libxlsxwriter C API semi-directly to use those functions without an allocator, as shown in the following examples:

- tables_no_alloc.zig, using addTableNoAlloc()
- rich_strings_no_alloc.zig, using writeRichStringNoAlloc()
- chart_pie_no_alloc.zig, using setPointsNoAlloc()
- chart_data_labels_no_alloc.zig, using setLabelsCustomNoAlloc()

Finally, libxlsxwriter C API can be used directly as well:

```zig
const c = @import("xlsxwriter").c;
```

## Requirements

- [zig v0.14.0](https://ziglang.org/download)
- [libxlsxwriter](https://github.com/jmcnamara/libxlsxwriter)

## Example

Sample code from tutorial3.zig to generate excel file.

```zig
const Expense = struct {
    item: [:0]const u8,
    cost: f64,
    datetime: xwz.DateTime,
};

var expenses = [_]Expense{ .{
    .item = "Rent",
    .cost = 1000.0,
    .datetime = .{ .year = 2013, .month = 1, .day = 13 },
}, .{
    .item = "Gas",
    .cost = 100.0,
    .datetime = .{ .year = 2013, .month = 1, .day = 14 },
}, .{
    .item = "Food",
    .cost = 300.0,
    .datetime = .{ .year = 2013, .month = 1, .day = 16 },
}, .{
    .item = "Gym",
    .cost = 50.0,
    .datetime = .{ .year = 2013, .month = 1, .day = 20 },
} };

pub fn main() !void {
    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, "tutorial3.xlsx", null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

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
    try worksheet.setColumn(.{}, 15, .default, null);

    // Write some data header.
    try worksheet.writeString(.{}, "Item", bold);
    try worksheet.writeString(.{ .col = 2 }, "Cost", bold);

    // Iterate over the data and write it out element by element.
    for (expenses, 1..) |expense, row| {
        try worksheet.writeString(.{ .row = @intCast(row) }, expense.item, .default);
        try worksheet.writeDateTime(.{ .row = @intCast(row), .col = 1 }, expense.datetime, date_format);
        try worksheet.writeNumber(.{ .row = @intCast(row), .col = 2 }, expense.cost, money);
    }

    // Write a total using a formula.
    try worksheet.writeString(.{ .row = expenses.len + 1 }, "Total", bold);
    try worksheet.writeFormula(.{ .row = expenses.len + 1, .col = 2 }, "=SUM(C2:C5)", money);
}

const xwz = @import("xlsxwriter");
```

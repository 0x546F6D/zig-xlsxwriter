pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Allocate memory for the data validation structure
    var data_validation: xwz.DataValidation = .{};

    // Add a format to use to highlight the header cells.
    const format = try workbook.addFormat();
    format.setBorder(.thin);
    format.setFgColor(@enumFromInt(0xC6EFCE));
    format.setBold();
    format.setTextWrap();
    format.setAlign(.vertical_center);
    format.setIndent(1);

    // Write some data for the validations.
    try write_worksheet_data(worksheet, format);

    // Set up layout of the worksheet.
    try worksheet.setColumn(.{}, 55, .default, null);
    try worksheet.setColumn(.{ .first = 1, .last = 1 }, 15, .default, null);
    try worksheet.setColumn(.{ .first = 3, .last = 3 }, 15, .default, null);
    try worksheet.setRow(0, 36, .default, null);

    // Example 1. Limiting input to an integer in a fixed range.
    try worksheet.writeString(.{ .row = 2, .col = 0 }, "Enter an integer between 1 and 10", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 10;

    try worksheet.dataValidationCell(.{ .row = 2, .col = 1 }, data_validation);

    // Example 2. Limiting input to an integer outside a fixed range.
    try worksheet.writeString(.{ .row = 4, .col = 0 }, "Enter an integer that is not between 1 and 10 (using cell references)", .default);

    data_validation.validate = .integer_formula;
    data_validation.criteria = .not_between;
    data_validation.minimum_formula = "=E3";
    data_validation.maximum_formula = "=F3";

    try worksheet.dataValidationCell(.{ .row = 4, .col = 1 }, data_validation);

    // Example 3. Limiting input to an integer greater than a fixed value.
    try worksheet.writeString(.{ .row = 6, .col = 0 }, "Enter an integer greater than 0", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .greater_than;
    data_validation.value_number = 0;

    try worksheet.dataValidationCell(.{ .row = 6, .col = 1 }, data_validation);

    // Example 4. Limiting input to an integer less than a fixed value.
    try worksheet.writeString(.{ .row = 8, .col = 0 }, "Enter an integer less than 10", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .less_than;
    data_validation.value_number = 10;

    try worksheet.dataValidationCell(.{ .row = 8, .col = 1 }, data_validation);

    // Example 5. Limiting input to a decimal in a fixed range.
    try worksheet.writeString(.{ .row = 10, .col = 0 }, "Enter a decimal between 0.1 and 0.5", .default);

    data_validation.validate = .decimal;
    data_validation.criteria = .between;
    data_validation.minimum_number = 0.1;
    data_validation.maximum_number = 0.5;

    try worksheet.dataValidationCell(.{ .row = 10, .col = 1 }, data_validation);

    // Example 6. Limiting input to a value in a dropdown list.
    try worksheet.writeString(.{ .row = 12, .col = 0 }, "Select a value from a dropdown list", .default);

    const list: xwz.CStringArray = &.{ "open", "high", "close" };
    data_validation.validate = .list;
    data_validation.value_list = list;

    try worksheet.dataValidationCell(.{ .row = 12, .col = 1 }, data_validation);

    // Example 7. Limiting input to a value in a dropdown list.
    try worksheet.writeString(.{ .row = 14, .col = 0 }, "Select a value from a dropdown list (using a cell range)", .default);

    data_validation.validate = .list_formula;
    data_validation.value_formula = "=$E$4:$G$4";

    try worksheet.dataValidationCell(.{ .row = 14, .col = 1 }, data_validation);

    // Example 8. Limiting input to a date in a fixed range.
    try worksheet.writeString(.{ .row = 16, .col = 0 }, "Enter a date between 1/1/2024 and 12/12/2024", .default);

    const datetime1: xwz.DateTime = .{ .year = 2024, .month = 1, .day = 1, .hour = 0, .min = 0, .sec = 0 };
    const datetime2: xwz.DateTime = .{ .year = 2024, .month = 12, .day = 12, .hour = 0, .min = 0, .sec = 0 };

    data_validation.validate = .date;
    data_validation.criteria = .between;
    data_validation.minimum_datetime = datetime1;
    data_validation.maximum_datetime = datetime2;

    try worksheet.dataValidationCell(.{ .row = 16, .col = 1 }, data_validation);

    // Example 9. Limiting input to a time in a fixed range.
    try worksheet.writeString(.{ .row = 18, .col = 0 }, "Enter a time between 6:00 and 12:00", .default);

    const datetime3: xwz.DateTime = .{ .year = 0, .month = 0, .day = 0, .hour = 6, .min = 0, .sec = 0 };
    const datetime4: xwz.DateTime = .{ .year = 0, .month = 0, .day = 0, .hour = 12, .min = 0, .sec = 0 };

    data_validation.validate = .time;
    data_validation.criteria = .between;
    data_validation.minimum_datetime = datetime3;
    data_validation.maximum_datetime = datetime4;

    try worksheet.dataValidationCell(.{ .row = 18, .col = 1 }, data_validation);

    // Example 10. Limiting input to a string greater than a fixed length.
    try worksheet.writeString(.{ .row = 20, .col = 0 }, "Enter a string longer than 3 characters", .default);

    data_validation.validate = .length;
    data_validation.criteria = .greater_than;
    data_validation.value_number = 3;

    try worksheet.dataValidationCell(.{ .row = 20, .col = 1 }, data_validation);

    // Example 11. Limiting input based on a formula.
    try worksheet.writeString(.{ .row = 22, .col = 0 }, "Enter a value if the following is true \"=AND(F5=50,G5=60)\"", .default);

    data_validation.validate = .custom_formula;
    data_validation.value_formula = "=AND(F5=50,G5=60)";

    try worksheet.dataValidationCell(.{ .row = 22, .col = 1 }, data_validation);

    // Example 12. Displaying and modifying data validation messages.
    try worksheet.writeString(.{ .row = 24, .col = 0 }, "Displays a message when you select the cell", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";

    try worksheet.dataValidationCell(.{ .row = 24, .col = 1 }, data_validation);

    // Example 13. Displaying and modifying data validation messages.
    try worksheet.writeString(.{ .row = 26, .col = 0 }, "Display a custom error message when integer isn't between 1 and 100", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";

    try worksheet.dataValidationCell(.{ .row = 26, .col = 1 }, data_validation);

    // Example 14. Displaying and modifying data validation messages.
    try worksheet.writeString(.{ .row = 28, .col = 0 }, "Display a custom info message when integer isn't between 1 and 100", .default);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";
    data_validation.error_type = .information;

    try worksheet.dataValidationCell(.{ .row = 28, .col = 1 }, data_validation);
}

// Write some data to the worksheet.
fn write_worksheet_data(worksheet: WorkSheet, format: Format) !void {
    try worksheet.writeString(.{ .row = 0, .col = 0 }, "Some examples of data validation in libxwz", format);
    try worksheet.writeString(.{ .row = 0, .col = 1 }, "Enter values in this column", format);
    try worksheet.writeString(.{ .row = 0, .col = 3 }, "Sample Data", format);

    try worksheet.writeString(.{ .row = 2, .col = 3 }, "Integers", .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 4 }, 1, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 5 }, 10, .default);

    try worksheet.writeString(.{ .row = 3, .col = 3 }, "List data", .default);
    try worksheet.writeString(.{ .row = 3, .col = 4 }, "open", .default);
    try worksheet.writeString(.{ .row = 3, .col = 5 }, "high", .default);
    try worksheet.writeString(.{ .row = 3, .col = 6 }, "close", .default);

    try worksheet.writeString(.{ .row = 4, .col = 3 }, "Formula", .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 5 }, 50, .default);
    try worksheet.writeFormula(.{ .row = 4, .col = 4 }, "=AND(F5=50,G5=60)", .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 6 }, 60, .default);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");
const WorkSheet = xwz.WorkSheet;
const Format = xwz.Format;

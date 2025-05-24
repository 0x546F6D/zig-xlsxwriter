// Write some data to the worksheet.
fn write_worksheet_data(worksheet: WorkSheet, format: Format) !void {
    try worksheet.writeString(0, 0, "Some examples of data validation in libxlsxwriter", format);
    try worksheet.writeString(0, 1, "Enter values in this column", format);
    try worksheet.writeString(0, 3, "Sample Data", format);

    try worksheet.writeString(2, 3, "Integers", .none);
    try worksheet.writeNumber(2, 4, 1, .none);
    try worksheet.writeNumber(2, 5, 10, .none);

    try worksheet.writeString(3, 3, "List data", .none);
    try worksheet.writeString(3, 4, "open", .none);
    try worksheet.writeString(3, 5, "high", .none);
    try worksheet.writeString(3, 6, "close", .none);

    try worksheet.writeString(4, 3, "Formula", .none);
    try worksheet.writeFormula(4, 4, "=AND(F5=50,G5=60)", .none);
    try worksheet.writeNumber(4, 5, 50, .none);
    try worksheet.writeNumber(4, 6, 60, .none);
}

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add worksheets.
    const workbook = try xlsxwriter.initWorkBook(xlsx_path.ptr);
    defer workbook.deinit() catch {};

    const worksheet = try workbook.addWorkSheet(null);

    // Allocate memory for the data validation structure
    var data_validation: xlsxwriter.DataValidation = .{};

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
    try worksheet.setColumn(0, 0, 55, .none);
    try worksheet.setColumn(1, 1, 15, .none);
    try worksheet.setColumn(3, 3, 15, .none);
    try worksheet.setRow(0, 36, .none);

    // Example 1. Limiting input to an integer in a fixed range.
    try worksheet.writeString(2, 0, "Enter an integer between 1 and 10", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 10;

    try worksheet.dataValidationCell(2, 1, data_validation);

    // Example 2. Limiting input to an integer outside a fixed range.
    try worksheet.writeString(4, 0, "Enter an integer that is not between 1 and 10 (using cell references)", .none);

    data_validation.validate = .integer_formula;
    data_validation.criteria = .not_between;
    data_validation.minimum_formula = "=E3";
    data_validation.maximum_formula = "=F3";

    try worksheet.dataValidationCell(4, 1, data_validation);

    // Example 3. Limiting input to an integer greater than a fixed value.
    try worksheet.writeString(6, 0, "Enter an integer greater than 0", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .greater_than;
    data_validation.value_number = 0;

    try worksheet.dataValidationCell(6, 1, data_validation);

    // Example 4. Limiting input to an integer less than a fixed value.
    try worksheet.writeString(8, 0, "Enter an integer less than 10", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .less_than;
    data_validation.value_number = 10;

    try worksheet.dataValidationCell(8, 1, data_validation);

    // Example 5. Limiting input to a decimal in a fixed range.
    try worksheet.writeString(10, 0, "Enter a decimal between 0.1 and 0.5", .none);

    data_validation.validate = .decimal;
    data_validation.criteria = .between;
    data_validation.minimum_number = 0.1;
    data_validation.maximum_number = 0.5;

    try worksheet.dataValidationCell(10, 1, data_validation);

    // Example 6. Limiting input to a value in a dropdown list.
    try worksheet.writeString(12, 0, "Select a value from a dropdown list", .none);

    const list: xlsxwriter.FilterListType = &.{ "open", "high", "close" };
    data_validation.validate = .list;
    data_validation.value_list = list;

    try worksheet.dataValidationCell(12, 1, data_validation);

    // Example 7. Limiting input to a value in a dropdown list.
    try worksheet.writeString(14, 0, "Select a value from a dropdown list (using a cell range)", .none);

    data_validation.validate = .list_formula;
    data_validation.value_formula = "=$E$4:$G$4";

    try worksheet.dataValidationCell(14, 1, data_validation);

    // Example 8. Limiting input to a date in a fixed range.
    try worksheet.writeString(16, 0, "Enter a date between 1/1/2024 and 12/12/2024", .none);

    const datetime1: xlsxwriter.DateTime = .{ .year = 2024, .month = 1, .day = 1, .hour = 0, .min = 0, .sec = 0 };
    const datetime2: xlsxwriter.DateTime = .{ .year = 2024, .month = 12, .day = 12, .hour = 0, .min = 0, .sec = 0 };

    data_validation.validate = .date;
    data_validation.criteria = .between;
    data_validation.minimum_datetime = datetime1;
    data_validation.maximum_datetime = datetime2;

    try worksheet.dataValidationCell(16, 1, data_validation);

    // Example 9. Limiting input to a time in a fixed range.
    try worksheet.writeString(18, 0, "Enter a time between 6:00 and 12:00", .none);

    const datetime3: xlsxwriter.DateTime = .{ .year = 0, .month = 0, .day = 0, .hour = 6, .min = 0, .sec = 0 };
    const datetime4: xlsxwriter.DateTime = .{ .year = 0, .month = 0, .day = 0, .hour = 12, .min = 0, .sec = 0 };

    data_validation.validate = .time;
    data_validation.criteria = .between;
    data_validation.minimum_datetime = datetime3;
    data_validation.maximum_datetime = datetime4;

    try worksheet.dataValidationCell(18, 1, data_validation);

    // Example 10. Limiting input to a string greater than a fixed length.
    try worksheet.writeString(20, 0, "Enter a string longer than 3 characters", .none);

    data_validation.validate = .length;
    data_validation.criteria = .greater_than;
    data_validation.value_number = 3;

    try worksheet.dataValidationCell(20, 1, data_validation);

    // Example 11. Limiting input based on a formula.
    try worksheet.writeString(22, 0, "Enter a value if the following is true \"=AND(F5=50,G5=60)\"", .none);

    data_validation.validate = .custom_formula;
    data_validation.value_formula = "=AND(F5=50,G5=60)";

    try worksheet.dataValidationCell(22, 1, data_validation);

    // Example 12. Displaying and modifying data validation messages.
    try worksheet.writeString(24, 0, "Displays a message when you select the cell", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";

    try worksheet.dataValidationCell(24, 1, data_validation);

    // Example 13. Displaying and modifying data validation messages.
    try worksheet.writeString(26, 0, "Display a custom error message when integer isn't between 1 and 100", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";

    try worksheet.dataValidationCell(26, 1, data_validation);

    // Example 14. Displaying and modifying data validation messages.
    try worksheet.writeString(28, 0, "Display a custom info message when integer isn't between 1 and 100", .none);

    data_validation.validate = .integer;
    data_validation.criteria = .between;
    data_validation.minimum_number = 1;
    data_validation.maximum_number = 100;
    data_validation.input_title = "Enter an integer:";
    data_validation.input_message = "between 1 and 100";
    data_validation.error_title = "Input value is not valid!";
    data_validation.error_message = "It should be an integer between 1 and 100";
    data_validation.error_type = .information;

    try worksheet.dataValidationCell(28, 1, data_validation);

    // Cleanup.
    // std.heap.c_allocator.destroy(data_validation);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xlsxwriter = @import("xlsxwriter");
const WorkSheet = @import("xlsxwriter").WorkSheet;
const Format = @import("xlsxwriter").Format;

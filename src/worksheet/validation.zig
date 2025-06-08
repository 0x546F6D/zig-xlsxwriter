pub const ValidationBoolean = enum(u8) {
    default = c.LXW_VALIDATION_DEFAULT,
    off = c.LXW_VALIDATION_OFF,
    on = c.LXW_VALIDATION_ON,
};
pub const ValidationTypes = enum(u8) {
    none = c.LXW_VALIDATION_TYPE_NONE,
    integer = c.LXW_VALIDATION_TYPE_INTEGER,
    integer_formula = c.LXW_VALIDATION_TYPE_INTEGER_FORMULA,
    decimal = c.LXW_VALIDATION_TYPE_DECIMAL,
    decimal_formula = c.LXW_VALIDATION_TYPE_DECIMAL_FORMULA,
    list = c.LXW_VALIDATION_TYPE_LIST,
    list_formula = c.LXW_VALIDATION_TYPE_LIST_FORMULA,
    date = c.LXW_VALIDATION_TYPE_DATE,
    date_formula = c.LXW_VALIDATION_TYPE_DATE_FORMULA,
    date_number = c.LXW_VALIDATION_TYPE_DATE_NUMBER,
    time = c.LXW_VALIDATION_TYPE_TIME,
    time_formula = c.LXW_VALIDATION_TYPE_TIME_FORMULA,
    time_number = c.LXW_VALIDATION_TYPE_TIME_NUMBER,
    length = c.LXW_VALIDATION_TYPE_LENGTH,
    length_formula = c.LXW_VALIDATION_TYPE_LENGTH_FORMULA,
    custom_formula = c.LXW_VALIDATION_TYPE_CUSTOM_FORMULA,
    any = c.LXW_VALIDATION_TYPE_ANY,
};
pub const ValidationCriteria = enum(u8) {
    none = c.LXW_VALIDATION_CRITERIA_NONE,
    between = c.LXW_VALIDATION_CRITERIA_BETWEEN,
    not_between = c.LXW_VALIDATION_CRITERIA_NOT_BETWEEN,
    equal_to = c.LXW_VALIDATION_CRITERIA_EQUAL_TO,
    not_equal_to = c.LXW_VALIDATION_CRITERIA_NOT_EQUAL_TO,
    greater_than = c.LXW_VALIDATION_CRITERIA_GREATER_THAN,
    less_than = c.LXW_VALIDATION_CRITERIA_LESS_THAN,
    greater_than_or_equal_to = c.LXW_VALIDATION_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
    less_than_or_equal_to = c.LXW_VALIDATION_CRITERIA_LESS_THAN_OR_EQUAL_TO,
};
pub const ValidationErrorTypes = enum(u8) {
    stop = c.LXW_VALIDATION_ERROR_TYPE_STOP,
    warning = c.LXW_VALIDATION_ERROR_TYPE_WARNING,
    information = c.LXW_VALIDATION_ERROR_TYPE_INFORMATION,
};
pub const DataValidation = struct {
    validate: ValidationTypes = .none,
    criteria: ValidationCriteria = .none,
    ignore_blank: ValidationBoolean = .default,
    show_input: ValidationBoolean = .default,
    show_error: ValidationBoolean = .default,
    error_type: ValidationErrorTypes = .stop,
    dropdown: ValidationBoolean = .default,
    value_number: f64 = 0,
    value_formula: ?[*:0]const u8 = null,
    value_list: CStringArray = &.{},
    value_datetime: c.lxw_datetime = .{},
    minimum_number: f64 = 0,
    minimum_formula: ?[*:0]const u8 = null,
    minimum_datetime: c.lxw_datetime = .{},
    maximum_number: f64 = 0,
    maximum_formula: ?[*:0]const u8 = null,
    maximum_datetime: c.lxw_datetime = .{},
    input_title: ?[*:0]const u8 = null,
    input_message: ?[*:0]const u8 = null,
    error_title: ?[*:0]const u8 = null,
    error_message: ?[*:0]const u8 = null,

    pub const default = DataValidation{
        .validate = .none,
        .criteria = .none,
        .ignore_blank = .default,
        .show_input = .default,
        .show_error = .default,
        .error_type = .stop,
        .dropdown = .default,
        .value_number = 0,
        .value_formula = null,
        .value_list = &.{},
        .value_datetime = .{},
        .minimum_number = 0,
        .minimum_formula = null,
        .minimum_datetime = .{},
        .maximum_number = 0,
        .maximum_formula = null,
        .maximum_datetime = .{},
        .input_title = null,
        .input_message = null,
        .error_title = null,
        .error_message = null,
    };

    inline fn toC(self: DataValidation) c.lxw_data_validation {
        return c.lxw_data_validation{
            .validate = @intFromEnum(self.validate),
            .criteria = @intFromEnum(self.criteria),
            .ignore_blank = @intFromEnum(self.ignore_blank),
            .show_input = @intFromEnum(self.show_input),
            .show_error = @intFromEnum(self.show_error),
            .error_type = @intFromEnum(self.error_type),
            .dropdown = @intFromEnum(self.dropdown),
            .value_number = self.value_number,
            .value_formula = self.value_formula,
            .value_list = @ptrCast(@constCast(self.value_list)),
            .value_datetime = self.value_datetime,
            .minimum_number = self.minimum_number,
            .minimum_formula = self.minimum_formula,
            .minimum_datetime = self.minimum_datetime,
            .maximum_number = self.maximum_number,
            .maximum_formula = self.maximum_formula,
            .maximum_datetime = self.maximum_datetime,
            .input_title = self.input_title,
            .input_message = self.input_message,
            .error_title = self.error_title,
            .error_message = self.error_message,
        };
    }
};

// pub extern fn worksheet_data_validation_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
pub inline fn dataValidationCell(
    self: WorkSheet,
    cell: Cell,
    validation: DataValidation,
) XlsxError!void {
    try check(c.worksheet_data_validation_cell(
        self.worksheet_c,
        cell.row,
        cell.col,
        @constCast(&validation.toC()),
    ));
}

// pub extern fn worksheet_data_validation_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
pub inline fn dataValidationRange(
    self: WorkSheet,
    range: Range,
    validation: DataValidation,
) XlsxError!void {
    try check(c.worksheet_data_validation_range(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        @constCast(&validation.toC()),
    ));
}

const c = @import("lxw");
const WorkSheet = @import("../WorkSheet.zig");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const utility = @import("../utility.zig");
const Cell = utility.Cell;
const Range = utility.Range;
const CStringArray = @import("../xlsxwriter.zig").CStringArray;

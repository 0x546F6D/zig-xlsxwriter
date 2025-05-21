const WorkSheet = @This();

worksheet_c: ?*c.lxw_worksheet,

pub inline fn setColumn(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_column(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
    ));
}

pub inline fn setRow(
    self: WorkSheet,
    row: u16,
    height: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_row(
        self.worksheet_c,
        row,
        height,
        format.format_c,
    ));
}

pub const RowColOptions = c.lxw_row_col_options;
// pub extern fn worksheet_set_row_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setRowOpt(
    self: WorkSheet,
    row: u32,
    height: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_row_opt(
        self.worksheet_c,
        row,
        height,
        format.format_c,
        @constCast(&options),
    ));
}

pub inline fn setColumnOpt(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_column_opt(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
        @constCast(&options),
    ));
}
pub inline fn writeString(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_string(
        self.worksheet_c,
        row,
        col,
        string,
        format.format_c,
    ));
}

pub inline fn writeNumber(
    self: WorkSheet,
    row: u32,
    col: u16,
    number: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_number(
        self.worksheet_c,
        row,
        col,
        number,
        format.format_c,
    ));
}

pub inline fn writeDateTime(
    self: WorkSheet,
    row: u32,
    col: u16,
    datetime: c.lxw_datetime,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_datetime(
        self.worksheet_c,
        row,
        col,
        @constCast(&datetime),
        format.format_c,
    ));
}

pub inline fn writeFormula(
    self: WorkSheet,
    row: u32,
    col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_formula(
        self.worksheet_c,
        row,
        col,
        formula,
        format.format_c,
    ));
}

pub inline fn writeArrayFormula(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_array_formula(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        formula,
        format.format_c,
    ));
}

pub inline fn insertChart(
    self: WorkSheet,
    row: u32,
    col: u16,
    chart: Chart,
) XlsxError!void {
    try check(c.worksheet_insert_chart(
        self.worksheet_c,
        row,
        col,
        chart.chart_c,
    ));
}

pub const FilterRule = struct {
    criteria: FilterCriteria,
    value_string: [*c]const u8 = null,
    value: f64 = 0,
};

pub const FilterCriteria = enum(c_int) {
    none = c.LXW_FILTER_CRITERIA_NONE,
    equal_to = c.LXW_FILTER_CRITERIA_EQUAL_TO,
    not_equal_to = c.LXW_FILTER_CRITERIA_NOT_EQUAL_TO,
    greater_than = c.LXW_FILTER_CRITERIA_GREATER_THAN,
    less_than = c.LXW_FILTER_CRITERIA_LESS_THAN,
    greater_than_or_equal_to = c.LXW_FILTER_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
    less_than_or_equal_to = c.LXW_FILTER_CRITERIA_LESS_THAN_OR_EQUAL_TO,
    blanks = c.LXW_FILTER_CRITERIA_BLANKS,
    non_blanks = c.LXW_FILTER_CRITERIA_NON_BLANKS,
};

// pub extern fn worksheet_autofilter(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
pub inline fn autoFilter(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
) XlsxError!void {
    try check(c.worksheet_autofilter(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
    ));
}

// pub extern fn worksheet_filter_column(worksheet: [*c]lxw_worksheet, col: lxw_col_t, rule: [*c]lxw_filter_rule) lxw_error;
pub inline fn filterColumn(
    self: WorkSheet,
    col: u16,
    rule: FilterRule,
) XlsxError!void {
    try check(c.worksheet_filter_column(
        self.worksheet_c,
        col,
        @constCast(&c.lxw_filter_rule{
            .criteria = @intFromEnum(rule.criteria),
            .value_string = rule.value_string,
            .value = rule.value,
        }),
    ));
}

// pub extern fn worksheet_filter_column2(worksheet: [*c]lxw_worksheet, col: lxw_col_t, rule1: [*c]lxw_filter_rule, rule2: [*c]lxw_filter_rule, and_or: u8) lxw_error;
pub inline fn filterColumn2(
    self: WorkSheet,
    col: u16,
    rule1: FilterRule,
    rule2: FilterRule,
    and_or: u8,
) XlsxError!void {
    try check(c.worksheet_filter_column2(
        self.worksheet_c,
        col,
        @constCast(&c.lxw_filter_rule{
            .criteria = @intFromEnum(rule1.criteria),
            .value_string = rule1.value_string,
            .value = rule1.value,
        }),
        @constCast(&c.lxw_filter_rule{
            .criteria = @intFromEnum(rule2.criteria),
            .value_string = rule2.value_string,
            .value = rule2.value,
        }),
        and_or,
    ));
}

// pub extern fn worksheet_filter_list(worksheet: [*c]lxw_worksheet, col: lxw_col_t, list: [*c][*c]const u8) lxw_error;
pub inline fn filterList(
    self: WorkSheet,
    col: u16,
    list: []const [:0]const u8,
) XlsxError!void {
    var list_c: [64][*c]const u8 = undefined;
    for (list, 0..) |item, i| list_c[i] = item.ptr;
    list_c[list.len] = null;
    try check(c.worksheet_filter_list(
        self.worksheet_c,
        col,
        @ptrCast(&list_c),
    ));
}

const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");

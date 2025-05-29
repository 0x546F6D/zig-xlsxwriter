pub const filter_or = c.LXW_FILTER_OR;
pub const filter_and = c.LXW_FILTER_AND;

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

pub const FilterRule = struct {
    criteria: FilterCriteria,
    value_string: ?CString = null,
    value: f64 = 0,
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
    list: CStringArray,
) !void {
    try check(c.worksheet_filter_list(
        self.worksheet_c,
        col,
        @ptrCast(@constCast(list.ptr)),
    ));
}

const Allocator = @import("std").mem.Allocator;
const c = @import("lxw");
const CString = @import("xlsxwriter.zig").CString;
const CStringArray = @import("xlsxwriter.zig").CStringArray;
const WorkSheet = @import("WorkSheet.zig");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;

pub const filter_or = c.LXW_FILTER_OR;
pub const filter_and = c.LXW_FILTER_AND;

pub const FilterCriteria = enum(u8) {
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
    criteria: FilterCriteria = .none,
    value_string: ?[*:0]const u8 = null,
    value: f64 = 0,

    inline fn toC(self: FilterRule) c.lxw_filter_rule {
        return c.lxw_filter_rule{
            .criteria = @intFromEnum(self.criteria),
            .value_string = self.value_string,
            .value = self.value,
        };
    }
};

// pub extern fn worksheet_autofilter(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
pub inline fn autoFilter(
    self: WorkSheet,
    range: Range,
) XlsxError!void {
    try check(c.worksheet_autofilter(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
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
        @constCast(&rule.toC()),
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
        @constCast(&rule1.toC()),
        @constCast(&rule2.toC()),
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
const CStringArray = @import("xlsxwriter.zig").CStringArray;
const utility = @import("utility.zig");
const Cell = utility.Cell;
const Cols = utility.Cols;
const Range = utility.Range;
const WorkSheet = @import("WorkSheet.zig");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;

pub const ConditionalFormatTypes = enum(u8) {
    none = c.LXW_CONDITIONAL_TYPE_NONE,
    cell = c.LXW_CONDITIONAL_TYPE_CELL,
    text = c.LXW_CONDITIONAL_TYPE_TEXT,
    time_period = c.LXW_CONDITIONAL_TYPE_TIME_PERIOD,
    average = c.LXW_CONDITIONAL_TYPE_AVERAGE,
    duplicate = c.LXW_CONDITIONAL_TYPE_DUPLICATE,
    unique = c.LXW_CONDITIONAL_TYPE_UNIQUE,
    top = c.LXW_CONDITIONAL_TYPE_TOP,
    bottom = c.LXW_CONDITIONAL_TYPE_BOTTOM,
    blanks = c.LXW_CONDITIONAL_TYPE_BLANKS,
    no_blanks = c.LXW_CONDITIONAL_TYPE_NO_BLANKS,
    errors = c.LXW_CONDITIONAL_TYPE_ERRORS,
    no_errors = c.LXW_CONDITIONAL_TYPE_NO_ERRORS,
    formula = c.LXW_CONDITIONAL_TYPE_FORMULA,
    color_2_scale = c.LXW_CONDITIONAL_2_COLOR_SCALE,
    color_3_scale = c.LXW_CONDITIONAL_3_COLOR_SCALE,
    data_bar = c.LXW_CONDITIONAL_DATA_BAR,
    icon_sets = c.LXW_CONDITIONAL_TYPE_ICON_SETS,
    last = c.LXW_CONDITIONAL_TYPE_LAST,
};

pub const ConditionalCriteria = enum(u8) {
    none = c.LXW_CONDITIONAL_CRITERIA_NONE,
    equal_to = c.LXW_CONDITIONAL_CRITERIA_EQUAL_TO,
    not_equal_to = c.LXW_CONDITIONAL_CRITERIA_NOT_EQUAL_TO,
    greater_than = c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN,
    less_than = c.LXW_CONDITIONAL_CRITERIA_LESS_THAN,
    greater_than_or_equal_to = c.LXW_CONDITIONAL_CRITERIA_GREATER_THAN_OR_EQUAL_TO,
    less_than_or_equal_to = c.LXW_CONDITIONAL_CRITERIA_LESS_THAN_OR_EQUAL_TO,
    between = c.LXW_CONDITIONAL_CRITERIA_BETWEEN,
    not_between = c.LXW_CONDITIONAL_CRITERIA_NOT_BETWEEN,
    text_containing = c.LXW_CONDITIONAL_CRITERIA_TEXT_CONTAINING,
    text_not_containing = c.LXW_CONDITIONAL_CRITERIA_TEXT_NOT_CONTAINING,
    text_begins_with = c.LXW_CONDITIONAL_CRITERIA_TEXT_BEGINS_WITH,
    text_ends_with = c.LXW_CONDITIONAL_CRITERIA_TEXT_ENDS_WITH,
    time_period_yesterday = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_YESTERDAY,
    time_period_today = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TODAY,
    time_period_tomorrow = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_TOMORROW,
    time_period_last_7_days = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_7_DAYS,
    time_period_last_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_WEEK,
    time_period_this_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_WEEK,
    time_period_next_week = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_WEEK,
    time_period_last_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_LAST_MONTH,
    time_period_this_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_THIS_MONTH,
    time_period_next_month = c.LXW_CONDITIONAL_CRITERIA_TIME_PERIOD_NEXT_MONTH,
    average_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE,
    average_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW,
    average_above_or_equal = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_ABOVE_OR_EQUAL,
    average_below_or_equal = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_BELOW_OR_EQUAL,
    average_1_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_ABOVE,
    average_1_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_1_STD_DEV_BELOW,
    average_2_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_ABOVE,
    average_2_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_2_STD_DEV_BELOW,
    average_3_std_dev_above = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_ABOVE,
    average_3_std_dev_below = c.LXW_CONDITIONAL_CRITERIA_AVERAGE_3_STD_DEV_BELOW,
    top_or_bottom_percent = c.LXW_CONDITIONAL_CRITERIA_TOP_OR_BOTTOM_PERCENT,
};

pub const ConditionalFormatRuleTypes = enum(u8) {
    none = c.LXW_CONDITIONAL_RULE_TYPE_NONE,
    minimum = c.LXW_CONDITIONAL_RULE_TYPE_MINIMUM,
    number = c.LXW_CONDITIONAL_RULE_TYPE_NUMBER,
    percent = c.LXW_CONDITIONAL_RULE_TYPE_PERCENT,
    percentile = c.LXW_CONDITIONAL_RULE_TYPE_PERCENTILE,
    formula = c.LXW_CONDITIONAL_RULE_TYPE_FORMULA,
    maximum = c.LXW_CONDITIONAL_RULE_TYPE_MAXIMUM,
    auto_min = c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MIN,
    auto_max = c.LXW_CONDITIONAL_RULE_TYPE_AUTO_MAX,
};

pub const ConditionalFormatBarDirection = enum(u8) {
    context = c.LXW_CONDITIONAL_BAR_DIRECTION_CONTEXT,
    right_to_left = c.LXW_CONDITIONAL_BAR_DIRECTION_RIGHT_TO_LEFT,
    left_to_right = c.LXW_CONDITIONAL_BAR_DIRECTION_LEFT_TO_RIGHT,
};

pub const ConditionalBarAxisPosition = enum(u8) {
    automatic = c.LXW_CONDITIONAL_BAR_AXIS_AUTOMATIC,
    midpoint = c.LXW_CONDITIONAL_BAR_AXIS_MIDPOINT,
    none = c.LXW_CONDITIONAL_BAR_AXIS_NONE,
};

pub const ConditionalIconTypes = enum(u8) {
    icons_3_arrows_colored = c.LXW_CONDITIONAL_ICONS_3_ARROWS_COLORED,
    icons_3_arrows_gray = c.LXW_CONDITIONAL_ICONS_3_ARROWS_GRAY,
    icons_3_flags = c.LXW_CONDITIONAL_ICONS_3_FLAGS,
    icons_3_traffic_lights_unrimmed = c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_UNRIMMED,
    icons_3_traffic_lights_rimmed = c.LXW_CONDITIONAL_ICONS_3_TRAFFIC_LIGHTS_RIMMED,
    icons_3_signs = c.LXW_CONDITIONAL_ICONS_3_SIGNS,
    icons_3_symbols_circled = c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_CIRCLED,
    icons_3_symbols_uncircled = c.LXW_CONDITIONAL_ICONS_3_SYMBOLS_UNCIRCLED,
    icons_4_arrows_colored = c.LXW_CONDITIONAL_ICONS_4_ARROWS_COLORED,
    icons_4_arrows_gray = c.LXW_CONDITIONAL_ICONS_4_ARROWS_GRAY,
    icons_4_red_to_black = c.LXW_CONDITIONAL_ICONS_4_RED_TO_BLACK,
    icons_4_ratings = c.LXW_CONDITIONAL_ICONS_4_RATINGS,
    icons_4_traffic_lights = c.LXW_CONDITIONAL_ICONS_4_TRAFFIC_LIGHTS,
    icons_5_arrows_colored = c.LXW_CONDITIONAL_ICONS_5_ARROWS_COLORED,
    icons_5_arrows_gray = c.LXW_CONDITIONAL_ICONS_5_ARROWS_GRAY,
    icons_5_ratings = c.LXW_CONDITIONAL_ICONS_5_RATINGS,
    icons_5_quarters = c.LXW_CONDITIONAL_ICONS_5_QUARTERS,
};

pub const ConditionalFormat = struct {
    type: ConditionalFormatTypes = .none,
    criteria: ConditionalCriteria = .none,
    value: f64 = 0,
    value_string: ?[*:0]const u8 = null,
    format: Format = .default,
    min_value: f64 = 0,
    min_value_string: ?[*:0]const u8 = null,
    min_rule_type: ConditionalFormatRuleTypes = .none,
    min_color: Format.DefinedColors = .default,
    mid_value: f64 = 0,
    mid_value_string: ?[*:0]const u8 = null,
    mid_rule_type: ConditionalFormatRuleTypes = .none,
    mid_color: Format.DefinedColors = .default,
    max_value: f64 = 0,
    max_value_string: ?[*:0]const u8 = null,
    max_rule_type: ConditionalFormatRuleTypes = .none,
    max_color: Format.DefinedColors = .default,
    bar_color: Format.DefinedColors = .default,
    bar_only: bool = false,
    data_bar_2010: bool = false,
    bar_solid: bool = false,
    bar_negative_color: Format.DefinedColors = .default,
    bar_border_color: Format.DefinedColors = .default,
    bar_negative_border_color: Format.DefinedColors = .default,
    bar_negative_color_same: bool = false,
    bar_negative_border_color_same: bool = false,
    bar_no_border: bool = false,
    bar_direction: ConditionalFormatBarDirection = .context,
    bar_axis_position: ConditionalBarAxisPosition = .automatic,
    bar_axis_color: Format.DefinedColors = .default,
    icon_style: ConditionalIconTypes = .icons_3_arrows_colored,
    reverse_icons: bool = false,
    icons_only: bool = false,
    multi_range: ?[*:0]const u8 = null,
    stop_if_true: bool = false,

    pub const default: ConditionalFormat = .{
        .type = .none,
        .criteria = .none,
        .value = 0,
        .value_string = null,
        .format = .default,
        .min_value = 0,
        .min_value_string = null,
        .min_rule_type = .none,
        .min_color = .default,
        .mid_value = 0,
        .mid_value_string = null,
        .mid_rule_type = .none,
        .mid_color = .default,
        .max_value = 0,
        .max_value_string = null,
        .max_rule_type = .none,
        .max_color = .default,
        .bar_color = .default,
        .bar_only = false,
        .data_bar_2010 = false,
        .bar_solid = false,
        .bar_negative_color = .default,
        .bar_border_color = .default,
        .bar_negative_border_color = .default,
        .bar_negative_color_same = false,
        .bar_negative_border_color_same = false,
        .bar_no_border = false,
        .bar_direction = .context,
        .bar_axis_position = .automatic,
        .bar_axis_color = .default,
        .icon_style = .icons_3_arrows_colored,
        .reverse_icons = false,
        .icons_only = false,
        .multi_range = null,
        .stop_if_true = false,
    };

    inline fn toC(self: ConditionalFormat) c.lxw_conditional_format {
        return c.lxw_conditional_format{
            .type = @intFromEnum(self.type),
            .criteria = @intFromEnum(self.criteria),
            .value = self.value,
            .value_string = self.value_string,
            .format = self.format.format_c,
            .min_value = self.min_value,
            .min_value_string = self.min_value_string,
            .min_rule_type = @intFromEnum(self.min_rule_type),
            .min_color = @intFromEnum(self.min_color),
            .mid_value = self.mid_value,
            .mid_value_string = self.mid_value_string,
            .mid_rule_type = @intFromEnum(self.mid_rule_type),
            .mid_color = @intFromEnum(self.mid_color),
            .max_value = self.max_value,
            .max_value_string = self.max_value_string,
            .max_rule_type = @intFromEnum(self.max_rule_type),
            .max_color = @intFromEnum(self.max_color),
            .bar_color = @intFromEnum(self.bar_color),
            .bar_only = @intFromBool(self.bar_only),
            .data_bar_2010 = @intFromBool(self.data_bar_2010),
            .bar_solid = @intFromBool(self.bar_solid),
            .bar_negative_color = @intFromEnum(self.bar_negative_color),
            .bar_border_color = @intFromEnum(self.bar_border_color),
            .bar_negative_border_color = @intFromEnum(self.bar_negative_border_color),
            .bar_negative_color_same = @intFromBool(self.bar_negative_color_same),
            .bar_negative_border_color_same = @intFromBool(self.bar_negative_border_color_same),
            .bar_no_border = @intFromBool(self.bar_no_border),
            .bar_direction = @intFromEnum(self.bar_direction),
            .bar_axis_position = @intFromEnum(self.bar_axis_position),
            .bar_axis_color = @intFromEnum(self.bar_axis_color),
            .icon_style = @intFromEnum(self.icon_style),
            .reverse_icons = @intFromBool(self.reverse_icons),
            .icons_only = @intFromBool(self.icons_only),
            .multi_range = self.multi_range,
            .stop_if_true = @intFromBool(self.stop_if_true),
        };
    }
};

// pub extern fn worksheet_conditional_format_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;
pub inline fn formatRange(
    self: WorkSheet,
    range: Range,
    conditional_format: ConditionalFormat,
) XlsxError!void {
    try check(c.worksheet_conditional_format_range(
        self.worksheet_c,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
        @constCast(&conditional_format.toC()),
    ));
}

// pub extern fn worksheet_conditional_format_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;
pub inline fn formatCell(
    self: WorkSheet,
    cell: Cell,
    conditional_format: ConditionalFormat,
) XlsxError!void {
    try check(c.worksheet_conditional_format_cell(
        self.worksheet_c,
        cell.row,
        cell.col,
        @constCast(&conditional_format.toC()),
    ));
}

const c = @import("lxw");
const WorkSheet = @import("../WorkSheet.zig");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const Format = @import("../Format.zig");
const utility = @import("../utility.zig");
const Cell = utility.Cell;
const Range = utility.Range;

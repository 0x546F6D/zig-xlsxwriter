const Chart = @This();

chart_c: ?*c.lxw_chart,

pub const Font = struct {
    name: ?CString = null,
    size: f64 = 0,
    bold: bool = true,
    italic: bool = false,
    underline: bool = false,
    rotation: i32 = 0,
    color: DefinedColors = .default,
    pitch_family: u8 = 0,
    charset: u8 = 0,
    baseline: i8 = 0,

    pub const default = Font{
        .name = null,
        .size = 0,
        .bold = false,
        .italic = false,
        .underline = false,
        .rotation = 0,
        .color = .default,
        .pitch_family = 0,
        .charset = 0,
        .baseline = 0,
    };

    inline fn toC(self: Font) c.lxw_chart_font {
        return c.struct_lxw_chart_font{
            .name = self.name,
            .size = self.size,
            // bold by default, need to explicitely disable it
            .bold = if (self.bold) @intFromBool(self.bold) else @intFromEnum(Bool.explicit_false),
            .italic = @intFromBool(self.italic),
            .underline = @intFromBool(self.underline),
            .rotation = self.rotation,
            .color = @intFromEnum(self.color),
            .pitch_family = self.pitch_family,
            .charset = self.charset,
            .baseline = self.baseline,
        };
    }
};

// lxw_chart_layout
pub const Layout = struct {
    x: f64 = 0,
    y: f64 = 0,
    width: f64 = 0,
    height: f64 = 0,
    has_inner: bool = false,

    pub const default = Layout{
        .x = 0,
        .y = 0,
        .width = 0,
        .height = 0,
        .has_inner = false,
    };

    inline fn toC(self: Layout) c.lxw_chart_layout {
        return c.lxw_chart_layout{
            .x = self.x,
            .y = self.y,
            .width = self.width,
            .height = self.height,
            .has_inner = @intFromBool(self.has_inner),
        };
    }
};

// lxw_chart_line
pub const Line = struct {
    color: DefinedColors = .default,
    none: bool = false,
    width: f32 = 2.25,
    dash_type: LineDashType = .solid,
    transparency: u8 = 0,

    pub const default = Line{
        .color = .default,
        .none = false,
        .width = 2.25,
        .dash_type = .solid,
        .transparency = 0,
    };

    inline fn toC(self: Line) c.lxw_chart_line {
        return c.lxw_chart_line{
            .color = @intFromEnum(self.color),
            .none = @intFromBool(self.none),
            .width = self.width,
            .dash_type = @intFromEnum(self.dash_type),
            .transparency = self.transparency,
        };
    }
};

// lxw_chart_fill
pub const Fill = struct {
    color: DefinedColors = .default,
    none: bool = false,
    transparency: u8 = 100,

    pub const default = Fill{
        .color = .default,
        .none = false,
        .transparency = 0,
    };

    inline fn toC(self: Fill) c.lxw_chart_fill {
        return c.lxw_chart_fill{
            .color = @intFromEnum(self.color),
            .none = @intFromBool(self.none),
            .transparency = self.transparency,
        };
    }
};

// lxw_chart_pattern
pub const Pattern = struct {
    fg_color: DefinedColors = .default,
    bg_color: DefinedColors = .default,
    type: PatternType = .none,

    pub const default = Pattern{
        .fg_color = .default,
        .bg_color = .default,
        .type = .none,
    };

    inline fn toC(self: Pattern) c.lxw_chart_pattern {
        return c.lxw_chart_pattern{
            .fg_color = @intFromEnum(self.fg_color),
            .bg_color = @intFromEnum(self.bg_color),
            .type = @intFromEnum(self.type),
        };
    }
};

// lxw_chart_marker
pub const Marker = struct {
    type: MarkerType = .automatic,
    size: u8 = 0,
    line: ?*Line = null,
    fill: ?*Fill = null,
    pattern: ?*Pattern = null,

    pub const default = Marker{
        .type = .automatic,
        .size = 0,
        .line = null,
        .fill = null,
        .pattern = null,
    };

    inline fn toC(
        self: Marker,
        line: ?*c.lxw_chart_line,
        fill: ?*c.lxw_chart_fill,
        pattern: ?*c.lxw_chart_pattern,
    ) c.lxw_chart_marker {
        return c.lxw_chart_marker{
            .type = @intFromEnum(self.type),
            .size = self.size,
            .line = @ptrCast(line),
            .fill = @ptrCast(fill),
            .pattern = @ptrCast(pattern),
        };
    }
};

// lxw_chart_point
pub const Point = extern struct {
    line: ?*Line = null,
    fill: ?*Fill = null,
    pattern: ?*Pattern = null,

    pub const default = Marker{
        .line = null,
        .fill = null,
        .pattern = null,
    };

    inline fn toC(
        self: Point,
        line: ?*c.lxw_chart_line,
        fill: ?*c.lxw_chart_fill,
        pattern: ?*c.lxw_chart_pattern,
    ) c.lxw_chart_point {
        _ = self;
        return c.lxw_chart_point{
            .line = @ptrCast(line),
            .fill = @ptrCast(fill),
            .pattern = @ptrCast(pattern),
        };
    }
};

pub const Series = c.lxw_chart_series;

// pub const struct_lxw_chart_title = extern struct {
//     name: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     row: lxw_row_t = @import("std").mem.zeroes(lxw_row_t),
//     col: lxw_col_t = @import("std").mem.zeroes(lxw_col_t),
//     font: [*c]lxw_chart_font = @import("std").mem.zeroes([*c]lxw_chart_font),
//     off: u8 = @import("std").mem.zeroes(u8),
//     is_horizontal: u8 = @import("std").mem.zeroes(u8),
//     ignore_cache: u8 = @import("std").mem.zeroes(u8),
//     has_overlay: u8 = @import("std").mem.zeroes(u8),
//     range: [*c]lxw_series_range = @import("std").mem.zeroes([*c]lxw_series_range),
//     data_point: struct_lxw_series_data_point = @import("std").mem.zeroes(struct_lxw_series_data_point),
//     layout: [*c]lxw_chart_layout = @import("std").mem.zeroes([*c]lxw_chart_layout),
// };

// pub const struct_lxw_series_data_point = extern struct {
//     is_string: u8 = @import("std").mem.zeroes(u8),
//     number: f64 = @import("std").mem.zeroes(f64),
//     string: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     no_data: u8 = @import("std").mem.zeroes(u8),
//     list_pointers: struct_unnamed_9 = @import("std").mem.zeroes(struct_unnamed_9),
// };

// pub const struct_lxw_series_range = extern struct {
//     formula: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     sheetname: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     first_row: lxw_row_t = @import("std").mem.zeroes(lxw_row_t),
//     last_row: lxw_row_t = @import("std").mem.zeroes(lxw_row_t),
//     first_col: lxw_col_t = @import("std").mem.zeroes(lxw_col_t),
//     last_col: lxw_col_t = @import("std").mem.zeroes(lxw_col_t),
//     ignore_cache: u8 = @import("std").mem.zeroes(u8),
//     has_string_cache: u8 = @import("std").mem.zeroes(u8),
//     num_data_points: u16 = @import("std").mem.zeroes(u16),
//     data_cache: [*c]struct_lxw_series_data_points = @import("std").mem.zeroes([*c]struct_lxw_series_data_points),
// };

// pub const struct_lxw_chart_series = extern struct {
//     categories: [*c]lxw_series_range = @import("std").mem.zeroes([*c]lxw_series_range),
//     values: [*c]lxw_series_range = @import("std").mem.zeroes([*c]lxw_series_range),
//     title: lxw_chart_title = @import("std").mem.zeroes(lxw_chart_title),
//     line: [*c]lxw_chart_line = @import("std").mem.zeroes([*c]lxw_chart_line),
//     fill: [*c]lxw_chart_fill = @import("std").mem.zeroes([*c]lxw_chart_fill),
//     pattern: [*c]lxw_chart_pattern = @import("std").mem.zeroes([*c]lxw_chart_pattern),
//     marker: [*c]lxw_chart_marker = @import("std").mem.zeroes([*c]lxw_chart_marker),
//     points: [*c]lxw_chart_point = @import("std").mem.zeroes([*c]lxw_chart_point),
//     data_labels: [*c]lxw_chart_custom_label = @import("std").mem.zeroes([*c]lxw_chart_custom_label),
//     point_count: u16 = @import("std").mem.zeroes(u16),
//     data_label_count: u16 = @import("std").mem.zeroes(u16),
//     smooth: u8 = @import("std").mem.zeroes(u8),
//     invert_if_negative: u8 = @import("std").mem.zeroes(u8),
//     has_labels: u8 = @import("std").mem.zeroes(u8),
//     show_labels_value: u8 = @import("std").mem.zeroes(u8),
//     show_labels_category: u8 = @import("std").mem.zeroes(u8),
//     show_labels_name: u8 = @import("std").mem.zeroes(u8),
//     show_labels_leader: u8 = @import("std").mem.zeroes(u8),
//     show_labels_legend: u8 = @import("std").mem.zeroes(u8),
//     show_labels_percent: u8 = @import("std").mem.zeroes(u8),
//     label_position: u8 = @import("std").mem.zeroes(u8),
//     label_separator: u8 = @import("std").mem.zeroes(u8),
//     default_label_position: u8 = @import("std").mem.zeroes(u8),
//     label_num_format: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     label_font: [*c]lxw_chart_font = @import("std").mem.zeroes([*c]lxw_chart_font),
//     label_line: [*c]lxw_chart_line = @import("std").mem.zeroes([*c]lxw_chart_line),
//     label_fill: [*c]lxw_chart_fill = @import("std").mem.zeroes([*c]lxw_chart_fill),
//     label_pattern: [*c]lxw_chart_pattern = @import("std").mem.zeroes([*c]lxw_chart_pattern),
//     x_error_bars: [*c]lxw_series_error_bars = @import("std").mem.zeroes([*c]lxw_series_error_bars),
//     y_error_bars: [*c]lxw_series_error_bars = @import("std").mem.zeroes([*c]lxw_series_error_bars),
//     has_trendline: u8 = @import("std").mem.zeroes(u8),
//     has_trendline_forecast: u8 = @import("std").mem.zeroes(u8),
//     has_trendline_equation: u8 = @import("std").mem.zeroes(u8),
//     has_trendline_r_squared: u8 = @import("std").mem.zeroes(u8),
//     has_trendline_intercept: u8 = @import("std").mem.zeroes(u8),
//     trendline_type: u8 = @import("std").mem.zeroes(u8),
//     trendline_value: u8 = @import("std").mem.zeroes(u8),
//     trendline_forward: f64 = @import("std").mem.zeroes(f64),
//     trendline_backward: f64 = @import("std").mem.zeroes(f64),
//     trendline_value_type: u8 = @import("std").mem.zeroes(u8),
//     trendline_name: [*c]u8 = @import("std").mem.zeroes([*c]u8),
//     trendline_line: [*c]lxw_chart_line = @import("std").mem.zeroes([*c]lxw_chart_line),
//     trendline_intercept: f64 = @import("std").mem.zeroes(f64),
//     list_pointers: struct_unnamed_10 = @import("std").mem.zeroes(struct_unnamed_10),
// };

pub const Type = enum(u8) {
    area = c.LXW_CHART_AREA,
    area_stacked = c.LXW_CHART_AREA_STACKED,
    area_stacked_percent = c.LXW_CHART_AREA_STACKED_PERCENT,
    bar = c.LXW_CHART_BAR,
    bar_stacked = c.LXW_CHART_BAR_STACKED,
    bar_stacked_percent = c.LXW_CHART_BAR_STACKED_PERCENT,
    column = c.LXW_CHART_COLUMN,
    column_stacked = c.LXW_CHART_COLUMN_STACKED,
    column_stacked_percent = c.LXW_CHART_COLUMN_STACKED_PERCENT,
    doughnut = c.LXW_CHART_DOUGHNUT,
    line = c.LXW_CHART_LINE,
    line_stacked = c.LXW_CHART_LINE_STACKED,
    line_stacked_percent = c.LXW_CHART_LINE_STACKED_PERCENT,
    pie = c.LXW_CHART_PIE,
    scatter = c.LXW_CHART_SCATTER,
    scatter_straight = c.LXW_CHART_SCATTER_STRAIGHT,
    scatter_straight_with_markers = c.LXW_CHART_SCATTER_STRAIGHT_WITH_MARKERS,
    scatter_smooth = c.LXW_CHART_SCATTER_SMOOTH,
    scatter_smooth_with_markers = c.LXW_CHART_SCATTER_SMOOTH_WITH_MARKERS,
    radar = c.LXW_CHART_RADAR,
    radar_with_markers = c.LXW_CHART_RADAR_WITH_MARKERS,
    radar_filled = c.LXW_CHART_RADAR_FILLED,
};

pub const LegendPosition = enum(u8) {
    none = c.LXW_CHART_LEGEND_NONE,
    right = c.LXW_CHART_LEGEND_RIGHT,
    left = c.LXW_CHART_LEGEND_LEFT,
    top = c.LXW_CHART_LEGEND_TOP,
    bottom = c.LXW_CHART_LEGEND_BOTTOM,
    top_right = c.LXW_CHART_LEGEND_TOP_RIGHT,
    overlay_right = c.LXW_CHART_LEGEND_OVERLAY_RIGHT,
    overlay_left = c.LXW_CHART_LEGEND_OVERLAY_LEFT,
    overlay_top_right = c.LXW_CHART_LEGEND_OVERLAY_TOP_RIGHT,
};

pub const LineDashType = enum(u8) {
    solid = c.LXW_CHART_LINE_DASH_SOLID,
    round_dot = c.LXW_CHART_LINE_DASH_ROUND_DOT,
    square_dot = c.LXW_CHART_LINE_DASH_SQUARE_DOT,
    dash = c.LXW_CHART_LINE_DASH_DASH,
    dash_dot = c.LXW_CHART_LINE_DASH_DASH_DOT,
    long_dash = c.LXW_CHART_LINE_DASH_LONG_DASH,
    long_dash_dot = c.LXW_CHART_LINE_DASH_LONG_DASH_DOT,
    long_dash_dot_dot = c.LXW_CHART_LINE_DASH_LONG_DASH_DOT_DOT,
    dot = c.LXW_CHART_LINE_DASH_DOT,
    system_dash_dot = c.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT,
    system_dash_dot_dot = c.LXW_CHART_LINE_DASH_SYSTEM_DASH_DOT_DOT,
};

pub const MarkerType = enum(u8) {
    automatic = c.LXW_CHART_MARKER_AUTOMATIC,
    none = c.LXW_CHART_MARKER_NONE,
    square = c.LXW_CHART_MARKER_SQUARE,
    diamond = c.LXW_CHART_MARKER_DIAMOND,
    triangle = c.LXW_CHART_MARKER_TRIANGLE,
    x = c.LXW_CHART_MARKER_X,
    star = c.LXW_CHART_MARKER_STAR,
    short_dash = c.LXW_CHART_MARKER_SHORT_DASH,
    long_dash = c.LXW_CHART_MARKER_LONG_DASH,
    circle = c.LXW_CHART_MARKER_CIRCLE,
    plus = c.LXW_CHART_MARKER_PLUS,
};

pub const PatternType = enum(u8) {
    none = c.LXW_CHART_PATTERN_NONE,
    percent_5 = c.LXW_CHART_PATTERN_PERCENT_5,
    percent_10 = c.LXW_CHART_PATTERN_PERCENT_10,
    percent_20 = c.LXW_CHART_PATTERN_PERCENT_20,
    percent_25 = c.LXW_CHART_PATTERN_PERCENT_25,
    percent_30 = c.LXW_CHART_PATTERN_PERCENT_30,
    percent_40 = c.LXW_CHART_PATTERN_PERCENT_40,
    percent_50 = c.LXW_CHART_PATTERN_PERCENT_50,
    percent_60 = c.LXW_CHART_PATTERN_PERCENT_60,
    percent_70 = c.LXW_CHART_PATTERN_PERCENT_70,
    percent_75 = c.LXW_CHART_PATTERN_PERCENT_75,
    percent_80 = c.LXW_CHART_PATTERN_PERCENT_80,
    percent_90 = c.LXW_CHART_PATTERN_PERCENT_90,
    downward_diagonal = c.LXW_CHART_PATTERN_LIGHT_DOWNWARD_DIAGONAL,
    upward_diagonal = c.LXW_CHART_PATTERN_LIGHT_UPWARD_DIAGONAL,
    dark_downward_diagonal = c.LXW_CHART_PATTERN_DARK_DOWNWARD_DIAGONAL,
    dark_upward_diagonal = c.LXW_CHART_PATTERN_DARK_UPWARD_DIAGONAL,
    wide_downward_diagonal = c.LXW_CHART_PATTERN_WIDE_DOWNWARD_DIAGONAL,
    wide_upward_diagonal = c.LXW_CHART_PATTERN_WIDE_UPWARD_DIAGONAL,
    light_vertical = c.LXW_CHART_PATTERN_LIGHT_VERTICAL,
    light_horizontal = c.LXW_CHART_PATTERN_LIGHT_HORIZONTAL,
    narrow_vertical = c.LXW_CHART_PATTERN_NARROW_VERTICAL,
    narrow_horizontal = c.LXW_CHART_PATTERN_NARROW_HORIZONTAL,
    dark_vertical = c.LXW_CHART_PATTERN_DARK_VERTICAL,
    dark_horizontal = c.LXW_CHART_PATTERN_DARK_HORIZONTAL,
    dashed_downward_diagonal = c.LXW_CHART_PATTERN_DASHED_DOWNWARD_DIAGONAL,
    dashed_upward_diagonal = c.LXW_CHART_PATTERN_DASHED_UPWARD_DIAGONAL,
    dashed_horizontal = c.LXW_CHART_PATTERN_DASHED_HORIZONTAL,
    dashed_vertical = c.LXW_CHART_PATTERN_DASHED_VERTICAL,
    small_confetti = c.LXW_CHART_PATTERN_SMALL_CONFETTI,
    large_confetti = c.LXW_CHART_PATTERN_LARGE_CONFETTI,
    zigzag = c.LXW_CHART_PATTERN_ZIGZAG,
    wave = c.LXW_CHART_PATTERN_WAVE,
    diagonal_brick = c.LXW_CHART_PATTERN_DIAGONAL_BRICK,
    horizontal_brick = c.LXW_CHART_PATTERN_HORIZONTAL_BRICK,
    weave = c.LXW_CHART_PATTERN_WEAVE,
    plaid = c.LXW_CHART_PATTERN_PLAID,
    divot = c.LXW_CHART_PATTERN_DIVOT,
    dotted_grid = c.LXW_CHART_PATTERN_DOTTED_GRID,
    dotted_diamond = c.LXW_CHART_PATTERN_DOTTED_DIAMOND,
    shingle = c.LXW_CHART_PATTERN_SHINGLE,
    sphere = c.LXW_CHART_PATTERN_SPHERE,
    trellis = c.LXW_CHART_PATTERN_TRELLIS,
    small_grid = c.LXW_CHART_PATTERN_SMALL_GRID,
    large_grid = c.LXW_CHART_PATTERN_LARGE_GRID,
    small_check = c.LXW_CHART_PATTERN_SMALL_CHECK,
    large_check = c.LXW_CHART_PATTERN_LARGE_CHECK,
    outlined_diamond = c.LXW_CHART_PATTERN_OUTLINED_DIAMOND,
    solid_diamond = c.LXW_CHART_PATTERN_SOLID_DIAMOND,
};

pub const LabelPosition = enum(u8) {
    default = c.LXW_CHART_LABEL_POSITION_DEFAULT,
    center = c.LXW_CHART_LABEL_POSITION_CENTER,
    right = c.LXW_CHART_LABEL_POSITION_RIGHT,
    left = c.LXW_CHART_LABEL_POSITION_LEFT,
    above = c.LXW_CHART_LABEL_POSITION_ABOVE,
    below = c.LXW_CHART_LABEL_POSITION_BELOW,
    inside_base = c.LXW_CHART_LABEL_POSITION_INSIDE_BASE,
    inside_end = c.LXW_CHART_LABEL_POSITION_INSIDE_END,
    outside_end = c.LXW_CHART_LABEL_POSITION_OUTSIDE_END,
    best_fit = c.LXW_CHART_LABEL_POSITION_BEST_FIT,
};

pub const LabelSeparator = enum(u8) {
    comma = c.LXW_CHART_LABEL_SEPARATOR_COMMA,
    semicolon = c.LXW_CHART_LABEL_SEPARATOR_SEMICOLON,
    period = c.LXW_CHART_LABEL_SEPARATOR_PERIOD,
    newline = c.LXW_CHART_LABEL_SEPARATOR_NEWLINE,
    space = c.LXW_CHART_LABEL_SEPARATOR_SPACE,
};

pub const AxisType = enum(u8) {
    x = c.LXW_CHART_AXIS_TYPE_X,
    y = c.LXW_CHART_AXIS_TYPE_Y,
};

pub const SubType = enum(u8) {
    none = c.LXW_CHART_SUBTYPE_NONE,
    stacked = c.LXW_CHART_SUBTYPE_STACKED,
    stacked_percent = c.LXW_CHART_SUBTYPE_STACKED_PERCENT,
};

pub const Grouping = enum(u8) {
    clustered = c.LXW_GROUPING_CLUSTERED,
    standard = c.LXW_GROUPING_STANDARD,
    percent_stacked = c.LXW_GROUPING_PERCENTSTACKED,
    stacked = c.LXW_GROUPING_STACKED,
};

pub const AxisTickPosition = enum(u8) {
    default = c.LXW_CHART_AXIS_POSITION_DEFAULT,
    on_tick = c.LXW_CHART_AXIS_POSITION_ON_TICK,
    between = c.LXW_CHART_AXIS_POSITION_BETWEEN,
};

pub const AxisLabelPosition = enum(u8) {
    next_to = c.LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO,
    high = c.LXW_CHART_AXIS_LABEL_POSITION_HIGH,
    low = c.LXW_CHART_AXIS_LABEL_POSITION_LOW,
    none = c.LXW_CHART_AXIS_LABEL_POSITION_NONE,
};

pub const AxisLabelAlignment = enum(u8) {
    center = c.LXW_CHART_AXIS_LABEL_ALIGN_CENTER,
    left = c.LXW_CHART_AXIS_LABEL_ALIGN_LEFT,
    right = c.LXW_CHART_AXIS_LABEL_ALIGN_RIGHT,
};

pub const AxisDisplayUnit = enum(u8) {
    none = c.LXW_CHART_AXIS_UNITS_NONE,
    hundreds = c.LXW_CHART_AXIS_UNITS_HUNDREDS,
    thousands = c.LXW_CHART_AXIS_UNITS_THOUSANDS,
    ten_thousands = c.LXW_CHART_AXIS_UNITS_TEN_THOUSANDS,
    hundred_thousands = c.LXW_CHART_AXIS_UNITS_HUNDRED_THOUSANDS,
    millions = c.LXW_CHART_AXIS_UNITS_MILLIONS,
    ten_millions = c.LXW_CHART_AXIS_UNITS_TEN_MILLIONS,
    hundred_millions = c.LXW_CHART_AXIS_UNITS_HUNDRED_MILLIONS,
    billions = c.LXW_CHART_AXIS_UNITS_BILLIONS,
    trillions = c.LXW_CHART_AXIS_UNITS_TRILLIONS,
};

pub const TickMark = enum(u8) {
    default = c.LXW_CHART_AXIS_TICK_MARK_DEFAULT,
    none = c.LXW_CHART_AXIS_TICK_MARK_NONE,
    inside = c.LXW_CHART_AXIS_TICK_MARK_INSIDE,
    outside = c.LXW_CHART_AXIS_TICK_MARK_OUTSIDE,
    crossing = c.LXW_CHART_AXIS_TICK_MARK_CROSSING,
};

pub inline fn addSeries(self: Chart, categories: ?CString, values: ?CString) XlsxError!*Series {
    return c.chart_add_series(self.chart_c, categories, values) orelse XlsxError.ChartAddSeries;
}

pub inline fn titleSetName(self: Chart, name: ?CString) void {
    c.chart_title_set_name(self.chart_c, name);
}

pub inline fn titleSetNameFont(self: Chart, font: Font) void {
    c.chart_title_set_name_font(
        self.chart_c,
        @constCast(&font.toC()),
    );
}

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CString = xlsxwriter.CString;
const CStringArray = xlsxwriter.CStringArray;
const Bool = xlsxwriter.Boolean;
const DefinedColors = @import("format.zig").DefinedColors;
const XlsxError = @import("errors.zig").XlsxError;

const ChartSeries = @This();

// alloc: ?std.mem.Allocator,
chartseries_c: ?*c.lxw_chart_series,

// pub extern fn chart_series_set_categories(series: [*c]lxw_chart_series, sheetname: [*c]const u8, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) void;
pub inline fn setCategories(
    self: ChartSeries,
    sheetname: [:0]const u8,
    range: Range,
) void {
    c.chart_series_set_categories(
        self.chartseries_c,
        sheetname,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
    );
}

// pub extern fn chart_series_set_values(series: [*c]lxw_chart_series, sheetname: [*c]const u8, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) void;
pub inline fn setValues(
    self: ChartSeries,
    sheetname: [:0]const u8,
    range: Range,
) void {
    c.chart_series_set_values(
        self.chartseries_c,
        sheetname,
        range.first_row,
        range.first_col,
        range.last_row,
        range.last_col,
    );
}

// pub extern fn chart_series_set_name(series: [*c]lxw_chart_series, name: [*c]const u8) void;
pub inline fn setName(
    self: ChartSeries,
    name: [:0]const u8,
) void {
    c.chart_series_set_name(
        self.chartseries_c,
        name,
    );
}

// pub extern fn chart_series_set_name_range(series: [*c]lxw_chart_series, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setNameRange(
    self: ChartSeries,
    sheetname: [:0]const u8,
    cell: Cell,
) void {
    c.chart_series_set_name_range(
        self.chartseries_c,
        sheetname,
        cell.row,
        cell.col,
    );
}

// pub extern fn chart_series_set_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setLine(
    self: ChartSeries,
    line: ChartLine,
) void {
    c.chart_series_set_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setFill(
    self: ChartSeries,
    fill: ChartFill,
) void {
    c.chart_series_set_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_invert_if_negative(series: [*c]lxw_chart_series) void;
pub inline fn setInvertIfNegative(self: ChartSeries) void {
    c.chart_series_set_invert_if_negative(self.chartseries_c);
}

// pub extern fn chart_series_set_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setPattern(
    self: ChartSeries,
    pattern: ChartPattern,
) void {
    c.chart_series_set_pattern(
        self.chartseries_c,
        @intFromEnum(pattern),
    );
}

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

// pub extern fn chart_series_set_marker_type(series: [*c]lxw_chart_series, @"type": u8) void;
pub inline fn setMarkerType(
    self: ChartSeries,
    @"type": MarkerType,
) void {
    c.chart_series_set_marker_type(
        self.chartseries_c,
        @intFromEnum(@"type"),
    );
}

// pub extern fn chart_series_set_marker_size(series: [*c]lxw_chart_series, size: u8) void;
pub inline fn setMarkerSize(
    self: ChartSeries,
    size: u8,
) void {
    c.chart_series_set_marker_size(
        self.chartseries_c,
        size,
    );
}

// pub extern fn chart_series_set_marker_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setMarkerLine(
    self: ChartSeries,
    line: ChartLine,
) void {
    c.chart_series_set_marker_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_marker_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setMarkerFill(
    self: ChartSeries,
    fill: ChartFill,
) void {
    c.chart_series_set_marker_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_marker_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setMarkerPattern(
    self: ChartSeries,
    pattern: ChartPattern,
) void {
    c.chart_series_set_marker_pattern(
        self.chartseries_c,
        @intFromEnum(pattern),
    );
}

// lxw_chart_point
pub const Point = extern struct {
    line: ?*ChartLine = null,
    fill: ?*ChartFill = null,
    pattern: ?*ChartPattern = null,

    pub const default = Point{
        .line = null,
        .fill = null,
        .pattern = null,
    };

    inline fn toC(
        self: Point,
    ) c.lxw_chart_point {
        return c.lxw_chart_point{
            .line = @ptrCast(self.line),
            .fill = @ptrCast(self.fill),
            .pattern = @ptrCast(self.pattern),
        };
    }
};
pub const ChartPointArray = [:null]Point;
// pub extern fn chart_series_set_points(series: [*c]lxw_chart_series, points: [*c][*c]lxw_chart_point) lxw_error;
pub inline fn setPoints(self: ChartSeries, points: ChartPointArray) void {
    c.chart_series_set_points(
        self.chartseries_c,
        @ptrCast(@constCast(points)),
    );
}

// pub extern fn chart_series_set_smooth(series: [*c]lxw_chart_series, smooth: u8) void;
pub inline fn setSmooth(self: ChartSeries, smooth: bool) void {
    c.chart_series_set_smooth(
        self.chartseries_c,
        @intFromBool(smooth),
    );
}

// pub extern fn chart_series_set_labels(series: [*c]lxw_chart_series) void;
pub inline fn setLabels(self: ChartSeries) void {
    c.chart_series_set_labels(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_options(series: [*c]lxw_chart_series, show_name: u8, show_category: u8, show_value: u8) void;
pub inline fn setLabelsOptions(
    self: ChartSeries,
    show_name: bool,
    show_category: bool,
    show_value: bool,
) void {
    c.chart_series_set_labels_options(
        self.chartseries_c,
        @intFromBool(show_name),
        @intFromBool(show_category),
        @intFromBool(show_value),
    );
}

pub const ChartDataLabel = struct {
    value: ?[:0]const u8 = null,
    hide: bool = false,
    font: ?*ChartFont = null,
    line: ?*ChartLine = null,
    fill: ?*ChartFill = null,
    pattern: ?*ChartPattern = .none,

    pub const default = ChartDataLabel{
        .value = null,
        .hide = false,
        .font = null,
        .line = null,
        .fill = null,
        .pattern = .none,
    };

    inline fn toC(self: ChartDataLabel) c.lxw_chart_data_label {
        return c.lxw_chart_data_label{
            .value = self.value,
            .hide = @intFromBool(self.hide),
            .font = @ptrCast(self.font),
            .line = @ptrCast(self.line),
            .fill = @ptrCast(self.fill),
            .pattern = @intFromEnum(self.pattern),
        };
    }
};
pub const ChartDataLabelArray = [:null]ChartDataLabel;
// pub extern fn chart_series_set_labels_custom(series: [*c]lxw_chart_series, data_labels: [*c][*c]lxw_chart_data_label) lxw_error;
pub inline fn setLabelsCustom(
    self: ChartSeries,
    data_labels: ChartDataLabelArray,
) void {
    c.chart_series_set_labels_custom(
        self.chartseries_c,
        @ptrCast(@constCast(data_labels)),
    );
}

pub const LabelSeparator = enum(u8) {
    comma = c.LXW_CHART_LABEL_SEPARATOR_COMMA,
    semicolon = c.LXW_CHART_LABEL_SEPARATOR_SEMICOLON,
    period = c.LXW_CHART_LABEL_SEPARATOR_PERIOD,
    newline = c.LXW_CHART_LABEL_SEPARATOR_NEWLINE,
    space = c.LXW_CHART_LABEL_SEPARATOR_SPACE,
};

// pub extern fn chart_series_set_labels_separator(series: [*c]lxw_chart_series, separator: u8) void;
pub inline fn setLabelsSeparator(
    self: ChartSeries,
    separator: LabelSeparator,
) void {
    c.chart_series_set_labels_separator(
        self.chartseries_c,
        @intFromEnum(separator),
    );
}

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

// pub extern fn chart_series_set_labels_position(series: [*c]lxw_chart_series, position: u8) void;
pub inline fn setLabelsPosition(
    self: ChartSeries,
    position: LabelPosition,
) void {
    c.chart_series_set_labels_position(
        self.chartseries_c,
        @intFromEnum(position),
    );
}

// pub extern fn chart_series_set_labels_leader_line(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsLeaderLine(self: ChartSeries) void {
    c.chart_series_set_labels_leader_line(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_legend(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsLegend(self: ChartSeries) void {
    c.chart_series_set_labels_legend(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_percentage(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsPercentage(self: ChartSeries) void {
    c.chart_series_set_labels_percentage(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_num_format(series: [*c]lxw_chart_series, num_format: [*c]const u8) void;
pub inline fn setLabelsNumFormat(
    self: ChartSeries,
    num_format: [:0]const u8,
) void {
    c.chart_series_set_labels_position(
        self.chartseries_c,
        num_format,
    );
}

// pub extern fn chart_series_set_labels_font(series: [*c]lxw_chart_series, font: [*c]lxw_chart_font) void;
pub inline fn setLabelsFont(self: ChartSeries, font: ChartFont) void {
    c.chart_series_set_labels_font(
        self.chartseries_c,
        @constCast(&font.toC()),
    );
}

// pub extern fn chart_series_set_labels_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setLabelsLine(
    self: ChartSeries,
    line: ChartLine,
) void {
    c.chart_series_set_labels_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_labels_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setLabelsFill(
    self: ChartSeries,
    fill: ChartFill,
) void {
    c.chart_series_set_labels_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_labels_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setLabelsPattern(
    self: ChartSeries,
    pattern: ChartPattern,
) void {
    c.chart_series_set_labels_pattern(
        self.chartseries_c,
        @intFromEnum(pattern),
    );
}

pub const TrendlineType = enum(u8) {
    linear = c.LXW_CHART_TRENDLINE_TYPE_LINEAR,
    log = c.LXW_CHART_TRENDLINE_TYPE_LOG,
    poly = c.LXW_CHART_TRENDLINE_TYPE_POLY,
    power = c.LXW_CHART_TRENDLINE_TYPE_POWER,
    exp = c.LXW_CHART_TRENDLINE_TYPE_EXP,
    average = c.LXW_CHART_TRENDLINE_TYPE_AVERAGE,
};

// pub extern fn chart_series_set_trendline(series: [*c]lxw_chart_series, @"type": u8, value: u8) void;
pub inline fn setTrendline(
    self: ChartSeries,
    @"type": TrendlineType,
    value: u8,
) void {
    c.chart_series_set_trendline(
        self.chartseries_c,
        @intFromEnum(@"type"),
        value,
    );
}

// pub extern fn chart_series_set_trendline_forecast(series: [*c]lxw_chart_series, forward: f64, backward: f64) void;
pub inline fn setTrendlineForecast(
    self: ChartSeries,
    forward: f64,
    backward: f64,
) void {
    c.chart_series_set_trendline_forecast(
        self.chartseries_c,
        forward,
        backward,
    );
}

// pub extern fn chart_series_set_trendline_equation(series: [*c]lxw_chart_series) void;
pub inline fn setTrendlineEquation(self: ChartSeries) void {
    c.chart_series_set_trendline_equation(self.chartseries_c);
}

// pub extern fn chart_series_set_trendline_r_squared(series: [*c]lxw_chart_series) void;
pub inline fn setTrendlineRSquared(self: ChartSeries) void {
    c.chart_series_set_trendline_r_squared(self.chartseries_c);
}

// pub extern fn chart_series_set_trendline_intercept(series: [*c]lxw_chart_series, intercept: f64) void;
pub inline fn setTrendlineIntercept(
    self: ChartSeries,
    intercept: f64,
) void {
    c.chart_series_set_trendline_intercept(
        self.chartseries_c,
        intercept,
    );
}

// pub extern fn chart_series_set_trendline_name(series: [*c]lxw_chart_series, name: [*c]const u8) void;
pub inline fn setTrendlineName(
    self: ChartSeries,
    name: [:0]const u8,
) void {
    c.chart_series_set_trendline_intercept(
        self.chartseries_c,
        name,
    );
}

// pub extern fn chart_series_set_trendline_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setTrendlineLine(
    self: ChartSeries,
    line: ChartLine,
) void {
    c.chart_series_set_trendline_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_get_error_bars(series: [*c]lxw_chart_series, axis_type: lxw_chart_error_bar_axis) [*c]lxw_series_error_bars;
pub inline fn getErrorBars(
    self: ChartSeries,
    axis_type: ErrorBarsAxis,
) XlsxError!ErrorBars {
    return ErrorBars{
        .error_bars_c = c.chart_series_get_error_bars(
            self.chartseries_c,
            @intFromEnum(axis_type),
        ) orelse return XlsxError.GetErrorBars,
    };
}

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Cell = @import("utility.zig").Cell;
const Range = @import("utility.zig").Range;
const ChartFont = @import("Chart.zig").Font;
const ChartLine = @import("Chart.zig").Line;
const ChartFill = @import("Chart.zig").Fill;
const ChartPattern = @import("Chart.zig").PatternType;
const ErrorBars = @import("ErrorBars.zig");
const ErrorBarsAxis = ErrorBars.Axis;

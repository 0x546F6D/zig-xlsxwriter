const Series = @This();

alloc: ?std.mem.Allocator,
chartseries_c: ?*c.lxw_chart_series,
x_error_bars: ErrorBars,
y_error_bars: ErrorBars,

// pub extern fn chart_series_set_categories(series: [*c]lxw_chart_series, sheetname: [*c]const u8, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) void;
pub inline fn setCategories(
    self: Series,
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
    self: Series,
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
    self: Series,
    name: [:0]const u8,
) void {
    c.chart_series_set_name(
        self.chartseries_c,
        name,
    );
}

// pub extern fn chart_series_set_name_range(series: [*c]lxw_chart_series, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setNameRange(
    self: Series,
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
    self: Series,
    line: ChartLine,
) void {
    c.chart_series_set_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setFill(
    self: Series,
    fill: ChartFill,
) void {
    c.chart_series_set_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_invert_if_negative(series: [*c]lxw_chart_series) void;
pub inline fn setInvertIfNegative(self: Series) void {
    c.chart_series_set_invert_if_negative(self.chartseries_c);
}

// pub extern fn chart_series_set_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setPattern(
    self: Series,
    pattern: ChartPattern,
) void {
    c.chart_series_set_pattern(
        self.chartseries_c,
        @constCast(&pattern.toC()),
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
    self: Series,
    @"type": MarkerType,
) void {
    c.chart_series_set_marker_type(
        self.chartseries_c,
        @intFromEnum(@"type"),
    );
}

// pub extern fn chart_series_set_marker_size(series: [*c]lxw_chart_series, size: u8) void;
pub inline fn setMarkerSize(
    self: Series,
    size: u8,
) void {
    c.chart_series_set_marker_size(
        self.chartseries_c,
        size,
    );
}

// pub extern fn chart_series_set_marker_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setMarkerLine(
    self: Series,
    line: ChartLine,
) void {
    c.chart_series_set_marker_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_marker_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setMarkerFill(
    self: Series,
    fill: ChartFill,
) void {
    c.chart_series_set_marker_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_marker_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setMarkerPattern(
    self: Series,
    pattern: ChartPattern,
) void {
    c.chart_series_set_marker_pattern(
        self.chartseries_c,
        @constCast(&pattern.toC()),
    );
}

// lxw_chart_point
pub const Point = extern struct {
    line: ?*const ChartLine = null,
    fill: ?*const ChartFill = null,
    pattern: ?*const ChartPattern = null,

    pub const default = Point{
        .line = null,
        .fill = null,
        .pattern = null,
    };
};

// pub extern fn chart_series_set_points(series: [*c]lxw_chart_series, points: [*c][*c]lxw_chart_point) lxw_error;
pub inline fn setPoints(
    self: Series,
    points: []const Point,
) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.SetPointsNoAlloc;

    var line_array = try allocator.alloc(c.lxw_chart_line, points.len);
    defer allocator.free(line_array);
    var fill_array = try allocator.alloc(c.lxw_chart_fill, points.len);
    defer allocator.free(fill_array);
    var pattern_array = try allocator.alloc(c.lxw_chart_pattern, points.len);
    defer allocator.free(pattern_array);
    var point_array = try allocator.alloc(c.lxw_chart_point, points.len);
    defer allocator.free(point_array);
    var point_c = try allocator.allocSentinel(?*c.lxw_chart_point, points.len, null);
    defer allocator.free(point_c);

    for (points, 0..) |point, i| {
        if (point.line) |line| {
            line_array[i] = line.toC();
            point_array[i].line = &line_array[i];
        } else point_array[i].line = null;

        if (point.fill) |fill| {
            fill_array[i] = fill.toC();
            point_array[i].fill = &fill_array[i];
        } else point_array[i].fill = null;

        if (point.pattern) |pattern| {
            pattern_array[i] = pattern.toC();
            point_array[i].pattern = &pattern_array[i];
        } else point_array[i].pattern = null;

        point_c[i] = &point_array[i];
    }

    try check(c.chart_series_set_points(
        self.chartseries_c,
        @ptrCast(point_c.ptr),
    ));
}

pub const PointNoAlloc = c.lxw_chart_point;
pub const PointNoAllocArray = [:null]const ?*const PointNoAlloc;
pub inline fn setPointsNoAlloc(self: Series, points: PointNoAllocArray) !void {
    try check(c.chart_series_set_points(
        self.chartseries_c,
        @ptrCast(@constCast(points)),
    ));
}

// pub extern fn chart_series_set_smooth(series: [*c]lxw_chart_series, smooth: u8) void;
pub inline fn setSmooth(self: Series, smooth: bool) void {
    c.chart_series_set_smooth(
        self.chartseries_c,
        @intFromBool(smooth),
    );
}

// pub extern fn chart_series_set_labels(series: [*c]lxw_chart_series) void;
pub inline fn setLabels(self: Series) void {
    c.chart_series_set_labels(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_options(series: [*c]lxw_chart_series, show_name: u8, show_category: u8, show_value: u8) void;
pub inline fn setLabelsOptions(
    self: Series,
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

pub const DataLabel = struct {
    value: ?[*:0]const u8 = null,
    hide: bool = false,
    font: ?*const ChartFont = null,
    line: ?*const ChartLine = null,
    fill: ?*const ChartFill = null,
    pattern: ?*const ChartPattern = null,

    pub const default = DataLabel{
        .value = null,
        .hide = false,
        .font = null,
        .line = null,
        .fill = null,
        .pattern = null,
    };
};

// pub extern fn chart_series_set_labels_custom(series: [*c]lxw_chart_series, data_labels: [*c][*c]lxw_chart_data_label) lxw_error;
pub inline fn setLabelsCustom(self: Series, data_labels: []const DataLabel) !void {
    const allocator: std.mem.Allocator = if (self.alloc) |allocator|
        allocator
    else
        return XlsxError.SetLabelsCustomNoAlloc;

    var font_array = try allocator.alloc(c.lxw_chart_font, data_labels.len);
    defer allocator.free(font_array);
    var line_array = try allocator.alloc(c.lxw_chart_line, data_labels.len);
    defer allocator.free(line_array);
    var fill_array = try allocator.alloc(c.lxw_chart_fill, data_labels.len);
    defer allocator.free(fill_array);
    var pattern_array = try allocator.alloc(c.lxw_chart_pattern, data_labels.len);
    defer allocator.free(pattern_array);
    var label_array = try allocator.alloc(c.lxw_chart_data_label, data_labels.len);
    defer allocator.free(label_array);
    var label_c = try allocator.allocSentinel(?*c.lxw_chart_data_label, data_labels.len, null);
    defer allocator.free(label_c);

    for (data_labels, 0..) |label, i| {
        label_array[i].value = label.value;
        label_array[i].hide = @intFromBool(label.hide);

        if (label.font) |font| {
            font_array[i] = font.toC(false);
            label_array[i].font = &font_array[i];
        } else label_array[i].font = null;

        if (label.line) |line| {
            line_array[i] = line.toC();
            label_array[i].line = &line_array[i];
        } else label_array[i].line = null;

        if (label.fill) |fill| {
            fill_array[i] = fill.toC();
            label_array[i].fill = &fill_array[i];
        } else label_array[i].fill = null;

        if (label.pattern) |pattern| {
            pattern_array[i] = pattern.toC();
            label_array[i].pattern = &pattern_array[i];
        } else label_array[i].pattern = null;

        label_c[i] = &label_array[i];
    }
    try check(c.chart_series_set_labels_custom(
        self.chartseries_c,
        @ptrCast(label_c.ptr),
    ));
}

pub const DataLabelNoAlloc = c.lxw_chart_data_label;
pub const DataLabelNoAllocArray = [:null]const ?*const DataLabelNoAlloc;
pub inline fn setLabelsCustomNoAlloc(self: Series, data_labels: DataLabelNoAllocArray) !void {
    try check(c.chart_series_set_labels_custom(
        self.chartseries_c,
        @ptrCast(@constCast(data_labels)),
    ));
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
    self: Series,
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
    self: Series,
    position: LabelPosition,
) void {
    c.chart_series_set_labels_position(
        self.chartseries_c,
        @intFromEnum(position),
    );
}

// pub extern fn chart_series_set_labels_leader_line(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsLeaderLine(self: Series) void {
    c.chart_series_set_labels_leader_line(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_legend(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsLegend(self: Series) void {
    c.chart_series_set_labels_legend(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_percentage(series: [*c]lxw_chart_series) void;
pub inline fn setLabelsPercentage(self: Series) void {
    c.chart_series_set_labels_percentage(self.chartseries_c);
}

// pub extern fn chart_series_set_labels_num_format(series: [*c]lxw_chart_series, num_format: [*c]const u8) void;
pub inline fn setLabelsNumFormat(
    self: Series,
    num_format: [:0]const u8,
) void {
    c.chart_series_set_labels_position(
        self.chartseries_c,
        num_format,
    );
}

// pub extern fn chart_series_set_labels_font(series: [*c]lxw_chart_series, font: [*c]lxw_chart_font) void;
pub inline fn setLabelsFont(self: Series, font: ChartFont) void {
    c.chart_series_set_labels_font(
        self.chartseries_c,
        @constCast(&font.toC(false)),
    );
}

// pub extern fn chart_series_set_labels_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setLabelsLine(
    self: Series,
    line: ChartLine,
) void {
    c.chart_series_set_labels_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_set_labels_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
pub inline fn setLabelsFill(
    self: Series,
    fill: ChartFill,
) void {
    c.chart_series_set_labels_fill(
        self.chartseries_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_series_set_labels_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setLabelsPattern(
    self: Series,
    pattern: ChartPattern,
) void {
    c.chart_series_set_labels_pattern(
        self.chartseries_c,
        @constCast(&pattern.toC()),
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
    self: Series,
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
    self: Series,
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
pub inline fn setTrendlineEquation(self: Series) void {
    c.chart_series_set_trendline_equation(self.chartseries_c);
}

// pub extern fn chart_series_set_trendline_r_squared(series: [*c]lxw_chart_series) void;
pub inline fn setTrendlineRSquared(self: Series) void {
    c.chart_series_set_trendline_r_squared(self.chartseries_c);
}

// pub extern fn chart_series_set_trendline_intercept(series: [*c]lxw_chart_series, intercept: f64) void;
pub inline fn setTrendlineIntercept(
    self: Series,
    intercept: f64,
) void {
    c.chart_series_set_trendline_intercept(
        self.chartseries_c,
        intercept,
    );
}

// pub extern fn chart_series_set_trendline_name(series: [*c]lxw_chart_series, name: [*c]const u8) void;
pub inline fn setTrendlineName(
    self: Series,
    name: [:0]const u8,
) void {
    c.chart_series_set_trendline_intercept(
        self.chartseries_c,
        name,
    );
}

// pub extern fn chart_series_set_trendline_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
pub inline fn setTrendlineLine(
    self: Series,
    line: ChartLine,
) void {
    c.chart_series_set_trendline_line(
        self.chartseries_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_series_get_error_bars(series: [*c]lxw_chart_series, axis_type: lxw_chart_error_bar_axis) [*c]lxw_series_error_bars;
pub inline fn getErrorBars(
    self: Series,
    axis_type: ErrorBarsAxis,
) XlsxError!ErrorBars {
    return ErrorBars{
        .error_bars_c = c.chart_series_get_error_bars(
            self.chartseries_c,
            @intFromEnum(axis_type),
        ) orelse return XlsxError.GetErrorBars,
    };
}

const std = @import("std");
const c = @import("lxw");
const xlsxwriter = @import("../xlsxwriter.zig");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const Cell = @import("../utility.zig").Cell;
const Range = @import("../utility.zig").Range;
const Chart = @import("../Chart.zig");
const ChartFont = Chart.Font;
const ChartLayout = Chart.Layout;
const ChartLine = Chart.Line;
const ChartFill = Chart.Fill;
const ChartPattern = Chart.Pattern;
const ErrorBars = @import("ErrorBars.zig");
const ErrorBarsAxis = ErrorBars.Axis;

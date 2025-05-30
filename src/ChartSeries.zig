const ChartSeries = @This();

// alloc: ?std.mem.Allocator,
chartseries_c: ?*c.lxw_chart_series,

// pub extern fn chart_series_set_categories(series: [*c]lxw_chart_series, sheetname: [*c]const u8, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) void;
pub inline fn setCategories(
    self: ChartSeries,
    sheetname: ?CString,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
) void {
    c.chart_series_set_categories(
        self.chartseries_c,
        sheetname,
        first_row,
        first_col,
        last_row,
        last_col,
    );
}

// pub extern fn chart_series_set_values(series: [*c]lxw_chart_series, sheetname: [*c]const u8, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) void;
pub inline fn setValues(
    self: ChartSeries,
    sheetname: ?CString,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
) void {
    c.chart_series_set_values(
        self.chartseries_c,
        sheetname,
        first_row,
        first_col,
        last_row,
        last_col,
    );
}

// pub extern fn chart_series_set_name(series: [*c]lxw_chart_series, name: [*c]const u8) void;
pub inline fn setName(
    self: ChartSeries,
    name: ?CString,
) void {
    c.chart_series_set_name(
        self.chartseries_c,
        name,
    );
}

// pub extern fn chart_series_set_name_range(series: [*c]lxw_chart_series, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setNameRange(
    self: ChartSeries,
    sheetname: ?CString,
    row: u32,
    col: u16,
) void {
    c.chart_series_set_name_range(
        self.chartseries_c,
        sheetname,
        row,
        col,
    );
}

// pub extern fn chart_series_set_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
// pub extern fn chart_series_set_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
// pub extern fn chart_series_set_invert_if_negative(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
// pub extern fn chart_series_set_marker_type(series: [*c]lxw_chart_series, @"type": u8) void;
// pub extern fn chart_series_set_marker_size(series: [*c]lxw_chart_series, size: u8) void;
// pub extern fn chart_series_set_marker_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
// pub extern fn chart_series_set_marker_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
// pub extern fn chart_series_set_marker_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
// pub extern fn chart_series_set_points(series: [*c]lxw_chart_series, points: [*c][*c]lxw_chart_point) lxw_error;
// pub extern fn chart_series_set_smooth(series: [*c]lxw_chart_series, smooth: u8) void;
// pub extern fn chart_series_set_labels(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_labels_options(series: [*c]lxw_chart_series, show_name: u8, show_category: u8, show_value: u8) void;
// pub extern fn chart_series_set_labels_custom(series: [*c]lxw_chart_series, data_labels: [*c][*c]lxw_chart_data_label) lxw_error;
// pub extern fn chart_series_set_labels_separator(series: [*c]lxw_chart_series, separator: u8) void;
// pub extern fn chart_series_set_labels_position(series: [*c]lxw_chart_series, position: u8) void;
// pub extern fn chart_series_set_labels_leader_line(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_labels_legend(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_labels_percentage(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_labels_num_format(series: [*c]lxw_chart_series, num_format: [*c]const u8) void;
// pub extern fn chart_series_set_labels_font(series: [*c]lxw_chart_series, font: [*c]lxw_chart_font) void;
// pub extern fn chart_series_set_labels_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
// pub extern fn chart_series_set_labels_fill(series: [*c]lxw_chart_series, fill: [*c]lxw_chart_fill) void;
// pub extern fn chart_series_set_labels_pattern(series: [*c]lxw_chart_series, pattern: [*c]lxw_chart_pattern) void;
// pub extern fn chart_series_set_trendline(series: [*c]lxw_chart_series, @"type": u8, value: u8) void;
// pub extern fn chart_series_set_trendline_forecast(series: [*c]lxw_chart_series, forward: f64, backward: f64) void;
// pub extern fn chart_series_set_trendline_equation(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_trendline_r_squared(series: [*c]lxw_chart_series) void;
// pub extern fn chart_series_set_trendline_intercept(series: [*c]lxw_chart_series, intercept: f64) void;
// pub extern fn chart_series_set_trendline_name(series: [*c]lxw_chart_series, name: [*c]const u8) void;
// pub extern fn chart_series_set_trendline_line(series: [*c]lxw_chart_series, line: [*c]lxw_chart_line) void;
// pub extern fn chart_series_get_error_bars(series: [*c]lxw_chart_series, axis_type: lxw_chart_error_bar_axis) [*c]lxw_series_error_bars;
// pub extern fn chart_series_set_error_bars(error_bars: [*c]lxw_series_error_bars, @"type": u8, value: f64) void;
// pub extern fn chart_series_set_error_bars_direction(error_bars: [*c]lxw_series_error_bars, direction: u8) void;
// pub extern fn chart_series_set_error_bars_endcap(error_bars: [*c]lxw_series_error_bars, endcap: u8) void;
// pub extern fn chart_series_set_error_bars_line(error_bars: [*c]lxw_series_error_bars, line: [*c]lxw_chart_line) void;

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CString = xlsxwriter.CString;
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;

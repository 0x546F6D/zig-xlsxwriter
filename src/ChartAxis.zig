const ChartAxis = @This();

axis_c: ?*c.lxw_chart_axis,

// pub extern fn chart_axis_set_name(axis: [*c]lxw_chart_axis, name: [*c]const u8) void;
pub inline fn setName(
    self: ChartAxis,
    name: ?CString,
) void {
    c.chart_axis_set_name(
        self.axis_c,
        name,
    );
}

// pub extern fn chart_axis_get(chart: [*c]lxw_chart, axis_type: lxw_chart_axis_type) [*c]lxw_chart_axis;
// pub extern fn chart_axis_set_name_range(axis: [*c]lxw_chart_axis, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
// pub extern fn chart_axis_set_name_layout(axis: [*c]lxw_chart_axis, layout: [*c]lxw_chart_layout) void;
// pub extern fn chart_axis_set_name_font(axis: [*c]lxw_chart_axis, font: [*c]lxw_chart_font) void;
// pub extern fn chart_axis_set_num_font(axis: [*c]lxw_chart_axis, font: [*c]lxw_chart_font) void;
// pub extern fn chart_axis_set_num_format(axis: [*c]lxw_chart_axis, num_format: [*c]const u8) void;
// pub extern fn chart_axis_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;
// pub extern fn chart_axis_set_fill(axis: [*c]lxw_chart_axis, fill: [*c]lxw_chart_fill) void;
// pub extern fn chart_axis_set_pattern(axis: [*c]lxw_chart_axis, pattern: [*c]lxw_chart_pattern) void;
// pub extern fn chart_axis_set_reverse(axis: [*c]lxw_chart_axis) void;
// pub extern fn chart_axis_set_crossing(axis: [*c]lxw_chart_axis, value: f64) void;
// pub extern fn chart_axis_set_crossing_max(axis: [*c]lxw_chart_axis) void;
// pub extern fn chart_axis_set_crossing_min(axis: [*c]lxw_chart_axis) void;
// pub extern fn chart_axis_off(axis: [*c]lxw_chart_axis) void;
// pub extern fn chart_axis_set_position(axis: [*c]lxw_chart_axis, position: u8) void;
// pub extern fn chart_axis_set_label_position(axis: [*c]lxw_chart_axis, position: u8) void;
// pub extern fn chart_axis_set_label_align(axis: [*c]lxw_chart_axis, @"align": u8) void;
// pub extern fn chart_axis_set_min(axis: [*c]lxw_chart_axis, min: f64) void;
// pub extern fn chart_axis_set_max(axis: [*c]lxw_chart_axis, max: f64) void;
// pub extern fn chart_axis_set_log_base(axis: [*c]lxw_chart_axis, log_base: u16) void;
// pub extern fn chart_axis_set_major_tick_mark(axis: [*c]lxw_chart_axis, @"type": u8) void;
// pub extern fn chart_axis_set_minor_tick_mark(axis: [*c]lxw_chart_axis, @"type": u8) void;
// pub extern fn chart_axis_set_interval_unit(axis: [*c]lxw_chart_axis, unit: u16) void;
// pub extern fn chart_axis_set_interval_tick(axis: [*c]lxw_chart_axis, unit: u16) void;
// pub extern fn chart_axis_set_major_unit(axis: [*c]lxw_chart_axis, unit: f64) void;
// pub extern fn chart_axis_set_minor_unit(axis: [*c]lxw_chart_axis, unit: f64) void;
// pub extern fn chart_axis_set_display_units(axis: [*c]lxw_chart_axis, units: u8) void;
// pub extern fn chart_axis_set_display_units_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
// pub extern fn chart_axis_major_gridlines_set_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
// pub extern fn chart_axis_minor_gridlines_set_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
// pub extern fn chart_axis_major_gridlines_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;
// pub extern fn chart_axis_minor_gridlines_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const CString = xlsxwriter.CString;

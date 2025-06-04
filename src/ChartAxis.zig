const ChartAxis = @This();

axis_c: ?*c.lxw_chart_axis,

pub const Type = enum(u8) {
    x = c.LXW_CHART_AXIS_TYPE_X,
    y = c.LXW_CHART_AXIS_TYPE_Y,
};

pub const Position = enum(u8) {
    default = c.LXW_CHART_AXIS_POSITION_DEFAULT,
    on_tick = c.LXW_CHART_AXIS_POSITION_ON_TICK,
    between = c.LXW_CHART_AXIS_POSITION_BETWEEN,
};

pub const LabelPosition = enum(u8) {
    next_to = c.LXW_CHART_AXIS_LABEL_POSITION_NEXT_TO,
    high = c.LXW_CHART_AXIS_LABEL_POSITION_HIGH,
    low = c.LXW_CHART_AXIS_LABEL_POSITION_LOW,
    none = c.LXW_CHART_AXIS_LABEL_POSITION_NONE,
};

pub const LabelAlignment = enum(u8) {
    center = c.LXW_CHART_AXIS_LABEL_ALIGN_CENTER,
    left = c.LXW_CHART_AXIS_LABEL_ALIGN_LEFT,
    right = c.LXW_CHART_AXIS_LABEL_ALIGN_RIGHT,
};

pub const DisplayUnit = enum(u8) {
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

// pub extern fn chart_axis_set_name(axis: [*c]lxw_chart_axis, name: [*c]const u8) void;
pub inline fn setName(
    self: ChartAxis,
    name: [:0]const u8,
) void {
    c.chart_axis_set_name(
        self.axis_c,
        name,
    );
}

// pub extern fn chart_axis_set_name_range(axis: [*c]lxw_chart_axis, sheetname: [*c]const u8, row: lxw_row_t, col: lxw_col_t) void;
pub inline fn setNameRange(
    self: ChartAxis,
    sheetname: [:0]const u8,
    cell: Cell,
) void {
    c.chart_axis_set_name_range(
        self.axis_c,
        sheetname,
        cell.row,
        cell.col,
    );
}

// pub extern fn chart_axis_set_name_layout(axis: [*c]lxw_chart_axis, layout: [*c]lxw_chart_layout) void;
pub inline fn setNameLayout(self: ChartAxis, layout: ChartLayout) void {
    c.chart_axis_set_name_layout(
        self.axis_c,
        @constCast(&layout.toC()),
    );
}

// pub extern fn chart_axis_set_name_font(axis: [*c]lxw_chart_axis, font: [*c]lxw_chart_font) void;
pub inline fn setNameFont(self: ChartAxis, font: ChartFont) void {
    c.chart_axis_set_name_font(
        self.axis_c,
        @constCast(&font.toC()),
    );
}

// pub extern fn chart_axis_set_num_font(axis: [*c]lxw_chart_axis, font: [*c]lxw_chart_font) void;
pub inline fn setNumFont(self: ChartAxis, font: ChartFont) void {
    c.chart_axis_set_num_font(
        self.axis_c,
        @constCast(&font.toC()),
    );
}

// pub extern fn chart_axis_set_num_format(axis: [*c]lxw_chart_axis, num_format: [*c]const u8) void;
pub inline fn setNumFormat(self: ChartAxis, num_format: [:0]const u8) void {
    c.chart_axis_set_num_format(
        self.axis_c,
        num_format,
    );
}

// pub extern fn chart_axis_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;
pub inline fn setLine(self: ChartAxis, line: ChartLine) void {
    c.chart_axis_set_line(
        self.axis_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_axis_set_fill(axis: [*c]lxw_chart_axis, fill: [*c]lxw_chart_fill) void;
pub inline fn setFill(self: ChartAxis, fill: ChartFill) void {
    c.chart_axis_set_fill(
        self.axis_c,
        @constCast(&fill.toC()),
    );
}

// pub extern fn chart_axis_set_pattern(axis: [*c]lxw_chart_axis, pattern: [*c]lxw_chart_pattern) void;
pub inline fn setPattern(self: ChartAxis, pattern: ChartPattern) void {
    c.chart_axis_set_pattern(
        self.axis_c,
        @constCast(&pattern.toC()),
    );
}

// pub extern fn chart_axis_set_reverse(axis: [*c]lxw_chart_axis) void;
pub inline fn setReverse(self: ChartAxis) void {
    c.chart_axis_set_reverse(self.axis_c);
}

// pub extern fn chart_axis_set_crossing(axis: [*c]lxw_chart_axis, value: f64) void;
pub inline fn setCrossing(self: ChartAxis, value: f64) void {
    c.chart_axis_set_crossing(self.axis_c, value);
}

// pub extern fn chart_axis_set_crossing_max(axis: [*c]lxw_chart_axis) void;
pub inline fn setCrossingMax(self: ChartAxis) void {
    c.chart_axis_set_crossing_max(self.axis_c);
}

// pub extern fn chart_axis_set_crossing_min(axis: [*c]lxw_chart_axis) void;
pub inline fn setCrossingMin(self: ChartAxis) void {
    c.chart_axis_set_crossing_min(self.axis_c);
}

// pub extern fn chart_axis_off(axis: [*c]lxw_chart_axis) void;
pub inline fn off(self: ChartAxis) void {
    c.chart_axis_set_crossing_min(self.axis_c);
}

// pub extern fn chart_axis_set_position(axis: [*c]lxw_chart_axis, position: u8) void;
pub inline fn setPosition(self: ChartAxis, position: Position) void {
    c.chart_axis_set_position(
        self.axis_c,
        @intFromEnum(position),
    );
}

// pub extern fn chart_axis_set_label_position(axis: [*c]lxw_chart_axis, position: u8) void;
pub inline fn setLabelPosition(self: ChartAxis, position: LabelPosition) void {
    c.chart_axis_set_label_position(
        self.axis_c,
        @intFromEnum(position),
    );
}

// pub extern fn chart_axis_set_label_align(axis: [*c]lxw_chart_axis, @"align": u8) void;
pub inline fn setLabelAlign(self: ChartAxis, @"align": LabelAlignment) void {
    c.chart_axis_set_label_position(
        self.axis_c,
        @intFromEnum(@"align"),
    );
}

// pub extern fn chart_axis_set_min(axis: [*c]lxw_chart_axis, min: f64) void;
pub inline fn setMin(self: ChartAxis, min: f64) void {
    c.chart_axis_set_min(self.axis_c, min);
}

// pub extern fn chart_axis_set_max(axis: [*c]lxw_chart_axis, max: f64) void;
pub inline fn setMax(self: ChartAxis, max: f64) void {
    c.chart_axis_set_max(self.axis_c, max);
}

// pub extern fn chart_axis_set_log_base(axis: [*c]lxw_chart_axis, log_base: u16) void;
pub inline fn setLogBase(self: ChartAxis, log_base: f64) void {
    c.chart_axis_set_log_base(self.axis_c, log_base);
}

// pub extern fn chart_axis_set_major_tick_mark(axis: [*c]lxw_chart_axis, @"type": u8) void;
pub inline fn setMajorTickMark(self: ChartAxis, @"type": TickMark) void {
    c.chart_axis_set_major_tick_mark(self.axis_c, @"type");
}

// pub extern fn chart_axis_set_minor_tick_mark(axis: [*c]lxw_chart_axis, @"type": u8) void;
pub inline fn setMinorTickMark(self: ChartAxis, @"type": TickMark) void {
    c.chart_axis_set_minor_tick_mark(self.axis_c, @"type");
}

// pub extern fn chart_axis_set_interval_unit(axis: [*c]lxw_chart_axis, unit: u16) void;
pub inline fn setIntervalUnit(self: ChartAxis, unit: u16) void {
    c.chart_axis_set_interval_unit(self.axis_c, unit);
}

// pub extern fn chart_axis_set_interval_tick(axis: [*c]lxw_chart_axis, unit: u16) void;
pub inline fn setIntervalTick(self: ChartAxis, unit: u16) void {
    c.chart_axis_set_interval_tick(self.axis_c, unit);
}

// pub extern fn chart_axis_set_major_unit(axis: [*c]lxw_chart_axis, unit: f64) void;
pub inline fn setMajorUnit(self: ChartAxis, unit: f64) void {
    c.chart_axis_set_major_unit(self.axis_c, unit);
}

// pub extern fn chart_axis_set_minor_unit(axis: [*c]lxw_chart_axis, unit: f64) void;
pub inline fn setMinorUnit(self: ChartAxis, unit: f64) void {
    c.chart_axis_set_minor_unit(self.axis_c, unit);
}

// pub extern fn chart_axis_set_display_units(axis: [*c]lxw_chart_axis, units: u8) void;
pub inline fn setDisplayUnits(self: ChartAxis, units: DisplayUnit) void {
    c.chart_axis_set_display_units(
        self.axis_c,
        @intFromEnum(units),
    );
}

// pub extern fn chart_axis_set_display_units_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
pub inline fn setDisplayUnitsVisible(self: ChartAxis, visible: bool) void {
    c.chart_axis_set_display_units_visible(
        self.axis_c,
        @intFromBool(visible),
    );
}

// pub extern fn chart_axis_major_gridlines_set_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
pub inline fn setMajorGridLinesVisible(self: ChartAxis, visible: bool) void {
    c.chart_axis_major_gridlines_set_visible(
        self.axis_c,
        @intFromBool(visible),
    );
}

// pub extern fn chart_axis_minor_gridlines_set_visible(axis: [*c]lxw_chart_axis, visible: u8) void;
pub inline fn setMinorGridLinesVisible(self: ChartAxis, visible: bool) void {
    c.chart_axis_minor_gridlines_set_visible(
        self.axis_c,
        @intFromBool(visible),
    );
}

// pub extern fn chart_axis_major_gridlines_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;
pub inline fn setMajorGridLinesLine(self: ChartAxis, line: ChartLine) void {
    c.chart_axis_major_gridlines_set_line(
        self.axis_c,
        @constCast(&line.toC()),
    );
}

// pub extern fn chart_axis_minor_gridlines_set_line(axis: [*c]lxw_chart_axis, line: [*c]lxw_chart_line) void;
pub inline fn setMinorGridLinesLine(self: ChartAxis, line: ChartLine) void {
    c.chart_axis_minor_gridlines_set_line(
        self.axis_c,
        @constCast(&line.toC()),
    );
}

const c = @import("lxw");
const xlsxwriter = @import("xlsxwriter.zig");
const Cell = @import("utility.zig").Cell;
const ChartFont = @import("Chart.zig").Font;
const ChartLayout = @import("Chart.zig").Layout;
const ChartLine = @import("Chart.zig").Line;
const ChartFill = @import("Chart.zig").Fill;
const ChartPattern = @import("Chart.zig").Pattern;

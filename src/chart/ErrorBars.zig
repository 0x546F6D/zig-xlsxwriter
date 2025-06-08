const ErrorBars = @This();

error_bars_c: ?*c.lxw_series_error_bars,

pub const Type = enum(u8) {
    std_error = c.LXW_CHART_ERROR_BAR_TYPE_STD_ERROR,
    fixed = c.LXW_CHART_ERROR_BAR_TYPE_FIXED,
    percentage = c.LXW_CHART_ERROR_BAR_TYPE_PERCENTAGE,
    std_dev = c.LXW_CHART_ERROR_BAR_TYPE_STD_DEV,
};

pub const Direction = enum(u8) {
    both = c.LXW_CHART_ERROR_BAR_DIR_BOTH,
    plus = c.LXW_CHART_ERROR_BAR_DIR_PLUS,
    minus = c.LXW_CHART_ERROR_BAR_DIR_MINUS,
};

pub const Axis = enum(u8) {
    x = c.LXW_CHART_ERROR_BAR_AXIS_X,
    y = c.LXW_CHART_ERROR_BAR_AXIS_Y,
};

pub const Cap = enum(u8) {
    end_cap = c.LXW_CHART_ERROR_BAR_END_CAP,
    no_cap = c.LXW_CHART_ERROR_BAR_NO_CAP,
};

// pub extern fn chart_series_set_error_bars(error_bars: [*c]lxw_series_error_bars, @"type": u8, value: f64) void;
pub inline fn setType(
    self: ErrorBars,
    @"type": Type,
    value: f64,
) void {
    c.chart_series_set_error_bars(
        self.error_bars_c,
        @intFromEnum(@"type"),
        value,
    );
}

// pub extern fn chart_series_set_error_bars_direction(error_bars: [*c]lxw_series_error_bars, direction: u8) void;
pub inline fn setDirection(
    self: ErrorBars,
    direction: Direction,
) void {
    c.chart_series_set_error_bars_direction(
        self.error_bars_c,
        @intFromEnum(direction),
    );
}

// pub extern fn chart_series_set_error_bars_endcap(error_bars: [*c]lxw_series_error_bars, endcap: u8) void;
pub inline fn setEndCap(
    self: ErrorBars,
    endcap: Cap,
) void {
    c.chart_series_set_error_bars_endcap(
        self.error_bars_c,
        @intFromEnum(endcap),
    );
}

// pub extern fn chart_series_set_error_bars_line(error_bars: [*c]lxw_series_error_bars, line: [*c]lxw_chart_line) void;
pub inline fn setLine(
    self: ErrorBars,
    line: ChartLine,
) void {
    c.chart_series_set_error_bars_line(
        self.error_bars_c,
        @constCast(&line.toC()),
    );
}

const c = @import("lxw");
const ChartLine = @import("../Chart.zig").Line;

const Format = @This();

format_c: ?*c.lxw_format,

pub const none: Format = .{
    .format_c = null,
};

pub const DefinedColor = enum(c_int) {
    black = c.LXW_COLOR_BLACK,
    blue = c.LXW_COLOR_BLUE,
    brown = c.LXW_COLOR_BROWN,
    cyan = c.LXW_COLOR_CYAN,
    gray = c.LXW_COLOR_GRAY,
    green = c.LXW_COLOR_GREEN,
    lime = c.LXW_COLOR_LIME,
    magenta = c.LXW_COLOR_MAGENTA,
    navy = c.LXW_COLOR_NAVY,
    orange = c.LXW_COLOR_ORANGE,
    // pink = c.LXW_COLOR_PINK,
    purple = c.LXW_COLOR_PURPLE,
    red = c.LXW_COLOR_RED,
    silver = c.LXW_COLOR_SILVER,
    white = c.LXW_COLOR_WHITE,
    yellow = c.LXW_COLOR_YELLOW,
};

pub fn setBold(self: Format) void {
    c.format_set_bold(self.format_c);
}

pub fn setNumFormat(self: Format, num_format: [*c]const u8) void {
    c.format_set_num_format(self.format_c, num_format);
}

const c = @import("xlsxwriter_c");

// pub extern fn worksheet_write_comment(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8) lxw_error;
pub const CommentDisplayTypes = enum(u8) {
    default = c.LXW_COMMENT_DISPLAY_DEFAULT,
    hidden = c.LXW_COMMENT_DISPLAY_HIDDEN,
    visible = c.LXW_COMMENT_DISPLAY_VISIBLE,
};

pub const CommentOptions = extern struct {
    visible: CommentDisplayTypes = .default,
    author: ?[*:0]const u8 = null,
    width: u16 = 0,
    height: u16 = 0,
    x_scale: f64 = 1,
    y_scale: f64 = 1,
    color: Format.DefinedColors = .default,
    font_name: ?[*:0]const u8 = null,
    font_size: f64 = 8,
    font_family: u8 = 0,
    start_row: u32 = 0,
    start_col: u16 = 0,
    x_offset: i32 = 0,
    y_offset: i32 = 0,

    pub const default = CommentOptions{
        .visible = .default,
        .author = null,
        .width = 0,
        .height = 0,
        .x_scale = 1,
        .y_scale = 1,
        .color = .default,
        .font_name = null,
        .font_size = 8,
        .font_family = 0,
        .start_row = 0,
        .start_col = 0,
        .x_offset = 0,
        .y_offset = 0,
    };

    inline fn toC(self: CommentOptions) c.lxw_comment_options {
        return c.lxw_comment_options{
            .visible = @intFromEnum(self.visible),
            .author = self.author,
            .width = self.width,
            .height = self.height,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .color = @intFromEnum(self.color),
            .font_name = self.font_name,
            .font_size = self.font_size,
            .font_family = self.font_family,
            .start_row = self.start_row,
            .start_col = self.start_col,
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
        };
    }
};

// pub extern fn worksheet_write_comment(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_comment_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, options: [*c]lxw_comment_options) lxw_error;
pub inline fn writeComment(
    self: WorkSheet,
    cell: Cell,
    string: [:0]const u8,
    options: ?CommentOptions,
) XlsxError!void {
    try check(if (options) |c_options|
        c.worksheet_write_comment_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            string,
            @constCast(&c_options.toC()),
        )
    else
        c.worksheet_write_comment(
            self.worksheet_c,
            cell.row,
            cell.col,
            string,
        ));
}

// pub extern fn worksheet_show_comments(worksheet: [*c]lxw_worksheet) void;
pub inline fn showComments(self: WorkSheet) void {
    c.worksheet_show_comments(self.worksheet_c);
}

const c = @import("lxw");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const WorkSheet = @import("../WorkSheet.zig");
const Format = @import("../Format.zig");
const utility = @import("../utility.zig");
const Cell = utility.Cell;

// image functions
pub const ImageOptions = struct {
    x_offset: i32 = 0,
    y_offset: i32 = 0,
    x_scale: f64 = 0,
    y_scale: f64 = 0,
    object_position: u8 = 0,
    description: ?[*:0]const u8 = null,
    decorative: bool = false,
    url: ?[*:0]const u8 = null,
    tip: ?[*:0]const u8 = null,
    cell_format: Format = .default,

    pub const default = ImageOptions{
        .x_offset = 0,
        .y_offset = 0,
        .x_scale = 0,
        .y_scale = 0,
        .object_position = 0,
        .description = null,
        .decorative = false,
        .url = null,
        .tip = null,
        .cell_format = .default,
    };

    inline fn toC(self: ImageOptions) c.lxw_image_options {
        return c.lxw_image_options{
            .x_offset = self.x_offset,
            .y_offset = self.y_offset,
            .x_scale = self.x_scale,
            .y_scale = self.y_scale,
            .object_position = self.object_position,
            .description = self.description,
            .decorative = @intFromBool(self.decorative),
            .url = self.url,
            .tip = self.tip,
            .cell_format = self.cell_format.format_c,
        };
    }
};

// pub extern fn worksheet_insert_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
// pub extern fn worksheet_insert_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
pub inline fn insertImage(
    self: WorkSheet,
    cell: Cell,
    filename: [:0]const u8,
    options: ?ImageOptions,
) XlsxError!void {
    try check(if (options) |img_options|
        c.worksheet_insert_image_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            filename,
            @constCast(&img_options.toC()),
        )
    else
        c.worksheet_insert_image(
            self.worksheet_c,
            cell.row,
            cell.col,
            filename,
        ));
}

// pub extern fn worksheet_insert_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
// pub extern fn worksheet_insert_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;
pub inline fn insertImageBuffer(
    self: WorkSheet,
    cell: Cell,
    image_buffer: [:0]const u8,
    // image_size: usize,
    options: ?ImageOptions,
) XlsxError!void {
    try check(if (options) |img_options|
        c.worksheet_insert_image_buffer_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            image_buffer,
            image_buffer.len,
            @constCast(&img_options.toC()),
        )
    else
        c.worksheet_insert_image_buffer(
            self.worksheet_c,
            cell.row,
            cell.col,
            image_buffer,
            image_buffer.len,
        ));
}

// pub extern fn worksheet_embed_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
// pub extern fn worksheet_embed_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
pub inline fn embedImage(
    self: WorkSheet,
    cell: Cell,
    filename: [:0]const u8,
    options: ?ImageOptions,
) XlsxError!void {
    try check(if (options) |img_options|
        c.worksheet_embed_image_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            filename,
            @constCast(&img_options.toC()),
        )
    else
        c.worksheet_embed_image(
            self.worksheet_c,
            cell.row,
            cell.col,
            filename,
        ));
}

// pub extern fn worksheet_embed_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
// pub extern fn worksheet_embed_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;
pub inline fn embedImageBuffer(
    self: WorkSheet,
    cell: Cell,
    image_buffer: [:0]const u8,
    options: ?ImageOptions,
) XlsxError!void {
    try check(if (options) |img_options|
        c.worksheet_embed_image_buffer_opt(
            self.worksheet_c,
            cell.row,
            cell.col,
            image_buffer,
            image_buffer.len,
            @constCast(&img_options.toC()),
        )
    else
        c.worksheet_embed_image_buffer(
            self.worksheet_c,
            cell.row,
            cell.col,
            image_buffer,
            image_buffer.len,
        ));
}

const c = @import("lxw");
const WorkSheet = @import("../WorkSheet.zig");
const XlsxError = @import("../errors.zig").XlsxError;
const check = @import("../errors.zig").checkResult;
const Format = @import("../Format.zig");
const Cell = @import("../utility.zig").Cell;

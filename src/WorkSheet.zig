const WorkSheet = @This();

worksheet_c: ?*c.lxw_worksheet,

// set functions
pub const RowColOptions = c.lxw_row_col_options;

// pub extern fn worksheet_set_row(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format) lxw_error;
pub inline fn setRow(
    self: WorkSheet,
    row: u16,
    height: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_row(
        self.worksheet_c,
        row,
        height,
        format.format_c,
    ));
}

// pub extern fn worksheet_set_row_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, height: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setRowOpt(
    self: WorkSheet,
    row: u32,
    height: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_row_opt(
        self.worksheet_c,
        row,
        height,
        format.format_c,
        @constCast(&options),
    ));
}

// pub extern fn worksheet_set_row_pixels(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_row_pixels_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;

// pub extern fn worksheet_set_column(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format) lxw_error;
pub inline fn setColumn(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_set_column(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
    ));
}

// pub extern fn worksheet_set_column_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, width: f64, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
pub inline fn setColumnOpt(
    self: WorkSheet,
    first_col: u16,
    last_col: u16,
    width: f64,
    format: Format,
    options: RowColOptions,
) XlsxError!void {
    try check(c.worksheet_set_column_opt(
        self.worksheet_c,
        first_col,
        last_col,
        width,
        format.format_c,
        @constCast(&options),
    ));
}

// pub extern fn worksheet_set_column_pixels(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_set_column_pixels_opt(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t, pixels: u32, format: [*c]lxw_format, options: [*c]lxw_row_col_options) lxw_error;
// pub extern fn worksheet_set_background(worksheet: [*c]lxw_worksheet, filename: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_background_buffer(worksheet: [*c]lxw_worksheet, image_buffer: [*c]const u8, image_size: usize) lxw_error;
// pub extern fn worksheet_set_selection(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_set_top_left_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
// pub extern fn worksheet_set_landscape(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_portrait(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_page_view(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_paper(worksheet: [*c]lxw_worksheet, paper_type: u8) void;
// pub extern fn worksheet_set_margins(worksheet: [*c]lxw_worksheet, left: f64, right: f64, top: f64, bottom: f64) void;
// pub extern fn worksheet_set_header(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_footer(worksheet: [*c]lxw_worksheet, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_header_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
// pub extern fn worksheet_set_footer_opt(worksheet: [*c]lxw_worksheet, string: [*c]const u8, options: [*c]lxw_header_footer_options) lxw_error;
// pub extern fn worksheet_set_h_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_row_t) lxw_error;
// pub extern fn worksheet_set_v_pagebreaks(worksheet: [*c]lxw_worksheet, breaks: [*c]lxw_col_t) lxw_error;
// pub extern fn worksheet_set_zoom(worksheet: [*c]lxw_worksheet, scale: u16) void;
// pub extern fn worksheet_set_start_page(worksheet: [*c]lxw_worksheet, start_page: u16) void;
// pub extern fn worksheet_set_print_scale(worksheet: [*c]lxw_worksheet, scale: u16) void;
// pub extern fn worksheet_set_tab_color(worksheet: [*c]lxw_worksheet, color: lxw_color_t) void;
// pub extern fn worksheet_set_default_row(worksheet: [*c]lxw_worksheet, height: f64, hide_unused_rows: u8) void;
// pub extern fn worksheet_set_vba_name(worksheet: [*c]lxw_worksheet, name: [*c]const u8) lxw_error;
// pub extern fn worksheet_set_comments_author(worksheet: [*c]lxw_worksheet, author: [*c]const u8) void;
// pub extern fn worksheet_set_error_cell(worksheet: [*c]lxw_worksheet, object_props: [*c]lxw_object_properties, ref_id: u32) void;

// write functions
// pub extern fn worksheet_write_number(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, number: f64, format: [*c]lxw_format) lxw_error;
pub inline fn writeNumber(
    self: WorkSheet,
    row: u32,
    col: u16,
    number: f64,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_number(
        self.worksheet_c,
        row,
        col,
        number,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeString(
    self: WorkSheet,
    row: u32,
    col: u16,
    string: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_string(
        self.worksheet_c,
        row,
        col,
        string,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeFormula(
    self: WorkSheet,
    row: u32,
    col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_formula(
        self.worksheet_c,
        row,
        col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
pub inline fn writeArrayFormula(
    self: WorkSheet,
    first_row: u32,
    first_col: u16,
    last_row: u32,
    last_col: u16,
    formula: [*c]const u8,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_array_formula(
        self.worksheet_c,
        first_row,
        first_col,
        last_row,
        last_col,
        formula,
        format.format_c,
    ));
}

// pub extern fn worksheet_write_dynamic_array_formula(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_dynamic_formula(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_dynamic_array_formula_num(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_dynamic_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;

// pub extern fn worksheet_write_datetime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, datetime: [*c]lxw_datetime, format: [*c]lxw_format) lxw_error;
pub inline fn writeDateTime(
    self: WorkSheet,
    row: u32,
    col: u16,
    datetime: c.lxw_datetime,
    format: Format,
) XlsxError!void {
    try check(c.worksheet_write_datetime(
        self.worksheet_c,
        row,
        col,
        @constCast(&datetime),
        format.format_c,
    ));
}

// pub extern fn worksheet_write_unixtime(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, unixtime: i64, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_url(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_url_opt(worksheet: [*c]lxw_worksheet, row_num: lxw_row_t, col_num: lxw_col_t, url: [*c]const u8, format: [*c]lxw_format, string: [*c]const u8, tooltip: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_boolean(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, value: c_int, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_blank(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_formula_num(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: f64) lxw_error;
// pub extern fn worksheet_write_formula_str(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, formula: [*c]const u8, format: [*c]lxw_format, result: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_rich_string(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, rich_string: [*c][*c]lxw_rich_string_tuple, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_write_comment(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8) lxw_error;
// pub extern fn worksheet_write_comment_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, string: [*c]const u8, options: [*c]lxw_comment_options) lxw_error;

// image functions
// pub extern fn worksheet_insert_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
pub inline fn insertImage(
    self: WorkSheet,
    row: u32,
    col: u16,
    filename: [*c]const u8,
) XlsxError!void {
    try check(c.worksheet_insert_image(
        self.worksheet_c,
        row,
        col,
        filename,
    ));
}

// pub extern fn worksheet_insert_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
// pub extern fn worksheet_insert_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
pub inline fn insertImageBuffer(
    self: WorkSheet,
    row: u32,
    col: u16,
    image_buffer: [*c]const u8,
    image_size: usize,
) XlsxError!void {
    try check(c.worksheet_insert_image_buffer(
        self.worksheet_c,
        row,
        col,
        image_buffer,
        image_size,
    ));
}

// pub extern fn worksheet_insert_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;
// pub extern fn worksheet_embed_image(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8) lxw_error;
// pub extern fn worksheet_embed_image_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, filename: [*c]const u8, options: [*c]lxw_image_options) lxw_error;
// pub extern fn worksheet_embed_image_buffer(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize) lxw_error;
// pub extern fn worksheet_embed_image_buffer_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, image_buffer: [*c]const u8, image_size: usize, options: [*c]lxw_image_options) lxw_error;

// chart functions
// pub extern fn worksheet_insert_chart(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart) lxw_error;
pub inline fn insertChart(
    self: WorkSheet,
    row: u32,
    col: u16,
    chart: Chart,
) XlsxError!void {
    try check(c.worksheet_insert_chart(
        self.worksheet_c,
        row,
        col,
        chart.chart_c,
    ));
}

// pub extern fn worksheet_insert_chart_opt(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, chart: [*c]lxw_chart, user_options: [*c]lxw_chart_options) lxw_error;

// filter functions
pub const autoFilter = filter.autoFilter;
pub const filterColumn = filter.filterColumn;
pub const filterColumn2 = filter.filterColumn2;
pub const filterList = filter.filterList;

// panes functions
// pub extern fn worksheet_freeze_panes(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t) void;
// pub extern fn worksheet_split_panes(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64) void;
// pub extern fn worksheet_freeze_panes_opt(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, top_row: lxw_row_t, left_col: lxw_col_t, @"type": u8) void;
// pub extern fn worksheet_split_panes_opt(worksheet: [*c]lxw_worksheet, vertical: f64, horizontal: f64, top_row: lxw_row_t, left_col: lxw_col_t) void;

// other functions
// pub extern fn worksheet_merge_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, string: [*c]const u8, format: [*c]lxw_format) lxw_error;
// pub extern fn worksheet_data_validation_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
// pub extern fn worksheet_data_validation_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, validation: [*c]lxw_data_validation) lxw_error;
// pub extern fn worksheet_conditional_format_cell(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;
// pub extern fn worksheet_conditional_format_range(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, conditional_format: [*c]lxw_conditional_format) lxw_error;
// pub extern fn worksheet_insert_button(worksheet: [*c]lxw_worksheet, row: lxw_row_t, col: lxw_col_t, options: [*c]lxw_button_options) lxw_error;
// pub extern fn worksheet_add_table(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t, options: [*c]lxw_table_options) lxw_error;
// pub extern fn worksheet_activate(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_select(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_hide(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_set_first_sheet(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_print_across(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_gridlines(worksheet: [*c]lxw_worksheet, option: u8) void;
// pub extern fn worksheet_center_horizontally(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_center_vertically(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_print_row_col_headers(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_repeat_rows(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, last_row: lxw_row_t) lxw_error;
// pub extern fn worksheet_repeat_columns(worksheet: [*c]lxw_worksheet, first_col: lxw_col_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_print_area(worksheet: [*c]lxw_worksheet, first_row: lxw_row_t, first_col: lxw_col_t, last_row: lxw_row_t, last_col: lxw_col_t) lxw_error;
// pub extern fn worksheet_fit_to_pages(worksheet: [*c]lxw_worksheet, width: u16, height: u16) void;
// pub extern fn worksheet_print_black_and_white(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_right_to_left(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_hide_zero(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_protect(worksheet: [*c]lxw_worksheet, password: [*c]const u8, options: [*c]lxw_protection) void;
// pub extern fn worksheet_outline_settings(worksheet: [*c]lxw_worksheet, visible: u8, symbols_below: u8, symbols_right: u8, auto_style: u8) void;
// pub extern fn worksheet_show_comments(worksheet: [*c]lxw_worksheet) void;
// pub extern fn worksheet_ignore_errors(worksheet: [*c]lxw_worksheet, @"type": u8, range: [*c]const u8) lxw_error;

const c = @import("xlsxwriter_c");
const XlsxError = @import("errors.zig").XlsxError;
const check = @import("errors.zig").checkResult;
const Format = @import("Format.zig");
const Chart = @import("Chart.zig");
const filter = @import("filter.zig");

pub const c = @import("xlsxwriter_c");
pub const xw_true = c.LXW_TRUE;
pub const xw_false = c.LXW_FALSE;
pub const explicit_false = c.LXW_EXPLICIT_FALSE;
pub const row_max = c.LXW_ROW_MAX;
pub const col_max = c.LXW_COL_MAX;
pub const col_meta_max = c.LXW_COL_META_MAX;
pub const header_footer_max = c.LXW_HEADER_FOOTER_MAX;
pub const max_number_urls = c.LXW_MAX_NUMBER_URLS;
pub const pane_name_length = c.LXW_PANE_NAME_LENGTH;
pub const image_buffer_size = c.LXW_IMAGE_BUFFER_SIZE;
pub const header_footer_objs_max = c.LXW_HEADER_FOOTER_OBJS_MAX;
pub const breaks_max = c.LXW_BREAKS_MAX;
pub const def_col_width = c.LXW_DEF_COL_WIDTH;
pub const def_row_height = c.LXW_DEF_ROW_HEIGHT;
pub const def_col_width_pixels = c.LXW_DEF_COL_WIDTH_PIXELS;
pub const def_row_width_pixels = c.LXW_DEF_ROW_HEIGHT_PIXELS;
pub const filter_or = c.LXW_FILTER_OR;
pub const filter_and = c.LXW_FILTER_AND;

pub const DateTime = c.lxw_datetime;

pub const range = c.RANGE;
pub const cell = c.CELL;
pub const cols = c.COLS;

pub const WorkBook = @import("WorkBook.zig");
pub const initWorkBook = WorkBook.init;

pub const WorkSheet = @import("WorkSheet.zig");
pub const RowColOptions = WorkSheet.RowColOptions;
pub const FilterRule = WorkSheet.FilterRule;

pub const Format = @import("Format.zig");
pub const DefinedColor = Format.DefinedColor;

pub const Chart = @import("Chart.zig");

pub const ChartType = Chart.Type;

pub const ChartFont = Chart.Font;

pub const ChartSeries = Chart.Series;

pub const errors = @import("errors.zig");
pub const XlsxError = errors.XlsxError;
pub const checkResult = errors.checkResult;
pub const strError = errors.strError;

pub const utility = @import("utility.zig");
pub const nameToRow = utility.nameToRow;
pub const nameToRow2 = utility.nameToRow2;
pub const nameToCol = utility.nameToCol;
pub const nameToCol2 = utility.nameToCol2;

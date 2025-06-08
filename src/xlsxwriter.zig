pub const CStringArray: type = [:null]const ?[*:0]const u8;

pub const c = @import("lxw");
pub const Boolean = enum(u8) {
    true = c.LXW_TRUE,
    false = c.LXW_FALSE,
    explicit_false = c.LXW_EXPLICIT_FALSE,
};
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

pub const DateTime = c.lxw_datetime;

pub const WorkBook = @import("WorkBook.zig");
// pub const init = WorkBook.init;
pub const initWorkBook = WorkBook.initWorkBook;
pub const WorkBookOptions = WorkBook.WorkBookOptions;
pub const initWorkBookOpt = WorkBook.newOpt;
pub const DocProperties = WorkBook.DocProperties;

pub const WorkSheet = @import("WorkSheet.zig");
pub const RowColOptions = WorkSheet.RowColOptions;
pub const RichStringTuple = WorkSheet.RichStringTuple;
pub const RichStringTupleNoAlloc = WorkSheet.RichStringTupleNoAlloc;
pub const RichStringNoAllocArray = WorkSheet.RichStringNoAllocArray;
pub const HeaderFooterOptions = WorkSheet.HeaderFooterOptions;
pub const ButtonOptions = WorkSheet.ButtonOptions;
pub const ObjectProperties = WorkSheet.ObjectProperties;

pub const filter = @import("worksheet/filter.zig");
pub const FilterRule = filter.FilterRule;
pub const filter_or = filter.filter_or;
pub const filter_and = filter.filter_and;

pub const table = @import("worksheet/table.zig");
pub const TableColumn = table.TableColumn;
pub const TableOptions = table.TableOptions;
pub const TableColumnNoAlloc = table.TableColumnNoAlloc;
pub const TableColumnNoAllocArray = table.TableColumnNoAllocArray;
pub const TableOptionsNoAlloc = table.TableOptionsNoAlloc;

pub const ImageOptions = @import("worksheet/image.zig").ImageOptions;

pub const ConditionalFormat = @import("worksheet/conditional.zig").ConditionalFormat;

pub const DataValidation = @import("worksheet/validation.zig").DataValidation;

pub const CommentOptions = @import("worksheet/comment.zig").CommentOptions;

pub const Format = @import("Format.zig");
pub const DefinedColors = Format.DefinedColors;
pub const Alignments = Format.Alignments;
pub const Underlines = Format.Underlines;
pub const Scripts = Format.Scripts;

pub const Chart = @import("Chart.zig");

pub const ChartOptions = Chart.Options;
pub const ChartType = Chart.Type;
pub const ChartFont = Chart.Font;
pub const ChartLine = Chart.Line;
pub const ChartFill = Chart.Fill;
pub const ChartPattern = Chart.Pattern;
pub const PatternType = Chart.PatternType;
pub const ErrorBars = @import("chart/ErrorBars.zig");
pub const ErrorBarsAxis = ErrorBars.Axis;

pub const ChartSeries = Chart.ChartSeries;
pub const ChartPoint = ChartSeries.Point;
pub const ChartPointNoAlloc = ChartSeries.PointNoAlloc;
pub const ChartPointNoAllocArray = ChartSeries.PointNoAllocArray;
pub const ChartDataLabel = ChartSeries.DataLabel;
pub const ChartDataLabelNoAlloc = ChartSeries.DataLabelNoAlloc;
pub const ChartDataLabelNoAllocArray = ChartSeries.DataLabelNoAllocArray;
pub const ChartAxis = Chart.ChartAxis;
pub const ChartAxisType = ChartAxis.Type;

pub const errors = @import("errors.zig");
pub const XlsxError = errors.XlsxError;
pub const checkResult = errors.checkResult;
pub const strError = errors.strError;

pub const utility = @import("utility.zig");
pub const nameToRow = utility.nameToRow;
pub const nameToRow2 = utility.nameToRow2;
pub const nameToCol = utility.nameToCol;
pub const nameToCol2 = utility.nameToCol2;
pub const Cell = utility.Cell;
pub const cell = utility.cell;
pub const Rows = utility.Rows;
pub const rows = utility.rows;
pub const Cols = utility.Cols;
pub const cols = utility.cols;
pub const Range = utility.Range;
pub const range = utility.range;

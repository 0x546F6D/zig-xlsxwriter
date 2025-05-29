pub const CustomPropertyTypes = enum(c_int) {
    none = c.LXW_CUSTOM_NONE,
    string = c.LXW_CUSTOM_STRING,
    double = c.LXW_CUSTOM_DOUBLE,
    integer = c.LXW_CUSTOM_INTEGER,
    boolean = c.LXW_CUSTOM_BOOLEAN,
    datetime = c.LXW_CUSTOM_DATETIME,
};

const md5_size = c.LXW_MD5_SIZE;
const sheetname_max = c.LXW_SHEETNAME_MAX;
const max_sheetname_length = c.LXW_MAX_SHEETNAME_LENGTH;
const max_row_name_length = "$XFD".len; // c.LXW_MAX_COL_NAME_LENGTH;
const max_col_name_length = "$1048576".len; // c.LXW_MAX_ROW_NAME_LENGTH;
const max_cell_name_length = "$XFWD$1048576".len; // c.LXW_MAX_CELL_NAME_LENGTH;
const max_cell_range_length = max_cell_name_length * 2; // c.LXW_MAX_CELL_RANGE_LENGTH;
const max_formula_range_length = c.LXW_MAX_FORMULA_RANGE_LENGTH;
const datetime_length = "2016-12-12T23:00:00Z".len; // c.LXW_DATETIME_LENGTH;
const guid_length = "{12345678-1234-1234-1234-1234567890AB}\\0".len; // c.LXW_GUID_LENGTH;
const epoch_1900 = c.LXW_EPOCH_1900;
const epoch_1904 = c.LXW_EPOCH_1904;
const uint32_t_length = "4294967296".len; // c.LXW_UINT32_T_LENGTH;
const filename_length = c.LXW_FILENAME_LENGTH;
const ignore = c.LXW_IGNORE;
const portrait = c.LXW_PORTRAIT;
const landscape = c.LXW_LANDSCAPE;
const schema_ms = c.LXW_SCHEMA_MS;
const schema_root = c.LXW_SCHEMA_ROOT;
const schema_drawing = c.LXW_SCHEMA_DRAWING;
const schema_officedoc = c.LXW_SCHEMA_OFFICEDOC;
const schema_package = c.LXW_SCHEMA_PACKAGE;
const schema_document = c.LXW_SCHEMA_DOCUMENT;
const schema_content = c.LXW_SCHEMA_CONTENT;

const c = @import("lxw");

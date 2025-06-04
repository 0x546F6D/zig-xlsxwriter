// Control Characters used in the header/footer:
//    Control             Category            Description
//    =======             ========            ===========
//    &L                  Justification       Left
//    &C                                      Center
//    &R                                      Right
//
//    &P                  Information         Page number
//    &N                                      Total number of pages
//    &D                                      Date
//    &T                                      Time
//    &F                                      File name
//    &A                                      Worksheet name
//
//    &fontsize           Font                Font size
//    &"font,style"                           Font name and style
//    &U                                      Single underline
//    &E                                      Double underline
//    &S                                      Strikethrough
//    &X                                      Superscript
//    &Y                                      Subscript
//
//    &[Picture]          Images              Image placeholder
//    &G                                      Same as &[Picture]
//
//    &&                  Miscellaneous       Literal ampersand &

pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(null, xlsx_path, null);
    defer workbook.deinit() catch {};

    // get image path
    const asset_path = try h.getAssetPath(alloc, "logo_small.png");
    defer alloc.free(asset_path);

    const preview = "Select Print Preview to see the header and footer";

    var worksheet: xwz.WorkSheet = undefined;
    var header: [:0]const u8 = undefined;
    var footer: [:0]const u8 = undefined;

    // A simple example to start
    const worksheet1 = try workbook.addWorkSheet("Simple");
    worksheet = worksheet1;
    header = "&CHere is some centered text.";
    footer = "&LHere is some left aligned text.";

    try worksheet.setHeader(header, null);
    try worksheet.setFooter(footer, null);

    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);

    // An example with an image
    const worksheet2 = try workbook.addWorkSheet("Image");
    worksheet = worksheet2;
    const header_options = xwz.HeaderFooterOptions{
        .image_left = asset_path,
    };

    try worksheet.setHeader("&L&[Picture]", header_options);

    worksheet.setMargins(-1, -1, 1.3, -1);
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);

    // This is an example of some of the header/footer variables
    const worksheet3 = try workbook.addWorkSheet("Variables");
    worksheet = worksheet3;
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);
    try worksheet.writeString(.{ .row = 20 }, "Next page", .default);

    header = "&LPage &P of &N" ++ "&CFilename: &F" ++ "&RSheetname: &A";
    footer = "&LCurrent date: &D" ++ "&RCurrent time: &T";
    try worksheet.setHeader(header, null);
    try worksheet.setFooter(footer, null);

    try worksheet.setHPageBreaks(&.{20});

    // This example shows how to use more than one font
    const worksheet4 = try workbook.addWorkSheet("Mixed fonts");
    worksheet = worksheet4;
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);

    header = "&C&\"Courier New,Bold\"Hello &\"Arial,Italic\"World";
    footer = "&C&\"Symbol\"e&\"Arial\" = mc&X2";
    try worksheet.setHeader(header, null);
    try worksheet.setFooter(footer, null);

    // Example of line wrapping
    const worksheet5 = try workbook.addWorkSheet("Word wrap");
    worksheet = worksheet5;
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);

    header = "&CHeading 1\nHeading 2";
    try worksheet.setHeader(header, null);

    // Example of inserting a literal ampersand &
    const worksheet6 = try workbook.addWorkSheet("Ampersand");
    worksheet = worksheet6;
    try worksheet.setColumn(.{}, 50, .default, null);
    try worksheet.writeString(.{}, preview, .default);

    header = "&CCuriouser && Curiouser - Attorneys at Law";
    try worksheet.setHeader(header, null);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

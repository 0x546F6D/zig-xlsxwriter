pub fn main() !void {
    defer _ = dbga.deinit();

    const xlsx_path = try h.getXlsxPath(alloc, @src().file);
    defer alloc.free(xlsx_path);

    // Create a workbook and add a worksheet.
    const workbook = try xwz.initWorkBook(alloc, xlsx_path, null);
    defer workbook.deinit() catch {};
    const worksheet = try workbook.addWorkSheet(null);

    // Add a bold format to use to highlight the header cells
    const bold = try workbook.addFormat();
    bold.setBold();

    // Some chart positioning options.
    const options = xwz.ChartOptions{
        .x_offset = 25,
        .y_offset = 10,
    };

    // Write some data for the chart.
    try worksheet.writeString(.{}, "Number", bold);
    try worksheet.writeNumber(.{ .row = 1 }, 2, .default);
    try worksheet.writeNumber(.{ .row = 2 }, 3, .default);
    try worksheet.writeNumber(.{ .row = 3 }, 4, .default);
    try worksheet.writeNumber(.{ .row = 4 }, 5, .default);
    try worksheet.writeNumber(.{ .row = 5 }, 6, .default);
    try worksheet.writeNumber(.{ .row = 6 }, 7, .default);

    try worksheet.writeString(.{ .col = 1 }, "Data", bold);
    try worksheet.writeNumber(.{ .row = 1, .col = 1 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 2, .col = 1 }, 10, .default);
    try worksheet.writeNumber(.{ .row = 3, .col = 1 }, 20, .default);
    try worksheet.writeNumber(.{ .row = 4, .col = 1 }, 30, .default);
    try worksheet.writeNumber(.{ .row = 5, .col = 1 }, 40, .default);
    try worksheet.writeNumber(.{ .row = 6, .col = 1 }, 30, .default);

    try worksheet.writeString(.{ .col = 2 }, "Text", bold);
    try worksheet.writeString(.{ .row = 1, .col = 2 }, "Jan", .default);
    try worksheet.writeString(.{ .row = 2, .col = 2 }, "Feb", .default);
    try worksheet.writeString(.{ .row = 3, .col = 2 }, "Mar", .default);
    try worksheet.writeString(.{ .row = 4, .col = 2 }, "Apr", .default);
    try worksheet.writeString(.{ .row = 5, .col = 2 }, "May", .default);
    try worksheet.writeString(.{ .row = 6, .col = 2 }, "Jun", .default);

    // Chart 1. Example with standard data labels.
    var chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Chart with standard data labels");

    // Add a data series to the chart.
    var series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 1, .col = 3 }, chart, options);

    // Chart 2. Example with value and category data labels.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Category and Value data labels");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Turn on Value and Category labels.
    series.setLabelsOptions(false, true, true);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 17, .col = 3 }, chart, options);

    // Chart 3. Example with standard data labels with different font.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Data labels with user defined font");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    var font: xwz.ChartFont = .{
        .bold = true,
        .color = .red,
        .rotation = -30,
    };
    series.setLabelsFont(font);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 33, .col = 3 }, chart, options);

    // Chart 4. Example with standard data labels and formatting.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Data labels with formatting");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Set the border/line and fill for the data labels.
    const line = xwz.ChartLine{ .color = .red };
    const fill = xwz.ChartFill{ .color = .yellow };

    series.setLabelsLine(line);
    series.setLabelsFill(fill);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 49, .col = 3 }, chart, options);

    // Chart 5. Example with custom string data labels.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Chart with custom string data labels");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Create some custom labels.
    var data_label_1: xwz.ChartDataLabel = .{ .value = "Amy" };
    var data_label_2: xwz.ChartDataLabel = .{ .value = "Bea" };
    var data_label_3: xwz.ChartDataLabel = .{ .value = "Eva" };
    var data_label_4: xwz.ChartDataLabel = .{ .value = "Fay" };
    var data_label_5: xwz.ChartDataLabel = .{ .value = "Liv" };
    var data_label_6: xwz.ChartDataLabel = .{ .value = "Una" };

    // Create an array of label pointers.
    const data_labels5 = &.{
        data_label_1,
        data_label_2,
        data_label_3,
        data_label_4,
        data_label_5,
        data_label_6,
    };

    // Set the custom labels.
    try series.setLabelsCustom(data_labels5);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 65, .col = 3 }, chart, options);

    // // Chart 6. Example with custom data labels from cells.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Chart with custom data labels from cells");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Create some custom labels.
    data_label_1 = .{ .value = "=Sheet1!$C$2" };
    data_label_2 = .{ .value = "=Sheet1!$C$3" };
    data_label_3 = .{ .value = "=Sheet1!$C$4" };
    data_label_4 = .{ .value = "=Sheet1!$C$5" };
    data_label_5 = .{ .value = "=Sheet1!$C$6" };
    data_label_6 = .{ .value = "=Sheet1!$C$7" };

    // Create an array of label pointers.
    const data_labels6 = &.{
        data_label_1,
        data_label_2,
        data_label_3,
        data_label_4,
        data_label_5,
        data_label_6,
    };

    // Set the custom labels.
    try series.setLabelsCustom(data_labels6);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 81, .col = 3 }, chart, options);

    // Chart 7. Example with custom and default data labels.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Mixed custom and default data labels");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    font = .{ .color = .red };

    // Add the series data labels.
    series.setLabels();

    // Create some custom labels.
    data_label_1 = .{ .value = "=Sheet1!$C$2", .font = &font };
    data_label_2 = .default;
    data_label_3 = .{ .value = "=Sheet1!$C$4", .font = &font };
    data_label_4 = .{ .value = "=Sheet1!$C$5", .font = &font };

    // Create an array of label pointers.
    const data_labels7 = &.{
        data_label_1,
        data_label_2,
        data_label_3,
        data_label_4,
    };

    // Set the custom labels.
    try series.setLabelsCustom(data_labels7);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 97, .col = 3 }, chart, options);

    // Chart 8. Example with deleted/hidden data labels.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Chart with deleted data labels");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Create some custom labels.
    const hide: xwz.ChartDataLabel = .{ .hide = true };
    const keep: xwz.ChartDataLabel = .{ .hide = false };

    // Create an array of label pointers.
    const data_labels8 = &.{
        hide,
        keep,
        hide,
        hide,
        keep,
        hide,
    };

    // Set the custom labels.
    try series.setLabelsCustom(data_labels8);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 113, .col = 3 }, chart, options);

    // Chart 9. Example with custom string data labels and formatting.
    chart = try workbook.addChart(.column);

    // Add a chart title.
    chart.titleSetName("Chart with custom labels and formatting");

    // Add a data series to the chart.
    series = try chart.addSeries("=Sheet1!$A$2:$A$7", "=Sheet1!$B$2:$B$7");

    // Add the series data labels.
    series.setLabels();

    // Set the border/line and fill for the data labels.
    const line2: xwz.ChartLine = .{ .color = .red };
    const fill2: xwz.ChartFill = .{ .color = .yellow };
    const line3: xwz.ChartLine = .{ .color = .blue };
    const fill3: xwz.ChartFill = .{ .color = .green };

    // Create some custom labels.
    data_label_1 = .{ .value = "Amy", .line = &line3 };
    data_label_2 = .{ .value = "Bea" };
    data_label_3 = .{ .value = "Eva" };
    data_label_4 = .{ .value = "Fay" };
    data_label_5 = .{ .value = "Liv" };
    data_label_6 = .{ .value = "Una", .fill = &fill3 };

    // Set the default formatting for the data labels in the series.
    series.setLabelsLine(line2);
    series.setLabelsFill(fill2);

    // Create an array of label pointers.
    const data_labels9 = &.{
        data_label_1,
        data_label_2,
        data_label_3,
        data_label_4,
        data_label_5,
        data_label_6,
    };

    // Set the custom labels.
    try series.setLabelsCustom(data_labels9);

    // Turn off the legend.
    chart.legendSetPosition(.none);

    // Insert the chart into the worksheet.
    try worksheet.insertChart(.{ .row = 129, .col = 3 }, chart, options);
}

var dbga: @import("std").heap.DebugAllocator(.{}) = .init;
const alloc = dbga.allocator();
const h = @import("_helper.zig");
const xwz = @import("xlsxwriter");

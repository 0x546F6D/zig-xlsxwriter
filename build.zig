const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const xlsxwriter_dep = b.dependency("libxlsxwriter", .{
        .target = target,
        .optimize = optimize,
        .USE_SYSTEM_MINIZIP = true,
    });

    const xlsxwriter_c = b.addTranslateC(.{
        .root_source_file = xlsxwriter_dep.path("include/xlsxwriter.h"),
        .target = target,
        .optimize = optimize,
    });

    const xlsxwriter_c_mod = xlsxwriter_c.createModule();
    xlsxwriter_c_mod.linkLibrary(xlsxwriter_dep.artifact("xlsxwriter"));
    xlsxwriter_c_mod.link_libc = true;
    xlsxwriter_c_mod.addCMacro("struct_headname", "");

    const xlsxwriter_mod = b.addModule("xlsxwriter", .{
        .root_source_file = b.path("src/xlsxwriter.zig"),
        .target = target,
        .optimize = optimize,
    });
    xlsxwriter_mod.addImport("xlsxwriter_c", xlsxwriter_c_mod);

    const make = b.option([]const u8, "make", "Provide name of example to make example (ex: 'tutorial1'") orelse null;
    if (make) |name| {
        var dbga: std.heap.DebugAllocator(.{}) = .init;
        defer _ = dbga.deinit();
        const alloc = dbga.allocator();

        const example_path = std.fmt.allocPrint(alloc, "examples/{s}.zig", .{name}) catch "examples/tutorial1.zig";
        defer alloc.free(example_path);
        makeExample(b, .{
            .path = example_path,
            .zig_mod = xlsxwriter_mod,
            .target = target,
            .optimize = optimize,
        });
    }
}

fn makeExample(b: *std.Build, options: BuildInfo) void {
    const example = b.addExecutable(.{
        .name = options.filename(),
        .root_source_file = b.path(options.path),
        .target = options.target,
        .optimize = options.optimize,
    });

    example.root_module.addImport("xlsxwriter", options.zig_mod);

    b.installArtifact(example);

    const run_cmd = b.addRunArtifact(example);
    run_cmd.step.dependOn(b.getInstallStep());
    if (b.args) |args| {
        run_cmd.addArgs(args);
    }

    const descr = b.fmt("Run the {s} example", .{options.filename()});
    const run_step = b.step(options.filename(), descr);
    run_step.dependOn(&run_cmd.step);
}

const BuildInfo = struct {
    target: std.Build.ResolvedTarget,
    optimize: std.builtin.OptimizeMode,
    zig_mod: *std.Build.Module,
    path: []const u8,

    fn filename(self: BuildInfo) []const u8 {
        var split = std.mem.splitSequence(u8, std.fs.path.basename(self.path), ".");
        return split.first();
    }
};

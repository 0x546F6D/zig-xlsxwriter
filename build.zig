const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    const xlsxwriter_dep = b.dependency("libxlsxwriter", .{
        .target = target,
        .optimize = optimize,
        .USE_SYSTEM_MINIZIP = true,
    });

    const lxw = b.addTranslateC(.{
        .root_source_file = xlsxwriter_dep.path("include/xlsxwriter.h"),
        .target = target,
        .optimize = optimize,
    });

    const lxw_mod = lxw.createModule();
    lxw_mod.linkLibrary(xlsxwriter_dep.artifact("xlsxwriter"));
    lxw_mod.link_libc = true;
    lxw_mod.addCMacro("struct_headname", "");

    const xlsxwriter_mod = b.addModule("xlsxwriter", .{
        .root_source_file = b.path("src/xlsxwriter.zig"),
        .target = target,
        .optimize = optimize,
    });
    xlsxwriter_mod.addImport("lxw", lxw_mod);

    const make = b.option([]const u8, "make", "Provide name of example to make example (ex: 'tutorial1'") orelse null;
    if (make) |name| {
        var dbga: std.heap.DebugAllocator(.{}) = .init;
        defer _ = dbga.deinit();
        const alloc = dbga.allocator();

        const example_dir = "examples/wrapped";
        const example_path = std.fmt.allocPrint(alloc, "{s}/{s}.zig", .{ example_dir, name }) catch example_dir ++ "/tutorial1.zig";
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
    example.root_module.addAnonymousImport("assets", .{
        .root_source_file = b.path("examples/assets/embed.zig"),
    });

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

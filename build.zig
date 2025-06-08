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

    const run_step = b.step("run", "run the build example(s)");
    const make = b.option([]const u8, "make", "Provide name of example to make example (ex: 'tutorial1'") orelse null;
    if (make) |name| {
        var dbga: std.heap.DebugAllocator(.{}) = .init;
        defer _ = dbga.deinit();
        const alloc = dbga.allocator();

        const example_dir = "examples/wrapped";

        if (std.mem.eql(u8, "all", name)) {
            // build all the example
            const ex_dir_o = b.build_root.handle.openDir(example_dir, .{ .iterate = true }) catch null;
            var ex_dir = if (ex_dir_o) |dir| dir else return;
            defer ex_dir.close();
            var walker = ex_dir.walk(alloc) catch return;
            defer walker.deinit();
            while (walker.next() catch null) |entry| {
                switch (entry.kind) {
                    std.fs.File.Kind.file => {
                        if (!std.mem.eql(u8, entry.basename, "_helper.zig") and
                            std.mem.eql(u8, entry.basename[entry.basename.len - 4 .. entry.basename.len], ".zig"))
                        {
                            const example_path = std.fmt.allocPrint(alloc, "{s}/{s}", .{ example_dir, entry.basename }) catch example_dir ++ "/tutorial1.zig";
                            defer alloc.free(example_path);

                            makeExample(
                                b,
                                .{
                                    .path = example_path,
                                    .zig_mod = xlsxwriter_mod,
                                    .target = target,
                                    .optimize = optimize,
                                },
                                run_step,
                            );
                        }
                    },
                    else => {},
                }
            }
        } else {
            const example_path = std.fmt.allocPrint(alloc, "{s}/{s}.zig", .{ example_dir, name }) catch example_dir ++ "/tutorial1.zig";
            defer alloc.free(example_path);

            makeExample(
                b,
                .{
                    .path = example_path,
                    .zig_mod = xlsxwriter_mod,
                    .target = target,
                    .optimize = optimize,
                },
                run_step,
            );
        }
    }
}

fn makeExample(b: *std.Build, options: BuildInfo, run_step: *std.Build.Step) void {
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

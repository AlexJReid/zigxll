/// Discovers all Excel functions in a given module at comptime
pub fn getAllFunctions(comptime module: type) []const type {
    return discoverDecls(module, "is_excel_function");
}

/// Discovers all Excel macros (commands) in a given module at comptime
pub fn getAllMacros(comptime module: type) []const type {
    return discoverDecls(module, "is_excel_macro");
}

fn discoverDecls(comptime module: type, comptime marker: []const u8) []const type {
    const type_info = @typeInfo(module);
    const decls = switch (type_info) {
        .@"struct" => |s| s.decls,
        else => @compileError("Expected struct type"),
    };
    comptime var results: []const type = &.{};

    inline for (decls) |decl| {
        const field = @field(module, decl.name);
        const T = @TypeOf(field);
        if (@typeInfo(T) == .type) {
            const ActualType = field;
            if (@typeInfo(ActualType) == .@"struct") {
                if (@hasDecl(ActualType, marker)) {
                    results = results ++ [_]type{ActualType};
                }
            }
        }
    }

    return results;
}

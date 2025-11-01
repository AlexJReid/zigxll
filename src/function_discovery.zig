/// Discovers all Excel functions in a given module at comptime
pub fn getAllFunctions(comptime module: type) []const type {
    const type_info = @typeInfo(module);
    const decls = switch (type_info) {
        .@"struct" => |s| s.decls,
        else => @compileError("Expected struct type"),
    };
    comptime var excel_funcs: []const type = &.{};

    inline for (decls) |decl| {
        const field = @field(module, decl.name);
        const T = @TypeOf(field);
        if (@typeInfo(T) == .type) {
            const ActualType = field;
            if (@typeInfo(ActualType) == .@"struct") {
                if (@hasDecl(ActualType, "is_excel_function")) {
                    excel_funcs = excel_funcs ++ [_]type{ActualType};
                }
            }
        }
    }

    return excel_funcs;
}

#!/usr/bin/env lua
-- lua_introspect.lua — Parse LDoc-style annotations from Lua files, emit Zig
--
-- Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]
--
-- Emits a .zig file with LuaFunction declarations + lua_scripts to stdout.
--
-- Options:
--   --prefix PREFIX      Excel function name prefix (default: "Lua.")
--   --category CAT       Default category (default: "Lua Functions")
--   --embed-root DIR     Base directory for @embedFile paths
--
-- Annotation format (above each global function):
--
--   --- Description of the function
--   -- @param x number First number
--   -- @param y string Name to greet
--   -- @async
--   -- @category My Category
--   -- @name CustomExcelName
--   -- @help_url https://example.com/help
--   function add(x, y) ... end

local function usage()
    io.stderr:write([[
Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]

Options:
  --prefix PREFIX      Excel name prefix (default: "Lua.")
  --category CAT       Default category (default: "Lua Functions")
  --embed-root DIR     Base directory for @embedFile paths

Parses LDoc-style --- annotations and emits a .zig file with
LuaFunction declarations and a lua_scripts constant.
]])
    os.exit(1)
end

-- Parse args
local prefix = "Lua."
local default_category = "Lua Functions"
local embed_root = nil
local files = {}

local i = 1
while i <= #arg do
    local a = arg[i]
    if a == "--prefix" then
        i = i + 1
        prefix = arg[i] or usage()
    elseif a == "--category" then
        i = i + 1
        default_category = arg[i] or usage()
    elseif a == "--embed-root" then
        i = i + 1
        embed_root = arg[i] or usage()
    elseif a == "--help" or a == "-h" then
        usage()
    elseif a:sub(1, 1) == "-" then
        io.stderr:write("Unknown option: " .. a .. "\n")
        usage()
    else
        files[#files + 1] = a
    end
    i = i + 1
end

if #files == 0 then
    io.stderr:write("Error: no input files\n\n")
    usage()
end

-- Helpers
local function zig_str(s)
    return '"' .. s:gsub('\\', '\\\\'):gsub('"', '\\"') .. '"'
end

local function pascal_case(s)
    local parts = {}
    for part in s:gmatch("[^_]+") do
        parts[#parts + 1] = part:sub(1, 1):upper() .. part:sub(2)
    end
    return table.concat(parts)
end

local function script_name(path)
    local base = path:match("([^/\\]+)$") or path
    return base:match("^(.+)%.lua$") or base
end

local function relative_path(root, path)
    if not root then return path end
    if root:sub(-1) ~= "/" then root = root .. "/" end
    if path:sub(1, #root) == root then
        return path:sub(#root + 1)
    end
    return path
end

-- Parse a single file
local function parse_file(path)
    local fh, err = io.open(path, "r")
    if not fh then
        io.stderr:write("Error opening " .. path .. ": " .. err .. "\n")
        os.exit(1)
    end

    local functions = {}
    local doc = nil

    for line in fh:lines() do
        local desc = line:match("^%s*%-%-%-%s*(.*)")
        if desc then
            if not doc then
                doc = {
                    description = nil,
                    params = {},
                    is_async = false,
                    category = nil,
                    help_url = nil,
                    excel_name = nil,
                }
            end
            if not doc.description and desc ~= "" then
                doc.description = desc
            end
        elseif doc and line:match("^%s*%-%-") then
            local tag_line = line:match("^%s*%-%-%s*(.*)")
            if tag_line then
                local pname, rest = tag_line:match("^@param%s+(%S+)%s*(.*)")
                if pname then
                    local ptype = nil
                    local pdesc = nil
                    if rest and rest ~= "" then
                        local first, remainder = rest:match("^(%S+)%s*(.*)")
                        if first == "number" or first == "string" or first == "boolean" then
                            ptype = first
                            pdesc = remainder ~= "" and remainder or nil
                        else
                            pdesc = rest
                        end
                    end
                    doc.params[#doc.params + 1] = {
                        name = pname,
                        type = ptype,
                        description = pdesc,
                    }
                elseif tag_line:match("^@async") then
                    doc.is_async = true
                elseif tag_line:match("^@category%s+") then
                    doc.category = tag_line:match("^@category%s+(.*)")
                elseif tag_line:match("^@name%s+") then
                    doc.excel_name = tag_line:match("^@name%s+(%S+)")
                elseif tag_line:match("^@help_url%s+") then
                    doc.help_url = tag_line:match("^@help_url%s+(%S+)")
                end
            end
        else
            if doc then
                local fname = line:match("^%s*function%s+([%w_]+)%s*%(")
                if fname then
                    doc.id = fname
                    functions[#functions + 1] = doc
                end
                doc = nil
            end
        end
    end

    fh:close()
    return functions
end

-- Collect from all files
local all_functions = {}
for _, file in ipairs(files) do
    local funcs = parse_file(file)
    for _, f in ipairs(funcs) do
        all_functions[#all_functions + 1] = f
    end
end

table.sort(all_functions, function(a, b) return a.id < b.id end)

-- Emit Zig
local out = {}
out[#out + 1] = "// Auto-generated by lua_introspect.lua — do not edit"
out[#out + 1] = 'const xll = @import("xll");'
out[#out + 1] = "const LuaFunction = xll.LuaFunction;"
out[#out + 1] = "const LuaParam = xll.LuaParam;"
out[#out + 1] = ""

for _, func in ipairs(all_functions) do
    local excel_name = func.excel_name or (prefix .. pascal_case(func.id))
    local zig_id = func.id:gsub("[%.%-%s]", "_")

    out[#out + 1] = "pub const " .. zig_id .. " = LuaFunction(.{"
    out[#out + 1] = "    .name = " .. zig_str(excel_name) .. ","
    out[#out + 1] = "    .id = " .. zig_str(func.id) .. ","
    if func.description then
        out[#out + 1] = "    .description = " .. zig_str(func.description) .. ","
    end
    out[#out + 1] = "    .category = " .. zig_str(func.category or default_category) .. ","
    if func.help_url then
        out[#out + 1] = "    .help_url = " .. zig_str(func.help_url) .. ","
    end
    if func.is_async then
        out[#out + 1] = "    .is_async = true,"
    end

    if #func.params > 0 then
        out[#out + 1] = "    .params = &[_]LuaParam{"
        for _, p in ipairs(func.params) do
            local parts = { ".name = " .. zig_str(p.name) }
            if p.type == "string" then
                parts[#parts + 1] = ".type = .string"
            elseif p.type == "boolean" then
                parts[#parts + 1] = ".type = .boolean"
            end
            if p.description then
                parts[#parts + 1] = ".description = " .. zig_str(p.description)
            end
            out[#out + 1] = "        .{ " .. table.concat(parts, ", ") .. " },"
        end
        out[#out + 1] = "    },"
    end

    out[#out + 1] = "});"
    out[#out + 1] = ""
end

out[#out + 1] = "pub const lua_scripts = .{"
for _, file in ipairs(files) do
    local name = script_name(file)
    local embed_path = relative_path(embed_root, file)
    out[#out + 1] = "    .{ .name = " .. zig_str(name) .. ", .source = @embedFile(" .. zig_str(embed_path) .. ") },"
end
out[#out + 1] = "};"
out[#out + 1] = ""

io.write(table.concat(out, "\n") .. "\n")

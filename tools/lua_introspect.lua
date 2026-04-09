#!/usr/bin/env lua
-- lua_introspect.lua — Parse LDoc-style annotations from Lua files, emit Zig
--
-- Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]
--
-- Emits a .zig file with LuaFunction declarations + lua_scripts to stdout.
--
-- Options:
--   --prefix PREFIX            Excel function name prefix (default: "FUNCS.")
--   --category CAT             Default category (default: "Lua Functions")
--   --embed-root DIR           Base directory for @embedFile paths
--   --functions-json FILE      Also write Office JS functions.json to FILE
--
-- Annotation format (above each global function):
--
--   --- Description of the function
--   -- @param x number First number
--   -- @param y string Name to greet
--   -- @async
--   -- @thread_safe false
--   -- @category My Category
--   -- @name CustomExcelName
--   -- @help_url https://example.com/help
--   function add(x, y) ... end

local function usage()
    io.stderr:write([[
Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]

Options:
  --prefix PREFIX            Excel name prefix (default: "FUNCS.")
  --category CAT             Default category (default: "Lua Functions")
  --embed-root DIR           Base directory for @embedFile paths
  --functions-json FILE      Also write Office JS functions.json to FILE

Parses LDoc-style --- annotations and emits a .zig file with
LuaFunction declarations and a lua_scripts constant.
]])
    os.exit(1)
end

-- Parse args
local prefix = "FUNCS."
local default_category = "Lua Functions"
local embed_root = nil
local functions_json_path = nil
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
    elseif a == "--functions-json" then
        i = i + 1
        functions_json_path = arg[i] or usage()
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
                    is_rtd = false,
                    thread_safe = nil,
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
                elseif tag_line:match("^@rtd") then
                    doc.is_rtd = true
                elseif tag_line:match("^@async") then
                    doc.is_async = true
                elseif tag_line:match("^@thread_safe%s+") then
                    local val = tag_line:match("^@thread_safe%s+(%S+)")
                    doc.thread_safe = (val ~= "false")
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
out[#out + 1] = "const LuaRtdFunction = xll.LuaRtdFunction;"
out[#out + 1] = "const LuaParam = xll.LuaParam;"
out[#out + 1] = ""

for _, func in ipairs(all_functions) do
    local excel_name = func.excel_name or (prefix .. pascal_case(func.id))
    local zig_id = func.id:gsub("[%.%-%s]", "_")

    local zig_type = func.is_rtd and "LuaRtdFunction" or "LuaFunction"
    out[#out + 1] = "pub const " .. zig_id .. " = " .. zig_type .. "(.{"
    out[#out + 1] = "    .name = " .. zig_str(excel_name) .. ","
    out[#out + 1] = "    .id = " .. zig_str(func.id) .. ","
    if func.description then
        out[#out + 1] = "    .description = " .. zig_str(func.description) .. ","
    end
    out[#out + 1] = "    .category = " .. zig_str(func.category or default_category) .. ","
    if func.help_url then
        out[#out + 1] = "    .help_url = " .. zig_str(func.help_url) .. ","
    end
    if func.is_async and not func.is_rtd then
        out[#out + 1] = "    .is_async = true,"
    end
    if func.thread_safe ~= nil and not func.is_rtd then
        out[#out + 1] = "    .thread_safe = " .. (func.thread_safe and "true" or "false") .. ","
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

-- Emit Office JS functions.json
if functions_json_path then
    local function json_escape(s)
        return s:gsub('\\', '\\\\'):gsub('"', '\\"'):gsub('\n', '\\n'):gsub('\r', '\\r'):gsub('\t', '\\t')
    end

    local function json_string(s)
        return '"' .. json_escape(s) .. '"'
    end

    local type_map = { number = "number", string = "string", boolean = "boolean" }

    local j = {}
    j[#j + 1] = '{'
    j[#j + 1] = '  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",'
    j[#j + 1] = '  "functions": ['

    for fi, func in ipairs(all_functions) do
        local excel_name = func.excel_name or (prefix .. pascal_case(func.id))
        -- Office JS uses uppercase names by convention
        local display_name = excel_name:gsub("%.", ".")

        j[#j + 1] = '    {'
        j[#j + 1] = '      "id": ' .. json_string(func.id) .. ','
        j[#j + 1] = '      "name": ' .. json_string(display_name) .. ','
        if func.description then
            j[#j + 1] = '      "description": ' .. json_string(func.description) .. ','
        end
        if func.help_url then
            j[#j + 1] = '      "helpUrl": ' .. json_string(func.help_url) .. ','
        end

        -- result
        j[#j + 1] = '      "result": {'
        j[#j + 1] = '        "dimensionality": "scalar"'
        j[#j + 1] = '      },'

        -- parameters
        j[#j + 1] = '      "parameters": ['
        for pi, p in ipairs(func.params) do
            j[#j + 1] = '        {'
            j[#j + 1] = '          "name": ' .. json_string(p.name) .. ','
            if p.description then
                j[#j + 1] = '          "description": ' .. json_string(p.description) .. ','
            end
            j[#j + 1] = '          "type": ' .. json_string(type_map[p.type] or "any") .. ','
            j[#j + 1] = '          "dimensionality": "scalar"'
            j[#j + 1] = '        }' .. (pi < #func.params and ',' or '')
        end
        j[#j + 1] = '      ]'

        -- async / thread_safe
        local is_async = func.is_async
        -- match Zig logic: async forces thread_safe=false, otherwise default true
        local thread_safe
        if is_async then
            thread_safe = false
        elseif func.thread_safe ~= nil then
            thread_safe = func.thread_safe
        else
            thread_safe = true
        end
        if is_async or not thread_safe then
            j[#j] = '      ],'
            if is_async then
                j[#j + 1] = '      "async": true,'
            end
            j[#j + 1] = '      "threadSafe": ' .. (thread_safe and 'true' or 'false')
        end

        j[#j + 1] = '    }' .. (fi < #all_functions and ',' or '')
    end

    j[#j + 1] = '  ]'
    j[#j + 1] = '}'
    j[#j + 1] = ''

    local fh, err = io.open(functions_json_path, "w")
    if not fh then
        io.stderr:write("Error writing " .. functions_json_path .. ": " .. err .. "\n")
        os.exit(1)
    end
    fh:write(table.concat(j, "\n"))
    fh:close()
    io.stderr:write("Wrote " .. functions_json_path .. "\n")
end

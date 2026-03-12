#!/usr/bin/env lua
-- lua_introspect.lua — Introspect Lua functions and emit functions.json to stdout
--
-- Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]
--
-- Options:
--   --prefix PREFIX    Excel function name prefix (default: "Lua.")
--   --category CAT     Category for all functions (default: "Lua Functions")

local function usage()
    io.stderr:write([[
Usage: lua tools/lua_introspect.lua [options] <script.lua> [...]

Options:
  --prefix PREFIX    Excel name prefix (default: "Lua.")
  --category CAT    Category (default: "Lua Functions")

Introspects global functions from Lua scripts and emits JSON
compatible with ZigXLL's lua_json build option.
]])
    os.exit(1)
end

-- Parse args
local prefix = "Lua."
local category = "Lua Functions"
local files = {}

local i = 1
while i <= #arg do
    local a = arg[i]
    if a == "--prefix" then
        i = i + 1
        prefix = arg[i] or usage()
    elseif a == "--category" then
        i = i + 1
        category = arg[i] or usage()
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

-- Collect builtins to exclude
local builtins = {}
for k, v in pairs(_G) do
    if type(v) == "function" then
        builtins[k] = true
    end
end

-- Load scripts into a sandbox
local sandbox = {}
for k, v in pairs(_G) do
    sandbox[k] = v
end

for _, file in ipairs(files) do
    local chunk, err = loadfile(file, "t", sandbox)
    if not chunk then
        io.stderr:write("Error loading " .. file .. ": " .. err .. "\n")
        os.exit(1)
    end
    chunk()
end

-- Collect user-defined functions
local functions = {}
for name, val in pairs(sandbox) do
    if type(val) == "function" and not builtins[name] then
        local info = debug.getinfo(val, "u")
        local params = {}
        for j = 1, info.nparams do
            local pname = debug.getlocal(val, j)
            params[#params + 1] = pname
        end
        functions[#functions + 1] = {
            name = name,
            params = params,
        }
    end
end

-- Sort by name for stable output
table.sort(functions, function(a, b) return a.name < b.name end)

-- JSON helpers
local function json_str(s)
    return '"' .. s:gsub('\\', '\\\\'):gsub('"', '\\"'):gsub('\n', '\\n') .. '"'
end

-- snake_case -> PascalCase
local function pascal_case(s)
    local parts = {}
    for part in s:gmatch("[^_]+") do
        parts[#parts + 1] = part:sub(1, 1):upper() .. part:sub(2)
    end
    return table.concat(parts)
end

-- Emit JSON
local out = {}
out[#out + 1] = "["

for idx, func in ipairs(functions) do
    local excel_name = prefix .. pascal_case(func.name)

    out[#out + 1] = "  {"
    out[#out + 1] = '    "name": ' .. json_str(excel_name) .. ','
    out[#out + 1] = '    "lua_name": ' .. json_str(func.name) .. ','
    out[#out + 1] = '    "category": ' .. json_str(category) .. ','

    if #func.params > 0 then
        out[#out + 1] = '    "params": ['
        for pi, pname in ipairs(func.params) do
            local comma = pi < #func.params and "," or ""
            out[#out + 1] = '      { "name": ' .. json_str(pname) .. ' }' .. comma
        end
        out[#out + 1] = '    ]'
    else
        out[#out + 1] = '    "params": []'
    end

    local comma = idx < #functions and "," or ""
    out[#out + 1] = "  }" .. comma
end

out[#out + 1] = "]"

io.write(table.concat(out, "\n") .. "\n")

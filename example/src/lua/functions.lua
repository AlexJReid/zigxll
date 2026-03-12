function add(x, y)
    return x + y
end

function greet(name)
    return "Hello, " .. name .. "!"
end

function hypotenuse(a, b)
    return math.sqrt(a * a + b * b)
end

function is_even(n)
    return n % 2 == 0
end

function factorial(n)
    if n <= 1 then return 1 end
    local result = 1
    for i = 2, n do
        result = result * i
    end
    return result
end

function fib(n)
    if n <= 0 then return 0 end
    if n == 1 then return 1 end
    local a, b = 0, 1
    for _ = 2, n do
        a, b = b, a + b
    end
    return b
end

-- Thread-safe: called from multiple Excel calc threads concurrently.
-- Each thread gets its own Lua state so there's no contention.
function is_prime(n)
    if n < 2 then return false end
    if n == 2 then return true end
    if n % 2 == 0 then return false end
    for i = 3, math.floor(math.sqrt(n)), 2 do
        if n % i == 0 then return false end
    end
    return true
end

function sum_range(lo, hi)
    local total = 0
    for i = lo, hi do
        total = total + i
    end
    return total
end

-- Async: runs on thread pool, simulates slow work
function slow_fib(n)
    -- deliberate naive recursion to simulate expensive computation
    if n <= 0 then return 0 end
    if n == 1 then return 1 end
    local a, b = 0, 1
    for _ = 2, n do
        a, b = b, a + b
        -- burn some time per iteration
        for _ = 1, 100000 do end
    end
    return b
end

function slow_prime_count(limit)
    local count = 0
    for i = 2, limit do
        local prime = true
        for j = 2, math.floor(math.sqrt(i)) do
            if i % j == 0 then
                prime = false
                break
            end
        end
        if prime then count = count + 1 end
    end
    return count
end

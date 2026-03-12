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

-- Black-Scholes option pricing
-- Zelen & Severo approximation for normal CDF
local function normal_cdf(x)
    local b0 = 0.2316419
    local b1 = 0.319381530
    local b2 = -0.356563782
    local b3 = 1.781477937
    local b4 = -1.821255978
    local b5 = 1.330274429

    if x >= 0 then
        local t = 1.0 / (1.0 + b0 * x)
        local t2 = t * t
        local t3 = t2 * t
        local poly = b1 * t + b2 * t2 + b3 * t3 + b4 * t2 * t2 + b5 * t2 * t3
        return 1.0 - poly * math.exp(-x * x / 2.0) / math.sqrt(2.0 * math.pi)
    else
        return 1.0 - normal_cdf(-x)
    end
end

function bs_call(S, K, T, r, sigma)
    local d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
    local d2 = d1 - sigma * math.sqrt(T)
    return S * normal_cdf(d1) - K * math.exp(-r * T) * normal_cdf(d2)
end

function bs_put(S, K, T, r, sigma)
    local d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
    local d2 = d1 - sigma * math.sqrt(T)
    return K * math.exp(-r * T) * normal_cdf(-d2) - S * normal_cdf(-d1)
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

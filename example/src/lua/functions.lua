--- Add two numbers
-- @param x number First number
-- @param y number Second number
function add(x, y)
    return x + y
end

--- Greet someone by name
-- @param name string Name to greet
function greet(name)
    return "Hello, " .. name .. "!"
end

--- Calculate hypotenuse
-- @param a number Side a
-- @param b number Side b
function hypotenuse(a, b)
    return math.sqrt(a * a + b * b)
end

--- Check if a number is even
-- @param n number Number to check
function is_even(n)
    return n % 2 == 0
end

--- Calculate factorial
-- @param n number Number
function factorial(n)
    if n <= 1 then return 1 end
    local result = 1
    for i = 2, n do
        result = result * i
    end
    return result
end

--- Calculate Fibonacci number
-- @param n number Index
function fib(n)
    if n <= 0 then return 0 end
    if n == 1 then return 1 end
    local a, b = 0, 1
    for _ = 2, n do
        a, b = b, a + b
    end
    return b
end

--- Check if a number is prime
-- @param n number Number to check
function is_prime(n)
    if n < 2 then return false end
    if n == 2 then return true end
    if n % 2 == 0 then return false end
    for i = 3, math.floor(math.sqrt(n)), 2 do
        if n % i == 0 then return false end
    end
    return true
end

--- Sum integers from lo to hi
-- @param lo number Start of range
-- @param hi number End of range
function sum_range(lo, hi)
    local total = 0
    for i = lo, hi do
        total = total + i
    end
    return total
end

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

--- Black-Scholes call option price
-- @name Lua.BS_CALL
-- @param S number Current stock price
-- @param K number Strike price
-- @param T number Time to maturity (years)
-- @param r number Risk-free rate
-- @param sigma number Volatility
function bs_call(S, K, T, r, sigma)
    local d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
    local d2 = d1 - sigma * math.sqrt(T)
    return S * normal_cdf(d1) - K * math.exp(-r * T) * normal_cdf(d2)
end

--- Black-Scholes put option price
-- @name Lua.BS_PUT
-- @param S number Current stock price
-- @param K number Strike price
-- @param T number Time to maturity (years)
-- @param r number Risk-free rate
-- @param sigma number Volatility
function bs_put(S, K, T, r, sigma)
    local d1 = (math.log(S / K) + (r + 0.5 * sigma * sigma) * T) / (sigma * math.sqrt(T))
    local d2 = d1 - sigma * math.sqrt(T)
    return K * math.exp(-r * T) * normal_cdf(-d2) - S * normal_cdf(-d1)
end

--- Fibonacci with simulated delay
-- @param n number Index
-- @async
function slow_fib(n)
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

--- Count primes up to limit
-- @param limit number Upper bound
-- @async
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

--- Subscribe to a timer tick via the demo RTD server.
-- Returns a live-updating counter that increments every 2 seconds.
-- Demonstrates xllify.rtd_subscribe from Lua.
-- @name Lua.TimerTick
-- @thread_safe false
function timer_tick()
    return xllify.rtd_subscribe("zigxll.example.timer")
end

--- Subscribe to a timer tick with a named topic.
-- @name Lua.TimerTickLabeled
-- @param label string Topic label (ignored by the server, illustrates multi-topic)
-- @thread_safe false
function timer_tick_labeled(label)
    return xllify.rtd_subscribe("zigxll.example.timer", label)
end

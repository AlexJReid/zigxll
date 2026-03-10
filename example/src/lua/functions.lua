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

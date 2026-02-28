-- @test(70, 1.75)
-- @test(90, 1.80)
-- @test(50, 1.60)
-- @test(100, 2.00)
xllify.ExcelFunction({
    name = "BMI",
    description = "Calculates Body Mass Index from weight and height",
    category = "xllify.math",
    parameters = {
        { name = "weight", type = "number", description = "Weight in kilograms" },
        { name = "height", type = "number", description = "Height in meters" }
    }
}, function(weight, height)
    if height <= 0 then
        return 0
    end
    return weight / (height * height)
end)

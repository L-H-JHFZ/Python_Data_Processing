import turtle  
import random  
  
# 设置画布大小和背景颜色  
turtle.setup(800, 600)  
turtle.bgcolor("black")  
  
# 定义烟花的颜色和大小  
colors = ["red", "orange", "yellow", "green", "blue", "purple"]  
sizes = [10, 20, 30, 40, 50]  
  
# 定义绘制烟花的函数  
def draw_firework(x, y):  
    # 选择烟花的颜色和大小  
    color = random.choice(colors)  
    size = random.choice(sizes)  
      
    # 设置烟花的颜色和画笔宽度  
    turtle.pencolor(color)  
    turtle.pensize(size)  
      
    # 将画笔移动到烟花的起始位置，隐藏画笔，开始绘制烟花  
    turtle.penup()  
    turtle.goto(x, y)  
    turtle.pendown()  
    turtle.hideturtle()  
      
    # 绘制烟花的主体部分，使用循环绘制多个弧线，形成烟花的效果  
    for i in range(36):  
        turtle.forward(size)  
        turtle.right(10)  
        turtle.forward(size)  
        turtle.right(170)  
        turtle.forward(size)  
        turtle.right(170)  
        turtle.forward(size)  
        turtle.right(10)  
      
    # 绘制烟花的尾部，绘制一个弧形，形成烟花尾部的拖影效果  
    turtle.right(45)  
    turtle.forward(size * 2)  
    turtle.right(135)  
    turtle.forward(size * 2)  
    turtle.right(180)  
    turtle.forward(size * 2)  
    turtle.right(135)  
    turtle.forward(size * 2)  
    turtle.right(45)  
    turtle.forward(size * 2)  
      
# 在画布上随机绘制多个烟花，每个烟花的位置和大小随机生成  
for i in range(30):  
    x = random.randint(-400, 400)  
    y = random.randint(-300, 300)  
    draw_firework(x, y)  
      
# 显示绘制完成的烟花效果，保持画面显示状态，直到用户关闭画布窗口为止  
turtle.done()

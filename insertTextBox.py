import fitz

# 创建文档和页面
doc = fitz.open("./test_insertText.pdf")
# doc = fitz.open()
# page = doc.new_page()
page = doc.load_page(0)

# 定义方框的位置和大小 (x0, y0, x1, y1)
rect = fitz.Rect(100, 100, 200, 150)
text = "123"  # 要插入的数字

# --------------------------
# 1. 绘制带填充颜色的方框
# --------------------------
shape = page.new_shape()
shape.draw_rect(rect)  # 绘制矩形

# 设置填充色（RGB 格式，值范围 0-1）
fill_color = (0.8, 0.5, 0.2)  # 橙色填充
border_color = (0, 0, 0)       # 黑色边框
shape.finish(
    color=border_color,  # 边框颜色
    fill=fill_color,     # 填充颜色
    width=1.5            # 边框线宽
)
shape.commit()  # 提交到页面

# --------------------------
# 2. 在方框中插入居中的数字
# --------------------------
fontsize = 20
# 插入文本（align=1 表示水平居中）
overflow = page.insert_textbox(
    rect,  # 使用同一区域
    text,
    fontsize=fontsize,
    align=1,          # 水平居中
    color=(0.1, 0.2, 0.8),  # 白色字体（与填充色对比）
    fill=(0.3,0.5,0.3),
    stroke_opacity=0.5,
    fontname="helv"   # 字体类型
)

# 检查文本是否溢出
if overflow > 0:
    print("警告：文本溢出，请调整字体大小或方框尺寸！")

# 保存文档
doc.save("colored_textbox.pdf")
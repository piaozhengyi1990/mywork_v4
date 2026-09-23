# -*- coding: utf-8 -*-
"""生成 PdfToWord 应用图标 icon.ico / icon.png

设计理念: PDF 红 -> Word 蓝 渐变圆角背景, 中央一张白色文档页(折角+文字行),
右下角一个转换箭头, 表达"PDF 转成 Word"。
"""

import os
from PIL import Image, ImageDraw, ImageFilter

SIZE = 512
RADIUS = 112

PDF_RED = (229, 72, 77)
WORD_BLUE = (43, 87, 154)
WHITE = (255, 255, 255)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def build_page_layer():
    """白色文档页: 折角 + 文字行"""
    layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # 页面矩形(带折角)
    left, top, right, bottom = 128, 96, 384, 416
    fold = 72
    page = [
        (left, top),
        (right - fold, top),
        (right, top + fold),
        (right, bottom),
        (left, bottom),
    ]
    d.polygon(page, fill=WHITE)

    # 折角三角(浅灰)
    d.polygon([(right - fold, top), (right, top + fold), (right - fold, top + fold)],
              fill=(226, 232, 240))

    # 文字行(灰色条)
    line_color = (148, 163, 184)
    y = top + 120
    widths = [200, 220, 180, 210, 150]
    for w in widths:
        d.rounded_rectangle([left + 32, y, left + 32 + w, y + 22],
                            radius=10, fill=line_color)
        y += 44

    return layer


def build_arrow_layer():
    """右下角圆形箭头(表示转换)"""
    layer = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    cx, cy, r = 368, 388, 78
    # 白色圆底 + 蓝色描边
    d.ellipse([cx - r, cy - r, cx + r, cy + r], fill=WHITE, outline=WORD_BLUE, width=12)

    # 双向折线箭头: PDF->Word 语义用右向箭头 + 文档小方块
    blue = WORD_BLUE
    # 横向箭头
    d.line([(cx - 40, cy), (cx + 34, cy)], fill=blue, width=16)
    d.polygon([(cx + 44, cy), (cx + 16, cy - 24), (cx + 16, cy + 24)], fill=blue)

    return layer


def build_icon():
    img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))

    # 背景: 垂直渐变圆角矩形
    grad = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grad)
    for y in range(SIZE):
        color = lerp(PDF_RED, WORD_BLUE, y / SIZE)
        gd.line([(0, y), (SIZE, y)], fill=color)
    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, SIZE - 1, SIZE - 1],
                                           radius=RADIUS, fill=255)
    img.paste(grad, (0, 0), mask)

    # 文档页 + 柔和阴影
    page = build_page_layer()
    shadow = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    shadow.paste((0, 0, 0, 90), (0, 0), page)
    shadow = shadow.filter(ImageFilter.GaussianBlur(14))
    img.alpha_composite(shadow, (6, 8))
    img.alpha_composite(page)

    # 转换箭头
    img.alpha_composite(build_arrow_layer())

    return img


def build_toolbar_icons():
    """生成工具栏小图标 (icons/*.png), 供 GUI 按钮使用"""
    icons_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")
    os.makedirs(icons_dir, exist_ok=True)
    S = 16
    GRAY = (75, 85, 99, 255)
    WHITE = (255, 255, 255, 255)

    def new_canvas():
        im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
        return im, ImageDraw.Draw(im)

    def save(im, name):
        im.save(os.path.join(icons_dir, f"{name}.png"))

    # add: +
    im, d = new_canvas()
    d.line([(3, 8), (13, 8)], fill=GRAY, width=2)
    d.line([(8, 3), (8, 13)], fill=GRAY, width=2)
    save(im, "add")

    # folder
    im, d = new_canvas()
    d.polygon([(2, 5), (6, 5), (7, 3), (12, 3), (13, 5), (14, 5), (14, 13), (2, 13)],
              outline=GRAY, width=1)
    d.line([(2, 6), (14, 6)], fill=GRAY, width=1)
    save(im, "folder")

    # check-all
    im, d = new_canvas()
    d.rectangle([(2, 2), (14, 14)], outline=GRAY, width=1)
    d.line([(4, 8), (7, 11), (12, 5)], fill=GRAY, width=2)
    save(im, "check")

    # uncheck
    im, d = new_canvas()
    d.rectangle([(2, 2), (14, 14)], outline=GRAY, width=1)
    save(im, "uncheck")

    # play (white, for primary btn)
    im, d = new_canvas()
    d.polygon([(4, 2), (14, 8), (4, 14)], fill=WHITE)
    save(im, "play")

    # stop (white)
    im, d = new_canvas()
    d.rectangle([(3, 3), (13, 13)], fill=WHITE)
    save(im, "stop")

    # retry
    im, d = new_canvas()
    d.arc([(3, 3), (13, 13)], 40, 320, fill=GRAY, width=2)
    d.polygon([(10, 1), (15, 5), (10, 6)], fill=GRAY)
    save(im, "retry")

    # trash
    im, d = new_canvas()
    d.line([(3, 4), (13, 4)], fill=GRAY, width=1)
    d.line([(6, 2), (10, 2)], fill=GRAY, width=1)
    d.rectangle([(4, 5), (12, 14)], outline=GRAY, width=1)
    d.line([(7, 7), (7, 12)], fill=GRAY, width=1)
    d.line([(9, 7), (9, 12)], fill=GRAY, width=1)
    save(im, "trash")

    print(f"工具栏图标已生成: {icons_dir}")
    return icons_dir


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    img = build_icon()
    ico_path = os.path.join(here, "icon.ico")
    img.save(ico_path, sizes=[(256, 256), (128, 128), (64, 64),
                              (48, 48), (32, 32), (24, 24), (16, 16)])
    png_path = os.path.join(here, "icon.png")
    img.save(png_path)
    print(f"图标已生成: {ico_path}")
    print(f"PNG预览:    {png_path}")
    build_toolbar_icons()

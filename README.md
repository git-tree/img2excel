# img2excel

图片 → Excel 像素画转换器。将任意图片转换为 Excel 单元格彩色填充的像素画，使用 `asyncio` 协程 + 线程池并发处理像素行。

## 效果预览

| 原图 | 转换效果 |
| :--: | :--: |
| <img src="https://s1.ax1x.com/2022/03/22/qKrKW4.png" width="300"> | <img src="./result/a.xlsx" width="300"> |

## 环境要求

- Python >= 3.9
- 依赖库：
  - `Pillow` —— 图片处理
  - `openpyxl` —— Excel 写入

## 安装

```bash
pip install Pillow openpyxl
```

## 使用方法

### 命令行

```bash
python img2excel_user.py <图片路径> [-o 输出路径] [-w 宽度] [--height 高度]
```

**参数说明：**

| 参数 | 必填 | 默认值 | 说明 |
| --- | :-: | :-: | --- |
| `image` | 是 | — | 输入图片路径（支持 png / jpg / bmp 等常见格式） |
| `-o`, `--output` | 否 | `输入名_pixel.xlsx` | 输出 Excel 文件路径 |
| `-w`, `--width` | 否 | `100` | 最大像素宽度（越大越精细） |
| `--height` | 否 | `100` | 最大像素高度（越大越精细） |

**示例：**

```bash
# 最简用法
python img2excel_user.py photo.jpg

# 指定输出路径
python img2excel_user.py photo.jpg -o output.xlsx

# 更高精度（200x200 像素）
python img2excel_user.py photo.jpg -w 200 --height 200
```

### 代码调用

```python
import asyncio
from img2excel_user import image_to_excel

asyncio.run(image_to_excel("input.png", "output.xlsx", max_width=100, max_height=100))
```

## 实现原理

1. **图片加载与缩放**：使用 Pillow 将图片转换为 RGB 模式，并按比例缩放到目标尺寸。
2. **并发像素提取**：通过 `asyncio` + `ThreadPoolExecutor` 并行提取每一行的像素颜色值，充分利用多核 CPU。
3. **Excel 写入**：使用 `openpyxl` 创建工作簿，将每个单元格填充为对应像素的十六进制颜色（`PatternFill`）。
4. **样式调整**：设置列宽 2.5、行高 15，使单元格呈正方形，保证像素画效果。

## 注意事项

- 像素越大（宽高越大），Excel 文件越大，渲染时间越长。建议首次使用 `100x100` 测试。
- Excel 打开大文件可能卡顿，建议最终效果控制在 300×300 以内。
- Windows 下已自动修复终端 UTF-8 编码问题。

## 项目结构

```
img2excel/
├── img2excel_user.py   # 主程序
├── result/            # 示例输出目录
├── LICENSE
└── README.md
```

## License

MIT
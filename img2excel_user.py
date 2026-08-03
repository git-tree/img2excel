"""
图片 → Excel 像素画转换器
使用 asyncio 协程并发处理像素行，将彩色照片转为 Excel 单元格填充色。
"""

import asyncio
import os
import sys
from concurrent.futures import ThreadPoolExecutor

from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# Windows 终端 UTF-8 编码修复
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def extract_row(y: int, width: int, pixels) -> list[tuple[int, int, int, int, int]]:
    """提取一行像素的 RGB 值（CPU 密集型，在线程池中执行）"""
    return [(x, y, *pixels[x, y]) for x in range(width)]


def build_hex_color(r: int, g: int, b: int) -> str:
    """将 RGB 转为 openpyxl 使用的 6 位十六进制颜色字符串"""
    return f"{r:02X}{g:02X}{b:02X}"


async def image_to_excel(
    image_path: str,
    output_path: str,
    max_width: int = 100,
    max_height: int = 100,
) -> tuple[int, int]:
    """
    将图片转换为 Excel 像素画。

    Args:
        image_path: 输入图片路径
        output_path: 输出 xlsx 路径
        max_width: 最大像素宽度（越大越精细，但文件也越大）
        max_height: 最大像素高度
    """
    print(f"[1/4] 加载图片: {image_path}")
    img = Image.open(image_path).convert("RGB")
    original_size = img.size
    img.thumbnail((max_width, max_height), Image.LANCZOS)
    width, height = img.size
    print(f"  原始尺寸: {original_size[0]}x{original_size[1]} → 缩放后: {width}x{height} ({width * height} 像素)")

    pixels = img.load()

    # ---------- 阶段 1: 并发提取所有像素颜色 ----------
    print(f"[2/4] 并发提取像素颜色 ({height} 行)...")
    loop = asyncio.get_running_loop()
    max_workers = min(os.cpu_count() or 4, height)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        tasks = [
            loop.run_in_executor(executor, extract_row, y, width, pixels)
            for y in range(height)
        ]
        all_rows = await asyncio.gather(*tasks)

    print(f"  提取完成, 共 {sum(len(r) for r in all_rows)} 个像素点")

    # ---------- 阶段 2: 写 Excel ----------
    print("[3/4] 写入 Excel（填充单元格颜色）...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Pixel Art"

    # 设置单元格为方形
    for x in range(width):
        ws.column_dimensions[get_column_letter(x + 1)].width = 2.5
    for y in range(height):
        ws.row_dimensions[y + 1].height = 15

    # 批量填充颜色
    total = width * height
    report_every = max(total // 10, 1)
    count = 0
    for row_data in all_rows:
        for x, y, r, g, b in row_data:
            hex_color = build_hex_color(r, g, b)
            cell = ws.cell(row=y + 1, column=x + 1)
            cell.fill = PatternFill(
                start_color=hex_color, end_color=hex_color, fill_type="solid"
            )
            count += 1
            if count % report_every == 0:
                pct = count * 100 // total
                print(f"  进度: {count}/{total} ({pct}%)")

    print("[4/4] 保存文件...")
    wb.save(output_path)
    file_size = os.path.getsize(output_path)
    print(f"  完成! 输出: {output_path} ({file_size / 1024:.1f} KB)")
    return width, height


def main():
    import argparse

    parser = argparse.ArgumentParser(description="图片 → Excel 像素画转换器")
    parser.add_argument("image", help="输入图片路径")
    parser.add_argument("-o", "--output", default=None, help="输出 xlsx 路径（默认: 输入名_pixel.xlsx）")
    parser.add_argument("-w", "--width", type=int, default=100, help="最大像素宽度 (默认: 100)")
    parser.add_argument("--height", type=int, default=100, help="最大像素高度 (默认: 100)")
    args = parser.parse_args()

    if not os.path.exists(args.image):
        print(f"错误: 文件不存在 — {args.image}", file=sys.stderr)
        sys.exit(1)

    output = args.output or f"{os.path.splitext(args.image)[0]}_pixel.xlsx"

    asyncio.run(image_to_excel(args.image, output, args.width, args.height))


if __name__ == "__main__":
    main()

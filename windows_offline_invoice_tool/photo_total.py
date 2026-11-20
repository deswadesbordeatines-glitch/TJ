"""Offline OCR of the printed gross total. One input image produces one cell."""
import re
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill


def read_total(lines):
    boxes = []
    for points, text, score in lines or []:
        xs, ys = zip(*points)
        boxes.append((min(xs), min(ys), max(xs), max(ys), re.sub(r"\s+", "", text), score))
    candidates = []
    money = r"[¥￥]?(-?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2})元?"
    for x, y, right, bottom, label, confidence in boxes:
        if "小写" not in label:
            continue
        suffix = label.split("小写")[-1].lstrip(")）：:")
        possible = [(suffix, confidence)]
        height = max(1, bottom - y)
        for left, top, _, end, text, score in boxes:
            if right - height / 2 <= left <= right + 8 * height and abs(top + end - y - bottom) <= 1.4 * height:
                possible.append((text, min(score, confidence)))
        for text, score in possible:
            match = re.fullmatch(money, text)
            if match:
                candidates.append((Decimal(match[1].replace(",", "")), score))
    if len({value for value, _ in candidates}) != 1:
        return None, "未定位到唯一的小写价税合计，请对照原图填写"
    value, score = max(candidates, key=lambda pair: pair[1])
    return value, "小写金额 OCR 置信度偏低，请核对" if score < 0.98 else ""


def export_photos(paths, destination, report):
    from rapidocr_onnxruntime import RapidOCR
    engine = RapidOCR(intra_op_num_threads=2, inter_op_num_threads=1)
    book = Workbook()
    sheet = book.active
    sheet.title = "发票情况"
    sheet.append(["发票总金额（价税合计）"])
    sheet["A1"].font = Font(bold=True)
    sheet["A1"].fill = PatternFill("solid", fgColor="CDE6D0")
    sheet.column_dimensions["A"].width = 32
    review = 0
    for index, path in enumerate(paths, 2):
        try:
            lines, _ = engine(Path(path).read_bytes())
            value, reason = read_total(lines)
        except Exception as exc:
            value, reason = None, f"照片读取失败，请复核。{exc}"
        cell = sheet.cell(index, 1, value)
        cell.number_format = '#,##0.00'
        cell.comment = Comment(str(path) + "\n" + (reason or "取自票面小写价税合计"), "发票识别工具")
        if reason:
            review += 1
            cell.fill = PatternFill("solid", fgColor="FFF2A8")
        report(f"{Path(path).name}  {value if value is not None else '待填写'}  {reason}")
    book.save(destination)
    return review

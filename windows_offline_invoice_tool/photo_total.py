"""Offline photo ledger. One invoice occupies one row."""
import re
from decimal import Decimal
from pathlib import Path
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HEADERS = ["发票编号", "开票时间", "开票单位", "发票单价", "发票数量", "计量单位", "发票金额", "发票凭证编号", "合同编号"]


def extract_fields(lines):
    boxes = []
    for points, text, score in lines or []:
        xs, ys = zip(*points)
        boxes.append(((min(xs)+max(xs))/2, (min(ys)+max(ys))/2, max(ys)-min(ys),
                      re.sub(r"\s+", "", text), score))

    def labeled(pattern):
        matches = [(m[1], s) for _, _, _, t, s in boxes if (m := re.search(pattern, t, re.I))]
        if len({v for v, _ in matches}) == 1:
            value, score = max(matches, key=lambda item: item[1])
            return value, "" if score >= .98 else "OCR 置信度偏低"
        return None, "未识别到唯一的带标签字段"

    number = labeled(r"(?:发票号码|发票编号|No\.?)[：:]?(\d{20}|\d{8})(?!\d)")
    issued, issue_note = labeled(r"开票日期[：:]?(\d{4}年\d{1,2}月\d{1,2}日)")
    if issued:
        try:
            issued = date(*map(int, re.findall(r"\d+", issued)))
        except ValueError:
            issued, issue_note = None, "日期无效"
    seller = labeled(r"(?:销售方|销方|开票单位)(?:名称)?[：:]([^：:]+)")
    if not seller[0]:
        # 纸票销方在价税合计下方。位置推断始终提示复核。
        totals_y = [y for _, y, _, t, _ in boxes if "小写" in t]
        companies = {t.split("：")[-1] for _, y, _, t, _ in boxes
                     if totals_y and y > max(totals_y) and t.endswith("公司")}
        if len(companies) == 1:
            seller = companies.pop(), "按票面下方销方位置识别，请核对"
    headers = {t:(x,y,h) for x,y,h,t,_ in boxes if t in ("单位","数量","单价")}
    price = quantity = unit = (None, "明细列结构不清，请核对")
    if len(headers) == 3:
        ux, qx, px = [headers[t][0] for t in ("单位","数量","单价")]
        y0 = max(v[1] for v in headers.values())
        stop = min((y for _,y,_,t,_ in boxes if y > y0 and
                    (t in ("合计","计") or "小写" in t or t.startswith(("¥","￥")))), default=y0)
        rows = [(x,y,h,t,s) for x,y,h,t,s in boxes if y0 < y < stop]
        prices = [(y,h,Decimal(t)) for x,y,h,t,_ in rows
                  if (qx+px)/2 < x < px+(px-qx)*.8 and re.fullmatch(r"-?\d+(?:\.\d+)?", t)]
        if prices:
            price = (prices[0][2], "按明细位置识别，请核对") if len(prices)==1 else (None,"多条明细，请人工确认单价")
            quantities, units = [], []
            for y,h,_ in prices:
                peers = [(x,t) for x,yy,_,t,_ in rows if abs(yy-y)<h*.65]
                qs = [Decimal(t) for x,t in peers if (ux+qx)/2 < x <= (qx+px)/2 and re.fullmatch(r"-?\d+(?:\.\d+)?",t)]
                us = [t for x,t in peers if abs(x-ux)<abs(qx-ux)*.48 and re.fullmatch(r"[\u4e00-\u9fffA-Za-z]{1,4}",t)]
                if len(qs)==1:
                    quantities.extend(qs)
                if len(us)==1:
                    units.extend(us)
            if len(units)==len(prices) and len(set(units))==1:
                unit = units[0], "按明细位置识别，请核对"
                if len(quantities)==len(prices):
                    quantity = sum(quantities), "同单位明细数量合计，请核对"
    contract, note = labeled(r"合同(?:编号|号)[：:]?([^；;，,。]+)")
    if contract:
        note = "请对照原图备注区核实完整编号，尤其是印章遮挡处及末尾字符"
    return [number,(issued,issue_note),seller,price,quantity,unit,read_total(lines),
            labeled(r"发票凭证编号[：:]([A-Za-z0-9-]+)"), (contract, note)]


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
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    sheet["A1"] = "发票情况"
    sheet["A1"].alignment = Alignment(horizontal="center")
    sheet.append(HEADERS)
    sheet.freeze_panes = "A3"
    for i, header in enumerate(sheet[2],1):
        header.fill = PatternFill("solid", fgColor="CDE6D0")
        sheet.column_dimensions[get_column_letter(i)].width = 36 if i==3 else 22
    sheet["A1"].font = Font(bold=True)
    sheet["A1"].fill = PatternFill("solid", fgColor="CDE6D0")
    sheet.column_dimensions["A"].width = 32
    review = 0
    for index, path in enumerate(paths, 3):
        try:
            lines, _ = engine(Path(path).read_bytes())
            fields = extract_fields(lines)
        except Exception as exc:
            fields = [(None, f"照片读取失败，请复核。{exc}")] * len(HEADERS)
        pending = []
        for col,(value, reason) in enumerate(fields,1):
            cell = sheet.cell(index,col,value)
            if isinstance(value,str):
                cell.data_type = "s"
            if col == 7:
                cell.number_format = '#,##0.00'
            cell.comment = Comment(str(path) + "\n" + reason, "发票识别工具")
            if reason:
                pending.append(HEADERS[col-1])
                cell.fill = PatternFill("solid", fgColor="FFF2A8")
        review += bool(pending)
        report(f"{Path(path).name}  待复核字段：{'、'.join(pending) or '无'}")
    book.save(destination)
    return review

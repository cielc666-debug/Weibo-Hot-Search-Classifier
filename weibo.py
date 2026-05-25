"""
Weibo HTML → Free Google Translate → Excel (Catch ALL Rows Including Dots)
=============================================
Reads Weibo hot-search HTML files saved by SingleFile,
calls a free Google Translate interface to translate Chinese → English,
then appends data to a master Excel workbook with standard clean formatting.

Requirements:
    pip install beautifulsoup4 openpyxl pandas deep-translator
"""

import argparse
import glob
import sys
import time
from datetime import datetime
from pathlib import Path

from bs4 import BeautifulSoup, Comment
from deep_translator import GoogleTranslator
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

# ── 基础标准边框 ────────────────────────────────────────────────
THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# ════════════════════════════════════════════════════════════════
# 1. HTML PARSING (抓取绝对所有行，包含圆点行)
# ════════════════════════════════════════════════════════════════

def extract_date_from_html(soup) -> datetime | None:
    """Pull the SingleFile 'saved date' comment."""
    for c in soup.find_all(string=lambda t: isinstance(t, Comment)):
        import re
        m = re.search(r"saved date:\s*(.+?)(?:\n|$)", str(c))
        if m:
            raw = m.group(1).strip()
            raw = re.sub(r"\s+GMT[^\s]*.*$", "", raw)
            try:
                return datetime.strptime(raw, "%a %b %d %Y %H:%M:%S")
            except ValueError:
                pass
    return None


def parse_html_file(path: str) -> dict:
    """Return {'date': datetime, 'entries': [{rank, chinese, engagement}]}."""
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    date_val = extract_date_from_html(soup)
    if date_val is None:
        stem = Path(path).stem
        try:
            date_val = datetime.strptime(stem, "%Y-%m-%d_%H%M")
        except ValueError:
            date_val = datetime.now()

    entries = []
    current_rank = 1  # 从1开始，为抓到的每一行数据重新排顺序号
    
    # 抓取表格中的所有行
    for row in soup.select("tbody tr"):
        td02 = row.select_one(".td-02")
        if not td02:
            continue

        # 【核心修改】彻底删除了对 action-type="realtimehot_ad" 的过滤逻辑
        # 从而保证带圆点的商业热搜行也能被完整保留并抓取

        a_tag = td02.select_one("a")
        if not a_tag:
            continue
            
        chinese = a_tag.get_text(strip=True)
        if not chinese or chinese == "查看更多": # 过滤网页底部的翻页残留提示
            continue

        # 抓取热度数字
        span = td02.select_one("span")
        engagement_raw = span.get_text(strip=True) if span else ""
        engagement = None
        if engagement_raw.strip():
            try:
                import re
                num_match = re.search(r'\d+', engagement_raw.replace(",", ""))
                if num_match:
                    engagement = int(num_match.group())
            except ValueError:
                pass

        # 只要存在有效话题文本，就全数收录，重新生成顺序 Rank
        entries.append({
            "rank": current_rank,
            "chinese": chinese,
            "engagement": engagement,
        })
        current_rank += 1

    return {"date": date_val, "entries": entries, "source": Path(path).name}


# ════════════════════════════════════════════════════════════════
# 2. FREE TRANSLATION ENGINE (NO KEY)
# ════════════════════════════════════════════════════════════════

def translate_entries(entries: list[dict]) -> list[dict]:
    """Translate all entries using free Google Translate interface with Title Case."""
    enriched = []
    translator = GoogleTranslator(source='zh-CN', target='en')
    
    total = len(entries)
    print(f"  Starting free Google Translation for {total} topics…")
    
    for i, entry in enumerate(entries, 1):
        out = dict(entry)
        chinese_text = entry["chinese"]
        
        # 清洗热搜标签
        clean_text = chinese_text.strip('#')
        
        try:
            # 执行翻译并转换首字母大写
            translated_text = translator.translate(clean_text)
            if translated_text:
                translated_text = translated_text.title()
            
            out["english"] = translated_text
            print(f"    [{i}/{total}] Translated: {chinese_text} -> {translated_text}")
        except Exception as e:
            print(f"    ⚠ [{i}/{total}] Translation failed for '{chinese_text}': {e}")
            out["english"] = chinese_text  # 失败了用中文保底
            
        enriched.append(out)
        
        if i < total:
            time.sleep(0.3)
            
    return enriched


# ════════════════════════════════════════════════════════════════
# 3. EXCEL OUTPUT (STANDARD & APPEND MODE WITH 65-HEIGHT BLANK ROW)
# ════════════════════════════════════════════════════════════════

def _style_standard(ws, row, col, value=None, bold=False, align="left", wrap=False, fmt=None):
    """标准的 Excel 单元格基础样式，去除了所有背景颜色填充"""
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    c.font = Font(bold=bold, name="Arial", size=10)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    return c


def append_to_excel(html_data_list, output_path):
    """
    延续已有的 Excel 文件并继续往下生成。
    并在每份新的 HTML 数据之间插入一个行高为 65 的空行隔开。
    """
    file_path = Path(output_path)
    
    # 1. 延续已有文件或创建新文件
    if file_path.exists():
        print(f"📊  Found existing file. Appending to {output_path}...")
        wb = openpyxl.load_workbook(output_path)
        ws = wb.active
        
        # 寻找真正的、有文本内容的最大行，防止被空白边框行干扰
        real_max_row = 1
        for row in range(ws.max_row, 0, -1):
            if any(ws.cell(row=row, column=col).value is not None for col in range(1, 6)):
                real_max_row = row
                break
        is_empty_file = False
    else:
        print(f"📊  Creating a new workbook: {output_path}")
        wb = Workbook()
        ws = wb.active
        ws.title = "All Data"
        real_max_row = 1
        is_empty_file = True

    # 2. 如果是全新的空文件，先写表头
    if is_empty_file or real_max_row == 1:
        headers = [
            "Date", 
            "Rank", 
            "Topic Headline: Chinese", 
            "Topic Headline: English Translation", 
            "Engagement"
        ]
        for c, label in enumerate(headers, 1):
            _style_standard(ws, 1, c, label, bold=True, align="center", wrap=True)
        ws.row_dimensions[1].height = 28
        real_max_row = 1

    # 3. 循环追加每一份 HTML 的数据
    for index, file_data in enumerate(html_data_list):
        
        # 判断何时需要插入 65 高度的空行：
        if (real_max_row > 1 and index == 0) or (index > 0):
            empty_row = real_max_row + 1
            
            # 建立一个标准空行隔开
            for c in range(1, 6):
                _style_standard(ws, empty_row, c, value="")
            # 设置这一行空行的高度为 65
            ws.row_dimensions[empty_row].height = 65
            
            start_row = empty_row + 1
        else:
            start_row = real_max_row + 1

        # 写入该 HTML 文件的具体条目
        for r_offset, row in enumerate(file_data):
            er = start_row + r_offset
            
            _style_standard(ws, er, 1, row["date"], align="center", fmt="YYYY-MM-DD HH:MM")
            _style_standard(ws, er, 2, row["rank"], align="center")
            _style_standard(ws, er, 3, row["chinese"], wrap=True)
            _style_standard(ws, er, 4, row["english"], wrap=True)
            
            eng = row.get("engagement")
            _style_standard(ws, er, 5, int(eng) if eng else None, align="right", fmt="#,##0")
        
        # 更新当前表格真正的最大行，供下一个 HTML 文件判断位置使用
        real_max_row = start_row + len(file_data) - 1

    # 4. 设置标准列宽
    widths = {1: 18, 2: 8, 3: 35, 4: 50, 5: 16}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w

    ws.freeze_panes = "A2"
    wb.save(output_path)


# ════════════════════════════════════════════════════════════════
# 4. MAIN
# ════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="Weibo HTML → Free Google Translate → Excel (Catch All Rows)"
    )
    parser.add_argument("--input",  "-i", required=True,
                        help="HTML file or folder containing .html files")
    parser.add_argument("--output", "-o", default="weibo_master.xlsx",
                        help="Output Excel file")
    args = parser.parse_args()

    p = Path(args.input)
    if p.is_file():
        html_files = [str(p)]
    elif p.is_dir():
        html_files = sorted(glob.glob(str(p / "*.html")) + glob.glob(str(p / "*.htm")))
        if not html_files:
            sys.exit(f"No HTML files found in {args.input}")
    else:
        sys.exit(f"Path not found: {args.input}")

    print(f"\n📂  Found {len(html_files)} HTML file(s)\n")

    html_data_list = []
    
    for path in html_files:
        print(f"  Parsing: {Path(path).name}")
        parsed = parse_html_file(path)
        entries = parsed["entries"]
        print(f"    → {len(entries)} hot-search entries, date={parsed['date']}")

        if not entries:
            continue

        # 翻译
        translated = translate_entries(entries)

        # 补全日期
        for entry in translated:
            entry["date"] = parsed["date"]

        html_data_list.append(translated)
        print()

    if not html_data_list:
        sys.exit("No new data extracted.")

    # 执行写入/追加 Excel
    total_rows = sum(len(x) for x in html_data_list)
    print(f"📊  Appending {total_rows} total entries to Excel…")
    append_to_excel(html_data_list, args.output)
    print(f"\n✅  Saved/Appended successfully to: {args.output}")


if __name__ == "__main__":
    main()
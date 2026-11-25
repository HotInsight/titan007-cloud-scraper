"""8波北单 → 兼容原 Titan007 工作流的十列工作簿。"""
from __future__ import annotations

import argparse
import logging
import re
import sys
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from playwright.sync_api import expect, sync_playwright

URL = "https://8bo.com/football/"
ROOT = Path(__file__).resolve().parent
HEADERS = ["赛事名称", "单场编号", "主队", "客队", "亚盘主队水位", "欧赔主胜",
           "亚洲盘口", "欧赔平局", "亚盘客队水位", "欧赔客胜"]


def select(item, wanted=True):
    expect(item).to_have_count(1)
    if (item.get_attribute("data-ai-vis") == "1") != wanted:
        item.click(timeout=10000)
    assertion = expect(item)
    if wanted:
        assertion.to_have_attribute("data-ai-vis", "1")
    else:
        assertion.not_to_have_attribute("data-ai-vis", "1")


def configure(page):
    expect(page.locator('a[ai-link="football"][data-ai-vis="1"]').filter(has_text="比分").first).to_be_attached()
    select(page.locator('#live_mode_switch [ai-value="-4"]'))
    select(page.locator('#live_type_switch [ai-value="2"]'))
    crown = page.locator('#live_odds_list [ai-value="1"][data-title="全部"]')
    if crown.get_attribute("data-ai-vis") != "1" and not crown.is_visible():
        page.locator("#live_odds_list").click()
    select(crown)
    cfg = page.locator("#live_config")
    if not cfg.locator("#live_config_odds").is_visible():
        cfg.click()
    for group in ("set_goal", "set_cred"):
        for value in ("sound", "win"):
            select(cfg.locator(f'[ai-store-group="{group}"] label[ai-value="{value}"]'), False)
    for value in "set_cyellow set_rank set_half set_corner set_remark set_top set_seal".split():
        select(cfg.locator(f'#live_config_other label[ai-value="{value}"]'), False)
    for value in ("1x2", "ah", "ou"):
        select(cfg.locator(f'#live_config_odds label[ai-value="set_odds_{value}"]'), value != "ou")
    for group, value in (("set_odds_names", "ah"), ("set_pkstyle", "ah_num")):
        select(cfg.locator(f'[ai-store-radio="{group}"] label[ai-value="{value}"]'))
    cfg.locator("button.a0close").click()


READ = r"""() => {
  const total = document.querySelector('#live_total [z0-total2]')?.getAttribute('z0-total2');
  const rows = [];
  for (const r of document.querySelectorAll('#live_source ul.r0data[id^="ev_"]')) {
    const num = r.querySelector('.c0number')?.getAttribute('z0-type2');
    if (!num || !r.getClientRects().length) continue;
    const text = sel => r.querySelector(sel)?.textContent.trim() || '';
    const odds = type => {
      const group = r.querySelector('.c0odds [z0-odds="' + type + '"]');
      if (!group) return ['', '', ''];
      const slots = [...group.querySelectorAll(':scope > span')];
      if (!slots.length && !group.textContent.trim()) return ['', '', ''];
      if (slots.length !== 3) throw Error(num + ': unexpected odds slots: ' + type);
      const vals = slots.map(el => el.textContent.trim());
      if (type === 'ah' && vals[1] && !vals[1].includes('受')) {
        const line = slots[1].getAttribute('d-pk') || '';
        const before = getComputedStyle(slots[1], '::before').content;
        if ((Number(line) < 0 || before.includes('受')) && /\d/.test(vals[1]))
          vals[1] = '受' + vals[1];
      }
      return vals;
    };
    const a = odds('ah'), e = odds('1x2');
    rows.push([text('.c0match'),num,text('.c0home .team-name'),text('.c0away .team-name'),
               a[0],e[0],a[1],e[1],a[2],e[2]]);
  }
  return {total,rows};
}"""


def extract(page, timeout=60):
    end = time.monotonic() + timeout
    previous = None
    while True:
        snap = page.evaluate(READ)
        rows, count = snap["rows"], snap["total"]
        if rows and count == str(len(rows)):
            keys = []
            for row in rows:
                if not all(row[:4]) or not re.fullmatch(r"[0-9]+", row[1]):
                    raise ValueError(f"北单编号或球队不完整：{row[:4]}")
                keys.append(tuple(row[:4]))
            if len(keys) != len(set(keys)):
                raise ValueError("发现重复比赛，停止输出")
            # Require two consecutive complete snapshots, with both markets loaded somewhere.
            if keys == previous and all(any(r[i] not in ("", "-", "--") for r in rows)
                                        for i in (4, 5)):
                for row in rows:
                    for i in (4, 5, 7, 8, 9):
                        try:
                            row[i] = float(row[i])
                        except ValueError:
                            pass
                return rows
            previous = keys
        else:
            previous = None
        if time.monotonic() >= end:
            raise RuntimeError(f"北单名单或指数未就绪：网页场数={count}，读取={len(rows)}；未导出")
        page.wait_for_timeout(900)


def scrape(headless=True):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        try:
            page = browser.new_page(viewport={"width": 1440, "height": 1000}, locale="zh-CN")
            page.goto(URL, wait_until="domcontentloaded", timeout=90000)
            page.locator("#live_source ul.r0data").first.wait_for(state="attached", timeout=90000)
            configure(page)
            return extract(page)
        finally:
            browser.close()


def save_excel(rows, output_dir):
    output_dir.mkdir(parents=True, exist_ok=True)
    now = datetime.now().astimezone()
    path = output_dir / f"titan007_single_{now:%Y-%m-%d_%H%M}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "单场实时数据"
    ws.append(HEADERS)
    for row in rows:
        ws.append(row)
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    border = Border(bottom=Side(style="thin", color="D9E2F3"))
    for cell in ws[1]:
        cell.fill, cell.font, cell.border = fill, font, border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal="center" if cell.column > 1 else None, vertical="center")
        row[1].number_format = "@"
        for i in (4, 5, 7, 8, 9):
            row[i].number_format = "0.00"
    for i, w in enumerate([15,12,22,22,16,14,16,14,16,14], 1):
        ws.column_dimensions[ws.cell(1,i).column_letter].width = w
    ws.row_dimensions[1].height = 36
    for i in range(2, ws.max_row + 1):
        ws.row_dimensions[i].height = 24
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:J{ws.max_row}"
    tab = Table(displayName="Titan007SingleData", ref=ws.auto_filter.ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(tab)
    meta = wb.create_sheet("运行信息")
    for entry in [("项目","内容"), ("抓取时间（本机时区）", now.strftime("%Y-%m-%d %H:%M:%S %z")),
                  ("数据源",URL), ("筛选","比分 / 二合一 / 北单 / Crown（全部）/ 亚欧大 / 数字"),
                  ("比赛数量",len(rows))]:
        meta.append(entry)
    meta.column_dimensions["A"].width, meta.column_dimensions["B"].width = 24, 70
    for cell in meta[1]:
        cell.fill, cell.font = fill, font
    meta.freeze_panes = "A2"
    wb.save(path)
    return path


def main():
    parser = argparse.ArgumentParser(description="抓取8波北单十列数据")
    parser.add_argument("--show-browser", action="store_true")
    parser.add_argument("--output-dir", default=str(ROOT / "output"))
    args = parser.parse_args()
    log = ROOT / "scraper.log"
    logging.basicConfig(filename=log, level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", encoding="utf-8")
    try:
        rows = scrape(not args.show_browser)
        out = save_excel(rows, Path(args.output_dir))
        logging.info("成功抓取 %d 场：%s", len(rows), out)
        print(f"成功抓取 {len(rows)} 场比赛\nExcel：{out}")
        return 0
    except Exception as exc:
        logging.exception("运行失败")
        print(f"运行失败：{exc}\n详情见：{log}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

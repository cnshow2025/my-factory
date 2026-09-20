# -*- coding: utf-8 -*-
"""產生空白填寫範本：四個分頁（說明／貨批／機台／能力矩陣），附欄位說明與範例列。

版面與「模擬WIP_200批.xlsx」共用 mkcommon，改一處兩個檔都會跟著變。
範例貨批刻意全部 arrive_h = 0（整批都已經在 WIP），一載進模擬器就看得出
各排程法則的差別；貨批的產生邏輯故意不跟 gendata.py 共用，那份的亂數序列
一動，分析報告第 10~11 頁的數字就對不上了。
"""
import random
import mkcommon as C

random.seed(70701)

N_LOTS = 90                       # 全部都在 WIP，沒有等料的批

# 封裝 → 程式、單 site 單顆秒數、需求佔比%
PKGS = {
    "QFN48":   (["MCU01_V3", "MCU01_V5"], 2.1, 40),
    "QFN32":   (["PMIC07_V2"],            3.4, 25),
    "SOP16":   (["LED03_V1"],             1.2, 20),
    "TSSOP28": (["SENS11_V4"],            2.8, 15),
}
BASE_YIELD = {"MCU01_V3":96.5, "MCU01_V5":95.2, "PMIC07_V2":92.0,
              "LED03_V1":97.6, "SENS11_V4":90.5}

MACH_ROWS = [
    ["T01", 8, "QFN48",   "MCU01_V3",  "QFN48|QFN32",   0],
    ["T02", 8, "QFN48",   "MCU01_V5",  "QFN48|QFN32",   0],
    ["T03", 4, "QFN32",   "PMIC07_V2", None,            0],
    ["T04", 4, None,      None,        "SOP16|TSSOP28", 0],
    ["T05", 2, "SOP16",   "LED03_V1",  "SOP16",         4],
]
FLEET_SITES = sum(m[1] for m in MACH_ROWS)

names = list(PKGS)
weights = [PKGS[p][2] for p in names]

def pick_qty():
    """批量級距：小 60% / 中 30% / 大 10%"""
    r = random.random() * 100
    if r < 60:   lo, hi, step, size = 500, 2000, 100, "小批"
    elif r < 90: lo, hi, step, size = 2000, 10000, 500, "中批"
    else:        lo, hi, step, size = 10000, 25000, 500, "大批"
    return int(round(random.uniform(lo, hi) / step) * step), size

LOT_ROWS, prev = [], None
for i in range(1, N_LOTS + 1):
    if prev and random.random() < 0.50:        # 產品集中度：一半沿用前一批的料號
        pkg, prog = prev
    else:
        pkg  = random.choices(names, weights=weights)[0]
        prog = random.choice(PKGS[pkg][0])
    prev = (pkg, prog)

    qty, size = pick_qty()
    tpu = round(PKGS[pkg][1] * random.uniform(0.92, 1.08), 2)
    hot = random.random() < 0.10
    work_h = qty * tpu / 3600 / (FLEET_SITES / len(MACH_ROWS)) + 1.0
    slack  = random.uniform(2.0, 3.5) if hot else random.uniform(4.0, 12.0)
    y = BASE_YIELD[prog] + random.uniform(-2.5, 1.5)
    LOT_ROWS.append([
        "LOT-%04d" % i, pkg, prog, qty, tpu,
        0,                                      # 全部已經在 WIP
        round(work_h * slack, 1),
        "Y" if hot else "N",
        round(y, 1) if random.random() < 0.6 else None,
        size + ("・急件" if hot else ""),
    ])

OUT = "../待測貨批_填寫範本.xlsx"
C.build(OUT, LOT_ROWS, MACH_ROWS, example=True)

PKG_ORDER = []
for r in LOT_ROWS:
    if r[1] not in PKG_ORDER: PKG_ORDER.append(r[1])
for r in MACH_ROWS:
    for p in [r[2]] + ((r[4] or "").split("|") if r[4] else []):
        if p and p not in PKG_ORDER: PKG_ORDER.append(p)

C.add_matrix(OUT,
             [dict(zip(["id","pkg","prog","qty","tpu","arrive","due","hot","yield","note"], r))
              for r in LOT_ROWS],
             [dict(zip(["id","sites","kit","prog","can","avail"],
                       [r[0], r[1], r[2], r[3], (r[4] or "").split("|") if r[4] else [], r[5]]))
              for r in MACH_ROWS],
             PKG_ORDER, example=True)

tot = sum(r[3] * r[4] for r in LOT_ROWS)
print("saved %s：%d 批 / %s 顆 / 純測試工時 %.1f h / 機隊 %d site / 理論最短 %.2f h"
      % (OUT, len(LOT_ROWS), format(sum(r[3] for r in LOT_ROWS), ","),
         tot/3600, FLEET_SITES, tot/3600/FLEET_SITES))
print("封裝順序:", PKG_ORDER, "| HOT", sum(1 for r in LOT_ROWS if r[7] == "Y"), "批")

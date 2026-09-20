# -*- coding: utf-8 -*-
"""產生空白填寫範本：四個分頁（說明／貨批／機台／能力矩陣），附欄位說明與範例列。

版面與「模擬WIP_200批.xlsx」共用 mkcommon，改一處兩個檔都會跟著變。
"""
import mkcommon as C

# 範例列（藍字）：填好自己的資料後把它們刪掉
LOT_ROWS = [
 ["LOT-0001", "QFN48",   "MCU01_V3",  8000, 2.1, 0, 36, "N", 96.5, "主力機種"],
 ["LOT-0002", "QFN48",   "MCU01_V3", 12000, 2.1, 0, 48, "N", None, "同料號續批，不用改機"],
 ["LOT-0003", "QFN32",   "PMIC07_V2", 1500, 3.4, 0, 12, "Y", 92.0, "客戶催料"],
 ["LOT-0004", "SOP16",   "LED03_V1", 25000, 1.2, 6, 72, "N", None, "6 小時後才到廠"],
 ["LOT-0005", "QFN48",   "MCU01_V5",  5000, 2.3, 0, 24, "N", None, "同 Kit 換程式"],
]

MACH_ROWS = [
 ["T01", 8, "QFN48", "MCU01_V3",  "QFN48|QFN32",   0],
 ["T02", 8, "QFN48", "MCU01_V5",  "QFN48|QFN32",   0],
 ["T03", 4, "QFN32", "PMIC07_V2", None,            0],
 ["T04", 4, None,    None,        "SOP16|TSSOP28", 0],
 ["T05", 2, "SOP16", "LED03_V1",  "SOP16",         4],
]

OUT = "../待測貨批_填寫範本.xlsx"
C.build(OUT, LOT_ROWS, MACH_ROWS, example=True)

# 能力矩陣：封裝欄位照範例資料裡出現的順序帶出來
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
print("saved", OUT, PKG_ORDER)

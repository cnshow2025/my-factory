# -*- coding: utf-8 -*-
"""產生 200 批的模擬 WIP 檔，結構與填寫範本完全相同。"""
import random, math
import mkcommon as C

random.seed(20260920)

# ---- 產品結構：封裝 → 程式、單 site 單顆秒數、需求佔比 ----
PKGS = {
    #  package     programs                                   tpu   佔比%
    "QFN48":   (["MCU01_V3", "MCU01_V5", "MCU02_V1"],         2.1,  40),
    "QFN32":   (["PMIC07_V2", "PMIC09_V1"],                   3.4,  25),
    "SOP16":   (["LED03_V1"],                                 1.2,  20),
    "TSSOP28": (["SENS11_V4"],                                2.8,  10),
    "BGA144":  (["SOC22_V2"],                                 4.5,   5),
}
# 每支程式的典型良率（填進 yield_pct 的基準）
BASE_YIELD = {"MCU01_V3":96.5, "MCU01_V5":95.2, "MCU02_V1":93.8, "PMIC07_V2":92.0,
              "PMIC09_V1":94.4, "LED03_V1":97.6, "SENS11_V4":90.5, "SOC22_V2":88.7}

PKG_NAMES = list(PKGS)
PKG_W     = [PKGS[p][2] for p in PKG_NAMES]

N_LOTS   = 200
N_WIP    = 180          # 已經躺在 WIP 的批（arrive_h = 0），其餘在開工後 3 小時內陸續到
FLEET_SITES = 5*8 + 10*4 + 10*2      # 100，用來報交期

def pick_qty():
    """批量級距：小 60% / 中 30% / 大 10%，和現場結構一致。"""
    r = random.random()*100
    if r < 60:  lo, hi, step, size = 500, 2000, 100, "S"
    elif r < 90: lo, hi, step, size = 2000, 10000, 500, "M"
    else:        lo, hi, step, size = 10000, 25000, 500, "L"
    return int(round(random.uniform(lo, hi)/step)*step), size

lots, prev = [], None
for i in range(1, N_LOTS+1):
    # 產品集中度：一半的批沿用前一批的料號，改機才不會每批都發生
    if prev and random.random() < 0.50:
        pkg, prog = prev
    else:
        pkg  = random.choices(PKG_NAMES, weights=PKG_W)[0]
        prog = random.choice(PKGS[pkg][0])
    prev = (pkg, prog)

    qty, size = pick_qty()
    tpu = round(PKGS[pkg][1] * random.uniform(0.92, 1.08), 2)
    arrive = 0 if i <= N_WIP else round(random.uniform(0.5, 3.0), 1)
    hot = random.random() < 0.10

    # 交期報價：工作量 ×1.3 retest 餘裕 + 一次換 Kit，再乘寬鬆倍數
    work_h = qty * tpu / 3600 / (FLEET_SITES/25) + 1.0
    slack  = random.uniform(2.0, 3.5) if hot else random.uniform(4.0, 12.0)
    due    = round(arrive + work_h * slack, 1)

    y = BASE_YIELD[prog] + random.uniform(-2.5, 1.5)
    lots.append([
        "LOT-%04d" % i, pkg, prog, qty, tpu,
        arrive, due, "Y" if hot else "N",
        round(y, 1) if random.random() < 0.6 else None,     # 四成留白，走隨機良率
        {"S":"小批","M":"中批","L":"大批"}[size] + ("・急件" if hot else ""),
    ])

# ---- 機台：5 台彈性 + 20 台專用，專用台數依需求量分配 ----
DEDICATE = ["QFN48"]*8 + ["QFN32"]*5 + ["SOP16"]*4 + ["TSSOP28"]*2 + ["BGA144"]*1
SITES    = [8]*5 + [4]*10 + [2]*10          # T01–T05 快機，T06–T15 中速，T16–T25 慢機
machines = []
for i in range(25):
    mid, sites = "T%02d" % (i+1), SITES[i]
    if i < 5:                                # 彈性機：什麼都能測
        can = ""
        kit = PKG_NAMES[i % len(PKG_NAMES)]
    else:
        kit = DEDICATE[i-5]
        can = kit
    prog = random.choice(PKGS[kit][0])
    avail = 0 if random.random() < 0.88 else round(random.uniform(1, 6), 1)
    machines.append([mid, sites, kit, prog, can, avail])

C.build("../模擬WIP_200批.xlsx", lots, machines,
        note="這是模擬產生的 200 批測試資料，結構與填寫範本相同，可直接載入模擬器。")
print("lots=%d machines=%d" % (len(lots), len(machines)))
tot = sum(l[3]*l[4] for l in lots)
print("總顆數 %d，純測試工時(1 site) %.1f h，機隊 %d site，理論最短 %.2f h"
      % (sum(l[3] for l in lots), tot/3600, sum(m[1] for m in machines),
         tot/3600/sum(m[1] for m in machines)))

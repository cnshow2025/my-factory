# 測試排程分析報告 — 產生方式

`../測試排程分析報告.pptx` 是由這個資料夾裡的程式產生的，可以重新產生。

## 檔案

| 檔案 | 用途 |
|---|---|
| `charts/charts.html` | 圖表渲染器。用純 SVG 畫出 10 張圖表（`hbar` / `vbar` / `line` / `donut`），資料寫在檔案最下面 |
| `charts/shoot.mjs` | 用 Chromium 把每張圖表截成 3 倍解析度 PNG（`charts/c_*.png`） |
| `gen.js` | 用 pptxgenjs 組版面，把 `charts/c_*.png` 內嵌進 11 頁投影片 |
| `audit.py` | 版面稽核：檢查出界、邊距、文字溢出、圖文重疊 |
| `report_data.json` | 模擬跑分的原始結果（各情境的良品、逾期、稼動、改機次數等） |

圖表刻意輸出成點陣圖而非 PowerPoint 原生圖表物件，因為原生圖表在部分看圖工具／行動裝置上不會被繪製，會變成空白頁。

## 重新產生

```bash
cd docs/report
node charts/shoot.mjs          # 需要 playwright-core 與一份 Chromium
node gen.js                    # 需要 pptxgenjs
python3 audit.py               # 需要 python-pptx
```

改數字就改 `charts/charts.html` 底部的資料，以及 `gen.js` 裡卡片的文字。

## 報告結構（13 頁）

1. 封面
2. 結論摘要
3. 模型與校準
4. 影響產出的因素排序
5. 因素一：派工時機
6. 因素二：排程法則比較
7. 因素三：工程師人力與人機比
8. 因素四：機台專用化與工程改善槓桿
9. 累積效益路徑
10. 實測案例：清空 200 批 WIP 要多久
11. 人力改變了最佳排程法則
12. 專業建議
13. 實施路徑、追蹤指標與模型限制

第 10~11 頁的數字來自「📥 實際貨批分析」載入 `../模擬WIP_200批.xlsx` 跑出來的結果
（純排程比較模式），資料檔與產生方式見 [`../wip/README.md`](../wip/README.md)。

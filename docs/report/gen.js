const pptxgen = require("pptxgenjs");
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";                 // 13.33 x 7.5
pres.author = "MyFactory 排程模擬";
pres.title = "測試排程分析報告";

// ---- 色票 ----
const NAVY = "0E1A3C", PRIMARY = "1E2761", TEAL = "1C7293", ICE = "CADCFC";
const INK = "16213E", MUTED = "5A6785", LINE = "DDE3EE", WHITE = "FFFFFF";
const GOOD = "0F7B6C", WARN = "C2410C", BAD = "9B1C1C";
const HEAD = "Cambria", BODY = "Calibri";

const T = (s, txt, o) => s.addText(txt, Object.assign({ isTextBox:true, fontFace:BODY, color:INK }, o));

// 圖表以高解析度 PNG 內嵌，任何看圖工具都能完整顯示線條
const AR = { c_factors:2.5522, c_cum:2.5522, c_depth:1.3965, c_late:1.3965, c_eng:1.3965,
             c_time:1.3965, c_ded:1.3146, c_ratio:1.5023, c_lever:1.5960, c_policy:1.5677,
             c_clear:2.7431, c_curve:1.5023, c_engpol:1.5023, c_frozen:1.3965,
             c_fleet:1.4554, c_rematch:1.4554, c_switch:2.3553 };
function img(s, id, x, y, w){
  const h = w / AR[id];
  s.addImage({ path:`charts/${id}.png`, x, y, w, h });
  return y + h;                                  // 回傳圖片底緣，方便排版
}

function titleSlide(s, kicker, title, sub){
  T(s, kicker, { x:0.9, y:2.15, w:11.5, h:0.35, fontSize:13, color:ICE, charSpacing:3 });
  T(s, title,  { x:0.9, y:2.55, w:11.5, h:1.5, fontSize:44, bold:true, color:WHITE, fontFace:HEAD });
  T(s, sub,    { x:0.9, y:4.15, w:11.5, h:0.9, fontSize:15, color:ICE, lineSpacing:26 });
}
function head(s, num, title, sub){
  T(s, num, { x:0.62, y:0.42, w:0.55, h:0.5, fontSize:13, bold:true, color:TEAL });
  T(s, title, { x:1.25, y:0.34, w:11.4, h:0.62, fontSize:30, bold:true, color:PRIMARY, fontFace:HEAD });
  if (sub) T(s, sub, { x:1.25, y:0.98, w:11.4, h:0.42, fontSize:13.5, color:MUTED });
}
function stat(s, x, y, w, value, label, note, color){
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h:1.62, fill:{ color:"F5F7FB" },
    line:{ color:"FFFFFF", width:0 }, rectRadius:0.06 });
  T(s, value, { x:x+0.02, y:y+0.18, w:w-0.04, h:0.72, fontSize:31, bold:true,
    color:color||PRIMARY, align:"center", fontFace:HEAD, margin:0 });
  T(s, label, { x:x+0.02, y:y+0.86, w:w-0.04, h:0.3, fontSize:12.5, color:INK, align:"center", margin:0 });
  T(s, note,  { x:x+0.02, y:y+1.16, w:w-0.04, h:0.34, fontSize:10.5, color:MUTED, align:"center", margin:0 });
}
function card(s, x, y, w, h, title, lines, accent){
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill:{ color:"F5F7FB" },
    line:{ color:"FFFFFF", width:0 }, rectRadius:0.05 });
  s.addShape(pres.ShapeType.ellipse, { x:x+0.26, y:y+0.28, w:0.16, h:0.16,
    fill:{ color:accent||TEAL }, line:{ color:accent||TEAL, width:0 } });
  T(s, title, { x:x+0.54, y:y+0.17, w:w-0.8, h:0.36, fontSize:14.5, bold:true, color:PRIMARY, margin:0 });
  T(s, lines.map(t => ({ text:t, options:{ bullet:true, breakLine:true } })),
    { x:x+0.54, y:y+0.56, w:w-0.8, h:h-0.66, fontSize:12, color:INK,
      lineSpacing:18, paraSpaceAfter:4, margin:0 });
}

/* ===== 1 封面 ===== */
let s = pres.addSlide(); s.background = { color:NAVY };
titleSlide(s, "MYFACTORY 排程模擬 · 25 台測試機產線",
  "測試排程分析報告",
  "以離散事件模擬量化影響產出的所有因素，比較七種排程法則，\n並提出可落地的排程與資源配置建議");
T(s, "資料來源：MyFactory 模擬器（已用現場三項實測數據校準）　|　每組情境 20 個獨立模擬日取平均",
  { x:0.9, y:6.45, w:11.5, h:0.4, fontSize:11, color:"8FA2C8" });
s.addNotes("本報告所有數字來自同一個校準過的模擬模型，情境之間只改變單一變數，因此差異可直接歸因。");

/* ===== 2 結論摘要 ===== */
s = pres.addSlide();
head(s, "01", "結論摘要", "把典型現行做法換成建議做法，模擬顯示每日良品可以接近翻倍——而且前兩項不花錢");
stat(s, 0.62, 1.62, 3.0, "+93%", "每日良品提升", "77 萬 → 148 萬顆／日", GOOD);
stat(s, 3.79, 1.62, 3.0, "220 → 2", "每日逾期批數", "交期表現同步改善", GOOD);
stat(s, 6.96, 1.62, 3.0, "37% → 1%", "機台等工程師占比", "同樣人力、排程改變後", TEAL);
stat(s, 10.13, 1.62, 2.55, "1 : 2.5", "建議人機比", "25 台配 10 位工程師", PRIMARY);
card(s, 0.62, 3.55, 6.06, 2.35, "不花錢就能拿到的效益（前兩項）", [
  "排程法則改為「省改機優先」：+41 萬顆／日",
  "取消預先派工，機台做完當下才決定：+21 萬顆／日",
  "兩項合計 +62 萬顆／日（+80%），逾期批數 220 → 20"
], GOOD);
card(s, 6.96, 3.55, 5.72, 2.35, "需要投入的效益（後兩項）", [
  "工程師由 8 人補至 10 人：+7 萬顆／日",
  "換程式時間 30→15 分：+3 萬顆／日",
  "兩項合計 +10 萬顆／日，並把逾期壓到 2 批以下"
], TEAL);
T(s, "情境定義　典型現行做法 ＝ 最早交期排序 ＋ 預先押一批 ＋ 8 位工程師　|　建議做法 ＝ 省改機排序 ＋ 即時派工 ＋ 10 位工程師 ＋ 換程式 15 分",
  { x:0.62, y:6.16, w:12.05, h:0.4, fontSize:11.5, color:MUTED });
T(s, "每組情境皆為 20 個獨立模擬日的平均；情境之間僅改變指定變數，其餘參數完全相同。",
  { x:0.62, y:6.58, w:12.05, h:0.4, fontSize:11.5, color:MUTED });
s.addNotes("所有數字來自同一模型、同一組 20 個模擬日，情境之間只改變指定變數。");

/* ===== 3 模型與校準 ===== */
s = pres.addSlide();
head(s, "02", "模型與校準", "模型不是假設出來的：用現場三項實測值反推參數，三項同時對上才開始分析。右圖是校準後的基準情境");
s.addTable([
  [{ text:"校準項目", options:{ bold:true, color:WHITE, fill:{ color:PRIMARY } } },
   { text:"現場實測", options:{ bold:true, color:WHITE, fill:{ color:PRIMARY } } },
   { text:"模型產出", options:{ bold:true, color:WHITE, fill:{ color:PRIMARY } } },
   { text:"結果", options:{ bold:true, color:WHITE, fill:{ color:PRIMARY } } }],
  ["全廠每日換 Kit 次數", "約 50 次", "39～44 次", "一致"],
  ["當機次數 ÷ 換 Kit 次數", "1.5 倍", "1.47 倍", "一致"],
  ["機隊等效 site 數", "5×8 + 10×4 + 10×2 = 100", "100（平均 4）", "一致"]
], { x:0.62, y:1.6, w:7.3, colW:[2.7,1.75,1.55,1.3], fontSize:12, fontFace:BODY,
     color:INK, border:{ type:"solid", color:LINE, pt:1 }, rowH:0.42, valign:"middle" });
card(s, 0.62, 3.55, 7.3, 2.5, "模型參數（校準後）", [
  "25 台測試機：5 台 8 site、10 台 4 site、10 台 2 site",
  "批量結構：小批 0.5–2k 占 60%、中批 2–10k 占 30%、大批 10–25k 占 10%",
  "換 Kit 60 分、換程式 30 分、良率 80–99%、RT 門檻 99/98",
  "當機率 11.5%／機時、修復 10–120 分、一天 24 小時三班"
], PRIMARY);
img(s, "c_time", 8.1, 1.55, 4.6);          // 底緣 ≈ 4.84
card(s, 8.1, 4.98, 4.6, 1.75, "瓶頸一眼可見", [
  "真正在測試的時間只有 46.7%",
  "改機 + 維修 + 等工程師 = 38.7%",
  "這三項全部繞著「工程師」轉"
], WARN);
T(s, "右圖是基準情境（已採用省改機＋即時派工＋10 人）；現場典型做法的稼動只有 29.2%，見第 9 頁的起點 A。",
  { x:8.1, y:6.80, w:4.6, h:0.42, fontSize:10, color:MUTED });
T(s, "註　時間結構為 25 台機 × 1440 分鐘的合計分配，取 20 個模擬日平均。",
  { x:0.62, y:6.18, w:7.3, h:0.4, fontSize:11, color:MUTED });
s.addNotes("每台機每天：測試 673 分、改機 343 分、維修 161 分、等工程師 53 分、閒置 210 分。");

/* ===== 4 因素排序 ===== */
s = pres.addSlide();
head(s, "03", "影響產出的因素排序", "同一組機台與人力，只改變單一因素造成的每日良品差距（單位：萬顆）");
img(s, "c_factors", 0.62, 1.48, 12.05);    // 底緣 ≈ 6.20
card(s, 0.62, 6.30, 12.05, 1.02, "讀這張圖的方式", [
  "前三項是「排程決策」——不必花錢，改做法即可；後三項是「工程改善」——需要投資或人力"
], GOOD);
s.addNotes("當機率減半呈現負值，是因為工程師被釋放後排程做了更多改機，屬於資源再分配效果。");

/* ===== 5 因素一：派工時機 ===== */
s = pres.addSlide();
head(s, "04", "因素一：派工時機（最大的免費槓桿）", "「預約深度」＝ 提前把幾批貨綁定到機台上排隊");
img(s, "c_depth", 0.62, 1.55, 5.9);        // 底緣 ≈ 5.77
img(s, "c_late", 6.81, 1.55, 5.9);
card(s, 0.62, 5.90, 12.05, 1.28, "機制：提前綁定 ＝ 用舊資訊做決策", [
  "後來的急件、突發當機、良率變化都還沒發生，而被綁住的批已經不能改派",
  "押 3 批 vs 做完才派：每日良品少 27.7 萬顆（−19%），逾期批數由 5 批暴增到 106 批"
], WARN);
s.addNotes("這是本報告中投資報酬率最高的一項：不花錢、今天就能改。");

/* ===== 6 因素二：排程法則 ===== */
s = pres.addSlide();
head(s, "05", "因素二：排程法則比較", "同樣 25 台機、同樣 10 位工程師、同樣的貨，只改變「先做哪一批」的規則");
img(s, "c_policy", 0.62, 1.50, 7.6);       // 底緣 ≈ 6.35
card(s, 8.45, 1.50, 4.22, 1.80, "為什麼「省改機」贏", [
  "改機佔用工程師，工程師是瓶頸",
  "少改機 → 工程師被釋放 → 其他機台少等",
  "換 Kit 39 次 vs 最差法則 64 次"
], GOOD);
card(s, 8.45, 3.45, 4.22, 1.80, "同時贏產出與交期", [
  "省改機法則的逾期批數也是最低（5 批／日）",
  "是唯一不必在衝產出與保交期之間取捨的法則",
  "其餘法則都得在兩者之間取捨"
], TEAL);
card(s, 8.45, 5.40, 4.22, 1.15, "例外：排名會翻轉", [
  "機台或人力寬裕時改用「大批給快機」更快，判斷方式見第 13 頁"
], WARN);
T(s, "註　最佳與最差差距 50.2 萬顆／日（+52%）；七種法則皆在相同的 20 個模擬日上比較。",
  { x:0.62, y:6.62, w:7.6, h:0.4, fontSize:11, color:MUTED });
s.addNotes("省改機法則同時把逾期批數壓到 5 批，是唯一兼顧產出與交期的法則。");

/* ===== 7 因素三：工程師人力與人機比 ===== */
s = pres.addSlide();
head(s, "06", "因素三：工程師人力與人機比", "改機與維修都要佔用一位工程師；人不足時機台就停著等");
img(s, "c_eng", 0.62, 1.48, 5.3);          // 底緣 ≈ 5.28
img(s, "c_ratio", 6.42, 1.48, 6.1);        // 底緣 ≈ 5.54
card(s, 0.62, 5.64, 6.0, 1.30, "本廠建議值：10 人（1:2.5）", [
  "10 人是拐點，再加人只多 0.1%",
  "降到 8 人：停等 14.4%、逾期由 5 批升到 20 批"
], TEAL);
card(s, 6.75, 5.64, 5.92, 1.30, "跨規模通用的管理規則", [
  "1:4 以內產能幾乎不損失；1:6 直接掉一半",
  "1:8 以上系統性崩塌 → 50 台至少要 13 人"
], GOOD);
T(s, "註　右圖各規模的到貨量與開局在製依機台數等比例放大，使每台機負載一致；縱軸以各規模自身 1:1 的產出為 100% 正規化。",
  { x:0.62, y:6.96, w:12.05, h:0.28, fontSize:10, color:MUTED });
s.addNotes("每台機每天需要 504 分鐘的工程師工時（改機 343 + 維修 161），25 台合計 8.8 人的純工作量，加上排隊緩衝需 10 人。");

/* ===== 8 因素四：機台專用化與工程改善槓桿 ===== */
s = pres.addSlide();
head(s, "07", "因素四：機台專用化與工程改善槓桿", "先看專用化如何削減換 Kit，再比較各項工程改善的實際效益");
img(s, "c_ded", 0.62, 1.48, 5.0);          // 底緣 ≈ 5.28
img(s, "c_lever", 6.40, 1.48, 6.15);       // 底緣 ≈ 5.33
card(s, 0.62, 5.45, 6.0, 1.78, "彈性機的「台數」已接近最佳", [
  "5 台彈性機扛下所有換 Kit，現況 39 次／日",
  "完全專用再省 70 分／台／日，但急單無處插隊",
  "專用機「各綁哪一種封裝」是另一回事，見第 12 頁"
], TEAL);
card(s, 6.75, 5.45, 5.92, 1.78, "槓桿：換程式才是大宗", [
  "換程式 213 次／日 vs 換 Kit 39 次／日",
  "30→15 分是唯一同時改善產出、工時與交期的措施",
  "換 Kit 60→30 分反使逾期由 5 升到 12.6 批"
], WARN);
s.addNotes("改機變便宜後排程更常換機種，反而打散交期——這是「局部改善未必全域最佳」的典型例子。");

/* ===== 9 累積效益路徑 ===== */
s = pres.addSlide();
head(s, "08", "累積效益路徑", "從典型現行做法開始，依序加入四項建議，每一步的實際模擬結果");
img(s, "c_cum", 0.96, 1.46, 11.4);         // 底緣 ≈ 5.93
card(s, 0.62, 6.03, 12.05, 1.28, "效益集中在前兩步", [
  "起點 A：稼動 29.2%、逾期 220 批／日　→　終點 E：稼動 47.2%、逾期 1.9 批／日、等工程師 0.8%",
  "B 與 C 合計 +62 萬顆／日，占全部提升的 86%，且完全不需要增加人力或設備"
], GOOD);
s.addNotes("A→E 累積 +71.4 萬顆／日（+92.8%）。B +40.8 萬、C +20.8 萬、D +6.9 萬、E +3.0 萬。");


/* ===== 10 實測案例：清空一批真實 WIP ===== */
s = pres.addSlide();
head(s, "09", "實測案例：清空 200 批 WIP 要多久", "把待測清單直接餵進模型，跑到最後一批做完為止——問的不再是「一天做多少」，而是「這堆貨要多久清光」");
img(s, "c_clear", 0.62, 1.50, 7.5);        // 底緣 ≈ 4.23
img(s, "c_curve", 8.27, 1.50, 4.44);       // 底緣 ≈ 4.46
card(s, 0.62, 4.60, 6.0, 2.00, "這批貨的基本盤", [
  "200 批・66.2 萬顆・25 台機（機隊 100 site）",
  "理論最短 4 小時 01 分（零改機零當機）",
  "最快的法則 9 小時 41 分，仍比下限多 142%",
  "最快與最慢差 10 小時 20 分"
], PRIMARY);
card(s, 6.75, 4.60, 5.92, 2.00, "完成曲線看得出尾巴", [
  "省改機前 3 小時就做完 137 批，遙遙領先",
  "但 9 小時後停在 198 批，最後 2 批拖到 14.5 小時",
  "大批給快機起步慢，10 小時整批清光",
  "看整批何時清空，不能只看前段衝得多快"
], WARN);
T(s, "註　情境為「純排程比較」：關閉當機與良率波動，同一批貨每次跑結果完全相同，法則之間的差異純粹來自排程順序。",
  { x:0.62, y:6.70, w:12.05, h:0.36, fontSize:11, color:MUTED });
s.addNotes("這是模擬產生的 200 批 WIP，結構比照現場：5 種封裝、8 種程式、小中大批 60/30/10。");

/* ===== 11 人力改變最佳法則 ===== */
s = pres.addSlide();
head(s, "10", "人力改變了最佳排程法則", "同一批貨、同一組機台，只改工程師人數，冠軍法則就換人做——這是前面所有結論的邊界條件");
img(s, "c_engpol", 0.62, 1.46, 5.9);       // 底緣 ≈ 5.39
img(s, "c_frozen", 7.21, 1.46, 5.5);       // 底緣 ≈ 5.40
card(s, 0.62, 5.52, 6.0, 1.75, "交叉點在 5～10 人之間", [
  "3 人：省改機 13.5 h 勝出，大批給快機要 18.0 h",
  "10 人：大批給快機 9.7 h，省改機落到 14.5 h",
  "人少 → 改機排隊；人多 → 閒機排隊"
], GOOD);
card(s, 6.75, 5.52, 5.92, 1.75, "為什麼省改機在人多時會輸", [
  "不換 Kit ＝ 把機台配置凍結在現況",
  "這 200 批裡 QFN32 佔 38.6% 工時，卻只有 22/100 site 裝著它",
  "光這一種封裝就要 7 小時，其他機台只能閒著"
], WARN);
s.addNotes("3 人時省改機的等工程師時間是 32 分，其他法則是 5,000~13,000 分；加上當機後，3 人除了省改機全部拖到 36~52 小時。");


/* ===== 12 機隊規模與專用機配置 ===== */
s = pres.addSlide();
head(s, "11", "機隊規模與專用機配置也會換掉答案",
     "同一批 90 批 WIP，把 5 台機換成真實的 25 台機隊（機隊 100 site、10 位工程師）重跑一次");
img(s, "c_fleet", 0.62, 1.44, 5.5);        // 底緣 ≈ 5.22
img(s, "c_rematch", 7.21, 1.44, 5.5);
card(s, 0.62, 5.30, 6.0, 1.62, "機台越多，省改機越不是答案", [
  "5 台機：省改機 14.6 h 勝，大批給快機 29.4 h",
  "25 台機：大批給快機 6.4 h，省改機退到 7.9 h",
  "機台少 → 改機排隊；機台多 → 快機沒用滿"
], WARN);
card(s, 6.75, 5.30, 5.96, 1.62, "配置對齊買到的是容錯，不是速度", [
  "最佳法則幾乎沒變：6.4 h → 6.6 h",
  "最差法則大幅改善：19.8 h → 12.5 h",
  "選錯法則的代價從 13.4 h 降到 5.9 h，等工程師歸零"
], GOOD);
T(s, "註　這批 WIP 的 QFN32 佔 46.5% 工時，現場卻只有 5 台專用機（22 site）裝著它，QFN48 反而有 8 台（40 site）——配置和貨量是反的。「重新對齊」＝ 機台數與 site 結構完全不動，只改 20 台專用機各自綁哪一種封裝。",
  { x:0.62, y:6.98, w:12.09, h:0.26, fontSize:10, color:MUTED });
s.addNotes("5 台機時瓶頸是改機排隊，所以省改機贏；25 台機時機台過剩、瓶頸變成有沒有把大批塞進 8-site 快機。重新對齊後所有法則的等工程師時間都是 0。");


/* ===== 13 切換訊號 ===== */
s = pres.addSlide();
head(s, "12", "什麼時候該切換排程法則",
     "四次翻轉放在同一把尺上量：改機與等工程師合計吃掉多少機台時間，門檻落在 33% 附近");
img(s, "c_switch", 1.77, 1.36, 9.8);       // 底緣 ≈ 5.52
card(s, 0.62, 5.62, 6.0, 1.38, "怎麼量（用現場的實績，不用模擬）", [
  "（改機分鐘 ＋ 等工程師分鐘）÷（機台數 × 當班分鐘）",
  "用「切換前」的排程量一次就好，現場稼動報表就有這兩欄"
], PRIMARY);
card(s, 6.75, 5.62, 5.96, 1.38, "怎麼判斷", [
  "超過 35% → 改用「省改機優先」，11 組情境全對",
  "低於 30% → 改用「大批給快機」；30～35% 兩者差不到 15%"
], GOOD);
T(s, "註　① 兩項要合計才準：5 台機那三組是改機吃掉 35%（等人只有 0～2%），25 台機 200 批 3 人那組反過來是等人吃掉 36%（改機只有 10%）——只看一項會判錯。② 全部以現行的「追交期」排程量測。③ 切換到省改機之後這個數字會自己掉下來（43.3% → 27.5%），那是措施生效，不是要你切回去。",
  { x:0.62, y:7.04, w:12.09, h:0.26, fontSize:9, color:MUTED });
s.addNotes("門檻的實際分界在 31.7% 與 35.8% 之間；靠近門檻時兩種法則差距小，選錯的代價不大。");

/* ===== 14 專業建議 ===== */
s = pres.addSlide(); s.background = { color:NAVY };
T(s, "13", { x:0.62, y:0.42, w:0.55, h:0.5, fontSize:13, bold:true, color:TEAL });
T(s, "專業建議", { x:1.25, y:0.34, w:11.4, h:0.62, fontSize:30, bold:true, color:WHITE, fontFace:HEAD });
T(s, "依「投入成本 → 效益」排序，前兩項不需任何資本支出", { x:1.25, y:0.98, w:11.4, h:0.4, fontSize:13.5, color:ICE });
const recs = [
  ["1", "排程法則改為「省改機優先」", "同機種串著跑，工程師被釋放後全廠少等", "+41 萬顆／日", GOOD],
  ["2", "派工單不提前開，做完當下才決定", "取消預先押批，改為即時派工", "+21 萬顆／日", GOOD],
  ["3", "工程師由 8 人補至 10 人（1:2.5）", "8 人時有 14.4% 的機台時間在停等", "+7 萬顆／日", TEAL],
  ["4", "縮短換程式時間 30→15 分", "一天 213 次，槓桿最大的工程改善", "+3 萬顆／日", TEAL]
];
recs.forEach((r, i) => {
  const y = 1.55 + i*1.28;
  s.addShape(pres.ShapeType.roundRect, { x:0.62, y, w:12.05, h:1.15,
    fill:{ color:"17254C" }, line:{ color:"17254C", width:0 }, rectRadius:0.05 });
  s.addShape(pres.ShapeType.ellipse, { x:0.92, y:y+0.34, w:0.47, h:0.47,
    fill:{ color:r[4] }, line:{ color:r[4], width:0 } });
  T(s, r[0], { x:0.92, y:y+0.38, w:0.47, h:0.4, fontSize:16, bold:true, color:WHITE, align:"center", margin:0 });
  T(s, r[1], { x:1.62, y:y+0.2, w:7.2, h:0.4, fontSize:17, bold:true, color:WHITE, margin:0 });
  T(s, r[2], { x:1.62, y:y+0.62, w:7.2, h:0.38, fontSize:12.5, color:"A9BBDD", margin:0 });
  T(s, r[3], { x:9.1, y:y+0.34, w:3.35, h:0.5, fontSize:19, bold:true, color:r[4],
    align:"right", fontFace:HEAD, margin:0 });
});
T(s, "四項全做：每日良品 77 萬 → 148 萬顆（+93%），逾期批數 220 批 → 2 批以下",
  { x:0.62, y:6.60, w:12.05, h:0.34, fontSize:13, color:ICE, align:"center" });
T(s, "但第 1 項有前提：省改機是「改機＋等工程師吃掉機台時間超過 35%」時的最佳解；低於 30% 就該改用「大批給快機」（見第 13 頁）。班別與機隊狀況不同，規則就該不同。",
  { x:0.62, y:6.94, w:12.05, h:0.34, fontSize:11.5, color:"A9BBDD", align:"center" });
s.addNotes("前兩項是純管理決策，今天就能改；第三項是人力預算；第四項需要 IT 或設備工程投入。");

/* ===== 15 實施路徑・追蹤指標・模型限制 ===== */
s = pres.addSlide();
head(s, "14", "實施路徑、追蹤指標與模型限制", "建議以四週為一個驗證循環，每一步都有可量測的指標");
const steps = [
  ["第 1–2 週", "凍結預先派工", "改為機台做完才派下一批", "逾期批數／日"],
  ["第 3–4 週", "導入省改機排序", "同 Kit 同程式的批優先串接", "每台每日改機分鐘"],
  ["第 5–8 週", "人力補至 10 人", "或重新分配現有工程師班別", "等工程師時間占比"],
  ["第 9 週起", "換程式標準化", "程式預載、遠端下載、recipe 統一", "（改機＋等人）÷ 機台時間"]
];
steps.forEach((st, i) => {
  const x = 0.62 + i*3.12;
  s.addShape(pres.ShapeType.roundRect, { x, y:1.55, w:2.92, h:2.62, fill:{ color:"F5F7FB" },
    line:{ color:"FFFFFF", width:0 }, rectRadius:0.05 });
  s.addShape(pres.ShapeType.ellipse, { x:x+0.25, y:1.76, w:0.46, h:0.46,
    fill:{ color:PRIMARY }, line:{ color:PRIMARY, width:0 } });
  T(s, String(i+1), { x:x+0.25, y:1.81, w:0.46, h:0.38, fontSize:15, bold:true, color:WHITE, align:"center", margin:0 });
  T(s, st[0], { x:x+0.25, y:2.34, w:2.45, h:0.28, fontSize:11.5, color:TEAL, bold:true, margin:0 });
  T(s, st[1], { x:x+0.25, y:2.64, w:2.45, h:0.38, fontSize:15, bold:true, color:PRIMARY, margin:0 });
  T(s, st[2], { x:x+0.25, y:3.06, w:2.45, h:0.46, fontSize:12, color:INK, lineSpacing:17, margin:0 });
  T(s, "追蹤：" + st[3], { x:x+0.25, y:3.50, w:2.45, h:0.58, fontSize:11, color:MUTED, lineSpacing:15, margin:0 });
});
card(s, 0.62, 4.34, 6.0, 2.36, "模型尚未納入（避免過度外推）", [
  "人員技能差異與交接班",
  "Handler／Kit 數量限制（假設隨時可取得）",
  "物料搬運與 WIP 暫存空間限制",
  "同一機台多產品混測（假設一次一批）"
], WARN);
card(s, 6.75, 4.34, 5.92, 2.36, "建議補充的現場數據", [
  "實際的 WIP 清單（第 10~13 頁用模擬資料）",
  "每台機實際日完成批數（核對稼動率 46.7%）",
  "逾期一批的真實違約成本（現以 20,000 代入）",
  "換程式是否一律由工程師執行——若由操作員執行，人力需求可由 10 人降至 5～6 人"
], PRIMARY);
T(s, "一句話總結　產出的瓶頸不是機台，是改機與工程師；而排程方式決定了這兩者被浪費多少。　驗證用模擬器：cnshow2025.github.io/my-factory（參數組選「我的工廠」）",
  { x:0.62, y:6.80, w:12.05, h:0.40, fontSize:11.5, color:MUTED });
s.addNotes("四個階段刻意分開，避免同時變動多個變數而無法歸因。最大的單一不確定性是換程式是否佔用工程師。");

pres.writeFile({ fileName: "../測試排程分析報告.pptx" }).then(f => console.log("written:", f));

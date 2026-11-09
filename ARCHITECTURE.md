# 系統架構書 — Incision 叫序資料庫

> 架構照搬 [moscito-relay-bidding](https://github.com/Bonki1268/moscito-relay-bidding)
> （它又照搬 [bridge-relay-bidding](https://github.com/Bonki1268/bridge-relay-bidding)）。
> 本文件只記錄與 moscito 不同的地方；表頭偵測、巢狀接力、跳轉連結等共用細節
> 請看那兩邊的 ARCHITECTURE.md。

## 1. 專案是什麼

把 Daniel J. Neill《Incision: Forcing Club System》（2019，36 頁）整理成七份 Excel
叫序表，轉成純靜態、有 RWD 的網站，部署在 GitHub Pages。

```
原始資料/*.xlsx  ──  python3 build.py  ──▶  網站/*.html  ──  push  ──▶  GitHub Pages
```

`網站/` 底下的 `.html` 全部是產出物，不要手改；`網站/assets/style.css`、`site.js`
從 moscito 複製，手動維護，build 不會覆寫。

## 2. 目錄結構

```
incision-bidding/
├── build.py
├── ARCHITECTURE.md
├── .github/workflows/pages.yml   ← 與 moscito 相同，部署 網站/
├── 原始資料/
│   ├── 制度總覽.xlsx   一梅花開叫.xlsx  一方塊開叫.xlsx  高花開叫.xlsx
│   └── 一無王開叫.xlsx 二階開叫.xlsx    一般約定.xlsx
└── 網站/
    ├── index.html
    ├── assets/style.css, site.js
    └── 叫序資料/制度總覽.html … 一般約定.html
```

## 3. 與 moscito 的 build.py 差異

| 項目 | moscito-relay-bidding | 本專案 |
|---|---|---|
| `PAGES` | 4 欄（xlsx、標題、角標、簡介） | 多一欄「首頁分組」：`開叫叫序`／`總覽與約定` |
| 首頁卡片分組 | `cards[:5]`／`cards[5:]` 寫死 | 依 `PAGES` 的分組欄位產生 |
| 點數 | rp 換算成 HCP，頁尾註明 | 原文即 HCP；頁尾註明 8–12 修訂 |
| 跨頁連結 | 「接力方案檔」「競叫檔」等 | 「制度總覽檔」「1♣ 開叫檔」「1♦ 開叫檔」「高花開叫檔」「1NT 開叫檔」「二階開叫檔」「一般約定檔」（`PAGE_REF_ALIASES`） |

其餘（`analyze_sheet()`、各 renderer、`REF_RE`、巢狀接力）原封不動。

## 4. xlsx 的格式約定

與 moscito 相同：

- **表頭列／段落小標**：粗體＋深藍底 `1F3864`。表頭之前的列是標題與情境說明。
- **「說明」分頁**：表頭為「項目／內容」。
- **巢狀子接力**：上層列以 `▸ ` 開頭（淺藍底），子列以全形空白 `　` 開頭（淺黃底），預設收合。
- **分頁跳轉**：寫「見表 N」，分頁標題以「N.」開頭。
- **註記**：「＊」開頭、只寫第一欄（合併儲存格）。在表尾 = 頁尾附註；在表中 = 註記列。

本專案另外注意：

- 第一欄不要只寫單一英文字母（如 `A`、`K`、`X`），也不要寫「字母＋空白」開頭，
  `BRANCH_CODE_RE` 會把它當成分支代號畫成徽章。干擾欄一律寫成 `(X)`、`(1NT)`。
- 文字中避免出現「表＋數字」（如「代表 3 張」），`REF_RE` 會把它轉成跳轉連結。

## 5. 內容約定

- 第一、二家 1♦／1♥／1♠／2♣／2♦ 的開叫範圍採本站修訂的 **8–12**（原文 10–12）；
  應叫者範圍、步階回答「高限 11+／12」、低花轉移「非低限（10+）」等仍是原文數字，
  在各檔「說明」標註待重新校準。
- 原文缺漏、前後矛盾、寫法不清整理在 `制度總覽.xlsx` 的「3.待討論」，各表遇到時以
  「見制度總覽檔」連過去。

## 6. 維護 SOP

```bash
# 1. 改 原始資料/*.xlsx（不要改表頭與小標的底色）
python3 build.py          # 2. 重新產生網站
open 網站/index.html      # 3. 目測
git add -A && git commit  # 4. push 後 GitHub Actions 自動部署
```

新增一份叫序表：把 xlsx 放進 `原始資料/`，在 `build.py` 的 `PAGES` 加一筆。

首次部署：push 後到 Settings → Pages 把 Source 設成「GitHub Actions」
（或 `gh api -X POST repos/<owner>/<repo>/pages -f build_type=workflow`）。
推送 workflow 檔需要 token 有 `workflow` scope。

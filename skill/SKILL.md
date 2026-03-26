---
name: project-showcase
description: >
  自動掃描程式專案並生成演講用互動式 HTML 簡報（可轉 PDF），含講者備註、SVG 架構圖、
  導覽列、即時編輯與儲存回 MD 功能。適用於需要展示程式專案成果給非技術受眾
  （如建築師、設計師、業主）的場合。當使用者提到「準備簡報」、「演講」、「展示專案」、
  「做投影片」、「PPT」、「簡報」、「分享專案」、「talk」、「presentation」等關鍵字，
  或需要將專案轉換為可展示格式時，請使用此技能。即使使用者只是提到要跟別人分享專案成果，
  也應主動建議使用。支援多種時長（5/20/45 分鐘）、5 種可即時切換的視覺主題、
  卡片式排版大字體、SVG 架構圖/流程圖、佔位符與自動嵌入兩種圖片模式。
---

# Project Showcase — 專案展示簡報生成器

將程式專案自動轉換為**故事導向**的演講簡報。

**產出鏈**：掃描專案 → 生成 Presentation MD → 輸出互動式 HTML → 瀏覽器列印 PDF

## 核心理念

你的受眾是**建築師、設計師、業主**——不是工程師。他們在意的是：
- 這個工具**解決了什麼設計痛點**？
- 實際成果**長什麼樣**？（截圖、Demo、Before/After）
- **我能不能也用**？

不要做「技術文件投影片」，要做「**故事**」。技術名詞一律提供中英對照。

---

## 工作流程

### Step 1：理解需求

與使用者確認以下項目（如果使用者已提供綱要，從綱要中提取答案，缺的再問）：

| 項目 | 說明 | 預設值 |
|------|------|--------|
| **時長** | 5 / 20 / 45 分鐘 | 20 分鐘 |
| **受眾** | 誰會聽？他們的技術程度？ | 建築師（非程式背景） |
| **核心訊息** | 觀眾離開時要記住的 1-3 件事 | （必填） |
| **主題風格** | 5 種內建主題，可即時切換 | 淺色主題（minimal-white） |
| **圖片模式** | `placeholder`（佔位符）或 `auto`（自動嵌入） | placeholder |
| **使用者綱要** | 使用者認為重要的重點 | （選填） |

### Step 2：掃描專案

用工具（Glob、Read、Grep、Bash）掃描目前專案。

### Step 3：生成 Presentation MD

- **格式規範**：詳見 `references/md-format.md`
- **框架選擇**：根據時長自動選擇，詳見 `references/frameworks.md`
- **存檔位置**：`{專案目錄}/output/presentation_{YYYYMMDD}.md`
- **重要**：修改 MD 前必須先 Read 最新版本，避免覆蓋使用者的手動修改

**生成原則**：
1. 每張投影片只放**一個核心概念**
2. 用類比取代術語（如「就像 BIM 自動產生施工圖」）
3. 架構/流程性內容用 `<!-- diagram: xxx -->` 標記，生成 SVG 圖表
4. 每個技術功能後面加一句「**所以你可以...**」
5. 講者備註要寫成**可直接唸的講稿草稿**
6. 姓名/案名需要去識別化處理

### Step 4：與使用者確認

將 MD 存檔讓使用者直接在編輯器中修改。使用者確認後進入 Step 5。

### Step 5：生成互動式 HTML 簡報

```bash
python "{skill_path}/scripts/generate_html.py" "{md_path}" --output "{output_dir}/presentation.html"
```

其中 `{skill_path}` 是這個 `skill/` 資料夾的絕對路徑（例如 `C:/path/to/your/skill`）。

**HTML 簡報功能**：
- **卡片式排版**：大字體、大量留白，投影時清晰可見
- **5 種主題即時切換**：右上角色點，不需重新生成
- **SVG 架構圖**：用 `<!-- diagram: xxx -->` 在 MD 中標記，自動生成（目前支援 `architecture-6`、`refinery-protocol`、`golden-journey-pipeline`）
- **右側導覽列**：Notion 風格，點 ☰ 展開，可跳到任意頁面
- **多種導覽方式**：← → 鍵盤、滾輪上下（自動排除在導覽列上的捲動）、觸控滑動（點擊不換頁）
- **講者備註面板**：按 N 切換
- **自動拆頁**：多個 code blocks（如 `**案例 A/B/C**`）或多個表格自動各拆一頁；彈點過多也會二分拆頁；不捲動、不溢出
- **多表格支援**：同一頁多個 Markdown 表格按 `**案例 X**` 標記各拆一頁
- **Markdown 渲染**：`**粗體**`、`` `code` ``、`*斜體*` 正確轉 HTML
- **括弧文字縮小**：全形括號 `（...）` 自動縮小為 0.78em 次要色（`.aside-text`），形成視覺層次
- **Code block URL 可點選**：code block 內的裸 `https://...` 自動包成可點擊 `<a>` 連結
- **多圖 Gallery**：同一頁放多張圖片時，自動排列為 flex 縮圖格，點擊任一張開 lightbox
- **影片支援**：`![alt](video.mp4)` 渲染為 `<video controls>`，原生進度條＋全螢幕，點播不自動播放；影片不進 lightbox
- **影片鍵盤控制**：當前投影片有影片時，Space = 暫停/繼續，← = 倒退 5 秒，→ = 快進 5 秒；無影片時退回正常換頁行為
- **Lightbox 方向鍵換頁**：燈箱開啟時，← → 關閉燈箱並跳至上/下頁
- **滾輪隔離**：wheel listener 掛在 `.deck`（而非 `document`），lightbox 為兄弟元素不在冒泡路徑，加上 `lightboxOpen` flag 雙重隔離
- **動態導覽列**：側邊欄由 JS 從 slide 標題動態建立，支援 `section-title`、`cover-title`、`closing-title`、`slide-title`、`toc-title`
- **多卡片佈局**：`<!-- layout: top+columns -->` / `<!-- layout: columns -->` + `<!-- card -->` 分隔符，將單一投影片拆成多張並排視覺卡片；每張卡片各自 autofit（`autoFitCards`），字體從 CSS 最大值自動縮到不溢出
- **表格欄位折行策略**：非末欄 `white-space: nowrap`，短標籤欄縮到剛好、末欄承擔所有換行
- **自動縮小溢出文字**（`autoFitSlides`）：內容超出投影片高度時，自動從 0.88em 縮到 0.45em（步距 0.04）；跳過 section / cover / toc / closing 頁
- **頁尾**：底部中央顯示作者+日期（frontmatter `author` + `date`），右下顯示頁碼
- **列印 PDF**：Ctrl+P，列印時隱藏所有 UI 元素
- **即時編輯模式**：按左下角 ✏ 或 E 鍵進入（需搭配 server）

### Step 6：啟動編輯伺服器（選用）

如果使用者想在瀏覽器中直接編輯簡報並儲存回 MD：

```bash
python "{skill_path}/scripts/presentation_server.py" "{md_path}"
```

啟動後自動開啟瀏覽器 `http://localhost:8080`。

**編輯模式功能**：
- 按 ✏ 進入編輯模式，文字出現虛線框可直接修改
- `+ 新增頁` / `🗑 刪除此頁` / `📝 備註` / `💾 儲存`
- 儲存時自動：HTML → MD 反向轉換 → 寫回 .md 檔 → 重新生成 HTML → 頁面重載

### Step 7：匯出 PDF

**方式一：瀏覽器列印（推薦）**
1. 開啟 HTML → Ctrl+P → 另存為 PDF → A4 橫向、無邊距

**方式二：Playwright 自動化**
```bash
pip install playwright && playwright install chromium
python "{skill_path}/scripts/convert_to_pdf.py" "{html_path}"
```

---

## SVG 架構圖系統

在 MD 中用 `<!-- diagram: xxx -->` 標記的投影片會自動生成 SVG 圖表。

**目前支援的圖表**：
| diagram ID | 說明 |
|------------|------|
| `architecture-6` | Alex Diary 6 模組環狀架構圖 |
| `refinery-protocol` | Refinery Protocol 漏斗流程圖 |
| `golden-journey-pipeline` | Golden Journey 照片→知識管線圖 |

**新增圖表**：在 `generate_html.py` 的 `render_diagram_svg()` 中加入新的 SVG 渲染函式。

**設計原則**：
- SVG 用 CSS 變數（`var(--accent)` 等）確保跟主題連動
- 深色方塊內的文字用 `style="fill:#fff"` 強制白色
- 連線用直線/直角折線，不要斜線

---

## 圖片處理策略

### Placeholder 模式（預設）
帶虛線邊框的佔位區塊 + 描述文字。描述要具體。

### Auto 模式
自動搜尋並嵌入圖片（base64 內嵌）。

```markdown
![系統主畫面](output/screenshot.png)           <!-- 真實圖片 -->
![placeholder: 建築師操作介面的截圖]()          <!-- placeholder 模式 -->
```

### 多圖 Gallery（同頁多張圖）

同一頁放多個 `![](...)` 會自動轉為 flex 縮圖格，點擊放大：

```markdown
## 成果展示

![排煙窗對話框](image/排煙窗_對話框.jpg)
![立面標示](image/排煙窗_立面標示.jpg)
![Excel 輸出](image/排煙excel_房間明細.png)
```

### 影片支援

副檔名為 `.mp4`、`.webm`、`.ogg`、`.mov` 時自動渲染為原生 `<video controls>`：

```markdown
![帷幕排列預覽](image/帷幕排列預覽介面.mp4)
```

> **注意**：影片不進 lightbox，使用瀏覽器原生播放器（含進度條、全螢幕）。

### 括弧補充說明縮小

全形括號內的補充文字自動縮小為次要色（0.78em）：

```markdown
- 處理時間從 2 小時縮短為 5 分鐘（實測於 10 層樓以上專案）
```

圓括號 `（...）` → 自動套用 `.aside-text` 樣式，不需手動標記。

---

## 簡報設計原則

1. **領先展示成果**：先 demo 再解釋原理
2. **用建築類比**：「版本控制就像圖面版次管理」
3. **大量留白**：不要塞滿文字
4. **架構/流程用圖表**：不要只列點，用 SVG diagram
5. **一組資訊一頁**：Before/After/Tips 是一組，不要跨頁拆開
6. **表格分類呈現**：如斜線指令用分類表格（工時/知識/回顧）
7. **最後一張不是「謝謝」**：放核心 Takeaway

---

## 可用主題

所有主題都可在 HTML 中即時切換，右上角色點選擇。

| 主題 ID | 名稱 | 適合場景 |
|---------|------|----------|
| `minimal-white` | 極簡白 | 正式場合、學術分享（預設） |
| `architect-dark` | 建築深色 | 設計類展示、作品集 |
| `blueprint` | 藍圖 | 建築/工程/結構主題 |
| `concrete` | 清水模 | 材料/結構/工業風 |
| `nature` | 自然 | 景觀/綠建築/永續 |

---

## 資料夾結構

```
project-showcase/
├── SKILL.md                      ← 你正在看的這個
├── scripts/
│   ├── generate_html.py          ← MD → HTML 核心腳本（卡片式互動簡報 + SVG 圖表）
│   ├── presentation_server.py    ← 編輯伺服器（即時編輯 + 儲存回 MD）
│   ├── generate_pptx.py          ← MD → PPTX 備用腳本
│   └── convert_to_pdf.py         ← HTML/PPTX → PDF 轉換
└── references/
    ├── md-format.md              ← Presentation MD 格式規範（含 gallery、影片、括弧語法）
    ├── frameworks.md             ← 簡報框架（by 時長）
    ├── themes.md                 ← 主題定義與色彩配置
    └── dev-notes.md              ← generate_html.py 開發陷阱與設計決策（維護必讀）
```

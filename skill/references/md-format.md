# Presentation MD 格式規範

本格式是 `generate_pptx.py` 的輸入格式。用 Markdown 撰寫簡報內容，腳本會自動轉換為 PPTX。

## 基本結構

```markdown
---
title: "專案名稱"
subtitle: "一句話描述"
author: "你的名字"
date: "2026-03-21"
duration: 20
theme: "minimal-white"
image_mode: "placeholder"
project_path: "C:/Project/MyProject"
contact: "email@example.com"
---

（第一張投影片內容）

---

（第二張投影片內容）

---

（第三張投影片內容）
```

## Frontmatter 欄位

| 欄位 | 必填 | 說明 | 預設值 |
|------|------|------|--------|
| `title` | 是 | 簡報標題 | — |
| `subtitle` | 否 | 副標題 | 空 |
| `author` | 否 | 講者名稱 | 空 |
| `date` | 否 | 日期 | 空 |
| `duration` | 否 | 演講時長（分鐘） | 20 |
| `theme` | 否 | 主題 ID | minimal-white |
| `image_mode` | 否 | `placeholder` 或 `auto` | placeholder |
| `project_path` | 否 | 專案根目錄（圖片相對路徑基準） | MD 檔所在目錄 |
| `contact` | 否 | 聯絡方式（顯示在結語頁） | 空 |

## 投影片分隔

用獨佔一行的 `---` 分隔投影片：

```markdown
## 第一張投影片

內容...

---

## 第二張投影片

內容...
```

## 多卡片佈局

同一張投影片可分割為多張並排的視覺卡片，適合呈現「條件 A / 條件 B」或「三個平行概念」：

```markdown
<!-- layout: top+columns -->
## 投影片標題

<!-- card -->
頂部卡片（佔滿寬度）
- 主要說明
- 次要說明

<!-- card -->
### 欄位卡片 1 標題
欄位內容

<!-- card -->
### 欄位卡片 2 標題
欄位內容
```

支援兩種版型：

| layout 值 | 說明 |
|-----------|------|
| `top+columns` | 第一張卡片佔全寬，其餘卡片水平並排 |
| `columns` | 所有卡片水平並排（等寬） |

- 卡片分隔符：`<!-- card -->`
- 每張卡片支援 `### 子標題`、`- 條列`、`+ 條列`、純文字
- `<!-- card -->` 之前、`## ` 標題行之前的空白區段會自動過濾

---

## 投影片類型

### 自動推斷

腳本會根據內容自動判斷投影片類型：

| 條件 | 推斷為 |
|------|--------|
| `#` 標題（H1） | `section`（區段標題） |
| `##` 標題 + 條列 | `content`（一般內容） |
| 只有圖片，無條列 | `image`（大圖） |
| 圖片 + 條列都有 | `split`（左文右圖） |

### 手動指定

用 HTML 註解強制指定類型：

```markdown
<!-- type: comparison -->
## Before vs After

| Before | After |
|--------|-------|
| 手動 2 小時 | 自動 5 分鐘 |
```

可用類型：
- `cover` — 封面（通常不需手動建立，frontmatter 會自動生成）
- `section` — 區段標題（大字置中）
- `content` — 一般內容（標題 + 條列）
- `image` — 大圖（圖片佔滿）
- `split` — 左文右圖
- `comparison` — 對比（兩欄或表格）
- `quote` — 引言
- `demo` — Live Demo（帶徽章標記）
- `closing` — 結語

## 元素語法

### 標題

```markdown
# 區段標題          → 產生 section 投影片
## 投影片標題        → 投影片的主標題
### 副標題          → 投影片的副標題
```

### 條列項目

```markdown
- 第一點
- 第二點
  - 子項目（前面加兩個空格）
  - 另一個子項目
- 第三點
```

### 圖片

```markdown
<!-- 真實圖片 -->
![系統截圖](output/screenshot.png)
![操作畫面](assets/demo.gif)

<!-- 佔位符圖片（顯示為灰色方塊 + 描述）-->
![placeholder: 系統主畫面操作截圖]()
![placeholder: Before/After 對比圖]()
```

### 多圖 Gallery（同頁多張圖）

同一投影片放多個 `![](...)` 時，自動排列為 flex 縮圖格，點任一張可放大：

```markdown
## 排煙窗法規檢討成果

![對話框成果](image/排煙窗_對話框.jpg)
![立面標示](image/排煙窗_立面標示.jpg)
![案例](image/排煙窗_案例.jpg)
```

> **注意**：圖片路徑必須使用 `![alt](path)` 完整語法，不能只寫 `file.jpg`。

### 影片

副檔名為 `.mp4`、`.webm`、`.ogg`、`.mov` 時自動渲染為 `<video controls>`：

```markdown
## 帷幕牆排列展示

![帷幕排列預覽介面](image/帷幕排列預覽介面.mp4)
```

- 點播，不自動播放
- 原生瀏覽器控制列（進度條、音量、全螢幕）
- **影片不進 lightbox**，直接在投影片內播放

### 括弧補充說明

全形括號 `（...）` 自動縮小為次要色（0.78em），形成視覺層次：

```markdown
- 處理時間從 2 小時縮短為 5 分鐘（實測於 10 層樓以上專案）
- 支援 Revit 2022–2026（需安裝對應版本的外掛）
```

不需手動加 class，`（...）` 自動套用。半形括號 `(...)` 不受影響。

### 講者備註

```markdown
> 講者備註：這裡寫你要講的內容。建議寫成可直接唸出來的講稿草稿，而不是簡略的提示詞。

> Note: English notes are also supported.
```

### 程式碼區塊

````markdown
```python
# 會顯示在投影片上的程式碼（建議 ≤ 10 行）
result = process_building_model(input_file)
print(f"完成：{result.floor_count} 層樓")
```
````

### 表格

```markdown
| 項目 | 手動流程 | 自動化後 |
|------|----------|----------|
| 時間 | 2 小時 | 5 分鐘 |
| 錯誤率 | 15% | < 1% |
```

## 完整範例

```markdown
---
title: "BuildingSync — 建築圖面自動同步工具"
subtitle: "改一處，全部同步更新"
author: "Alex Chen"
date: "2026-03-25"
duration: 20
theme: "architect-dark"
image_mode: "placeholder"
---

# 問題與動機

---

## 建築師的日常痛點

- 一個窗戶尺寸改了，要手動更新 10 張圖
- 平立剖面不一致，被工地打電話追殺
- 設計變更歷史無法追溯（Version Control 版本控制）

![placeholder: 建築師在桌前對著大量圖紙的場景]

> 講者備註：先問觀眾：「在座有多少人，改過一個窗戶尺寸，結果要改十張圖？」等待舉手後再繼續。

---

# 解決方案

---

## BuildingSync 做了什麼

- 定義一次參數，所有圖面自動同步
- 就像 BIM 的「連動更新」，但更輕量
- 支援 AutoCAD / Revit / SketchUp 格式

![placeholder: 系統操作介面截圖]

> 講者備註：這裡用類比：「就像 Excel 的公式連動，改一個儲存格，所有引用它的地方都自動更新。」

---

<!-- type: demo -->
## 現場示範

- 打開一個包含 10 張圖的專案
- 修改窗戶 W1 的寬度從 120cm 改為 150cm
- 看所有圖面即時同步更新

> 講者備註：如果 Demo 失敗，切到下一頁的預錄截圖。

---

<!-- type: comparison -->
## 效率對比

| 項目 | 手動流程 | 使用 BuildingSync |
|------|----------|-------------------|
| 修改時間 | 2 小時 | 5 分鐘 |
| 遺漏率 | 約 15% | < 1% |
| 版本追溯 | 無 | 完整歷史 |

> 講者備註：強調「2小時 vs 5分鐘」的差距，讓數字說話。

---

<!-- type: closing -->
## 改一處，全部自動同步

- 省下重複修改的時間，專注設計本身
- 減少圖面不一致的風險
- 完整的設計變更歷史

> 講者備註：結尾回扣開頭的問題：「現在改一個窗戶，不用再改十張圖了。」
```

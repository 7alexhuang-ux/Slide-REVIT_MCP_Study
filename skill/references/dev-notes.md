# generate_html.py 開發筆記

維護 `generate_html.py` 時必讀。記錄踩過的坑與設計決策。

---

## 坑 1：generate_css vs generate_js 的大括號規則

這是最容易出錯的地方。**兩者都是 plain `'''` string，但大括號寫法完全不同。**

### generate_css()

```python
def generate_css() -> str:
    return '''
    * {{ margin: 0; padding: 0; }}   ← CSS 大括號必須寫 {{}}
    body {{ color: var(--text); }}
    '''
```

原因：Line 2729 會做後處理：
```python
css_text = generate_css().replace('{{', '{').replace('}}', '}')
```
`{{}}` 在 plain string 裡是字面值 `{{}}`，`.replace()` 後才變成合法 CSS 的 `{}`。
然後 `{css_text}` 插入外層 f-string 時，`css_text` 的值（已是 `{}`）直接嵌入，不再解析。

### generate_js()

```python
def generate_js(default_theme: str) -> str:
    return '''
    const themes = { 'dark': { bg: '#000' } };   ← JS 大括號直接寫 {}
    function foo() { return 1; }
    '''
```

原因：外層 f-string 是 `<script>{generate_js(default_theme)}</script>`。
f-string 先呼叫函式取得返回值，再把字串值嵌入。返回的字串裡的 `{` 不會被再次解析為 f-string placeholder。

### 快速記憶法

| 函數 | 大括號 | 理由 |
|------|--------|------|
| `generate_css()` | `{{}}` | 經過 `.replace()` 後處理 |
| `generate_js()` | `{}` | 函式返回值直接插入，不再解析 |

---

## 坑 2：scrollHeight 在 flexbox + overflow:hidden 不可靠

**錯誤做法**：
```javascript
if (slide.scrollHeight > slide.clientHeight) { /* 縮小 */ }
```

`overflow: hidden` 時，`scrollHeight` 不反映真實溢出高度，flexbox 子元素的高度計算也可能不準。

**正確做法**（`autoFitSlides` 採用）：
```javascript
function totalH() {
    let h = 0;
    for (const el of slide.children) {
        if (el === footer) continue;
        const cs = getComputedStyle(el);
        h += el.offsetHeight + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
    }
    return h;
}
if (totalH() > available) { /* 縮小 */ }
```

---

## 坑 3：JS 要在 layout 完成後才能計算高度

`autoFitSlides` 若在 DOM 插入後同步執行，`offsetHeight` 可能都是 0（因為 browser 還沒 layout）。

**解法**：雙層 `requestAnimationFrame`：
```javascript
requestAnimationFrame(function() {
    requestAnimationFrame(autoFitSlides);
});
```
第一層排進 paint 佇列，第二層確保 layout 真正完成。

---

## 坑 4：autoFitSlides 要跳過特定投影片類型

section、cover、toc、closing 這類頁面是全畫面設計，字體很大是刻意的，不應縮小。

```javascript
const SKIP = ['slide--section', 'slide--cover', 'slide--toc', 'slide--closing'];
if (SKIP.some(c => slide.classList.contains(c))) return;
```

縮放範圍：`0.88em` → `0.45em`，步距 `0.04`。

---

## 坑 5：圖片語法一定要完整的 Markdown 格式

使用者在 MD 裡直接寫 `file.jpg` 或 `image/file.jpg`（沒有 `![](...)`）不會被解析為圖片，
會被當成普通文字。

**正確**：
```markdown
![排煙窗案例](image/排煙窗_案例.jpg)
```

**錯誤（不會顯示圖片）**：
```
排煙窗_案例.jpg
image/排煙窗_案例.jpg
```

---

## 坑 6：多個 ## 段落沒有 --- 分隔符會塞進同一頁

```markdown
## 工作流程的改變
表格...

## 學到的Lesson          ← 沒有 --- → 跟上一頁合在一起！
條列...

## 三件值得帶走的事       ← 同上
條列...
```

**正確寫法**：每個 `##` 前都要有 `---`：
```markdown
## 工作流程的改變
表格...

---

## 學到的Lesson
條列...

---

## 三件值得帶走的事
條列...
```

---

## 坑 7：行末反斜線 `\` 會合併下一行

MD 裡行末的 `\` 在某些解析器裡代表「繼續」，會把下一行連過來導致排版異常。
如果只是排版換行需求，直接讓文字自然折行即可，不需要 `\`。

---

## 坑 8：編輯 MD 前必須先重新讀取

Claude Code 的 Edit 工具要求：**自上次 Read 後若檔案有被修改，Edit 會失敗**（`File has been modified since read`）。

使用者常在 IDE 裡直接修改 MD，然後要求更新 HTML。此時必須先 `Read` 再 `Edit`，而不是直接 `Edit`。

---

## 設計決策：影片不進 Lightbox

影片用原生 `<video controls playsinline preload="metadata">` 而非自訂 lightbox。

**理由**：
- 影片進 lightbox 需要自訂播放器（播放/暫停、進度條、音量、全螢幕）
- 瀏覽器原生 `<video controls>` 已內建這些功能
- 工程量差距大，原生方案效果更好

影片 CSS（`generate_css()` 中）：
```css
.slide-video {
    background: #000;
    width: 100%;
}
```

---

## 設計決策：Lightbox 與換頁的隔離（重要，多次踩坑）

### 根本原則：不要用 `stopPropagation()`，要用架構隔離

`.image-modal` 是 `.deck` 的**兄弟元素**，不是子元素。
把換頁 wheel handler 掛在 `.deck` 而非 `document`，
當 lightbox 開啟時，滾輪事件的冒泡路徑是：
`modal-img → modal-viewport → image-modal → body → document`，
永遠不會經過 `.deck`，天然隔離。

```javascript
document.querySelector('.deck').addEventListener('wheel', e => {
    if (editMode || wheelLock || lightboxOpen) return;
    ...
}, { passive: true });
```

### lightboxOpen flag（belt-and-suspenders）

不要只靠 `imageModal.classList.contains('visible')` 判斷，
要在 `openLightbox()`/`closeLightbox()` 明確設定 `lightboxOpen` boolean：

```javascript
let lightboxOpen = false;
function openLightbox(src) { ...; imageModal.classList.add('visible'); lightboxOpen = true; }
function closeLightbox() { imageModal.classList.remove('visible'); lightboxOpen = false; }
```

### 點擊換頁已完全停用

過去用 `document.addEventListener('click', ...)` 做點擊換頁，造成：
- 點圖片開 lightbox → 同時觸發換頁
- 點影片播放 → 同時觸發換頁
- 點關閉 lightbox → 同時觸發換頁

**解法：完全移除點擊換頁**，只保留滾輪和鍵盤左右鍵換頁。

### toc-title 需加入 buildNav() 選擇器

目錄頁的標題用 `.toc-title`，不在 `buildNav()` 預設選擇器裡，
會 fallback 為「投影片 X」。正確的選擇器：

```javascript
slide.querySelector('.section-title, .cover-title, .closing-title, .slide-title, .toc-title')
```

---

## 坑 9：需要對齊的多行文字要用 fenced code block

MD 裡縮排的 continuation line 不會保留空白，會被壓成一行。
若有需要對齊的多欄範例（如面板排列記號），必須用 fenced code block：

**錯誤（對齊會消失）**：
```markdown
+ 說明
  121212　　ABBABBA　　ABCD
  121212　　ABBABBA　　BCDA
```

**正確（保留對齊）**：
```markdown
+ 說明

  \```
  121212　　ABBABBA　　ABCD
  121212　　ABBABBA　　BCDA
  \```
```

---

## 坑 10：多卡片佈局的字體繼承

`.bullet-list > li` 有全域的 `font-size: clamp(1.4rem, 2.6vw, 2.22rem)`。
如果只在 `.card-column .bullet-list` 設 `font-size`，會因為 CSS 特異性而被 `.bullet-list > li` 覆蓋。

**正確做法**：在 `.card-column`（容器本身）設基準 `font-size`，然後用 `.card-column .bullet-list > li { font-size: 1em; }` 讓子元素繼承容器值。這樣 `autoFitCards` 只需要控制容器的 `fontSize` 就能縮放所有內容。

```css
/* 正確 */
.card-column { font-size: clamp(1.0rem, 1.8vw, 1.5rem); }
.card-column .bullet-list > li { font-size: 1em; }  /* 繼承 */

/* 錯誤（被全域 .bullet-list > li 覆蓋） */
.card-column .bullet-list { font-size: ...; }
```

---

## 設計決策：autoFitCards — per-card 字體自動縮放

多卡片佈局的問題：`autoFitSlides` 是整頁縮放，但每張卡片內容量不同，需要各自縮放。

**解法**：`autoFitCards()` 函式，在 `autoFitSlides` 執行完後再跑：

```javascript
function autoFitCards() {
    document.querySelectorAll('.card-column, .card-top').forEach(card => {
        card.style.fontSize = '';  // reset 到 CSS 最大值
        const padV = parseFloat(getComputedStyle(card).paddingTop) + ...;
        const available = card.clientHeight - padV;
        // 用 offsetHeight 累加子元素高度（不受 overflow:hidden 影響）
        function totalH() { ... }
        if (totalH() <= available) return;
        const startPx = parseFloat(getComputedStyle(card).fontSize);
        for (let px = startPx - step; px >= minPx; px -= step) {
            card.style.fontSize = px + 'px';
            if (totalH() <= available) break;
        }
    });
}
```

**執行順序**：
```javascript
requestAnimationFrame(() => requestAnimationFrame(() => {
    autoFitSlides();           // 1. 整頁縮放
    requestAnimationFrame(autoFitCards);  // 2. 等 reflow 後 per-card 縮放
}));
```

**結果**：每張卡片各自取「剛好塞得下的最大字體」，文字少的卡片字大，文字多的卡片字小。

**為什麼不選其他方案**：
- 選項 A（統一字體）：簡單，但文字多的卡片可能溢出，文字少的浪費空間
- 選項 B（依欄數縮放）：只考慮欄數，不考慮實際內容量，不夠精準
- 選項 C（per-card autofit）：每張卡片取各自最大可用字體，最符合「盡可能大但不斷行太多」需求

同排卡片字體大小不同是**預期行為**（各自的內容量本來就不同），不是 bug。

---

## 設計決策：表格非末欄不折行

表格的第一欄（標籤）和中間欄（短項目）不應折行，折行應由最後一欄（說明文字）承擔。

```css
.card-table td:not(:last-child),
.card-table th:not(:last-child) { white-space: nowrap; }
```

這是一個通則，適用於大多數比較表格（前面是標籤/短項，最後一欄是說明）。

---

## 資料夾結構更新（SKILL.md 的補充）

```
references/
├── md-format.md     ← 語法規範（含 gallery、影片、括弧語法）
├── frameworks.md    ← 簡報框架（by 時長）
├── themes.md        ← 主題定義
└── dev-notes.md     ← 你正在看的這個
```

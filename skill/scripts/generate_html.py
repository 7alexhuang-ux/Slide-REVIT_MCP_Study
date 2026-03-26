#!/usr/bin/env python3
"""
Project Showcase — HTML 卡片式簡報生成器
從 Presentation Markdown 生成互動式 HTML 簡報

特色：
  - 卡片式佈局，大字體清楚易讀
  - 全螢幕投影片導航（鍵盤 ←→ / 點擊）
  - 講者備註面板（按 N 切換）
  - 進度條
  - 5 種內建主題
  - @media print 支援直接列印成 PDF

用法:
    python generate_html.py input.md [--output output.html]

依賴:
    pip install PyYAML
"""

import argparse
import io
import json
import os
import re
import sys
import yaml
from pathlib import Path

if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# ============================================================
# 主題定義（CSS 變數）
# ============================================================

THEMES = {
    'minimal-white': {
        'name': '極簡白',
        'bg': '#FFFFFF',
        'bg_card': '#F7F7F7',
        'bg_card_hover': '#F0F0F0',
        'text_primary': '#3A3A3A',
        'text_secondary': '#666666',
        'accent': '#007ACC',
        'accent_light': 'rgba(0, 122, 204, 0.08)',
        'muted': '#999999',
        'border': '#E5E5E5',
        'shadow': 'rgba(0,0,0,0.06)',
        'font': "'Noto Serif TC', 'Microsoft JhengHei', serif",
    },
    'architect-dark': {
        'name': '建築深色',
        'bg': '#0F0F0F',
        'bg_card': '#1A1A1A',
        'bg_card_hover': '#222222',
        'text_primary': '#E0E0E0',
        'text_secondary': '#AAAAAA',
        'accent': '#D4A853',
        'accent_light': 'rgba(212, 168, 83, 0.10)',
        'muted': '#666666',
        'border': '#2A2A2A',
        'shadow': 'rgba(0,0,0,0.3)',
        'font': "'Noto Serif TC', 'Microsoft JhengHei', serif",
    },
    'blueprint': {
        'name': '藍圖',
        'bg': '#1B2A4A',
        'bg_card': '#223356',
        'bg_card_hover': '#2A3D66',
        'text_primary': '#E0E8F0',
        'text_secondary': '#B0C8E0',
        'accent': '#5DADE2',
        'accent_light': 'rgba(93, 173, 226, 0.10)',
        'muted': '#6680A0',
        'border': '#2E4570',
        'shadow': 'rgba(0,0,0,0.3)',
        'font': "'Noto Serif TC', 'Microsoft JhengHei', serif",
    },
    'concrete': {
        'name': '清水模',
        'bg': '#E8E4E1',
        'bg_card': '#D8D3CE',
        'bg_card_hover': '#CEC8C2',
        'text_primary': '#444444',
        'text_secondary': '#666666',
        'accent': '#8B4513',
        'accent_light': 'rgba(139, 69, 19, 0.08)',
        'muted': '#8B8B8B',
        'border': '#C0BAB5',
        'shadow': 'rgba(0,0,0,0.08)',
        'font': "'Noto Serif TC', 'Microsoft JhengHei', serif",
    },
    'nature': {
        'name': '自然',
        'bg': '#FAFAF5',
        'bg_card': '#F0EDE5',
        'bg_card_hover': '#E8E4DA',
        'text_primary': '#3D4446',
        'text_secondary': '#555555',
        'accent': '#27AE60',
        'accent_light': 'rgba(39, 174, 96, 0.08)',
        'muted': '#95A5A6',
        'border': '#D5D0C8',
        'shadow': 'rgba(0,0,0,0.06)',
        'font': "'Noto Serif TC', 'Microsoft JhengHei', serif",
    },
}


# ============================================================
# Markdown 解析（同 generate_pptx.py）
# ============================================================

def parse_frontmatter(content: str):
    if content.startswith('---'):
        parts = content.split('---', 2)
        if len(parts) >= 3:
            try:
                config = yaml.safe_load(parts[1]) or {}
            except yaml.YAMLError:
                config = {}
            return config, parts[2]
    return {}, content


def parse_slides(body: str):
    raw_slides = re.split(r'\n---\s*\n', body)
    slides = []
    for raw in raw_slides:
        raw = raw.strip()
        if not raw:
            continue
        slide = parse_single_slide(raw)
        if slide:
            slides.append(slide)
    return slides


def parse_single_slide(raw: str) -> dict:
    slide = {
        'type': None,
        'title': '',
        'subtitle': '',
        'bullets': [],
        'images': [],
        'speaker_notes': '',
        'code_blocks': [],
        'raw_text': '',
        'table': None,
        'diagram': None,
        'multi_layout': None,
        '_raw_cards': [],
    }

    type_match = re.search(r'<!--\s*(?:type|slide):\s*(\S+)\s*-->', raw)
    if type_match:
        slide['type'] = type_match.group(1)
        raw = re.sub(r'<!--\s*(?:type|slide):\s*\S+\s*-->', '', raw)

    diagram_match = re.search(r'<!--\s*diagram:\s*(\S+)\s*-->', raw)
    if diagram_match:
        slide['diagram'] = diagram_match.group(1)
        slide['type'] = 'diagram'
        raw = re.sub(r'<!--\s*diagram:\s*\S+\s*-->', '', raw)

    layout_match = re.search(r'<!--\s*layout:\s*(\S+)\s*-->', raw)
    if layout_match:
        slide['multi_layout'] = layout_match.group(1)
        raw = re.sub(r'<!--\s*layout:\s*\S+\s*-->', '', raw)
        card_sep = re.compile(r'\n?<!--\s*card\s*-->\n?')
        slide['_raw_cards'] = [p.strip() for p in card_sep.split(raw) if p.strip()]

    code_blocks = re.findall(r'```(\w*)\n(.*?)```', raw, re.DOTALL)
    slide['code_blocks'] = [{'lang': lang, 'code': code.strip()} for lang, code in code_blocks]
    raw_no_code = re.sub(r'```\w*\n.*?```', '', raw, flags=re.DOTALL)

    table_matches = re.findall(r'(\|.+\|\n\|[-| :]+\|\n(?:\|.+\|\n?)+)', raw_no_code)
    if table_matches:
        slide['table'] = [t.strip() for t in table_matches]
        for t in table_matches:
            raw_no_code = raw_no_code.replace(t, '')

    lines = raw_no_code.strip().split('\n')
    text_lines = []

    for line in lines:
        stripped = line.strip()
        if stripped.startswith('# ') and not stripped.startswith('## '):
            slide['title'] = stripped[2:].strip()
            if slide['type'] is None:
                slide['type'] = 'section'
        elif stripped.startswith('## '):
            slide['title'] = stripped[3:].strip()
        elif stripped.startswith('### '):
            slide['subtitle'] = stripped[4:].strip()
            text_lines.append(stripped)  # 保留在內容流中以便渲染為子標題
        elif stripped.startswith('- '):
            slide['bullets'].append(stripped[2:].strip())
        elif stripped.startswith('  - '):
            slide['bullets'].append('  ' + stripped[4:].strip())
        elif re.match(r'!\[', stripped):
            img_match = re.match(r'!\[(.*?)\]\((.*?)\)', stripped)
            if img_match:
                alt = img_match.group(1)
                src = img_match.group(2)
                if alt.lower().startswith('placeholder:'):
                    slide['images'].append({
                        'type': 'placeholder',
                        'description': alt.split(':', 1)[1].strip(),
                    })
                else:
                    slide['images'].append({
                        'type': 'file',
                        'path': src,
                        'alt': alt,
                    })
        elif stripped.startswith('> 講者備註:') or stripped.startswith('> 講者備註：'):
            note_text = re.split(r'[:：]', stripped, maxsplit=1)[1].strip() if re.search(r'[:：]', stripped) else ''
            slide['speaker_notes'] = note_text
        elif stripped.startswith('> Note:'):
            slide['speaker_notes'] = stripped.split(':', 1)[1].strip()
        elif stripped and not stripped.startswith('<!--'):
            text_lines.append(stripped)

    slide['raw_text'] = '\n'.join(text_lines)

    if slide['type'] is None:
        if slide['images'] and not slide['bullets']:
            slide['type'] = 'image'
        elif slide['images'] and slide['bullets']:
            slide['type'] = 'split'
        elif slide['table']:
            slide['type'] = 'content'
        else:
            slide['type'] = 'content'

    return slide


# 投影片內容太長時自動拆成多頁
MAX_ITEMS = 6  # 降低上限，讓較大字體有發揮空間

def _content_weight(slide: dict) -> int:
    """估算投影片內容量"""
    w = len(slide['bullets'])
    if slide['code_blocks']:
        w += sum(2 + cb['code'].count('\n') // 3 for cb in slide['code_blocks'])
    if slide['table']:
        tables = slide['table'] if isinstance(slide['table'], list) else [slide['table']]
        w += sum(2 + t.count('\n') for t in tables)
    if slide['images']:
        w += 2 * len(slide['images'])
    if slide['raw_text']:
        w += max(1, len(slide['raw_text'].split('\n')) // 2)
    return w


def _find_case_splits(raw_lines: list, n: int) -> list:
    """找到案例分割點（**案例 X** 標記），確保返回 n+1 個點"""
    case_indices = []
    for idx, line in enumerate(raw_lines):
        if re.match(r'\*\*案例\s*[A-Z]', line.strip()):
            case_indices.append(idx)
    # 用第 2 個及以後的標記作為分割點（第 1 個屬於第 1 段的一部分）
    split_points = [0] + case_indices[1:] + [len(raw_lines)]
    if len(split_points) < n + 1:
        step = max(1, len(raw_lines) // n)
        split_points = [j * step for j in range(n)] + [len(raw_lines)]
    return split_points[:n] + [split_points[-1]]


def _split_multi_codeblocks(slide: dict) -> list:
    """多個 code blocks → 每個獨立一頁（按案例標記對應 raw_text）"""
    n = len(slide['code_blocks'])
    raw_lines = slide['raw_text'].split('\n')
    splits = _find_case_splits(raw_lines, n)
    pages = []
    for j in range(n):
        s = dict(slide)
        s['code_blocks'] = [slide['code_blocks'][j]]
        s['table'] = None
        is_last = (j == n - 1)
        s['images'] = slide['images'] if is_last else []
        s['speaker_notes'] = slide['speaker_notes'] if is_last else ''
        s['bullets'] = slide['bullets'] if is_last else []
        s['raw_text'] = '\n'.join(raw_lines[splits[j]:splits[j + 1]]).strip()
        pages.append(s)
    return pages


def _split_multi_tables(slide: dict) -> list:
    """多個 tables → 每個獨立一頁（按案例標記對應 raw_text）"""
    tables = slide['table'] if isinstance(slide['table'], list) else [slide['table']]
    n = len(tables)
    raw_lines = slide['raw_text'].split('\n')
    splits = _find_case_splits(raw_lines, n)
    pages = []
    for j in range(n):
        s = dict(slide)
        s['table'] = [tables[j]]
        s['code_blocks'] = []
        is_last = (j == n - 1)
        s['images'] = slide['images'] if is_last else []
        s['speaker_notes'] = slide['speaker_notes'] if is_last else ''
        s['bullets'] = slide['bullets'] if j == 0 else []
        s['raw_text'] = '\n'.join(raw_lines[splits[j]:splits[j + 1]]).strip()
        pages.append(s)
    return pages


def split_long_slides(slides: list) -> list:
    """將過長的投影片拆成多頁（queue-based，支援多輪次拆分）"""
    result = []
    queue = list(slides)

    while queue:
        slide = queue.pop(0)

        if slide['type'] in ('section', 'closing', 'cover', 'quote', 'diagram'):
            result.append(slide)
            continue

        n_cbs = len(slide['code_blocks'])
        tables = slide['table'] if isinstance(slide['table'], list) else ([slide['table']] if slide['table'] else [])
        n_tables = len(tables)

        # 優先處理：多 code blocks 或多 tables（無論 weight 一律拆分）
        if n_cbs >= 2:
            queue[:0] = _split_multi_codeblocks(slide)
            continue

        if n_tables >= 2:
            queue[:0] = _split_multi_tables(slide)
            continue

        weight = _content_weight(slide)
        if weight <= MAX_ITEMS:
            result.append(slide)
            continue

        n_bullets = len(slide['bullets'])

        if n_bullets > MAX_ITEMS:
            mid = (n_bullets + 1) // 2
            s1 = dict(slide)
            s1['bullets'] = slide['bullets'][:mid]
            s1['code_blocks'] = []
            s1['table'] = None
            s2 = dict(slide)
            s2['title'] = slide['title'] + '（續）'
            s2['bullets'] = slide['bullets'][mid:]
            s2['images'] = []
            s2['speaker_notes'] = ''
            queue[:0] = [s1, s2]
            continue

        if n_bullets > 0 and (slide['code_blocks'] or slide['table']):
            s1 = dict(slide)
            s1['code_blocks'] = []
            s1['table'] = None
            s2 = dict(slide)
            s2['title'] = slide['title'] + '（續）'
            s2['bullets'] = []
            s2['images'] = []
            s2['speaker_notes'] = ''
            queue[:0] = [s1, s2]
            continue

        raw_lines_list = slide['raw_text'].split('\n') if slide['raw_text'] else []
        if len(raw_lines_list) > 8:
            mid = len(raw_lines_list) // 2
            for i in range(mid, len(raw_lines_list)):
                if not raw_lines_list[i].strip() or re.match(r'\*\*案例\s*[A-Z]', raw_lines_list[i].strip()):
                    mid = i
                    break
            s1 = dict(slide)
            s1['raw_text'] = '\n'.join(raw_lines_list[:mid])
            s2 = dict(slide)
            s2['title'] = slide['title'] + '（續）'
            s2['raw_text'] = '\n'.join(raw_lines_list[mid:])
            s2['bullets'] = []
            s2['images'] = []
            s2['speaker_notes'] = ''
            queue[:0] = [s1, s2]
            continue

        result.append(slide)

    return result


# ============================================================
# HTML 生成
# ============================================================

def escape_html(text: str) -> str:
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def autolink_urls(html_text: str) -> str:
    """將純文字 URL 轉為可點擊的 <a> 連結（在 escape_html 之後呼叫）"""
    return re.sub(
        r'(https?://[^\s<>&"]+)',
        r'<a href="\1" target="_blank" rel="noopener" class="code-link">\1</a>',
        html_text
    )


def md_inline(text: str) -> str:
    """將 inline Markdown 轉為 HTML（粗體、斜體、行內程式碼、連結）並支援 \n 換行"""
    t = escape_html(text).replace('\n', '<br>')
    # 粗體 **text**
    t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    # 斜體 *text*（避免與粗體衝突）
    t = re.sub(r'(?<!\*)\*([^*]+?)\*(?!\*)', r'<em>\1</em>', t)
    # 行內程式碼 `code`
    t = re.sub(r'`([^`]+?)`', r'<code>\1</code>', t)
    # 全形括號內補充說明縮小
    t = re.sub(r'（([^）]+)）', r'<span class="aside-text">（\1）</span>', t)
    return t


def render_table_html(table_md: str) -> str:
    lines = [l.strip() for l in table_md.strip().split('\n') if l.strip()]
    if len(lines) < 2:
        return ''
    headers = [c.strip() for c in lines[0].strip('|').split('|')]
    rows = []
    for line in lines[2:]:
        cells = [c.strip() for c in line.strip('|').split('|')]
        rows.append(cells)

    html = '<table class="card-table"><thead><tr>'
    for h in headers:
        html += f'<th>{md_inline(h)}</th>'
    html += '</tr></thead><tbody>'
    for row in rows:
        html += '<tr>'
        for cell in row:
            html += f'<td>{md_inline(cell)}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    return html


def render_card_section(raw: str) -> str:
    """將單一卡片的 Markdown 文字轉為 HTML 內容（不含外層 card div）"""
    if not raw.strip():
        return ''
    # Extract code blocks
    code_blocks = re.findall(r'```(\w*)\n(.*?)```', raw, re.DOTALL)
    raw_no_code = re.sub(r'```\w*\n.*?```', '', raw, flags=re.DOTALL)
    # Extract tables
    tables = re.findall(r'(\|.+\|\n\|[-| :]+\|\n(?:\|.+\|\n?)+)', raw_no_code)
    for t in tables:
        raw_no_code = raw_no_code.replace(t, '')

    lines = raw_no_code.strip().split('\n')
    parts = []
    bullets = []
    text_lines = []

    def flush_bullets():
        if bullets:
            items_html = ''.join(
                f'<li>{"<ul><li>" + md_inline(b[2:]) + "</li></ul>" if b.startswith("  ") else md_inline(b)}</li>'
                for b in bullets
            )
            parts.append(f'<ul class="bullet-list">{items_html}</ul>')
            bullets.clear()

    def flush_text():
        if text_lines:
            parts.append(''.join(
                f'<p class="body-text">{md_inline(l)}</p>' for l in text_lines if l.strip()
            ))
            text_lines.clear()

    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith('<!--'):
            continue
        if stripped.startswith('> 講者備註') or stripped.startswith('> Note:') or stripped.startswith('> note:'):
            continue
        if stripped.startswith('## ') or stripped.startswith('# '):
            continue  # 投影片標題由外層渲染
        if stripped.startswith('### '):
            flush_bullets()
            flush_text()
            parts.append(f'<h3 class="card-heading">{md_inline(stripped[4:])}</h3>')
        elif stripped.startswith('- '):
            flush_text()
            bullets.append(stripped[2:].strip())
        elif stripped.startswith('  - '):
            flush_text()
            bullets.append('  ' + stripped[4:].strip())
        elif stripped.startswith('+ '):
            flush_text()
            bullets.append(stripped[2:].strip())
        elif re.match(r'\d+\.\s', stripped):
            flush_text()
            bullets.append(re.sub(r'^\d+\.\s+', '', stripped))
        else:
            flush_bullets()
            text_lines.append(stripped)

    flush_bullets()
    flush_text()

    for lang, code in code_blocks:
        parts.append(f'<pre><code>{autolink_urls(escape_html(code.strip()))}</code></pre>')
    for t in tables:
        parts.append(render_table_html(t))

    return '\n'.join(parts)


def _render_multi_card(slide: dict) -> str:
    """渲染多卡片佈局（layout: top+columns 或 layout: columns）"""
    layout = slide['multi_layout']
    raw_cards = slide['_raw_cards']

    # 渲染每張卡片，過濾掉只有標題行的空卡片
    def card_html(raw):
        content = render_card_section(raw)
        return content.strip()

    cards = [card_html(r) for r in raw_cards]
    cards = [c for c in cards if c]  # 過濾空卡片

    if not cards:
        return f'<h2 class="slide-title">{md_inline(slide["title"])}</h2>'

    if layout == 'top+columns' and len(cards) >= 2:
        top = f'<div class="card card-top">{cards[0]}</div>'
        cols = ''.join(f'<div class="card card-column">{c}</div>' for c in cards[1:])
        layout_html = f'''<div class="multi-card-layout layout-top-columns">
            {top}
            <div class="columns-row">{cols}</div>
        </div>'''
    elif layout == 'columns' or (layout == 'top+columns' and len(cards) == 1):
        cols = ''.join(f'<div class="card card-column">{c}</div>' for c in cards)
        layout_html = f'''<div class="multi-card-layout layout-columns">
            <div class="columns-row">{cols}</div>
        </div>'''
    else:
        stacked = ''.join(f'<div class="card content-card">{c}</div>' for c in cards)
        layout_html = f'<div class="multi-card-layout">{stacked}</div>'

    return f'''
        <h2 class="slide-title">{md_inline(slide["title"])}</h2>
        {layout_html}'''


def render_slide_html(slide: dict, index: int, total: int) -> str:
    """將單一投影片轉為 HTML"""
    slide_type = slide['type']
    classes = f'slide slide--{slide_type}'

    inner = ''

    if slide.get('multi_layout') and slide.get('_raw_cards'):
        inner = _render_multi_card(slide)

    elif slide_type == 'toc':
        title = slide.get('title') or '今天的路線圖'
        flow_html = ''
        if slide.get('_sections'):
            svg = render_chapter_flow_svg(slide['_sections'])
            flow_html = f'<div class="chapter-flow-toc">{svg}</div>'
        inner = f'''
        <h2 class="toc-title">{md_inline(title)}</h2>
        {flow_html}'''

    elif slide_type == 'diagram':
        svg = render_diagram_svg(slide.get('diagram', ''))
        diagram_id = slide.get('diagram', '')
        inner = f'''
        <h2 class="slide-title">{md_inline(slide['title'])}</h2>
        <div class="diagram-container">{svg}</div>'''
        # 加 data-diagram 屬性以便編輯模式反轉回 MD
        extra_attrs = f' data-diagram="{diagram_id}"'

    elif slide_type == 'section':
        flow_html = ''
        if slide.get('_sections'):
            svg = render_chapter_flow_svg(slide['_sections'], highlight=slide.get('_section_num', -1), compact=True)
            flow_html = f'<div class="chapter-flow-compact">{svg}</div>'
        part_num = slide.get('_section_num', 0)
        display_title = f'Part {part_num}' if part_num else md_inline(slide['title'])
        inner = f'''
        <div class="section-with-flow">
            <div class="section-content">
                <div class="section-accent-line"></div>
                <h1 class="section-title">{display_title}</h1>
                {'<p class="section-subtitle">' + md_inline(slide["subtitle"]) + '</p>' if slide.get("subtitle") else ''}
            </div>
            {flow_html}
        </div>'''

    elif slide_type == 'demo':
        bullets_html = ''.join(
            f'<li><span class="step-num">Step {i+1}</span> {md_inline(b)}</li>'
            for i, b in enumerate(slide['bullets']) if not b.startswith('  ')
        )
        inner = f'''
        <div class="demo-badge">&#9654;  LIVE DEMO</div>
        <h2 class="slide-title">{md_inline(slide['title'])}</h2>
        <div class="card demo-card">
            <ol class="demo-steps">{bullets_html}</ol>
        </div>'''

    elif slide_type == 'comparison':
        comp_tables = ''
        if slide['table']:
            tables = slide['table'] if isinstance(slide['table'], list) else [slide['table']]
            comp_tables = ''.join(render_table_html(t) for t in tables)
        inner = f'''
        <h2 class="slide-title">{md_inline(slide['title'])}</h2>
        <div class="card comparison-card">{comp_tables}</div>'''

    elif slide_type == 'quote':
        inner = f'''
        <div class="quote-content">
            <span class="quote-mark">&ldquo;</span>
            <p class="quote-text">{md_inline(slide.get('raw_text', slide.get('title', '')))}</p>
        </div>'''

    elif slide_type == 'closing':
        bullets_html = ''.join(
            f'<li>{md_inline(b)}</li>' for b in slide['bullets']
        )
        inner = f'''
        <div class="closing-content">
            <h1 class="closing-title">{md_inline(slide['title'])}</h1>
            {'<ul class="closing-points">' + bullets_html + '</ul>' if bullets_html else ''}
        </div>'''

    elif slide_type == 'image':
        imgs = slide['images']
        img_html = _render_gallery(imgs) if len(imgs) > 1 else _render_image(imgs[0] if imgs else None)
        inner = f'''
        <h2 class="slide-title">{md_inline(slide['title'])}</h2>
        <div class="image-full">{img_html}</div>'''

    elif slide_type == 'split':
        img_html = _render_image(slide['images'][0] if slide['images'] else None)
        bullets_html = ''.join(
            f'<li>{"<ul><li>" + md_inline(b[2:]) + "</li></ul>" if b.startswith("  ") else md_inline(b)}</li>'
            for b in slide['bullets']
        )
        inner = f'''
        <h2 class="slide-title">{md_inline(slide['title'])}</h2>
        <div class="split-layout">
            <div class="card split-text">
                <ul class="bullet-list">{bullets_html}</ul>
            </div>
            <div class="split-image">{img_html}</div>
        </div>'''

    else:  # content (default)
        bullets_html = ''.join(
            f'<li>{"<ul><li>" + md_inline(b[2:]) + "</li></ul>" if b.startswith("  ") else md_inline(b)}</li>'
            for b in slide['bullets']
        )
        img_html = ''
        if slide['images']:
            imgs = slide['images']
            img_html = _render_gallery(imgs) if len(imgs) > 1 else _render_image(imgs[0])

        code_html = ''
        if slide['code_blocks']:
            code = slide['code_blocks'][0]
            code_html = f'<div class="card code-card"><pre><code>{autolink_urls(escape_html(code["code"]))}</code></pre></div>'

        table_html = ''
        if slide['table']:
            tables = slide['table'] if isinstance(slide['table'], list) else [slide['table']]
            table_html = ''.join(f'<div class="card table-card">{render_table_html(t)}</div>' for t in tables)

        text_part = ''
        if bullets_html:
            text_part = f'<div class="card content-card"><ul class="bullet-list">{bullets_html}</ul></div>'
        elif slide['raw_text']:
            raw_lines = slide['raw_text'].split('\n')
            def _raw_line(l):
                if l.startswith('### '):
                    return f'<h3 class="slide-subheading">{md_inline(l[4:])}</h3>'
                return f'<p class="body-text">{md_inline(l)}</p>'
            raw_paras = ''.join(_raw_line(l) for l in raw_lines if l.strip())
            text_part = f'<div class="card content-card">{raw_paras}</div>'

        if img_html and text_part:
            inner = f'''
            <h2 class="slide-title">{md_inline(slide['title'])}</h2>
            <div class="split-layout">
                <div class="split-text">{text_part}</div>
                <div class="split-image">{img_html}</div>
            </div>'''
        else:
            inner = f'''
            <h2 class="slide-title">{md_inline(slide['title'])}</h2>
            {text_part}{code_html}{table_html}'''
            if img_html and not text_part:
                inner += f'<div class="image-full">{img_html}</div>'

    notes_attr = f' data-notes="{escape_html(slide["speaker_notes"])}"' if slide.get('speaker_notes') else ''
    page_num = f'<div class="slide-footer"><span class="footer-left"></span><span class="footer-author">{escape_html(slide.get("_author", ""))}</span><span class="page-num">{index} / {total}</span></div>' if slide_type not in ('section', 'closing') else ''
    diagram_attr = f' data-diagram="{slide.get("diagram", "")}"' if slide.get('diagram') else ''

    return f'<section class="{classes}"{notes_attr}{diagram_attr}>{inner}{page_num}</section>'


def _render_image(img: dict, gallery: bool = False) -> str:
    if not img:
        return ''
    if img['type'] == 'placeholder':
        return f'''<div class="placeholder-img">
            <div class="placeholder-icon">&#128247;</div>
            <div class="placeholder-text">請插入：{escape_html(img['description'])}</div>
        </div>'''
    path = img['path']
    alt = img.get('alt', '')
    is_video = path.lower().endswith(('.mp4', '.webm', '.ogg', '.mov'))
    if is_video:
        media = (f'<video class="slide-img slide-video{" gallery-img" if gallery else ""}"'
                 f' controls playsinline preload="metadata">'
                 f'<source src="{escape_html(path)}"></video>')
    else:
        media = (f'<img src="{escape_html(path)}" alt="{escape_html(alt)}"'
                 f' class="slide-img{" gallery-img" if gallery else ""}">')
    if gallery:
        caption = f'<figcaption class="gallery-caption">{escape_html(alt)}</figcaption>' if alt else ''
        return f'<figure class="gallery-item">{media}{caption}</figure>'
    return media


def _render_gallery(images: list) -> str:
    """多張圖片排為縮圖格，每張可點擊放大"""
    items = ''.join(_render_image(img, gallery=True) for img in images)
    return f'<div class="image-gallery">{items}</div>'


# ============================================================
# SVG 圖表
# ============================================================

def render_diagram_svg(diagram_id: str) -> str:
    """根據 diagram ID 回傳對應的 SVG"""
    renderers = {
        'architecture-6': _svg_architecture_6,
        'refinery-protocol': _svg_refinery_protocol,
        'golden-journey-pipeline': _svg_golden_journey_pipeline,
    }
    fn = renderers.get(diagram_id)
    if fn:
        return fn()
    return f'<div class="placeholder-img"><div class="placeholder-text">未定義的圖表：{diagram_id}</div></div>'


def _svg_architecture_6() -> str:
    """Alex Diary 6 模組環狀架構圖"""
    return '''<svg viewBox="0 0 800 500" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-height:65vh">
  <style>
    .box{rx:12;fill:var(--bg-card,#f7f7f7);stroke:var(--border,#e5e5e5);stroke-width:1.5}
    .box-center{rx:16;fill:var(--accent,#007acc);fill-opacity:.12;stroke:var(--accent,#007acc);stroke-width:2}
    .label{{font-family:var(--font,sans-serif);fill:var(--text-primary,#1a1a1a);font-size:16px;text-anchor:middle;dominant-baseline:central}}
    .label-sm{{font-family:var(--font,sans-serif);fill:var(--text-secondary,#555);font-size:12px;text-anchor:middle;dominant-baseline:central}}
    .label-center{{font-family:var(--font,sans-serif);fill:var(--accent,#007acc);font-size:20px;font-weight:700;text-anchor:middle;dominant-baseline:central}}
    .conn{stroke:var(--accent,#007acc);stroke-width:1.2;stroke-dasharray:6 3;fill:none;opacity:.4}
    .icon{font-size:22px;text-anchor:middle;dominant-baseline:central}
  </style>
  <!-- 中心 -->
  <rect class="box-center" x="310" y="195" width="180" height="110"/>
  <text class="icon" x="400" y="230">🤖</text>
  <text class="label-center" x="400" y="258">AI 核心引擎</text>
  <text class="label-sm" x="400" y="278">Claude + 20 Skills</text>
  <!-- 連線 -->
  <line class="conn" x1="400" y1="195" x2="400" y2="78"/>
  <line class="conn" x1="310" y1="220" x2="168" y2="138"/>
  <line class="conn" x1="310" y1="275" x2="168" y2="370"/>
  <line class="conn" x1="490" y1="220" x2="632" y2="138"/>
  <line class="conn" x1="490" y1="275" x2="632" y2="370"/>
  <line class="conn" x1="400" y1="305" x2="400" y2="420"/>
  <!-- 任務管理 -->
  <rect class="box" x="310" y="28" width="180" height="60"/>
  <text class="icon" x="345" y="58">📋</text>
  <text class="label" x="420" y="50">任務管理</text>
  <text class="label-sm" x="420" y="70">待辦池 · 日計畫 · 多輪追蹤</text>
  <!-- 工時追蹤 -->
  <rect class="box" x="58" y="88" width="180" height="60"/>
  <text class="icon" x="93" y="118">⏱️</text>
  <text class="label" x="170" y="110">工時追蹤</text>
  <text class="label-sm" x="170" y="130">自動計時 · 15min 量化</text>
  <!-- 人脈 CRM -->
  <rect class="box" x="562" y="88" width="180" height="60"/>
  <text class="icon" x="597" y="118">👥</text>
  <text class="label" x="672" y="110">人脈 CRM</text>
  <text class="label-sm" x="672" y="130">72 人 · 互動紀錄 · 自動提醒</text>
  <!-- 知識庫 -->
  <rect class="box" x="562" y="340" width="180" height="60"/>
  <text class="icon" x="597" y="370">📚</text>
  <text class="label" x="672" y="362">知識庫</text>
  <text class="label-sm" x="672" y="382">32 條 Lessons · SOP</text>
  <!-- 財務追蹤 -->
  <rect class="box" x="58" y="340" width="180" height="60"/>
  <text class="icon" x="93" y="370">💰</text>
  <text class="label" x="170" y="362">財務追蹤</text>
  <text class="label-sm" x="170" y="382">收支分析 · 訂閱管理</text>
  <!-- 每日回顧 -->
  <rect class="box" x="310" y="410" width="180" height="60"/>
  <text class="icon" x="345" y="440">🔄</text>
  <text class="label" x="420" y="432">每日回顧</text>
  <text class="label-sm" x="420" y="452">結算 · 反思 · 明日規劃</text>
</svg>'''


def _svg_refinery_protocol() -> str:
    """Refinery Protocol 漏斗流程圖"""
    return '''<svg viewBox="0 0 800 480" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-height:65vh">
  <style>
    .rbox{rx:10;stroke-width:1.5}
    .rlabel{{font-family:var(--font,sans-serif);fill:var(--text-primary,#1a1a1a);font-size:16px;text-anchor:middle;dominant-baseline:central}}
    .rlabel-sm{{font-family:var(--font,sans-serif);fill:var(--text-secondary,#555);font-size:13px;text-anchor:middle;dominant-baseline:central}}
    .rlabel-b{{font-family:var(--font,sans-serif);fill:#fff;font-size:18px;font-weight:600;text-anchor:middle;dominant-baseline:central}}
    .arrow{fill:var(--accent,#007acc);opacity:.5}
    .input-box{fill:var(--bg-card,#f7f7f7);stroke:var(--border,#e5e5e5)}
    .process-box{fill:var(--accent,#007acc);stroke:none}
    .output-box{fill:var(--bg-card,#f7f7f7);stroke:var(--accent,#007acc);stroke-width:2}
    .icon{font-size:20px;text-anchor:middle;dominant-baseline:central}
  </style>
  <!-- 輸入源 (top) -->
  <rect class="rbox input-box" x="40" y="20" width="140" height="50"/><text class="icon" x="70" y="45">💬</text><text class="rlabel" x="125" y="45">LINE 對話</text>
  <rect class="rbox input-box" x="220" y="20" width="140" height="50"/><text class="icon" x="250" y="45">📸</text><text class="rlabel" x="305" y="45">現場照片</text>
  <rect class="rbox input-box" x="440" y="20" width="140" height="50"/><text class="icon" x="470" y="45">🎙️</text><text class="rlabel" x="525" y="45">會議紀錄</text>
  <rect class="rbox input-box" x="620" y="20" width="140" height="50"/><text class="icon" x="650" y="45">📧</text><text class="rlabel" x="705" y="45">Email / 名片</text>
  <!-- 箭頭往下 -->
  <polygon class="arrow" points="400,90 380,80 420,80"/>
  <rect fill="var(--accent,#007acc)" opacity=".15" x="200" y="88" width="400" height="2"/>
  <!-- Step 1: 全量捕捉 -->
  <rect class="rbox process-box" x="250" y="110" width="300" height="52" rx="26"/>
  <text class="rlabel-b" x="400" y="130">① 全量捕捉</text>
  <text class="rlabel-sm" x="400" y="148" style="fill:#fff">所有原始資料完整存入 Journal</text>
  <polygon class="arrow" points="400,180 388,168 412,168"/>
  <!-- Step 2: 資產分流 -->
  <rect class="rbox process-box" x="250" y="190" width="300" height="52" rx="26"/>
  <text class="rlabel-b" x="400" y="210">② 資產分流</text>
  <text class="rlabel-sm" x="400" y="228" style="fill:#fff">AI 判斷歸類 → 工時 / 任務 / 人脈 / 財務</text>
  <polygon class="arrow" points="400,260 388,248 412,248"/>
  <!-- Step 3: 智慧精煉 -->
  <rect class="rbox process-box" x="250" y="270" width="300" height="52" rx="26"/>
  <text class="rlabel-b" x="400" y="290">③ 智慧精煉</text>
  <text class="rlabel-sm" x="400" y="308" style="fill:#fff">提煉經驗法則 · 標記行動項目</text>
  <polygon class="arrow" points="400,340 388,328 412,328"/>
  <!-- Step 4: 排程整合 -->
  <rect class="rbox process-box" x="250" y="350" width="300" height="52" rx="26"/>
  <text class="rlabel-b" x="400" y="370">④ 排程整合</text>
  <text class="rlabel-sm" x="400" y="388" style="fill:#fff">對照 Google Calendar · 檢查衝突</text>
  <!-- 輸出 (bottom) -->
  <rect fill="var(--accent,#007acc)" opacity=".15" x="100" y="418" width="600" height="2"/>
  <rect class="rbox output-box" x="40" y="428" width="130" height="42"/><text class="icon" x="65" y="449">⏱️</text><text class="rlabel" x="118" y="449">工時</text>
  <rect class="rbox output-box" x="200" y="428" width="130" height="42"/><text class="icon" x="225" y="449">📋</text><text class="rlabel" x="278" y="449">任務</text>
  <rect class="rbox output-box" x="360" y="428" width="130" height="42"/><text class="icon" x="385" y="449">👥</text><text class="rlabel" x="438" y="449">人脈</text>
  <rect class="rbox output-box" x="520" y="428" width="130" height="42"/><text class="icon" x="545" y="449">📚</text><text class="rlabel" x="598" y="449">知識</text>
  <rect class="rbox output-box" x="680" y="428" width="90" height="42"/><text class="icon" x="705" y="449">💰</text><text class="rlabel" x="738" y="449">財務</text>
</svg>'''


def _svg_golden_journey_pipeline() -> str:
    """Golden Journey 照片 → 知識管線圖"""
    return '''<svg viewBox="0 0 820 360" xmlns="http://www.w3.org/2000/svg" style="width:100%;max-height:60vh">
  <style>
    .gbox{rx:12;fill:var(--bg-card,#f7f7f7);stroke:var(--border,#e5e5e5);stroke-width:1.5}
    .gbox-accent{rx:12;fill:var(--accent,#007acc);fill-opacity:.1;stroke:var(--accent,#007acc);stroke-width:2}
    .glabel{{font-family:var(--font,sans-serif);fill:var(--text-primary,#1a1a1a);font-size:16px;font-weight:600;text-anchor:middle;dominant-baseline:central}}
    .glabel-sm{{font-family:var(--font,sans-serif);fill:var(--text-secondary,#555);font-size:13px;text-anchor:middle;dominant-baseline:central}}
    .garrow{{fill:none;stroke:var(--accent,#007acc);stroke-width:2;marker-end:url(#arrowG)}}
    .gnum{{font-family:var(--font,sans-serif);fill:var(--accent,#007acc);font-size:13px;font-weight:700;text-anchor:middle}}
    .icon{font-size:26px;text-anchor:middle;dominant-baseline:central}
  </style>
  <defs><marker id="arrowG" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="var(--accent,#007acc)"/></marker></defs>
  <!-- Step 1: 照片 -->
  <rect class="gbox" x="20" y="60" width="120" height="100"/>
  <text class="icon" x="80" y="95">📷</text>
  <text class="glabel" x="80" y="125">原始照片</text>
  <text class="glabel-sm" x="80" y="142">2800+ 張</text>
  <line class="garrow" x1="140" y1="110" x2="170" y2="110"/>
  <!-- Step 2: EXIF -->
  <rect class="gbox" x="175" y="60" width="120" height="100"/>
  <text class="icon" x="235" y="95">📍</text>
  <text class="glabel" x="235" y="125">GPS 提取</text>
  <text class="glabel-sm" x="235" y="142">EXIF 掃描</text>
  <line class="garrow" x1="295" y1="110" x2="325" y2="110"/>
  <!-- Step 3: 聚類 -->
  <rect class="gbox" x="330" y="60" width="120" height="100"/>
  <text class="icon" x="390" y="95">🗺️</text>
  <text class="glabel" x="390" y="125">地理聚類</text>
  <text class="glabel-sm" x="390" y="142">500m 分組 → 45 點</text>
  <line class="garrow" x1="450" y1="110" x2="480" y2="110"/>
  <!-- Step 4: 標註 -->
  <rect class="gbox" x="485" y="60" width="120" height="100"/>
  <text class="icon" x="545" y="95">✍️</text>
  <text class="glabel" x="545" y="125">互動標註</text>
  <text class="glabel-sm" x="545" y="142">選片 · 命名 · 筆記</text>
  <line class="garrow" x1="605" y1="110" x2="635" y2="110"/>
  <!-- Step 5: AI 精修 -->
  <rect class="gbox-accent" x="640" y="60" width="140" height="100"/>
  <text class="icon" x="710" y="95">🤖</text>
  <text class="glabel" x="710" y="125">AI 精修</text>
  <text class="glabel-sm" x="710" y="142">三層深度筆記</text>
  <!-- 輸出行 -->
  <line class="garrow" x1="710" y1="160" x2="710" y2="210" style="fill:none"/>
  <rect class="gbox" x="80" y="250" width="160" height="80"/>
  <text class="icon" x="160" y="278">🌐</text>
  <text class="glabel" x="160" y="302">互動網頁報告</text>
  <text class="glabel-sm" x="160" y="318">Scrollytelling · 3 主題</text>
  <rect class="gbox" x="310" y="250" width="160" height="80"/>
  <text class="icon" x="390" y="278">📽️</text>
  <text class="glabel" x="390" y="302">PowerPoint</text>
  <text class="glabel-sm" x="390" y="318">506MB · 精選照片</text>
  <rect class="gbox" x="540" y="250" width="200" height="80"/>
  <text class="icon" x="640" y="278">📦</text>
  <text class="glabel" x="640" y="302">離線策展套件</text>
  <text class="glabel-sm" x="640" y="318">5.1GB · 完整攜帶</text>
  <!-- 連線到輸出（直角折線） -->
  <line class="garrow" x1="710" y1="210" x2="710" y2="230" style="fill:none"/>
  <line class="garrow" x1="160" y1="230" x2="710" y2="230" style="fill:none;marker-end:none"/>
  <line class="garrow" x1="640" y1="230" x2="640" y2="250" style="fill:none"/>
  <line class="garrow" x1="390" y1="230" x2="390" y2="250" style="fill:none"/>
  <line class="garrow" x1="160" y1="230" x2="160" y2="250" style="fill:none"/>
  <!-- 步驟編號 -->
  <text class="gnum" x="80" y="50">Step 1</text>
  <text class="gnum" x="235" y="50">Step 2</text>
  <text class="gnum" x="390" y="50">Step 3</text>
  <text class="gnum" x="545" y="50">Step 4</text>
  <text class="gnum" x="710" y="50">Step 5</text>
</svg>'''


# ============================================================
# 章節心智圖 / 路線圖
# ============================================================

def build_section_map(slides_data: list, duration: int) -> list:
    """識別各章節及其包含的投影片數量，估算講述時間"""
    sections = []
    current_idx = -1
    for i, slide in enumerate(slides_data):
        if slide.get('type') == 'section':
            sections.append({
                'slide_idx': i,
                'title': slide.get('title', f'章節 {len(sections)+1}'),
                'slide_count': 0,
                'section_num': len(sections) + 1, # 從 1 開始編號
                'minutes': 0,
            })
            current_idx = len(sections) - 1
        elif current_idx >= 0:
            sections[current_idx]['slide_count'] += 1
    if not sections:
        return sections
    secs_per_slide = (duration * 60) / max(len(slides_data), 1)
    for sec in sections:
        sec['minutes'] = max(1, round(sec['slide_count'] * secs_per_slide / 60))
    return sections


def _strip_part_prefix(title: str) -> str:
    """去除 'Part X：' 前綴，保留核心標題"""
    stripped = re.sub(r'^Part\s*\d+[：:]\s*', '', title).strip()
    return stripped if stripped else title


def _auto_wrap_title(title: str, max_chars: int) -> tuple:
    """若 title 超過 max_chars，依語意斷行，回傳 (line1, line2)；
    否則回傳 (title, None)。斷點優先找中點附近的空格或標點。"""
    if len(title) <= max_chars:
        return title, None
    mid = len(title) // 2
    BREAKS = set(' \u3001\uff0c\uff1a\uff1b/-\u2014\u30fb')  # 、，：；/—・
    for delta in range(len(title)):
        for pos in sorted({mid - delta, mid + delta}):
            if 1 <= pos < len(title):
                if title[pos - 1] in BREAKS:
                    return title[:pos].rstrip(), title[pos:].lstrip()
                if title[pos] in BREAKS:
                    return title[:pos + 1].rstrip(), title[pos + 1:].lstrip()
    # 找不到標點 → 正中切開
    return title[:mid], title[mid:]


def render_chapter_flow_svg(sections: list, highlight: int = -1, compact: bool = False) -> str:
    """產生水平章節流程 SVG（全尺寸或緊湊版）
    full  : 節點 160×110，字 16px，附時間標籤，viewBox 900×320
    compact: 節點 150×82，字 15px，無時間標籤，viewBox 900×140
    """
    if not sections:
        return ''
    n = len(sections)

    # ── 尺寸定義 ──────────────────────────────────
    if compact:
        W, H        = 2600, 700   # 加大寬度，讓長標題不碰邊
        GAP         = 30
        nh          = 420
        fs_num      = 28
        fs_title    = 58          # 與 TOC 版字體一致
        show_time   = False
        max_height  = 0           # compact 模式不限制高度，由 CSS flex 控制
    else:
        W, H        = 2600, 700
        GAP         = 30
        nh          = 420
        fs_num      = 28
        fs_title    = 58
        show_time   = True
        max_height  = 850

    # 節點寬：提供極致的呼吸空間
    nw = min(700, max(350, (W - 120 - (n - 1) * GAP) // n))

    total_nw = n * nw + (n - 1) * GAP
    x0 = (W - total_nw) / 2
    cy = H * 0.44 if show_time else H / 2

    mh_style = f'max-height:{max_height}px' if max_height else ''
    parts = [
        f'<svg viewBox="0 0 {W} {H}" xmlns="http://www.w3.org/2000/svg"'
        f' style="width:100%;{mh_style}">'
    ]

    # ── 連線與箭頭 ──────────────────────────────────
    arrow_y = cy
    for i in range(n - 1):
        lx = x0 + i * (nw + GAP) + nw
        rx = x0 + (i + 1) * (nw + GAP)
        mx = (lx + rx) / 2
        parts.append(
            f'<line x1="{lx:.0f}" y1="{arrow_y:.0f}" x2="{rx-9:.0f}" y2="{arrow_y:.0f}"'
            f' stroke="var(--muted,#555)" stroke-width="2"/>'
        )
        parts.append(
            f'<polygon points="{rx:.0f},{arrow_y:.0f} {rx-10:.0f},{arrow_y-5:.0f} {rx-10:.0f},{arrow_y+5:.0f}"'
            f' fill="var(--muted,#555)"/>'
        )

    # ── 節點 ──────────────────────────────────────
    for i, sec in enumerate(sections):
        cx  = x0 + i * (nw + GAP) + nw / 2
        nx  = cx - nw / 2
        ny  = cy - nh / 2
        # 是否高亮？（匹配 1-indexed 的 section_num）
        is_active = (sec['section_num'] == highlight)
        dim    = (highlight >= 0 and not is_active)

        if is_active:
            fill, stroke_c, sw = 'var(--accent,#D4A853)', 'var(--accent,#D4A853)', '3'
            tc, op, num_c = '#ffffff', '1', '#ffffff'
        elif dim:
            fill, stroke_c, sw = 'var(--bg-card,#1a1a1a)', 'var(--border,#333)', '1.5'
            tc, op, num_c = 'var(--muted,#666)', '0.38', 'var(--muted,#666)'
        else:
            fill, stroke_c, sw = 'var(--bg-card,#1a1a1a)', 'var(--accent,#D4A853)', '2'
            tc, op, num_c = 'var(--text-primary,#fff)', '1', 'var(--muted,#888)'

        # 當前章節：上方三角指示器
        if is_active:
            parts.append(
                f'<polygon points="{cx:.0f},{ny-16:.0f} {cx-8:.0f},{ny-3:.0f} {cx+8:.0f},{ny-3:.0f}"'
                f' fill="var(--accent,#D4A853)"/>'
            )

        # clipPath：防止文字溢出方塊邊界
        clip_id = f'clip_sec_{i}'
        parts.append(
            f'<defs><clipPath id="{clip_id}">'
            f'<rect x="{nx:.1f}" y="{ny:.1f}" width="{nw}" height="{nh}" rx="10"/>'
            f'</clipPath></defs>'
        )

        # 節點方塊
        parts.append(
            f'<rect x="{nx:.1f}" y="{ny:.1f}" width="{nw}" height="{nh}"'
            f' rx="10" fill="{fill}" stroke="{stroke_c}" stroke-width="{sw}" opacity="{op}"/>'
        )

        # Part X 小標
        parts.append(
            f'<text x="{cx:.1f}" y="{ny + nh*0.28:.1f}" text-anchor="middle"'
            f' font-family="var(--font,sans-serif)" font-size="{fs_num}"'
            f' fill="{num_c}" opacity="{op}" clip-path="url(#{clip_id})">Part {i+1}</text>'
        )

        # 主標題：超出單行寬度時依語意自動斷行
        display_title = _strip_part_prefix(sec['title'])
        max_chars_line = max(4, int(nw / (fs_title * 0.7)))
        line1, line2 = _auto_wrap_title(display_title, max_chars_line)
        lh = int(fs_title * 1.2)   # 行距
        if line2 is not None:
            parts.append(
                f'<text x="{cx:.1f}" y="{ny + nh*0.52:.1f}" text-anchor="middle"'
                f' font-family="var(--font,sans-serif)" font-size="{fs_title}" font-weight="700"'
                f' fill="{tc}" opacity="{op}" clip-path="url(#{clip_id})">'
                f'<tspan x="{cx:.1f}">{escape_html(line1)}</tspan>'
                f'<tspan x="{cx:.1f}" dy="{lh}">{escape_html(line2)}</tspan>'
                f'</text>'
            )
        else:
            parts.append(
                f'<text x="{cx:.1f}" y="{ny + nh*0.65:.1f}" text-anchor="middle"'
                f' font-family="var(--font,sans-serif)" font-size="{fs_title}" font-weight="700"'
                f' fill="{tc}" opacity="{op}" clip-path="url(#{clip_id})">{escape_html(display_title)}</text>'
            )

        # 時間標籤（僅全尺寸，加大間距避免與方塊疊合）
        if show_time and sec.get('minutes'):
            parts.append(
                f'<text x="{cx:.1f}" y="{ny + nh + 75:.1f}" text-anchor="middle"'
                f' font-family="var(--font,sans-serif)" font-size="{fs_num + 1}"'
                f' fill="var(--muted,#888)" opacity="{op}">約 {sec["minutes"]} 分鐘</text>'
            )

    parts.append('</svg>')
    return '\n'.join(parts)


def generate_css() -> str:
    return '''
    /* CSS 變數由 JS 動態注入 */

    * {{ margin: 0; padding: 0; box-sizing: border-box; }}

    html, body {{
        font-family: var(--font);
        background: var(--bg);
        color: var(--text-primary);
        overflow: hidden;
        height: 100vh;
        width: 100vw;
    }}

    /* === 投影片容器 === */
    .deck {{
        height: 100vh;
        width: 100vw;
        position: relative;
    }}

    .slide {{
        position: absolute;
        top: 0; left: 0;
        width: 100vw;
        height: 100vh;
        padding: 5vh 6vw;
        display: flex;
        flex-direction: column;
        justify-content: center;
        overflow: hidden;
        opacity: 0;
        pointer-events: none;
        transition: opacity 0.4s ease;
    }}
    .slide.active {{
        opacity: 1;
        pointer-events: auto;
    }}

    /* === 標題 === */
    .slide-title {{
        font-size: clamp(2.2rem, 4.4vw, 3.8rem);
        font-weight: 700;
        margin-bottom: 2rem;
        line-height: 1.2;
        position: relative;
        padding-bottom: 0.8rem;
        color: var(--accent);
    }}
    .slide-title::after {{
        content: '';
        position: absolute;
        bottom: 0; left: 0;
        width: 3rem;
        height: 4px;
        background: var(--accent);
        border-radius: 2px;
    }}

    /* === 卡片 === */
    .card {{
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 16px;
        padding: 2rem 2.5rem;
        box-shadow: 0 4px 20px var(--shadow);
        transition: background 0.2s;
    }}
    .card:hover {{
        background: var(--bg-card-hover);
    }}

    /* === 條列 === */
    .bullet-list {{
        list-style: none;
        padding: 0;
    }}
    .bullet-list > li {{
        font-size: clamp(1.4rem, 2.6vw, 2.22rem);
        line-height: 1.5;
        color: var(--text-secondary);
        padding: 0.8rem 0;
        padding-left: 1.8rem;
        position: relative;
    }}
    .bullet-list > li::before {{
        content: '';
        position: absolute;
        left: 0;
        top: 1.1rem;
        width: 0.6rem;
        height: 3px;
        background: var(--accent);
        border-radius: 2px;
    }}
    .bullet-list > li > ul {{
        list-style: none;
        padding-left: 1rem;
        margin-top: 0.3rem;
    }}
    .bullet-list > li > ul > li {{
        font-size: 0.88em;
        color: var(--muted);
        padding: 0.2rem 0;
    }}
    .bullet-list > li > ul > li::before {{
        content: '\u2013';
        margin-right: 0.5rem;
        color: var(--accent);
    }}

    .body-text {{
        font-size: clamp(1.3rem, 2.5vw, 2rem);
        line-height: 1.6;
        margin-bottom: 1rem;
        color: var(--text-secondary);
    }}

    /* === 圖文分割 === */
    .split-layout {{
        display: grid;
        grid-template-columns: 1fr 1.3fr;
        gap: 2rem;
        flex: 1;
        min-height: 0;
        align-items: start;
    }}
    .split-text {{
        display: flex;
        flex-direction: column;
        justify-content: center;
    }}
    .split-image {{
        display: flex;
        align-items: center;
        justify-content: center;
        min-height: 0;
    }}

    /* === 多卡片佈局 === */
    .multi-card-layout {{
        display: flex;
        flex-direction: column;
        flex: 1;
        min-height: 0;
        gap: 0.8rem;
    }}
    .columns-row {{
        display: flex;
        flex-direction: row;
        gap: 0.8rem;
        flex: 1;
        min-height: 0;
    }}
    .card-top {{
        width: 100%;
        flex-shrink: 0;
        font-size: clamp(1.1rem, 2vw, 1.6rem);
    }}
    .card-column {{
        flex: 1;
        min-width: 0;
        overflow: hidden;
        padding: 1rem 1.2rem;
        font-size: clamp(1.0rem, 1.8vw, 1.5rem);
    }}
    /* 卡片內文字全部繼承 card 的 font-size，讓 autoFitCards 統一控制 */
    .card-column .bullet-list > li,
    .card-top .bullet-list > li {{
        font-size: 1em;
        padding: 0.25rem 0 0.25rem 1.4rem;
    }}
    .card-column .bullet-list > li::before,
    .card-top .bullet-list > li::before {{
        top: 0.65em;
    }}
    .card-column .body-text,
    .card-top .body-text {{
        font-size: 1em;
        margin-bottom: 0.4rem;
        line-height: 1.5;
    }}
    .card-heading {{
        font-size: 1.05em;
        font-weight: 700;
        color: var(--accent);
        margin-bottom: 0.5rem;
        padding-bottom: 0.3rem;
        border-bottom: 2px solid var(--accent);
        opacity: 0.9;
    }}

    /* === 圖表容器 === */
    .diagram-container {{
        display: flex;
        align-items: center;
        justify-content: center;
        flex: 1;
        min-height: 0;
        padding: 1rem 0;
    }}
    .diagram-container svg {{
        filter: drop-shadow(0 2px 8px var(--shadow));
    }}

    /* === 佔位符圖片 === */
    .placeholder-img {{
        background: var(--bg-card);
        border: 2px dashed var(--border);
        border-radius: 16px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 3rem;
        min-height: 20vh;
        width: 100%;
    }}
    .placeholder-icon {{
        font-size: 3rem;
        margin-bottom: 1rem;
        opacity: 0.5;
    }}
    .placeholder-text {{
        font-size: 1.1rem;
        color: var(--muted);
        text-align: center;
    }}
    .slide-img {{
        max-width: 100%;
        max-height: 60vh;
        border-radius: 12px;
        object-fit: contain;
    }}
    .slide-video {{
        background: #000;
        width: 100%;
    }}

    /* === 圖片 / 影片並排 Gallery === */
    .image-gallery {{
        display: flex;
        flex-wrap: wrap;
        gap: 1rem;
        justify-content: center;
        align-items: flex-start;
        width: 100%;
    }}
    .gallery-item {{
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0.4rem;
        flex: 1 1 200px;
        min-width: 0;
    }}
    .gallery-img {{
        max-height: 40vh;
        max-width: 100%;
        border-radius: 8px;
        object-fit: contain;
        cursor: pointer;
        transition: opacity 0.2s;
    }}
    .gallery-img:hover {{ opacity: 0.88; }}
    .gallery-caption {{
        font-size: 0.72em;
        color: var(--text-secondary);
        text-align: center;
        opacity: 0.8;
    }}

    /* === Section 投影片 === */
    .slide--section {{
        justify-content: center;
        align-items: center;
        text-align: center;
    }}
    .section-content {{
        flex-shrink: 1;
        margin-bottom: 1rem;
    }}
    .section-accent-line {{
        width: 4rem;
        height: 4px;
        background: var(--accent);
        margin: 0 auto 2rem;
        border-radius: 2px;
    }}
    .section-part-num {{
        font-size: 0.5em;
        color: var(--accent);
        font-weight: 400;
        letter-spacing: 0.05em;
    }}
    .section-title {{
        font-size: clamp(3.5rem, 8vw, 6.5rem);
        font-weight: 700;
        line-height: 1.2;
        letter-spacing: 0.05em;
        white-space: nowrap;
        color: var(--accent);
    }}
    .section-subtitle {{
        font-size: clamp(1.6rem, 3.2vw, 2.4rem);
        color: var(--text-secondary);
        margin-top: 1.2rem;
    }}
    .section-with-flow {{
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 0;
        height: 100%;
        width: 100%;
    }}
    .chapter-flow-compact {{
        width: 100%;
        max-width: 1650px;
        flex: 1;
        min-height: 0;
        margin: 0 auto;
    }}
    .chapter-flow-compact svg {{
        width: 100%;
        height: 100%;
    }}

    /* === 目錄投影片（TOC） === */
    .slide--toc {{
        justify-content: center;
        align-items: center;
        flex-direction: column;
        gap: 1.8rem;
    }}
    .toc-title {{
        font-size: clamp(2rem, 4.2vw, 3.8rem);
        font-weight: 700;
        color: var(--text-primary);
        letter-spacing: 0.04em;
    }}
    .chapter-flow-toc {{
        width: 100%;
        max-width: 1650px;
    }}

    /* === Demo 投影片 === */
    .demo-badge {{
        display: inline-block;
        background: var(--accent);
        color: var(--bg);
        font-size: 1.2rem;
        font-weight: 700;
        padding: 0.5rem 1.5rem;
        border-radius: 2rem;
        margin-bottom: 1.5rem;
    }}
    .demo-steps {{
        list-style: none;
        padding: 0;
    }}
    .demo-steps li {{
        font-size: clamp(1.35rem, 2.4vw, 2rem);
        padding: 1rem 0;
        color: var(--text-secondary);
        border-bottom: 1px solid var(--border);
    }}
    .demo-steps li:last-child {{ border-bottom: none; }}
    .step-num {{
        display: inline-block;
        background: var(--accent-light);
        color: var(--accent);
        font-weight: 700;
        padding: 0.2rem 0.6rem;
        border-radius: 6px;
        margin-right: 0.5rem;
        font-size: 0.85em;
    }}

    /* === 對比/表格 === */
    .card-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: clamp(1.2rem, 2.2vw, 1.8rem);
    }}
    .card-table th {{
        text-align: center;
        padding: 1rem;
        border-bottom: 3px solid var(--accent);
        color: var(--text-primary);
        font-weight: 700;
    }}
    .card-table td {{
        text-align: center;
        padding: 1rem 1.2rem;
        border-bottom: 1px solid var(--border);
        color: var(--text-secondary);
    }}
    .card-table tr:last-child td {{ border-bottom: none; }}
    /* 最後一欄以外不折行，讓短欄縮到剛好，長欄承擔換行 */
    .card-table td:not(:last-child),
    .card-table th:not(:last-child) {{ white-space: nowrap; }}

    /* === 引言 === */
    .slide--quote {{ align-items: center; text-align: center; }}
    .quote-content {{ max-width: 70%; }}
    .quote-mark {{
        font-size: 6rem;
        color: var(--accent);
        line-height: 1;
        opacity: 0.3;
    }}
    .quote-text {{
        font-size: clamp(2rem, 4.5vw, 3.5rem);
        line-height: 1.5;
        color: var(--text-secondary);
        font-style: italic;
    }}

    /* === Closing === */
    .slide--closing {{ align-items: center; text-align: center; }}
    .closing-content {{ max-width: 70%; }}
    .closing-title {{
        font-size: clamp(2.5rem, 6vw, 5.2rem);
        font-weight: 700;
        margin-bottom: 2.5rem;
    }}
    .closing-points {{
        list-style: none;
        text-align: left;
        display: inline-block;
    }}
    .closing-points li {{
        font-size: clamp(1.4rem, 2.8vw, 2rem);
        color: var(--text-secondary);
        padding: 0.5rem 0;
        padding-left: 1.5rem;
        position: relative;
    }}
    .closing-points li::before {{
        content: '';
        position: absolute;
        left: 0; top: 1rem;
        width: 0.6rem; height: 3px;
        background: var(--accent);
        border-radius: 2px;
    }}

    /* === 程式碼 === */
    .code-card pre {{
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 1.15rem;
        line-height: 1.5;
        overflow-x: auto;
        color: var(--text-secondary);
    }}

    /* === 投影片子標題（### heading） === */
    .slide-subheading {{
        font-size: clamp(1rem, 1.8vw, 1.5rem);
        font-weight: 600;
        color: var(--accent);
        margin-top: 1.1rem;
        margin-bottom: 0.3rem;
        padding-left: 0.6rem;
        border-left: 3px solid var(--accent);
        line-height: 1.3;
    }}

    /* 行內 code：小框框樣式 */
    :not(pre) > code {{
        font-family: 'Consolas', 'Courier New', monospace;
        font-size: 0.82em;
        background: var(--accent-light);
        border: 1px solid var(--border);
        border-radius: 4px;
        padding: 0.1em 0.45em;
        color: var(--accent);
        white-space: nowrap;
    }}

    /* 全形括號補充說明縮小 */
    .aside-text {{
        font-size: 0.78em;
        color: var(--text-secondary);
        opacity: 0.85;
    }}

    /* code block 內的超連結 */
    .code-link {{
        color: var(--accent);
        text-decoration: underline;
        word-break: break-all;
    }}
    .code-link:hover {{ opacity: 0.75; }}

    /* === 頁尾 === */
    .slide-footer {{
        position: absolute;
        bottom: 2vh;
        left: 4vw;
        right: 4vw;
        display: grid;
        grid-template-columns: 1fr auto 1fr;
        align-items: center;
        font-size: 0.85rem;
        color: var(--muted);
    }}
    .footer-left {{ }}
    .footer-author {{ text-align: center; }}
    .page-num {{ text-align: right; }}

    /* === 進度條 === */
    .progress-bar {{
        position: fixed;
        top: 0; left: 0;
        height: 3px;
        background: var(--accent);
        transition: width 0.3s ease;
        z-index: 100;
    }}

    /* === 講者備註面板 === */
    .notes-panel {{
        position: fixed;
        bottom: 0; left: 0; right: 0;
        background: rgba(0,0,0,0.92);
        color: #eee;
        padding: 1.5rem 3rem;
        font-size: 1.1rem;
        line-height: 1.6;
        transform: translateY(100%);
        transition: transform 0.3s ease;
        z-index: 200;
        max-height: 30vh;
        overflow-y: auto;
    }}
    .notes-panel.visible {{
        transform: translateY(0);
    }}
    .notes-label {{
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: var(--accent);
        margin-bottom: 0.5rem;
    }}

    /* === 封面投影片 === */
    .slide--cover {{
        justify-content: center;
        align-items: center;
        text-align: center;
    }}
    .cover-title {{
        font-size: clamp(3.2rem, 7.5vw, 6.5rem);
        font-weight: 700;
        line-height: 1.1;
        margin-bottom: 1.5rem;
    }}
    .cover-subtitle {{
        font-size: clamp(1.8rem, 3.8vw, 3rem);
        color: var(--text-secondary);
        margin-bottom: 3rem;
    }}
    .cover-meta {{
        font-size: 1rem;
        color: var(--muted);
    }}

    /* === 主題切換器 === */
    .theme-switcher {
        position: fixed;
        top: 1.2rem;
        right: 1.5rem;
        z-index: 300;
        display: flex;
        gap: 0.5rem;
        align-items: center;
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: 2rem;
        padding: 0.35rem 0.6rem;
        box-shadow: 0 2px 12px var(--shadow);
        opacity: 0.2;
        transition: opacity 0.3s;
    }
    .theme-switcher:hover { opacity: 1; }
    .theme-dot {
        width: 1.1rem;
        height: 1.1rem;
        border-radius: 50%;
        border: 2px solid transparent;
        cursor: pointer;
        transition: border-color 0.2s, transform 0.2s;
    }
    .theme-dot:hover { transform: scale(1.25); }
    .theme-dot.active { border-color: var(--text-primary); transform: scale(1.2); }

    /* === 編輯模式 === */
    .edit-toolbar {
        display: none;
        position: fixed;
        bottom: 1.5rem; left: 50%;
        transform: translateX(-50%);
        gap: 0.5rem;
        align-items: center;
        background: var(--bg-card);
        border: 2px solid var(--accent);
        border-radius: 1rem;
        padding: 0.4rem 0.8rem;
        z-index: 400;
        box-shadow: 0 4px 20px var(--shadow);
    }
    .edit-toolbar button {
        background: none;
        border: 1px solid var(--border);
        color: var(--text-primary);
        padding: 0.4rem 0.8rem;
        border-radius: 0.5rem;
        cursor: pointer;
        font-family: var(--font);
        font-size: 0.85rem;
        transition: background 0.2s;
    }
    .edit-toolbar button:hover { background: var(--accent-light); }
    .edit-toolbar .btn-save {
        background: var(--accent);
        color: #fff;
        border-color: var(--accent);
        font-weight: 600;
    }
    .edit-toolbar .btn-save:hover {{ opacity: 0.85; }}
    .edit-toolbar .btn-snapshot {{
        background: none;
        border-color: var(--muted);
        color: var(--text-secondary);
    }}
    .edit-toolbar .btn-snapshot:hover {{ background: var(--accent-light); color: var(--text-primary); }}
    .snapshot-modal {{
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.6);
        z-index: 500;
        align-items: center;
        justify-content: center;
    }}
    .snapshot-modal.visible {{ display: flex; }}
    .edit-btn {
        position: fixed;
        bottom: 1.5rem; left: 1.5rem;
        background: var(--bg-card);
        border: 1px solid var(--border);
        color: var(--text-secondary);
        width: 2.5rem; height: 2.5rem;
        border-radius: 50%;
        cursor: pointer;
        font-size: 1.1rem;
        z-index: 300;
        box-shadow: 0 2px 8px var(--shadow);
        display: flex; align-items: center; justify-content: center;
        transition: background 0.2s;
    }
    .edit-btn:hover { background: var(--accent-light); }
    body.edit-mode .edit-btn { background: var(--accent); color: #fff; border-color: var(--accent); }
    body.edit-mode [contenteditable="true"] {
        outline: 2px dashed var(--accent);
        outline-offset: 4px;
        border-radius: 4px;
        min-height: 1em;
    }
    body.edit-mode [contenteditable="true"]:focus {
        outline-style: solid;
        background: var(--accent-light);
    }

    /* === 浮動圖片（可拖曳排版） === */
    .floating-img-wrapper {{
        position: absolute;
        cursor: move;
        border: 2px solid transparent;
        border-radius: 6px;
        user-select: none;
        z-index: 20;
        display: flex;
        align-items: center;
        justify-content: center;
        min-width: 60px;
        min-height: 60px;
    }}
    .floating-img-wrapper.selected {{
        border-color: var(--accent);
        box-shadow: 0 0 0 3px var(--accent-light);
    }}
    .floating-img-wrapper img,
    .floating-img-wrapper video {{
        display: block;
        width: 100%;
        height: 100%;
        object-fit: contain;
        pointer-events: none;
        border-radius: 4px;
    }}
    .floating-img-wrapper video {{
        pointer-events: auto;
    }}
    body.edit-mode .floating-img-wrapper video {{
        pointer-events: none;
    }}
    .resize-handle {{
        position: absolute;
        width: 10px;
        height: 10px;
        background: var(--accent);
        border: 2px solid #fff;
        border-radius: 2px;
        display: none;
        z-index: 21;
        box-shadow: 0 1px 4px rgba(0,0,0,0.3);
    }}
    .floating-img-wrapper.selected .resize-handle {{ display: block; }}
    .resize-handle.nw {{ top: -5px; left: -5px; cursor: nw-resize; }}
    .resize-handle.n  {{ top: -5px; left: calc(50% - 5px); cursor: n-resize; }}
    .resize-handle.ne {{ top: -5px; right: -5px; cursor: ne-resize; }}
    .resize-handle.e  {{ top: calc(50% - 5px); right: -5px; cursor: e-resize; }}
    .resize-handle.se {{ bottom: -5px; right: -5px; cursor: se-resize; }}
    .resize-handle.s  {{ bottom: -5px; left: calc(50% - 5px); cursor: s-resize; }}
    .resize-handle.sw {{ bottom: -5px; left: -5px; cursor: sw-resize; }}
    .resize-handle.w  {{ top: calc(50% - 5px); left: -5px; cursor: w-resize; }}
    .img-delete-btn {{
        position: absolute;
        top: -14px; right: -14px;
        width: 24px; height: 24px;
        background: #e74c3c;
        color: #fff;
        border: 2px solid #fff;
        border-radius: 50%;
        display: none;
        align-items: center;
        justify-content: center;
        font-size: 14px;
        cursor: pointer;
        line-height: 1;
        z-index: 22;
        box-shadow: 0 2px 6px rgba(0,0,0,0.3);
    }}
    .floating-img-wrapper.selected .img-delete-btn {{ display: flex; }}
    body.edit-mode .placeholder-img {{
        cursor: pointer;
        transition: opacity 0.2s, border-color 0.2s;
    }}
    body.edit-mode .placeholder-img:hover {{
        opacity: 0.7;
        border-color: var(--accent);
    }}
    body.edit-mode .placeholder-img::after {{
        content: '點擊上傳圖片';
        position: absolute;
        bottom: 0.5rem;
        font-size: 0.75rem;
        color: var(--accent);
        opacity: 0.8;
    }}
    body.edit-mode .placeholder-img {{ position: relative; }}

    .notes-edit-modal {
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.6);
        z-index: 500;
        align-items: center;
        justify-content: center;
    }
    .notes-edit-modal.visible { display: flex; }
    .notes-edit-box {
        background: var(--bg-card);
        border-radius: 1rem;
        padding: 2rem;
        width: 90%;
        max-width: 600px;
        box-shadow: 0 8px 40px rgba(0,0,0,0.3);
    }
    .notes-edit-box textarea {
        width: 100%;
        min-height: 120px;
        border: 1px solid var(--border);
        border-radius: 0.5rem;
        padding: 0.8rem;
        font-family: var(--font);
        font-size: 1rem;
        resize: vertical;
        background: var(--bg);
        color: var(--text-primary);
    }
    .notes-edit-box h3 { margin-bottom: 0.8rem; color: var(--text-primary); }
    .notes-edit-actions { margin-top: 1rem; display: flex; gap: 0.5rem; justify-content: flex-end; }
    .notes-edit-actions button { padding: 0.5rem 1.2rem; border-radius: 0.5rem; cursor: pointer; border: 1px solid var(--border); background: var(--bg-card); color: var(--text-primary); font-family: var(--font); }
    .notes-edit-actions .btn-ok { background: var(--accent); color: #fff; border-color: var(--accent); }
    .save-toast {
        position: fixed;
        top: 2rem; left: 50%;
        transform: translateX(-50%) translateY(-100px);
        background: var(--accent);
        color: #fff;
        padding: 0.8rem 2rem;
        border-radius: 2rem;
        font-weight: 600;
        z-index: 600;
        transition: transform 0.4s ease;
        pointer-events: none;
    }
    .save-toast.show { transform: translateX(-50%) translateY(0); }

    /* === 側邊導覽列 === */
    .nav-toggle {
        position: fixed;
        right: 1.5rem;
        top: 50%;
        transform: translateY(-50%);
        background: var(--bg-card);
        border: 1px solid var(--border);
        color: var(--text-secondary);
        width: 2rem; height: 2rem;
        border-radius: 50%;
        cursor: pointer;
        font-size: 0.8rem;
        z-index: 350;
        box-shadow: 0 2px 8px var(--shadow);
        display: flex; align-items: center; justify-content: center;
        transition: background 0.2s, right 0.3s;
        opacity: 0.5;
    }
    .nav-toggle:hover { opacity: 1; }
    body.nav-open .nav-toggle { right: 17rem; opacity: 1; }
    .nav-sidebar {
        position: fixed;
        top: 0; right: 0;
        width: 16rem;
        height: 100vh;
        background: var(--bg-card);
        border-left: 1px solid var(--border);
        box-shadow: -4px 0 20px var(--shadow);
        z-index: 340;
        transform: translateX(100%);
        transition: transform 0.3s ease;
        display: flex;
        flex-direction: column;
        overflow: hidden;
    }
    body.nav-open .nav-sidebar { transform: translateX(0); }
    .nav-header {
        padding: 1rem 1.2rem 0.5rem;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        color: var(--muted);
        border-bottom: 1px solid var(--border);
    }
    .nav-list {
        flex: 1;
        overflow-y: auto;
        padding: 0.5rem 0;
        list-style: none;
    }
    .nav-list::-webkit-scrollbar { width: 3px; }
    .nav-list::-webkit-scrollbar-thumb { background: var(--muted); border-radius: 2px; }
    .nav-item {
        padding: 0.45rem 1.2rem;
        font-size: 0.82rem;
        color: var(--text-secondary);
        cursor: pointer;
        transition: background 0.15s;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        border-left: 3px solid transparent;
    }
    .nav-item:hover { background: var(--accent-light); }
    .nav-item.active { color: var(--accent); border-left-color: var(--accent); font-weight: 600; }
    .nav-item.nav-section {
        font-weight: 700;
        color: var(--text-primary);
        margin-top: 0.3rem;
        font-size: 0.85rem;
    }

    /* === 圖片縮放預覽 (Lightbox) === */
    .image-modal {{
        display: none;
        position: fixed;
        inset: 0;
        background: rgba(0,0,0,0.96);
        z-index: 1000;
        align-items: center;
        justify-content: center;
        cursor: crosshair;
        user-select: none;
    }}
    .image-modal.visible {{ display: flex; }}
    .image-modal-close {{
        position: absolute;
        top: 1.5rem; right: 2rem;
        color: #fff; font-size: 2.5rem;
        cursor: pointer; z-index: 1001;
        opacity: 0.6; transition: opacity 0.2s;
    }}
    .image-modal-close:hover {{ opacity: 1; }}
    .modal-viewport {{
        width: 100%; height: 100%;
        display: flex; align-items: center; justify-content: center;
        overflow: hidden;
    }}
    .modal-img {{
        max-width: 90%; max-height: 90%;
        object-fit: contain;
        transition: transform 0.1s ease-out;
        cursor: grab;
        transform-origin: center center;
    }}
    .modal-img.dragging {{ cursor: grabbing; transition: none; }}
    .zoom-hint {{
        position: absolute; bottom: 1.5rem; left: 50%;
        transform: translateX(-50%);
        color: #888; font-size: 0.9rem;
        background: rgba(0,0,0,0.5);
        padding: 0.4rem 1.2rem; border-radius: 2rem;
        pointer-events: none;
    }}

    /* === 列印 PDF === */
    @media print {{
        html, body {{ overflow: visible; height: auto; }}
        .progress-bar, .notes-panel, .theme-switcher, .edit-btn, .edit-toolbar, .notes-edit-modal, .save-toast, .nav-sidebar, .nav-toggle, .image-modal {{ display: none !important; }}
        .slide {{
            position: relative;
            opacity: 1;
            pointer-events: auto;
            page-break-after: always;
            page-break-inside: avoid;
            min-height: 100vh;
            width: 100vw;
        }}
        .page-num {{ position: absolute; }}
    }}
    '''


def generate_js(default_theme: str) -> str:
    return '''
    // === 主題定義 ===
    const THEMES = {
        'minimal-white': {
            name: '極簡白', dot: '#FFFFFF',
            bg: '#FFFFFF', 'bg-card': '#F7F7F7', 'bg-card-hover': '#F0F0F0',
            'text-primary': '#3A3A3A', 'text-secondary': '#666666',
            accent: '#007ACC', 'accent-light': 'rgba(0,122,204,0.08)',
            muted: '#999999', border: '#E5E5E5', shadow: 'rgba(0,0,0,0.06)',
            font: "'Noto Serif TC','Microsoft JhengHei',serif"
        },
        'architect-dark': {
            name: '建築深色', dot: '#1A1A1A',
            bg: '#0F0F0F', 'bg-card': '#1A1A1A', 'bg-card-hover': '#222222',
            'text-primary': '#E0E0E0', 'text-secondary': '#AAAAAA',
            accent: '#D4A853', 'accent-light': 'rgba(212,168,83,0.10)',
            muted: '#666666', border: '#2A2A2A', shadow: 'rgba(0,0,0,0.3)',
            font: "'Noto Serif TC','Microsoft JhengHei',serif"
        },
        'blueprint': {
            name: '藍圖', dot: '#1B2A4A',
            bg: '#1B2A4A', 'bg-card': '#223356', 'bg-card-hover': '#2A3D66',
            'text-primary': '#E0E8F0', 'text-secondary': '#B0C8E0',
            accent: '#5DADE2', 'accent-light': 'rgba(93,173,226,0.10)',
            muted: '#6680A0', border: '#2E4570', shadow: 'rgba(0,0,0,0.3)',
            font: "'Noto Serif TC','Microsoft JhengHei',serif"
        },
        'concrete': {
            name: '清水模', dot: '#D8D3CE',
            bg: '#E8E4E1', 'bg-card': '#D8D3CE', 'bg-card-hover': '#CEC8C2',
            'text-primary': '#444444', 'text-secondary': '#666666',
            accent: '#8B4513', 'accent-light': 'rgba(139,69,19,0.08)',
            muted: '#8B8B8B', border: '#C0BAB5', shadow: 'rgba(0,0,0,0.08)',
            font: "'Noto Serif TC','Microsoft JhengHei',serif"
        },
        'nature': {
            name: '自然', dot: '#27AE60',
            bg: '#FAFAF5', 'bg-card': '#F0EDE5', 'bg-card-hover': '#E8E4DA',
            'text-primary': '#3D4446', 'text-secondary': '#555555',
            accent: '#27AE60', 'accent-light': 'rgba(39,174,96,0.08)',
            muted: '#95A5A6', border: '#D5D0C8', shadow: 'rgba(0,0,0,0.06)',
            font: "'Noto Serif TC','Microsoft JhengHei',serif"
        }
    };

    // === 主題切換 ===
    function applyTheme(id) {
        const t = THEMES[id];
        if (!t) return;
        const r = document.documentElement.style;
        for (const [k, v] of Object.entries(t)) {
            if (k === 'name' || k === 'dot' || k === 'font') continue;
            r.setProperty('--' + k, v);
        }
        r.setProperty('--font', t.font);
        // 更新 dot active
        document.querySelectorAll('.theme-dot').forEach(d => {
            d.classList.toggle('active', d.dataset.theme === id);
        });
        localStorage.setItem('presentation-theme', id);
    }

    // 初始化主題（localStorage 優先 > frontmatter 預設）
    const savedTheme = localStorage.getItem('presentation-theme');
    applyTheme(savedTheme || ''' + f"'{default_theme}'" + ''');

    // 綁定 dot 點擊（避免冒泡到翻頁）
    document.querySelectorAll('.theme-dot').forEach(dot => {
        dot.addEventListener('click', e => {
            e.stopPropagation();
            applyTheme(dot.dataset.theme);
        });
    });
    document.querySelector('.theme-switcher').addEventListener('click', e => e.stopPropagation());

    // === 投影片導航 ===
    const slides = document.querySelectorAll('.slide');
    const progress = document.querySelector('.progress-bar');
    const notesPanel = document.querySelector('.notes-panel');
    const notesContent = document.querySelector('.notes-content');
    let current = 0;
    let notesVisible = false;

    function showSlide(n) {
        if (n < 0 || n >= slides.length) return;
        slides[current].classList.remove('active');
        current = n;
        slides[current].classList.add('active');
        progress.style.width = ((current + 1) / slides.length * 100) + '%';
        const notes = slides[current].getAttribute('data-notes');
        notesContent.textContent = notes || '(no notes)';
    }

    function next() { showSlide(current + 1); }
    function prev() { showSlide(current - 1); }

    document.addEventListener('keydown', e => {
        if (imageModal && imageModal.classList.contains('visible')) return; // lightbox 開啟時不換頁
        // 影片控制：若當前投影片有 <video>，Space／←→ 優先控制影片
        const curVideo = slides[current] && slides[current].querySelector('video');
        if (curVideo) {
            if (e.key === ' ') { e.preventDefault(); curVideo.paused ? curVideo.play() : curVideo.pause(); return; }
            if (e.key === 'ArrowRight') { e.preventDefault(); curVideo.currentTime = Math.min(curVideo.duration || Infinity, curVideo.currentTime + 5); return; }
            if (e.key === 'ArrowLeft') { e.preventDefault(); curVideo.currentTime = Math.max(0, curVideo.currentTime - 5); return; }
        }
        if (e.key === 'ArrowRight' || e.key === ' ') { e.preventDefault(); next(); }
        else if (e.key === 'ArrowLeft') { e.preventDefault(); prev(); }
        else if (e.key === 'n' || e.key === 'N') {
            notesVisible = !notesVisible;
            notesPanel.classList.toggle('visible', notesVisible);
        }
        else if (e.key === 'f' || e.key === 'F') {
            if (!document.fullscreenElement) document.documentElement.requestFullscreen();
            else document.exitFullscreen();
        }
    });

    // 點擊換頁已停用，改用滾輪或鍵盤左右鍵換頁

    let touchStartX = 0;
    document.addEventListener('touchstart', e => { touchStartX = e.touches[0].clientX; });
    document.addEventListener('touchend', e => {
        if (lightboxOpen) return; // lightbox 開啟時不換頁
        const diff = e.changedTouches[0].clientX - touchStartX;
        if (Math.abs(diff) > 50) { diff < 0 ? next() : prev(); }
    });

    // === 導覽列 ===
    document.querySelector('.nav-toggle').addEventListener('click', e => {
        e.stopPropagation();
        document.body.classList.toggle('nav-open');
    });
    document.querySelector('.nav-sidebar').addEventListener('click', e => e.stopPropagation());

    // 動態建立導覽列（從 slide 標題讀取，確保與內容永遠同步）
    function buildNav() {
        const navList = document.querySelector('.nav-list');
        navList.innerHTML = '';
        slides.forEach((slide, idx) => {
            const li = document.createElement('li');
            li.className = 'nav-item';
            if (slide.classList.contains('slide--section')) li.classList.add('nav-section');
            li.dataset.slide = idx;
            const titleEl = slide.querySelector('.section-title, .cover-title, .closing-title, .slide-title, .toc-title');
            li.textContent = titleEl ? titleEl.textContent.trim() : `投影片 ${idx + 1}`;
            li.addEventListener('click', e => {
                e.stopPropagation();
                showSlide(idx);
                if (window.innerWidth < 1200) document.body.classList.remove('nav-open');
            });
            navList.appendChild(li);
        });
    }

    function updateNav() {
        document.querySelectorAll('.nav-item').forEach(item => {
            item.classList.toggle('active', parseInt(item.dataset.slide) === current);
        });
        const activeNav = document.querySelector('.nav-item.active');
        if (activeNav) activeNav.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
    buildNav();
    // 覆蓋 showSlide 加入 nav 更新
    const _origShowSlide = showSlide;
    showSlide = function(n) { _origShowSlide(n); updateNav(); };
    updateNav();

    // 滾輪換頁（加 debounce 避免連續觸發）
    // 掛在 .deck 而非 document，讓 lightbox/nav/notes 等兄弟元素的滾輪事件
    // 永遠不會進入此冒泡路徑；lightboxOpen 為額外保險
    let wheelLock = false;
    document.querySelector('.deck').addEventListener('wheel', e => {
        if (editMode || wheelLock || lightboxOpen) return;
        wheelLock = true;
        if (e.deltaY > 0) next(); else if (e.deltaY < 0) prev();
        setTimeout(() => { wheelLock = false; }, 400);
    }, { passive: true });

    // === 圖片縮放預覽 (Lightbox) ===
    const imageModal = document.querySelector('.image-modal');
    const modalImg = document.querySelector('.modal-img');
    let scale = 1;
    let posX = 0, posY = 0;
    let isDraggingImg = false;
    let startDragX, startDragY;
    let lightboxSrcs = [];
    let lightboxIdx = 0;
    let lightboxOpen = false;

    function openLightbox(src) {
        // 收集當前投影片所有靜態圖片（排除 video）
        const curSlide = slides[current];
        lightboxSrcs = [...curSlide.querySelectorAll('img.slide-img')]
            .map(i => i.src).filter(Boolean);
        lightboxIdx = lightboxSrcs.indexOf(src);
        if (lightboxIdx < 0) { lightboxSrcs = [src]; lightboxIdx = 0; }
        modalImg.src = src;
        scale = 1; posX = 0; posY = 0;
        updateModalTransform();
        imageModal.classList.add('visible');
        lightboxOpen = true;
    }

    function closeLightbox() {
        imageModal.classList.remove('visible');
        lightboxOpen = false;
    }

    function lightboxNav(dir) {
        if (lightboxSrcs.length < 2) return;
        lightboxIdx = (lightboxIdx + dir + lightboxSrcs.length) % lightboxSrcs.length;
        modalImg.src = lightboxSrcs[lightboxIdx];
        scale = 1; posX = 0; posY = 0;
        updateModalTransform();
    }

    function updateModalTransform() {
        modalImg.style.transform = `translate(${posX}px, ${posY}px) scale(${scale})`;
    }

    // 點擊投影片內的圖片開啟
    document.addEventListener('click', e => {
        const img = e.target.closest('.slide-img, .placeholder-img, .floating-img-wrapper img');
        if (img && !editMode) {
            if (img.src) openLightbox(img.src);
        }
    });

    imageModal.addEventListener('wheel', e => {
        if (!imageModal.classList.contains('visible')) return;
        e.preventDefault();
        e.stopPropagation();
        const delta = -e.deltaY;
        const factor = delta > 0 ? 1.1 : 0.9;
        scale = Math.min(Math.max(0.5, scale * factor), 10);
        updateModalTransform();
    }, { passive: false });

    modalImg.addEventListener('mousedown', e => {
        if (!imageModal.classList.contains('visible')) return;
        isDraggingImg = true;
        modalImg.classList.add('dragging');
        startDragX = e.clientX - posX;
        startDragY = e.clientY - posY;
        e.preventDefault();
    });

    document.addEventListener('mousemove', e => {
        if (!isDraggingImg) return;
        posX = e.clientX - startDragX;
        posY = e.clientY - startDragY;
        updateModalTransform();
    });

    document.addEventListener('mouseup', () => {
        isDraggingImg = false;
        modalImg.classList.remove('dragging');
    });

    imageModal.addEventListener('click', e => {
        if (e.target === imageModal || e.target.closest('.image-modal-close') || e.target.closest('.modal-viewport')) {
            if (!isDraggingImg && e.target !== modalImg) closeLightbox();
        }
    });

    document.addEventListener('keydown', e => {
        if (!imageModal.classList.contains('visible')) return;
        if (e.key === 'Escape') { closeLightbox(); return; }
        if (e.key === 'ArrowRight') { e.preventDefault(); closeLightbox(); next(); }
        else if (e.key === 'ArrowLeft') { e.preventDefault(); closeLightbox(); prev(); }
    });

    showSlide(0);

    // === 自動縮小溢出投影片的字體 ===
    // 用 children offsetHeight 加總判斷溢出，比 scrollHeight 對 flexbox 更可靠
    function autoFitSlides() {
        const SKIP = ['slide--section', 'slide--cover', 'slide--toc', 'slide--closing'];
        document.querySelectorAll('.slide').forEach(slide => {
            if (SKIP.some(c => slide.classList.contains(c))) return;
            slide.style.fontSize = '';
            const footer = slide.querySelector('.slide-footer');
            const style = getComputedStyle(slide);
            const padV = parseFloat(style.paddingTop) + parseFloat(style.paddingBottom);
            const available = slide.clientHeight - padV;
            function totalH() {
                let h = 0;
                for (const el of slide.children) {
                    if (el === footer) continue;
                    const cs = getComputedStyle(el);
                    h += el.offsetHeight + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
                }
                return h;
            }
            if (totalH() <= available) return;
            for (let scale = 0.88; scale >= 0.45; scale -= 0.04) {
                slide.style.fontSize = scale + 'em';
                if (totalH() <= available) break;
            }
        });
    }
    function autoFitCards() {
        document.querySelectorAll('.card-column, .card-top').forEach(card => {
            card.style.fontSize = '';  // 先 reset，從 CSS 定義的最大值開始
            const style = getComputedStyle(card);
            const padV = parseFloat(style.paddingTop) + parseFloat(style.paddingBottom);
            const available = card.clientHeight - padV;
            if (available <= 0) return;
            function totalH() {
                let h = 0;
                for (const el of card.children) {
                    const cs = getComputedStyle(el);
                    h += el.offsetHeight + parseFloat(cs.marginTop) + parseFloat(cs.marginBottom);
                }
                return h;
            }
            if (totalH() <= available) return;
            const startPx = parseFloat(getComputedStyle(card).fontSize);
            const minPx = startPx * 0.45;
            const stepPx = startPx * 0.04;
            for (let px = startPx - stepPx; px >= minPx; px -= stepPx) {
                card.style.fontSize = px + 'px';
                if (totalH() <= available) break;
            }
        });
    }

    // 頁面渲染完成後執行（避免 flexbox 尺寸未計算問題）
    // autoFitSlides 先跑整頁縮放，再跑 autoFitCards 做 per-card 縮放
    requestAnimationFrame(function() {
        requestAnimationFrame(function() {
            autoFitSlides();
            requestAnimationFrame(autoFitCards);
        });
    });

    // === 編輯模式 ===
    let editMode = false;
    const EDITABLE_SEL = '.slide-title, .section-title, .cover-title, .closing-title, .cover-subtitle, .bullet-list li, .body-text, .card-table td, .card-table th, .demo-steps li, .closing-points li, .placeholder-text, .quote-text';

    function toggleEdit() {
        editMode = !editMode;
        document.body.classList.toggle('edit-mode', editMode);
        document.querySelector('.edit-toolbar').style.display = editMode ? 'flex' : 'none';
        document.querySelectorAll(EDITABLE_SEL).forEach(el => { el.contentEditable = editMode; });
        // 編輯模式時停用點擊換頁
    }

    document.querySelector('.edit-btn').addEventListener('click', e => { e.stopPropagation(); toggleEdit(); });
    document.addEventListener('keydown', e => {
        if (e.key === 'e' || e.key === 'E') {
            if (!editMode && !e.target.isContentEditable) toggleEdit();
        }
    });

    // === 圖片插入與拖曳排版 ===
    const imgFileInput = document.getElementById('img-file-input');
    let imgTargetPlaceholder = null;

    function readImgFile(file, cb) {
        const reader = new FileReader();
        reader.onload = e => cb(e.target.result, file.name);
        reader.readAsDataURL(file);
    }

    function createFloatingImg(src, slideEl, x, y, w, h, relPath) {
        const wrapper = document.createElement('div');
        wrapper.className = 'floating-img-wrapper';
        wrapper.style.left   = Math.round(x) + 'px';
        wrapper.style.top    = Math.round(y) + 'px';
        wrapper.style.width  = Math.round(w) + 'px';
        wrapper.style.height = Math.round(h) + 'px';
        if (relPath) wrapper.dataset.relPath = relPath;

        const img = document.createElement('img');
        img.src = src;
        img.draggable = false;
        wrapper.appendChild(img);

        ['nw','n','ne','e','se','s','sw','w'].forEach(dir => {
            const h = document.createElement('div');
            h.className = 'resize-handle ' + dir;
            h.dataset.dir = dir;
            wrapper.appendChild(h);
        });

        const delBtn = document.createElement('div');
        delBtn.className = 'img-delete-btn';
        delBtn.innerHTML = '&times;';
        delBtn.addEventListener('mousedown', e => { e.stopPropagation(); });
        delBtn.addEventListener('click', e => { e.stopPropagation(); wrapper.remove(); });
        wrapper.appendChild(delBtn);

        slideEl.appendChild(wrapper);
        setupImgDragResize(wrapper);
        return wrapper;
    }

    function createFloatingVideo(src, slideEl, x, y, w, h, relPath) {
        const wrapper = document.createElement('div');
        wrapper.className = 'floating-img-wrapper';
        wrapper.dataset.mediaType = 'video';
        wrapper.style.left   = Math.round(x) + 'px';
        wrapper.style.top    = Math.round(y) + 'px';
        wrapper.style.width  = Math.round(w) + 'px';
        wrapper.style.height = Math.round(h) + 'px';
        if (relPath) wrapper.dataset.relPath = relPath;

        const vid = document.createElement('video');
        vid.src = src;
        vid.controls = true;
        vid.draggable = false;
        wrapper.appendChild(vid);

        ['nw','n','ne','e','se','s','sw','w'].forEach(dir => {
            const h = document.createElement('div');
            h.className = 'resize-handle ' + dir;
            h.dataset.dir = dir;
            wrapper.appendChild(h);
        });
        const delBtn = document.createElement('div');
        delBtn.className = 'img-delete-btn';
        delBtn.innerHTML = '&times;';
        delBtn.addEventListener('mousedown', e => { e.stopPropagation(); });
        delBtn.addEventListener('click', e => { e.stopPropagation(); wrapper.remove(); });
        wrapper.appendChild(delBtn);

        slideEl.appendChild(wrapper);
        setupImgDragResize(wrapper);
        return wrapper;
    }

    function setupImgDragResize(wrapper) {
        let isDragging = false, isResizing = false;
        let startX, startY, startL, startT, startW, startH, resizeDir;

        wrapper.addEventListener('mousedown', e => {
            if (e.target.classList.contains('img-delete-btn')) return;
            e.stopPropagation();
            document.querySelectorAll('.floating-img-wrapper.selected')
                .forEach(el => el.classList.remove('selected'));
            wrapper.classList.add('selected');

            if (e.target.classList.contains('resize-handle')) {
                isResizing = true;
                resizeDir = e.target.dataset.dir;
            } else {
                isDragging = true;
            }
            startX = e.clientX; startY = e.clientY;
            startL = parseFloat(wrapper.style.left)  || 0;
            startT = parseFloat(wrapper.style.top)   || 0;
            startW = wrapper.offsetWidth;
            startH = wrapper.offsetHeight;

            const onMove = e => {
                const dx = e.clientX - startX, dy = e.clientY - startY;
                if (isDragging) {
                    wrapper.style.left = (startL + dx) + 'px';
                    wrapper.style.top  = (startT + dy) + 'px';
                } else if (isResizing) {
                    let nL = startL, nT = startT, nW = startW, nH = startH;
                    if (resizeDir.includes('e')) nW = Math.max(60, startW + dx);
                    if (resizeDir.includes('s')) nH = Math.max(60, startH + dy);
                    if (resizeDir.includes('w')) { nW = Math.max(60, startW - dx); nL = startL + startW - nW; }
                    if (resizeDir.includes('n')) { nH = Math.max(60, startH - dy); nT = startT + startH - nH; }
                    wrapper.style.left   = Math.round(nL) + 'px';
                    wrapper.style.top    = Math.round(nT) + 'px';
                    wrapper.style.width  = Math.round(nW) + 'px';
                    wrapper.style.height = Math.round(nH) + 'px';
                }
            };
            const onUp = () => {
                isDragging = false; isResizing = false;
                document.removeEventListener('mousemove', onMove);
                document.removeEventListener('mouseup', onUp);
            };
            document.addEventListener('mousemove', onMove);
            document.addEventListener('mouseup', onUp);
        });
    }

    // 點擊 placeholder → 上傳取代
    document.addEventListener('click', e => {
        if (!editMode) return;
        const ph = e.target.closest('.placeholder-img');
        if (!ph) return;
        e.stopPropagation();
        imgTargetPlaceholder = ph;
        imgFileInput.click();
    }, true);

    // 工具列「插入圖片」按鈕
    document.getElementById('btn-insert-img').addEventListener('click', e => {
        e.stopPropagation();
        imgTargetPlaceholder = null;
        imgFileInput.click();
    });

    // 點選空白處取消選取
    document.addEventListener('click', e => {
        if (!e.target.closest('.floating-img-wrapper'))
            document.querySelectorAll('.floating-img-wrapper.selected')
                .forEach(el => el.classList.remove('selected'));
    });

    // 處理選取的檔案
    imgFileInput.addEventListener('change', e => {
        const file = e.target.files[0];
        if (!file) { return; }
        const isVideo = file.type.startsWith('video/');
        const slideEl = slides[current];

        if (isVideo) {
            // 影片：用 binary 上傳，不做 base64
            fetch('/api/save-video', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/octet-stream',
                    'X-Filename': encodeURIComponent(file.name),
                    'Content-Length': file.size,
                },
                body: file,
            })
            .then(r => r.ok ? r.json() : Promise.reject(r))
            .then(data => {
                if (data.ok) {
                    _placeVideo(data.path, data.path, slideEl);
                } else {
                    alert('影片上傳失敗：' + (data.error || '未知錯誤'));
                }
            })
            .catch(() => alert('影片需要伺服器模式才能上傳。\\n請先用 presentation_server.py 啟動。'));
        } else {
            readImgFile(file, (dataUrl, filename) => {
                fetch('/api/save-image', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({base64: dataUrl, filename: filename})
                })
                .then(r => r.ok ? r.json() : Promise.reject())
                .then(data => {
                    if (data.ok) { _placeImg(data.path, data.path, slideEl); }
                    else { _placeImg(dataUrl, null, slideEl); }
                })
                .catch(() => _placeImg(dataUrl, null, slideEl));
            });
        }
        imgFileInput.value = '';
    });

    function _placeImg(src, relPath, slideEl) {
        if (imgTargetPlaceholder) {
            const rect = imgTargetPlaceholder.getBoundingClientRect();
            const sr   = slideEl.getBoundingClientRect();
            createFloatingImg(src, slideEl,
                rect.left - sr.left, rect.top - sr.top,
                rect.width, rect.height, relPath);
            imgTargetPlaceholder.style.display = 'none';
            imgTargetPlaceholder = null;
        } else {
            const sw = slideEl.offsetWidth, sh = slideEl.offsetHeight;
            createFloatingImg(src, slideEl, sw/2 - 150, sh/2 - 100, 300, 200, relPath);
        }
    }

    function _placeVideo(src, relPath, slideEl) {
        const sw = slideEl.offsetWidth, sh = slideEl.offsetHeight;
        createFloatingVideo(src, slideEl, sw/2 - 200, sh/2 - 130, 400, 260, relPath);
    }

    // 編輯模式下攔截點擊換頁
    const origClick = document.onclick;
    document.addEventListener('click', e => {
        if (editMode && !e.target.closest('.edit-toolbar') && !e.target.closest('.edit-btn') && !e.target.closest('.theme-switcher') && !e.target.closest('.notes-edit-modal')) {
            e.stopImmediatePropagation();
        }
    }, true);

    // 新增頁面
    document.getElementById('btn-add-slide').addEventListener('click', e => {
        e.stopPropagation();
        const tpl = document.createElement('section');
        tpl.className = 'slide slide--content';
        tpl.setAttribute('data-notes', '');
        tpl.innerHTML = '<h2 class="slide-title" contenteditable="true">新投影片標題</h2><div class="card content-card"><ul class="bullet-list"><li contenteditable="true">內容</li></ul></div><div class="page-num"></div>';
        const deck = document.querySelector('.deck');
        const ref = slides[current];
        ref.after(tpl);
        // 重新整理 slides NodeList — 要重新查詢
        location.reload();
    });

    // 刪除當前頁
    document.getElementById('btn-del-slide').addEventListener('click', e => {
        e.stopPropagation();
        if (slides.length <= 1) return;
        if (!confirm('確定刪除此頁？')) return;
        slides[current].remove();
        location.reload();
    });

    // 編輯備註
    document.getElementById('btn-edit-notes').addEventListener('click', e => {
        e.stopPropagation();
        const modal = document.querySelector('.notes-edit-modal');
        const ta = modal.querySelector('textarea');
        ta.value = slides[current].getAttribute('data-notes') || '';
        modal.classList.add('visible');
    });
    document.getElementById('notes-cancel').addEventListener('click', e => {
        document.querySelector('.notes-edit-modal').classList.remove('visible');
    });
    document.getElementById('notes-ok').addEventListener('click', e => {
        const ta = document.querySelector('.notes-edit-modal textarea');
        slides[current].setAttribute('data-notes', ta.value);
        notesContent.textContent = ta.value;
        document.querySelector('.notes-edit-modal').classList.remove('visible');
    });

    // === 儲存（HTML → MD → POST） ===
    document.getElementById('btn-save').addEventListener('click', e => {
        e.stopPropagation();
        const md = extractMarkdown();
        fetch('/api/save', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({markdown: md})
        })
        .then(r => r.json())
        .then(data => {
            const toast = document.querySelector('.save-toast');
            toast.textContent = data.ok ? '✓ 已儲存並更新 MD' : '✕ 儲存失敗';
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 2500);
            if (data.ok && data.reload) {
                setTimeout(() => location.reload(), 800);
            }
        })
        .catch(() => {
            const toast = document.querySelector('.save-toast');
            toast.textContent = '✕ 無法連線伺服器（請先啟動 presentation_server.py）';
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 3500);
        });
    });

    // === 快照功能 ===
    const snapshotModal = document.querySelector('.snapshot-modal');
    const snapshotInput = document.getElementById('snapshot-name-input');

    document.getElementById('btn-snapshot').addEventListener('click', e => {
        e.stopPropagation();
        snapshotInput.value = '';
        snapshotModal.classList.add('visible');
        setTimeout(() => snapshotInput.focus(), 50);
    });
    document.getElementById('snapshot-cancel').addEventListener('click', e => {
        snapshotModal.classList.remove('visible');
    });
    snapshotInput.addEventListener('keydown', e => {
        if (e.key === 'Enter') document.getElementById('snapshot-ok').click();
        if (e.key === 'Escape') snapshotModal.classList.remove('visible');
    });
    document.getElementById('snapshot-ok').addEventListener('click', e => {
        const name = snapshotInput.value.trim();
        if (!name) { snapshotInput.focus(); return; }
        snapshotModal.classList.remove('visible');
        fetch('/api/snapshot', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({name: name})
        })
        .then(r => r.json())
        .then(data => {
            const toast = document.querySelector('.save-toast');
            toast.textContent = data.ok
                ? '✓ 快照已儲存：snapshots/' + name
                : '✕ 快照失敗：' + (data.error || '');
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 3000);
        })
        .catch(() => {
            const toast = document.querySelector('.save-toast');
            toast.textContent = '✕ 需要啟動 presentation_server.py 才能儲存快照';
            toast.classList.add('show');
            setTimeout(() => toast.classList.remove('show'), 3500);
        });
    });

    function extractMarkdown() {
        const allSlides = document.querySelectorAll('.slide');
        let md = document.body.getAttribute('data-frontmatter') || '';
        md += '\\n\\n';

        allSlides.forEach((slide, i) => {
            if (i === 0) return; // cover handled by frontmatter
            const cls = slide.className;
            const notes = slide.getAttribute('data-notes') || '';

            if (cls.includes('slide--section')) {
                const t = slide.querySelector('.section-title');
                md += '# ' + (t ? t.textContent.trim() : '') + '\\n\\n';
            } else if (cls.includes('slide--closing')) {
                md += '<!-- type: closing -->\\n';
                const t = slide.querySelector('.closing-title');
                md += '## ' + (t ? t.textContent.trim() : '') + '\\n\\n';
                slide.querySelectorAll('.closing-points li').forEach(li => {
                    md += '- ' + li.textContent.trim() + '\\n';
                });
            } else if (cls.includes('slide--diagram')) {
                const t = slide.querySelector('.slide-title');
                md += '## ' + (t ? t.textContent.trim() : '') + '\\n\\n';
                const svg = slide.querySelector('svg');
                // 保留 diagram 標記（無法從 SVG 反推，所以用 data 屬性）
                const diagramId = slide.getAttribute('data-diagram') || '';
                if (diagramId) md += '<!-- diagram: ' + diagramId + ' -->\\n\\n';
            } else {
                const t = slide.querySelector('.slide-title');
                if (cls.includes('slide--comparison')) md += '<!-- type: comparison -->\\n';
                if (cls.includes('slide--demo')) md += '<!-- type: demo -->\\n';
                md += '## ' + (t ? t.textContent.trim() : '') + '\\n\\n';

                // bullets
                slide.querySelectorAll('.bullet-list > li').forEach(li => {
                    const sub = li.querySelector('ul');
                    const mainText = sub ? li.childNodes[0].textContent.trim() : li.textContent.trim();
                    md += '- ' + mainText + '\\n';
                    if (sub) {
                        sub.querySelectorAll('li').forEach(s => {
                            md += '  - ' + s.textContent.trim() + '\\n';
                        });
                    }
                });

                // code blocks
                slide.querySelectorAll('pre code').forEach(code => {
                    md += '\\n```\\n' + code.textContent + '\\n```\\n';
                });

                // tables
                slide.querySelectorAll('.card-table').forEach(table => {
                    const ths = [...table.querySelectorAll('th')].map(th => th.textContent.trim());
                    md += '\\n| ' + ths.join(' | ') + ' |\\n';
                    md += '|' + ths.map(() => '------').join('|') + '|\\n';
                    table.querySelectorAll('tbody tr').forEach(tr => {
                        const tds = [...tr.querySelectorAll('td')].map(td => td.textContent.trim());
                        md += '| ' + tds.join(' | ') + ' |\\n';
                    });
                });

                // placeholders（未被取代的）
                slide.querySelectorAll('.placeholder-img').forEach(ph => {
                    if (ph.style.display === 'none') return;
                    const pt = ph.querySelector('.placeholder-text');
                    const desc = pt ? pt.textContent.replace(/^請插入[：:]\\s*/, '') : '';
                    md += '\\n![placeholder: ' + desc + ']()\\n';
                });

                // 浮動圖片 / 影片（序列化位置與路徑）
                slide.querySelectorAll('.floating-img-wrapper').forEach(fw => {
                    const img = fw.querySelector('img');
                    const vid = fw.querySelector('video');
                    const media = img || vid;
                    if (!media) return;
                    const x = Math.round(parseFloat(fw.style.left)  || 0);
                    const y = Math.round(parseFloat(fw.style.top)   || 0);
                    const w = Math.round(fw.offsetWidth);
                    const h = Math.round(fw.offsetHeight);
                    const relPath = fw.dataset.relPath || '';
                    const src = relPath || media.src;
                    if (vid) {
                        md += '\\n<!-- float-video: {"x":' + x + ',"y":' + y + ',"w":' + w + ',"h":' + h + '} -->\\n';
                        md += '![影片](' + src + ')\\n';
                    } else {
                        md += '\\n<!-- float-img: {"x":' + x + ',"y":' + y + ',"w":' + w + ',"h":' + h + '} -->\\n';
                        md += '![圖片](' + src + ')\\n';
                    }
                });
            }

            if (notes) md += '\\n> 講者備註：' + notes + '\\n';
            md += '\\n---\\n\\n';
        });
        return md;
    }
    '''


def generate_html(md_path: str, output_path: str = None, theme_override: str = None):
    md_file = Path(md_path)
    if not md_file.exists():
        print(f'Error: file not found {md_path}', file=sys.stderr)
        sys.exit(1)

    content = md_file.read_text(encoding='utf-8')
    config, body = parse_frontmatter(content)

    default_theme = theme_override or config.get('theme', 'minimal-white')
    if default_theme not in THEMES:
        default_theme = 'minimal-white'

    slides_data = split_long_slides(parse_slides(body))
    if not slides_data:
        print('Error: no slides found', file=sys.stderr)
        sys.exit(1)

    # 封面
    title = config.get('title', 'Presentation')
    subtitle = config.get('subtitle', '')
    author = config.get('author', '')
    date = config.get('date', '')
    meta = ' | '.join(filter(None, [author, date]))

    cover_html = f'''
    <section class="slide slide--cover active">
        <div>
            <h1 class="cover-title">{escape_html(title).replace('\n', '<br>')}</h1>
            {'<p class="cover-subtitle">' + escape_html(subtitle) + '</p>' if subtitle else ''}
            {'<p class="cover-meta">' + escape_html(meta) + '</p>' if meta else ''}
        </div>
    </section>'''

    # 主題切換器 UI — 5 個色點
    theme_dots = []
    for tid, t in THEMES.items():
        dot_color = t.get('dot', t.get('accent', '#999'))
        theme_dots.append(
            f'<div class="theme-dot" data-theme="{tid}" '
            f'style="background:{dot_color};'
            f'{"border-color:#ccc;" if dot_color == "#FFFFFF" else ""}" '
            f'title="{t["name"]}"></div>'
        )
    switcher_html = f'<div class="theme-switcher">{"".join(theme_dots)}</div>'

    # 章節地圖 + 標注
    duration = config.get('duration', 20)
    sections = build_section_map(slides_data, duration)
    for sec in sections:
        slides_data[sec['slide_idx']]['_section_num'] = sec['section_num']
        slides_data[sec['slide_idx']]['_sections'] = sections

    # 若 MD 裡已有 type:toc 投影片，填入 sections 並跳過自動插入
    md_has_toc = any(sd.get('type') == 'toc' for sd in slides_data)
    if md_has_toc:
        for sd in slides_data:
            if sd.get('type') == 'toc':
                sd['_sections'] = sections
        injected_toc_html = ''
    else:
        # 自動插入目錄（封面之後，MD 無 toc 時的舊行為）
        if sections:
            toc_svg = render_chapter_flow_svg(sections)
            injected_toc_html = f'<section class="slide slide--toc"><h2 class="toc-title">今天的路線圖</h2><div class="chapter-flow-toc">{toc_svg}</div></section>'
        else:
            injected_toc_html = ''

    # 投影片
    has_injected = bool(injected_toc_html)
    total = len(slides_data) + 1 + (1 if has_injected else 0)
    slides_html = [cover_html]
    if has_injected:
        slides_html.append(injected_toc_html)
    footer_text = ' | '.join(filter(None, [author, date]))
    slide_offset = 2 if has_injected else 1
    for i, sd in enumerate(slides_data, slide_offset):
        sd['_author'] = footer_text
        slides_html.append(render_slide_html(sd, i, total - 1))

    # 導覽列
    nav_items = [f'<li class="nav-item" data-slide="0">封面</li>']
    if has_injected:
        nav_items.append('<li class="nav-item" data-slide="1">路線圖</li>')
    for i, sd in enumerate(slides_data, slide_offset):
        t = sd.get('title', f'第 {i} 頁') or f'第 {i} 頁'
        cls = 'nav-item nav-section' if sd.get('type') == 'section' else 'nav-item'
        nav_items.append(f'<li class="{cls}" data-slide="{i}">{escape_html(t)}</li>')
    nav_html = f'<nav class="nav-sidebar"><div class="nav-header">投影片導覽</div><ul class="nav-list"></ul></nav>'

    # 組裝
    css_text = generate_css().replace('{{', '{').replace('}}', '}')
    html = f'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;700&display=swap" rel="stylesheet">
<title>{escape_html(title)}</title>
<style>{css_text}</style>
</head>
<body data-frontmatter="{escape_html(content.split('---', 2)[0] + '---' + content.split('---', 2)[1] + '---') if '---' in content else ''}">
{switcher_html}
<div class="progress-bar" style="width: {100/total:.1f}%"></div>
<div class="deck">
{''.join(slides_html)}
</div>
<div class="image-modal">
    <div class="image-modal-close">&times;</div>
    <div class="modal-viewport">
        <img src="" class="modal-img">
    </div>
    <div class="zoom-hint">滾輪縮放 · 拖曳平移 · ESC 關閉</div>
</div>
<div class="notes-panel">
    <div class="notes-label">SPEAKER NOTES (N)</div>
    <div class="notes-content"></div>
</div>
{nav_html}
<button class="nav-toggle" title="導覽列">&#9776;</button>
<button class="edit-btn" title="編輯模式 (E)">&#9998;</button>
<div class="edit-toolbar">
    <button id="btn-insert-img">&#128247; 插入圖片/影片</button>
    <button id="btn-add-slide">+ 新增頁</button>
    <button id="btn-del-slide">&#128465; 刪除此頁</button>
    <button id="btn-edit-notes">&#128221; 備註</button>
    <button id="btn-save" class="btn-save">&#128190; 儲存</button>
    <button id="btn-snapshot" class="btn-snapshot">&#128247; 存快照</button>
</div>
<input type="file" id="img-file-input" accept="image/*,video/*" style="display:none">
<div class="snapshot-modal">
    <div class="notes-edit-box">
        <h3>儲存快照</h3>
        <p style="font-size:0.85rem;color:var(--text-secondary);margin-bottom:0.8rem">
            快照會複製到 <code>snapshots/</code> 資料夾，不影響工作版本。
        </p>
        <input id="snapshot-name-input" type="text"
               placeholder="例如：v01_SD初版、v02_業主審查"
               style="width:100%;padding:0.5rem 0.8rem;border:1px solid var(--border);
                      border-radius:0.5rem;background:var(--bg);color:var(--text-primary);
                      font-family:var(--font);font-size:0.9rem;margin-bottom:1rem">
        <div class="notes-edit-actions">
            <button id="snapshot-cancel">取消</button>
            <button id="snapshot-ok" class="btn-ok">確認儲存</button>
        </div>
    </div>
</div>
<div class="notes-edit-modal">
    <div class="notes-edit-box">
        <h3>編輯講者備註</h3>
        <textarea placeholder="輸入此頁的講者備註..."></textarea>
        <div class="notes-edit-actions">
            <button id="notes-cancel">取消</button>
            <button id="notes-ok" class="btn-ok">確定</button>
        </div>
    </div>
</div>
<div class="save-toast"></div>
<script>{generate_js(default_theme)}</script>
</body>
</html>'''

    if output_path is None:
        output_path = str(md_file.with_suffix('.html'))
    Path(output_path).write_text(html, encoding='utf-8')
    print(f'OK - HTML: {output_path}')
    print(f'   Theme: {THEMES[default_theme]["name"]} ({default_theme})')
    print(f'   Slides: {total}')
    print(f'   Keys: Arrow keys navigate | N = notes | F = fullscreen')

    return output_path


def main():
    parser = argparse.ArgumentParser(description='Project Showcase - HTML presentation generator')
    parser.add_argument('input', help='Presentation MD file')
    parser.add_argument('--output', '-o', help='Output HTML path')
    parser.add_argument('--theme', '-t', help=f'Theme ({", ".join(THEMES.keys())})')
    args = parser.parse_args()
    generate_html(args.input, args.output, theme_override=args.theme)


if __name__ == '__main__':
    main()

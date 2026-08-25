# -*- coding: utf-8 -*-
"""
小红书公开搜索采集脚本（本机 Chrome + Playwright）

用法:
    python collect_xhs.py

流程:
    1. 优先连接本机 9222 调试端口；否则用本机 Chrome 默认配置启动。
    2. 默认配置启动失败时，改用独立配置 .xhs-chrome-profile，首次需扫码登录一次。
    3. 自动搜索 5 个关键词，选择平台公开的“最热”排序，滚动加载。
    4. 对候选笔记打开详情页，读取公开的点赞/收藏/评论数和评论原文。
    5. 保存 JSON，并把结果回写到 ai-digital-product-market-signals.html。
"""
import datetime
import html
import json
import os
import pathlib
import re
import sys
import time
import urllib.parse

from playwright.sync_api import sync_playwright

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

REPO = pathlib.Path(__file__).resolve().parent
DATA_DIR = REPO / "xhs_data"
RAW_FILE = DATA_DIR / "raw_results.json"
HTML_FILE = REPO / "ai-digital-product-market-signals.html"
PROFILE_DIR = REPO / ".xhs-chrome-profile"

KEYWORDS = ["AI提示词", "AI文案代写", "小红书代运营", "Prompt模板", "AI工具测评"]
TOP_N = 20
CANDIDATES = 30
MAX_COMMENTS = 30
LOGIN_TIMEOUT_SEC = 300


def log(msg):
    print(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def chrome_path():
    p = pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "Application" / "chrome.exe"
    return str(p) if p.exists() else "chrome"


def default_profile():
    return str(pathlib.Path(os.environ.get("LOCALAPPDATA", "")) / "Google" / "Chrome" / "User Data")


def parse_count(text):
    if not text:
        return None
    text = text.strip().replace(",", "").replace(" ", "")
    m = re.match(r"^([\d.]+)\s*([万wW]?)$", text)
    if not m:
        return None
    num = float(m.group(1))
    if m.group(2):
        num *= 10000
    return int(num)


def is_logged_in(context):
    try:
        cookies = context.cookies("https://www.xiaohongshu.com")
        names = {c["name"] for c in cookies}
        return "web_session" in names
    except Exception:
        return False


def wait_login(page, context, timeout_sec=LOGIN_TIMEOUT_SEC):
    start = time.time()
    log("如页面显示登录二维码，请用小红书 App 扫码；登录后脚本会自动继续。")
    while time.time() - start < timeout_sec:
        if is_logged_in(context):
            log("检测到登录状态，继续执行。")
            return True
        try:
            logged = page.evaluate(
                "() => !!document.querySelector('.user-info, .user, img[src*=\"sns-avatar\"], .avatar')"
            )
            if logged:
                log("检测到登录状态，继续执行。")
                return True
        except Exception:
            pass
        page.wait_for_timeout(5000)
    return False


def launch_context(pw):
    try:
        browser = pw.chromium.connect_over_cdp("http://127.0.0.1:9222")
        if browser.contexts:
            log("已连接本机 Chrome 调试端口 9222。")
            return browser, browser.contexts[0], True
    except Exception:
        pass

    try:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=default_profile(),
            headless=False,
            executable_path=chrome_path(),
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
            args=["--disable-blink-features=AutomationControlled"],
        )
        log("已用本机 Chrome 默认配置启动（若你已登录小红书，无需再扫码）。")
        return None, ctx, False
    except Exception as exc:
        log(f"默认配置启动失败（Chrome 可能正在运行）：{exc}")
        log("改用独立采集配置 .xhs-chrome-profile，首次需要扫码登录一次。")

    ctx = pw.chromium.launch_persistent_context(
        user_data_dir=str(PROFILE_DIR),
        headless=False,
        executable_path=chrome_path(),
        viewport={"width": 1440, "height": 900},
        locale="zh-CN",
        args=["--disable-blink-features=AutomationControlled"],
    )
    return None, ctx, False


JS_CARDS = """
() => {
  const out = [];
  const seen = new Set();
  const cards = document.querySelectorAll('section.note-item, div.note-item, a[href*="/explore/"]');
  for (const node of cards) {
    let a = node.tagName === 'A' ? node : (node.querySelector('a[href*="/explore/"]') || node.closest('a[href*="/explore/"]'));
    if (!a) continue;
    const href = a.getAttribute('href') || '';
    const m = href.match(/\\/explore\\/([0-9a-zA-Z]+)/);
    if (!m) continue;
    const id = m[1];
    if (seen.has(id)) continue;
    seen.add(id);
    const card = node.tagName === 'A' ? (node.parentElement || node) : node;
    const titleEl = card.querySelector('[class*="title"]');
    const authorEl = card.querySelector('[class*="author"] [class*="name"], [class*="user"] [class*="name"]');
    const likeEl = card.querySelector('[class*="like"] [class*="count"], [class*="interact"] [class*="count"]');
    out.push({
      id: id,
      href: href.startsWith('http') ? href : 'https://www.xiaohongshu.com' + href,
      title: titleEl ? titleEl.innerText.trim() : '',
      author: authorEl ? authorEl.innerText.trim() : '',
      likes_text: likeEl ? likeEl.innerText.trim() : ''
    });
  }
  return out;
}
"""

JS_NOTE = """
() => {
  const out = {title:'', author:'', likes:null, collects:null, comments_count:null, comments:[]};
  try {
    const state = (typeof window.__INITIAL_STATE__ !== 'undefined') ? window.__INITIAL_STATE__ : null;
    if (state && state.note && state.note.noteDetailMap) {
      const key = Object.keys(state.note.noteDetailMap)[0];
      const note = state.note.noteDetailMap[key] && state.note.noteDetailMap[key].note;
      if (note) {
        out.title = note.title || out.title;
        out.author = (note.user && note.user.nickname) || out.author;
        const it = note.interactInfo || {};
        if (typeof it.likedCount !== 'undefined') out.likes = it.likedCount;
        if (typeof it.collectedCount !== 'undefined') out.collects = it.collectedCount;
        if (typeof it.commentCount !== 'undefined') out.comments_count = it.commentCount;
      }
    }
  } catch (e) {}
  const els = document.querySelectorAll('div.comment-item, div[class*="comment-item"], div[class*="commentItem"]');
  const seen = new Set();
  for (const el of els) {
    const c = el.querySelector('[class*="content"]') || el.querySelector('[class*="note-text"]') || el;
    const txt = (c.innerText || '').trim();
    if (txt && !seen.has(txt)) {
      seen.add(txt);
      out.comments.push(txt);
    }
  }
  return out;
}
"""


def click_hot(page):
    try:
        clicked = page.evaluate(
            """() => {
              const els = [...document.querySelectorAll('div,span,li,button')].filter(e => e.textContent.trim() === '最热' && e.offsetParent !== null);
              if (els.length) { els[0].click(); return true; }
              return false;
            }"""
        )
        if clicked:
            page.wait_for_timeout(2500)
            log("已点击搜索页“最热”排序。")
            return True
    except Exception:
        pass
    log("未找到“最热”按钮，继续使用页面当前排序。")
    return False


def scroll_load(page, min_cards=60, max_rounds=30):
    count = 0
    for i in range(max_rounds):
        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        except Exception:
            pass
        page.wait_for_timeout(1600)
        try:
            count = page.evaluate(
                "document.querySelectorAll('section.note-item, div.note-item, a[href*=\"/explore/\"]').length"
            )
        except Exception:
            count = 0
        if count >= min_cards:
            break
    log(f"滚动加载完成，当前页面卡片约 {count} 张。")
    return count


def collect_search(page, keyword):
    url = (
        "https://www.xiaohongshu.com/search_result?keyword="
        + urllib.parse.quote(keyword)
        + "&source=web_explore_feed"
    )
    log(f"开始采集关键词：{keyword}")
    page.goto(url, timeout=60000, wait_until="domcontentloaded")
    page.wait_for_timeout(3000)
    try:
        page.wait_for_selector(
            "section.note-item, div.note-item, a[href*='/explore/']", timeout=25000
        )
    except Exception:
        log("搜索页未出现笔记卡片，可能被风控或页面结构变化。")
        return []
    click_hot(page)
    scroll_load(page)
    cards = page.evaluate(JS_CARDS)
    seen = {}
    for c in cards:
        seen.setdefault(c["id"], c)
    unique = list(seen.values())
    for c in unique:
        c["likes"] = parse_count(c.get("likes_text", ""))
    unique.sort(key=lambda x: (x["likes"] if x["likes"] is not None else -1), reverse=True)
    log(f"关键词「{keyword}」解析到 {len(unique)} 张卡片。")
    return unique[:CANDIDATES]


def collect_note(page, card):
    try:
        page.goto(card["href"], timeout=60000, wait_until="domcontentloaded")
        page.wait_for_timeout(2500)
        for _ in range(5):
            try:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            except Exception:
                pass
            page.wait_for_timeout(1200)
        data = page.evaluate(JS_NOTE)
        data["comments"] = data["comments"][:MAX_COMMENTS]
        return {**card, **data, "error": None}
    except Exception as exc:
        return {**card, "title": card.get("title", ""), "author": card.get("author", ""),
                "likes": None, "collects": None, "comments_count": None,
                "comments": [], "error": str(exc)}


def esc(text):
    return html.escape(str(text if text is not None else ""))


def build_html_section(results, collected_at):
    rows_html = []
    for keyword, notes in results.items():
        detail_rows = []
        for i, n in enumerate(notes[:TOP_N], 1):
            likes = n.get("likes")
            collects = n.get("collects")
            cc = n.get("comments_count")
            fmt = lambda v: "未获取" if v is None else f"{v:,}"
            detail_rows.append(
                "<tr>"
                f"<td>{i}</td>"
                f"<td>{esc(n.get('title',''))}</td>"
                f"<td>{esc(n.get('author',''))}</td>"
                f"<td>{fmt(likes)}</td>"
                f"<td>{fmt(collects)}</td>"
                f"<td>{fmt(cc)}</td>"
                "</tr>"
            )
        comment_html = ""
        all_comments = []
        for n in notes[:TOP_N]:
            for c in n.get("comments", []):
                all_comments.append(c)
        if all_comments:
            items = "".join(f"<li>{esc(c)}</li>" for c in all_comments[:30])
            comment_html = f"<h4>评论原文样例（每篇首屏最多 30 条，共列出前 30 条）</h4><ul>{items}</ul>"
        rows_html.append(
            "<div class='card'>"
            f"<h3>关键词：{esc(keyword)}</h3>"
            "<div class='table-wrap'><table>"
            "<thead><tr><th>序号</th><th>标题</th><th>账号</th><th>点赞</th><th>收藏</th><th>评论数</th></tr></thead>"
            f"<tbody>{''.join(detail_rows)}</tbody></table></div>"
            f"{comment_html}</div>"
        )
    summary_rows = "".join(
        f"<tr><td>{esc(k)}</td><td>{len(v[:TOP_N])}</td><td>{collected_at}</td></tr>"
        for k, v in results.items()
    )
    section = (
        "<section class='module' id='xhs-real'>"
        "<div class='module-head'><span class='m-no'>七</span><h2>本机实测数据（小红书登录采集）</h2></div>"
        "<p>采集方式：本机 Chrome + Playwright，登录态下按搜索页公开“最热”排序滚动加载，再逐篇打开笔记页读取公开互动数与首屏评论原文。</p>"
        f"<p><strong>采集时间：</strong>{esc(collected_at)}　<strong>排序说明：</strong>小红书公开排序为“最热”，平台没有独立“按点赞”排序，本表同时记录每篇实际点赞数。</p>"
        "<div class='table-wrap'><table><thead><tr><th>关键词</th><th>有效笔记数</th><th>采集时间</th></tr></thead>"
        f"<tbody>{summary_rows}</tbody></table></div>"
        f"{''.join(rows_html)}"
        "<div class='note'><strong>诚实说明：</strong>点赞、收藏、评论数均来自笔记页公开数据；评论为页面首屏能加载到的公开评论，不代表全部评论。如某篇被风控拦截，表中会标注“未获取”。</div>"
        "</section>"
    )
    return section


def update_html(results, collected_at):
    if not HTML_FILE.exists():
        log(f"未找到 HTML 文件：{HTML_FILE}")
        return False
    text = HTML_FILE.read_text(encoding="utf-8")
    section = build_html_section(results, collected_at)
    if "<!-- XHS_REAL_DATA -->" in text:
        text = text.replace("<!-- XHS_REAL_DATA -->", section)
    else:
        text = text.replace("  <footer class=\"footer\">", section + "\n  <footer class=\"footer\">")

    status_top5 = (
        "<div class='tip'><strong>已更新：</strong>已通过本机脚本于 "
        f"{esc(collected_at)} 采集。小红书公开排序为“最热”，无独立按点赞排序；各关键词前 5 条见下方“本机实测数据”章节。</div>"
    )
    status_comments = (
        "<div class='tip'><strong>已更新：</strong>已通过本机脚本采集笔记页公开首屏评论原文，最多每篇记录 30 条；见下方“本机实测数据”章节。</div>"
    )
    pattern = re.compile(r'<div class="unavail"><strong>本项无法执行：</strong>.*?</div>', re.S)
    text, n1 = pattern.subn(status_top5, text, count=1)
    text, n2 = pattern.subn(status_comments, text, count=1)

    old_badge = '<td><span class="badge b-low">低</span> 仅能引用第三方内容，无法直接验证</td>'
    new_badge = '<td><span class="badge b-mid">中</span> 已用本机登录采集补充实测数据</td>'
    text = text.replace(old_badge, new_badge)

    HTML_FILE.write_text(text, encoding="utf-8")
    log(f"HTML 已更新：{HTML_FILE}")
    log(f"替换了 {n1} 个“无法执行”提示（TOP5）、{n2} 个“无法执行”提示（评论）。")
    return True


def main():
    DATA_DIR.mkdir(exist_ok=True)
    results = {}
    if RAW_FILE.exists():
        try:
            old = json.loads(RAW_FILE.read_text(encoding="utf-8"))
            results = old.get("results", {})
            log(f"检测到已有数据文件，已保留 {len(results)} 个关键词的结果。")
        except Exception:
            results = {}

    with sync_playwright() as pw:
        browser, context, attached = launch_context(pw)
        page = context.new_page()
        try:
            page.goto("https://www.xiaohongshu.com/explore", timeout=60000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            if not is_logged_in(context):
                ok = wait_login(page, context)
                if not ok:
                    log("等待登录超时，已中止。请重新运行脚本并完成扫码。")
                    return 1

            for keyword in KEYWORDS:
                if results.get(keyword) and len(results[keyword]) >= TOP_N:
                    log(f"「{keyword}」已有数据，跳过。")
                    continue
                cards = collect_search(page, keyword)
                notes = []
                for idx, card in enumerate(cards, 1):
                    log(f"「{keyword}」打开第 {idx}/{len(cards)} 篇：{card.get('id')}")
                    note = collect_note(page, card)
                    notes.append(note)
                    page.wait_for_timeout(1800)
                notes.sort(key=lambda x: (x.get("likes") if x.get("likes") is not None else -1), reverse=True)
                results[keyword] = notes[:TOP_N]
                RAW_FILE.write_text(
                    json.dumps({"collected_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "results": results}, ensure_ascii=False, indent=2),
                    encoding="utf-8",
                )
                log(f"「{keyword}」完成，已保存 {len(notes[:TOP_N])} 条。")
        finally:
            try:
                page.close()
            except Exception:
                pass
            if attached:
                log("保留你已打开的 Chrome，脚本不关闭它。")
            else:
                try:
                    context.close()
                except Exception:
                    pass

    collected_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    update_html(results, collected_at)
    log("全部完成。结果 JSON：" + str(RAW_FILE))
    return 0


if __name__ == "__main__":
    sys.exit(main())
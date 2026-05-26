#!/usr/bin/env python3
"""ShowZ Store Daily Inventory Check

Rules (applied only when stock == 0):
  No prefix (In Stock)  -> trigger sales check if switch is OFF
  Prefix + SKU in Sheet -> set stock = Sheet qty (skip if qty == 0)
  Prefix + not in Excel + [Pre-Order]   -> set stock = 20
  Prefix + not in Excel + [Coming Soon] -> skip

Notification: reports all products where stock == 0 at scan time.
"""

import asyncio
import csv
import io
import os
import re
import traceback
import requests
from datetime import datetime
from urllib.parse import urljoin, quote
from playwright.async_api import async_playwright

# ── Configuration ─────────────────────────────────────────────────────────────
MANAGE_BASE    = "https://showzstore.com/manage/"
LOGIN_URL      = MANAGE_BASE + "?m=products&a=products"
USERNAME       = "chantia@showz.store"
PASSWORD       = os.environ.get("SHOWZ_PASSWORD", "SS27650942")
SPREADSHEET_ID = "1JJKAvZh-bE2JlICqhBY17gwMbPG19_FzSInjU2Q-Z5M"
SHEET_NAMES    = ["APC Toys", "Iron Factory", "Gear Factory"]
# Each entry: (display_name, [search_keyword, ...])
# Iron Factory uses both "Iron Factory" and "IronFactory" in product names
BRANDS = [
    ("APC Toys",     ["APC Toys"]),
    ("Iron Factory", ["Iron Factory", "IronFactory"]),
    ("Gear Factory", ["Gear Factory"]),
]
TG_TOKEN   = os.environ.get("TELEGRAM_TOKEN",   "8841015387:AAEJUhOZDKgHp84GZ0NwujqI-e-2Ao5Q71I")
TG_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "8965386696")
LOG_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inventory_check.log")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


# ── Logging ───────────────────────────────────────────────────────────────────

def log(msg: str):
    ts   = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ── Windows proxy detection ───────────────────────────────────────────────────

def _get_win_proxy() -> dict | None:
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings",
        )
        enabled = winreg.QueryValueEx(key, "ProxyEnable")[0]
        server  = winreg.QueryValueEx(key, "ProxyServer")[0]
        winreg.CloseKey(key)
        if enabled and server:
            proxy_url = f"http://{server}"
            return {"http": proxy_url, "https": proxy_url}
    except Exception:
        pass
    return None


# ── Telegram ──────────────────────────────────────────────────────────────────

def send_telegram(msg: str):
    url     = f"https://api.telegram.org/bot{TG_TOKEN}/sendMessage"
    payload = {"chat_id": TG_CHAT_ID, "text": msg}
    proxies = _get_win_proxy()
    for attempt, px in enumerate([proxies, None]):
        try:
            requests.post(url, json=payload, timeout=20,
                          proxies=px if px else {})
            log(f"Telegram sent {'(via proxy)' if px else '(direct)'}")
            return
        except Exception as e:
            if attempt == 1:
                log(f"[WARN] Telegram failed: {e}")
                log("[INFO] All changes are recorded in inventory_check.log")


# ── Google Sheet reader ───────────────────────────────────────────────────────

def read_sku_qty(spreadsheet_id: str) -> dict:
    """Fetch {SKU: 剩余可加库存} from all sheets of a public Google Spreadsheet."""
    proxies = _get_win_proxy()
    data = {}
    for sheet in SHEET_NAMES:
        url = (
            f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
            f"/gviz/tq?tqx=out:csv&sheet={quote(sheet)}"
        )
        try:
            resp = requests.get(url, timeout=30, proxies=proxies or {})
            resp.raise_for_status()
        except Exception as e:
            log(f"[WARN] Sheet '{sheet}' fetch failed: {e}")
            continue

        rows = list(csv.reader(io.StringIO(resp.text)))
        if not rows:
            continue
        header = [str(c).strip() for c in rows[0]]
        sku_i = next(
            (i for i, h in enumerate(header) if re.search(r"SKU|编号|型号", h, re.I)),
            None,
        )
        qty_i = next(
            (i for i, h in enumerate(header) if "剩余可加库存" in h or "可加库存" in h),
            None,
        )
        if sku_i is None or qty_i is None:
            log(f"[WARN] Sheet '{sheet}': SKU or 剩余可加库存 column not found")
            continue
        for row in rows[1:]:
            sku = str(row[sku_i]).strip() if sku_i < len(row) and row[sku_i] else ""
            if not sku or sku.lower() == "none":
                continue
            try:
                data[sku] = int(float(row[qty_i] or 0))
            except (ValueError, TypeError):
                data[sku] = 0
    log(f"Google Sheet loaded: {len(data)} SKU(s) — {list(data.items())[:5]}")
    return data


# ── Browser helpers ────────────────────────────────────────────────────────────

async def safe_goto(page, url: str, wait_selector: str = "body"):
    for attempt in range(3):
        try:
            await page.goto(url, wait_until="commit", timeout=60_000)
            await page.wait_for_selector(wait_selector, timeout=90_000)
            return
        except Exception as e:
            if attempt < 2:
                log(f"[WARN] Page slow (attempt {attempt+1}), retrying: {url[:80]}")
                await page.wait_for_timeout(4000)
            else:
                raise e


async def do_login(page):
    await safe_goto(page, LOGIN_URL, wait_selector="input[placeholder='用户名']")
    await page.fill("input[placeholder='用户名']", USERNAME)
    await page.fill("input[type='password']", PASSWORD)
    await page.locator("input[type='submit'], .login-btn, button").first.click()
    await page.wait_for_selector("table tbody tr", timeout=60_000)
    log("Logged in successfully")



async def parse_product_rows(page) -> list:
    return await page.evaluate(r"""
        () => {
            const rows = document.querySelectorAll('table tbody tr');
            const result = [];
            for (const row of rows) {
                const cells = row.querySelectorAll('td');
                if (cells.length <= 14) continue;

                const cb    = cells[0].querySelector('input[type="checkbox"]');
                const proid = cb ? cb.value : null;

                const links = cells[2].querySelectorAll('a');
                const name  = links[0] ? links[0].innerText.trim() : '';
                const sku   = links[1] ? links[1].innerText.trim() : '';

                const rawStock = cells[7].innerText.trim();
                let stock = -1;
                const m = rawStock.match(/^(\d+)/);
                if (m) stock = parseInt(m[1]);

                const sw = cells[10].querySelector('.switchery');
                const salesCheckOn = sw ? sw.classList.contains('checked') : false;

                const editEl = cells[14].querySelector('a:first-child');
                const href   = editEl ? editEl.getAttribute('href') : null;

                if (name) result.push({ proid, name, sku, stock, salesCheckOn, href });
            }
            return result;
        }
    """)


async def edit_stock(page, href: str, new_stock: int, sku: str, back_url: str) -> bool:
    full_url = urljoin(MANAGE_BASE, href)
    try:
        await safe_goto(page, full_url, wait_selector="a[data-name='sales_info']")
    except Exception as e:
        log(f"[ERROR] Cannot load edit page for {sku}: {e}")
        try:
            await safe_goto(page, back_url, wait_selector="table tbody tr")
        except Exception:
            pass
        return False

    tab = await page.query_selector("a[data-name='sales_info']")
    if tab:
        await tab.click()

    try:
        await page.wait_for_selector("input[name='Stock']", state="visible", timeout=15_000)
    except Exception:
        log(f"[ERROR] input[name='Stock'] never became visible for {sku}")
        try:
            await safe_goto(page, back_url, wait_selector="table tbody tr")
        except Exception:
            pass
        return False

    inp = await page.query_selector("input[name='Stock']")
    if not inp:
        log(f"[ERROR] input[name='Stock'] not found for {sku}")
        try:
            await safe_goto(page, back_url, wait_selector="table tbody tr")
        except Exception:
            pass
        return False

    await inp.fill(str(new_stock))

    # Mark the target button with a known ID so we can click it via Playwright
    # (JS .click() inside evaluate doesn't wait for the resulting navigation)
    btn_val = await page.evaluate(r"""
        () => {
            const btns = [...document.querySelectorAll("input[name='submit_button'][type='submit']")];
            const preferred = btns.find(b => b.value.includes('前台'));
            const fallback  = btns.find(b => !b.value.includes('返回'));
            const btn = preferred || fallback;
            if (btn) { btn.id = '__sz_submit_btn__'; return btn.value; }
            return null;
        }
    """)
    if not btn_val:
        log(f"[ERROR] Submit button not found for {sku}")
        await safe_goto(page, back_url)
        return False

    # The CMS saves via jQuery $.post() AJAX — no page navigation occurs.
    # Intercept the POST response and verify data.ret == 1 for real success.
    try:
        async with page.expect_response(
            lambda r: r.request.method == "POST" and "manage" in r.url,
            timeout=60_000
        ) as resp_info:
            await page.click("#__sz_submit_btn__")

        resp = await resp_info.value
        try:
            result = await resp.json()
            if result.get("ret") != 1:
                log(f"[ERROR] {sku}: server rejected save: {result.get('msg', '')}")
                try:
                    await safe_goto(page, back_url, wait_selector="table tbody tr")
                except Exception:
                    pass
                return False
        except Exception:
            pass  # non-JSON response — treat as success
    except Exception as e:
        log(f"[ERROR] Submit AJAX failed for {sku}: {e}")
        try:
            await safe_goto(page, back_url, wait_selector="table tbody tr")
        except Exception:
            pass
        return False

    await page.wait_for_timeout(500)
    log(f"[OK] {sku}: stock set to {new_stock} (btn: {btn_val})")
    await safe_goto(page, back_url, wait_selector="table tbody tr")
    return True


async def trigger_sales_check(page, proid: str, sku: str) -> bool:
    for attempt in range(3):
        try:
            status = await page.evaluate(
                """async (args) => {
                    const fd = new FormData();
                    fd.append('do_action', 'products.check_products');
                    fd.append('ProId', args.proid);
                    fd.append('IsCheck', '1');
                    fd.append('Type', 'Blue');
                    const r = await fetch(args.base, {method: 'POST', body: fd,
                                                       credentials: 'include'});
                    return r.status;
                }""",
                {"proid": str(proid), "base": MANAGE_BASE},
            )
            ok = status == 200
            log(f"{'[OK]' if ok else '[WARN]'} {sku}: sales-check HTTP {status}")
            return ok
        except Exception as e:
            if attempt < 2:
                log(f"[WARN] sales-check attempt {attempt+1} failed for {sku}: {e}, retrying…")
                await asyncio.sleep(3)
            else:
                log(f"[ERROR] sales-check failed for {sku} after 3 attempts: {e}")
                return False
    return False


# ── Brand processor ───────────────────────────────────────────────────────────

def product_status(name: str) -> str:
    if "[Pre-Order]" in name:
        return "Pre-Order"
    if "[Coming Soon]" in name:
        return "Coming Soon"
    return "In Stock"


async def process_brand(
    page, brand: str, keywords: list, sku_qty: dict, zero_stock: list, modified: list
) -> int:
    """Process all pages for a brand using one or more search keywords.

    Multiple keywords are needed when the same brand uses different spellings in
    product names (e.g. "Iron Factory" vs "IronFactory").  Results are deduplicated
    by ProId so each product is processed exactly once.
    Returns number of products scanned.
    """
    log(f"─── Brand: {brand}")
    seen_proids: set = set()
    MAX_PAGES   = 30
    scanned     = 0

    for keyword in keywords:
        page_no = 1
        kw_seen_proids: set = set()  # wrap-around detection per keyword

        while page_no <= MAX_PAGES:
            url = (
                f"{MANAGE_BASE}?Keyword={quote(keyword)}"
                f"&CateId=&Other=0&m=products&a=products&page={page_no}"
            )
            await safe_goto(page, url, wait_selector="table tbody tr, td.dataTables_empty")

            rows = await parse_product_rows(page)
            if not rows:
                await page.wait_for_timeout(5000)
                rows = await parse_product_rows(page)
            if not rows:
                log(f"    [{keyword}] Page {page_no}: no rows → done")
                break

            page_proids = {p["proid"] for p in rows if p.get("proid")}
            if page_proids and page_proids.issubset(kw_seen_proids):
                log(f"    [{keyword}] Page {page_no}: all products already seen → done")
                break
            kw_seen_proids.update(page_proids)

            # Only process products not yet handled by a previous keyword pass
            new_rows = [p for p in rows if p.get("proid") not in seen_proids]
            seen_proids.update(page_proids)

            log(f"    [{keyword}] Page {page_no}: {len(rows)} product(s), {len(new_rows)} new")
            scanned += len(new_rows)

            for p in new_rows:
                name, sku, stock = p["name"], p["sku"], p["stock"]
                status        = product_status(name)
                has_prefix    = bool(re.match(r'^\[', name.strip()))
                is_preorder   = "[Pre-Order]" in name
                is_comingsoon = "[Coming Soon]" in name

                if stock > 0:
                    continue

                if stock != 0:
                    continue  # -1 or unparseable → skip

                # stock == 0 + sales-check OFF: record for notification
                if not p["salesCheckOn"]:
                    zero_stock.append(f"[{status}] {sku}: 0件")

                if not has_prefix:
                    # Rule 1: In Stock, 前台可售=0
                    if p["salesCheckOn"]:
                        log(f"    In Stock stock=0 but sales-check already ON → skip: {sku}")
                        continue
                    log(f"    In Stock stock=0 → 销售检查: {sku}")
                    ok = await trigger_sales_check(page, p["proid"], sku)
                    if ok:
                        modified.append(f"[{status}] {sku}：已开启销售检查")

                else:
                    # Has prefix ([Pre-Order] or [Coming Soon])
                    if sku in sku_qty:
                        # Rule 2: SKU in Excel
                        qty = sku_qty[sku]
                        if qty > 0:
                            log(f"    {status} stock=0, Excel qty={qty} → set stock: {sku}")
                            ok = await edit_stock(page, p["href"], qty, sku, url)
                            if ok:
                                modified.append(f"[{status}] {sku}：加库存{qty}")
                        else:
                            log(f"    SKU in Excel but qty=0 → skip: {sku}")

                    elif is_preorder:
                        # Rule 3: Pre-Order + not in Excel → set 20
                        log(f"    Pre-Order stock=0, not in Excel → set 20: {sku}")
                        ok = await edit_stock(page, p["href"], 20, sku, url)
                        if ok:
                            modified.append(f"[{status}] {sku}：加库存20")

                    elif is_comingsoon:
                        # Rule 4: Coming Soon + not in Excel → skip
                        log(f"    Coming Soon stock=0, not in Excel → skip: {sku}")

            page_no += 1

    return scanned


# ── Browser lifecycle ─────────────────────────────────────────────────────────

async def _new_page(pw, win_proxy):
    kwargs: dict = {
        "headless": True,
        "args": ["--no-sandbox", "--disable-dev-shm-usage"],
    }
    if win_proxy:
        kwargs["proxy"] = {"server": win_proxy["https"]}
    browser = await pw.chromium.launch(**kwargs)
    ctx  = await browser.new_context(user_agent=UA)
    page = await ctx.new_page()
    return browser, page


# ── Main ───────────────────────────────────────────────────────────────────────

async def main():
    log("=" * 60)
    log("Inventory check started")

    sku_qty        = read_sku_qty(SPREADSHEET_ID)
    zero_stock: list = []
    modified:   list = []
    total_scanned    = 0

    async with async_playwright() as pw:
        win_proxy = _get_win_proxy()
        if win_proxy:
            log(f"Browser proxy: {win_proxy['https']}")

        browser, page = await _new_page(pw, win_proxy)
        try:
            await do_login(page)
        except Exception:
            log(f"[ERROR] Login failed: {traceback.format_exc()}")
            try:
                await browser.close()
            except Exception:
                pass
            return

        for brand, keywords in BRANDS:
            for attempt in range(2):
                try:
                    n = await process_brand(page, brand, keywords, sku_qty, zero_stock, modified)
                    total_scanned += n
                    break
                except Exception as e:
                    err = str(e)
                    if any(k in err.lower() for k in ("connection closed", "target closed", "closed")):
                        log(f"[WARN] Browser crashed during '{brand}' (attempt {attempt+1}): {e}")
                        try:
                            await browser.close()
                        except Exception:
                            pass
                        if attempt == 0:
                            log("[INFO] Restarting browser and re-logging in…")
                            try:
                                browser, page = await _new_page(pw, win_proxy)
                                await do_login(page)
                            except Exception:
                                log(f"[ERROR] Re-login failed: {traceback.format_exc()}")
                                break
                        else:
                            log(f"[ERROR] '{brand}' still failing after restart, skipping.")
                    else:
                        log(f"[ERROR] '{brand}': {traceback.format_exc()}")
                        break

        try:
            await browser.close()
        except Exception:
            pass

    # ── Build notification ─────────────────────────────────────────────────────
    if zero_stock:
        zero_lines = "\n".join(zero_stock)
    else:
        zero_lines = "无需关注的产品"

    if modified:
        change_lines = "\n".join(modified)
    else:
        change_lines = "今日无改动"

    report = (
        f"⚠️ 库存为0且未开销售检查的产品：\n{zero_lines}\n\n"
        f"🔧 今日改动：\n{change_lines}\n\n"
        f"📊 共扫描 {total_scanned} 个产品"
    )

    send_telegram(report)
    log(f"Done. 共扫描 {total_scanned} 个产品. {len(zero_stock)} 件库存为0, {len(modified)} 项改动. Report sent.")
    log("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())

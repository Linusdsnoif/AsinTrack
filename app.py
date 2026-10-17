# Before running, ensure you install the stealth library:
# pip install playwright-stealth

import sqlite3
import os
import re
import time
import json
import asyncio
import requests
import nest_asyncio
import pandas as pd
from datetime import datetime
from io import BytesIO
from flask import Flask, request, send_file, jsonify
from pyvirtualdisplay import Display
from playwright.async_api import async_playwright
from playwright_stealth import Stealth
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from database import init_db, log_to_database
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker

# Allow async execution inside standard Flask routes
nest_asyncio.apply()

app = Flask(__name__)

# ==========================================
# 1. Credentials & Configuration
# ==========================================
# NOTE: move these to environment variables (os.environ) before sharing/deploying
# this file anywhere. Hardcoded LWA secrets + refresh tokens are a real exposure risk.
CLIENT_SECRET = os.environ.get("SP_API_CLIENT_SECRET")
CLIENT_ID = os.environ.get("SP_API_CLIENT_ID")
LWA_URL = "https://api.amazon.com/auth/o2/token"

CONFIG_JSON = {
  "sp_api_configurations": [
    {
      "account_name": "BundWest",
      "region": "Far East (FE)",
      "endpoint": "https://sellingpartnerapi-fe.amazon.com",
      "aws_region": "us-west-2",
      "refresh_token": "Atzr|IwEBIEMWwCh_GivH8L_felQ857L8hRUiE0dvK79QyAIxhmYHRdNf0mKD2Tn2T6M5RM_eZdDEWkBXEgRDxR4sKLOn1GmsDMNnbTXRhi7vp007_9H0HDF24c_RuN6iv3yxo7Mp9Jj2liqCGxfjm1Jy5ftO-3MqLriNzzAW-YJ_ud7KyvXjPmXSSf4k4SR1YakLG049iuvrwpuDIFVF0jnRgfPC1TgvpQj1U65XDk6JJiTFVCOBEcFqZ0I0W_j3ciyNscS72dMG7C9xyJIXMhIKPUxzkPqohSaBLtiEHsv2D9qkhCxRV8oGsbO5y5bm6b15iiJL0dTg-tEMcGi_-8fDEIVfGjtC",
      "marketplaces": {
        "Australia": "A39IBJ37TRP1C6"
      }
    },
    {
      "account_name": "BundWest",
      "region": "Europe (EU)",
      "endpoint": "https://sellingpartnerapi-eu.amazon.com",
      "aws_region": "eu-west-1",
      "refresh_token": "Atzr|IwEBIFq40drbKdAbotvgQsqIFoXtv9fA6qz_rUnfuiMwfL8ZCb8PsGqG1ddt4NjiAOyVBXcXpDACZpX7qMmvW1WsDiO2Bhu35Z5wafXhhFlIqlnFgPGqGaMDZik2MNS0Ddy5doUZ2jPggyLJ5dN1eMYL3s0KUyWLP4--TO4RciE9B-RPRgeZ8AyJYyPy5FezybEtQu2NolWc26didKweovoa3k14F0a_SJA2tSCefFW-fDjx-W8SlCcoutR68AJbuQNaoiDkgSRLaru99w6zQqE2oUSa9dWKPTZFAD8G7Xh7mmSkvQWG4jGim0-PPwsikZvgIl19RU1S6jaTpY1T0bTUriQt",
      "marketplaces": {
        "UAE": "A2VIGQ35RCS4UG"
      }
    },
    {
      "account_name": "BundWest",
      "region": "Europe (EU)",
      "endpoint": "https://sellingpartnerapi-eu.amazon.com",
      "aws_region": "eu-west-1",
      "refresh_token": "Atzr|IwEBIDJhNGj8yHTEkIM7kjrfH7U7Py-U9cCZPdiMl4H_mWgT9BFcN93tquTj3b9-iNkIIN3YFkDXyVUUIGtOJl3jZK99bYhcPz6egrLHKX6fmcujhZ8z2b5K3Q7J8ArT_0xmJTHN_VsWy2B3vTDjhYQI8aNKRG4GUSzt4gKkCHcsazfRuJmbZkRseF4eIiVFrfZ-bnS-w5sFMIMRgymlp9Od6pE9tjXLMdfIGig3lQzWSmDKYRMNDFlL2MDl47puf0L2oyBd6yw1kehVHQMSQENVcjV8tXiKgDJwO7rshYP6z7WQWUPcJFSHtip8Vb_SNji70i_FThUmv7XoOPxnwxrjYPY4",
      "marketplaces": {
        "Saudi Arabia": "A17E79C6D8DWNP"
      }
    },
    {
      "account_name": "BundWest",
      "region": "Europe (EU - Pan-EU Group)",
      "endpoint": "https://sellingpartnerapi-eu.amazon.com",
      "aws_region": "eu-west-1",
      "refresh_token": "Atzr|IwEBIEPGn8_nHhCWKEvLJpkcyuwHMR2M1LXqVv65OxnEMpOsOo_zZqaUA1sH4s3iVc9Y9PnUBD1dSdo6UxmtR8p92JRweDokqC7StlR_WmJxMAPJ-tI3VS02wI7J1FnTHWRenHNbnZ2IfxZKAvaU9YxKfwuixq-nWRltqnol9RKnEUYT3cM_8C7mEmuWbdeb475xbCXiNy7r9o8cJ6WU9qgf5i0suMoqCz07KFuyQfMjrbwmGrztK0S8s5TJPwIERwGhVF0AbtoaXb0VDFDsi3-o-3I-A2Il-mijW7pjsZmz6uD8N55kM7KqIT5EM-KK1ucPgNAMNeuvlza0S8M6Ea1CsLXX",
      "marketplaces": {
        "United Kingdom": "A1F83G8C2ARO7P",
        "Germany": "A1PA6795UKMFR9",
        "France": "A13V1IB3VIYZZH",
        "Italy": "APJ6JRA9NG5V4",
        "Spain": "A1RKKUPIHCS9HS",
        "Netherlands": "A1805IZSGTT6HS",
        "Sweden": "A2NODRKZP88ZB9",
        "Poland": "A1C3SOZRARQ6R3",
        "Belgium": "AMEN7PMS3EDWL",
        "Ireland": "A1VC38T7YXB528"
      }
    },
    {
      "account_name": "LUFFCABO",
      "region": "North America (NA)",
      "endpoint": "https://sellingpartnerapi-na.amazon.com",
      "aws_region": "us-east-1",
      "refresh_token": "Atzr|IwEBIM5gyn0eF29KHBDaZThzkJ9JA0jnfiRuQBgq6OmC2bgGWuzYmauKALiZu7sz2t6GaFxrKvf-LHsk6LBN7FBOCI9u1pR2TEoLiXSjOgnjVdIk6Kihamle8sIX1mKUxCPoDVhzCrW5k1ZRwmrRkYUoPFU2DkPU50Hui7IRhv-SO-7SSOy-GVhNEVSWddfyhB53DTQKgByvME3zovDKmT1oboYFc96MDQC6Tps8vXPKg4WzzYNSnUluLW_X35ulx67k1SmF2KT-Fikgfuo0iI_jrwwwDFzwL0rUqhQ_Fk0zbdPilgA83L65H9Zp2rgfWEMZDz1o3TKiLaXBRBPYYs4aO_GE",
      "marketplaces": {
        "United States": "ATVPDKIKX0DER",
        "Canada": "A2EUQ1WTGCTBG2",
        "Mexico": "A1AM78C64UM0Y8",
        "Brazil": "A2Q3Y263D00KWC"
      }
    }
  ]
}

def get_access_token(refresh_token):
    payload = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(LWA_URL, data=payload, headers=headers)
    response.raise_for_status()
    return response.json().get("access_token")

def get_target_config(account_name, marketplace_name):
    for config in CONFIG_JSON["sp_api_configurations"]:
        if config["account_name"] == account_name and marketplace_name in config["marketplaces"]:
            return config, config["marketplaces"][marketplace_name]
    return None, None

def get_amazon_domain(marketplace_name):
    domains = {
        "United States": "amazon.com",
        "Canada": "amazon.ca",
        "Mexico": "amazon.com.mx",
        "Brazil": "amazon.com.br",
        "United Kingdom": "amazon.co.uk",
        "Germany": "amazon.de",
        "France": "amazon.fr",
        "Italy": "amazon.it",
        "Spain": "amazon.es",
        "Netherlands": "amazon.nl",
        "Sweden": "amazon.se",
        "Poland": "amazon.pl",
        "Belgium": "amazon.com.be",
        "Ireland": "amazon.co.uk",
        "Australia": "amazon.com.au",
        "UAE": "amazon.ae",
        "Saudi Arabia": "amazon.sa"
    }
    # Defaults to US (.com) if the marketplace is missing from the dictionary
    return domains.get(marketplace_name, "amazon.com")

def get_locale(marketplace_name):
    locales = {
        "United States": "en-US", "Canada": "en-CA", "Mexico": "es-MX", "Brazil": "pt-BR",
        "United Kingdom": "en-GB", "Germany": "de-DE", "France": "fr-FR", "Italy": "it-IT",
        "Spain": "es-ES", "Netherlands": "nl-NL", "Sweden": "sv-SE", "Poland": "pl-PL",
        "Belgium": "fr-BE", "Ireland": "en-IE", "Australia": "en-AU",
        "UAE": "en-AE", "Saudi Arabia": "en-SA"
    }
    return locales.get(marketplace_name, "en-US")

def get_timezone(marketplace_name):
    timezones = {
        "United States": "America/New_York", "Canada": "America/Toronto", 
        "Mexico": "America/Mexico_City", "Brazil": "America/Sao_Paulo",
        "United Kingdom": "Europe/London", "Germany": "Europe/Berlin", 
        "France": "Europe/Paris", "Italy": "Europe/Rome",
        "Spain": "Europe/Madrid", "Netherlands": "Europe/Amsterdam", 
        "Sweden": "Europe/Stockholm", "Poland": "Europe/Warsaw",
        "Belgium": "Europe/Brussels", "Ireland": "Europe/Dublin", 
        "Australia": "Australia/Sydney", "UAE": "Asia/Dubai", 
        "Saudi Arabia": "Asia/Riyadh"
    }
    return timezones.get(marketplace_name, "America/New_York")

def get_currency_symbol(marketplace_name):
    currencies = {
        "United States": "$", "Canada": "CA$", "Mexico": "MX$", "Brazil": "R$",
        "United Kingdom": "£", "Germany": "€", "France": "€", "Italy": "€",
        "Spain": "€", "Netherlands": "€", "Sweden": "SEK ", "Poland": "PLN ",
        "Belgium": "€", "Ireland": "€", "Australia": "A$",
        "UAE": "AED ", "Saudi Arabia": "SAR "
    }
    return currencies.get(marketplace_name, "")



# ==========================================
# 1b. Price normalization helpers
# ==========================================
def parse_price_to_float(raw_text):
    if not raw_text or raw_text in ("N/A", "Error", "CAPTCHA Block"):
        return None

    cleaned = re.sub(r"[^\d.,]", "", raw_text)
    if not cleaned:
        return None

    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        parts = cleaned.split(",")
        if len(parts[-1]) == 2:
            cleaned = "".join(parts[:-1]) + "." + parts[-1]
        else:
            cleaned = cleaned.replace(",", "")

    try:
        return round(float(cleaned), 2)
    except ValueError:
        return None


def extract_currency(raw_text):
    if not raw_text or raw_text in ("N/A", "Error", "CAPTCHA Block"):
        return ""

    symbol_match = re.match(r"^\s*([^\d\s]+)", raw_text)
    if symbol_match:
        return symbol_match.group(1).strip()

    code_match = re.search(r"\b([A-Z]{3})\b", raw_text)
    if code_match:
        return code_match.group(1)

    return ""


def calculate_savings_percent(current_price, listing_price):
    if current_price is None or listing_price is None:
        return 0
    if listing_price <= 0 or current_price >= listing_price:
        return 0
    return round((1 - (current_price / listing_price)) * 100)


def get_sp_api_prices(endpoint, asin, marketplace_id, access_token, max_retries=4):
    """
    Checks SP-API Pricing Offers. Returns:
    (current_price, currency_code, state)
    state: "in_stock" | "out_of_stock" | "api_error"
    """
    headers = {"x-amz-access-token": access_token}
    price_url = f"{endpoint}/products/pricing/v0/items/{asin}/offers"

    for attempt in range(1, max_retries + 1):
        try:
            res = requests.get(
                price_url, headers=headers,
                params={"MarketplaceId": marketplace_id, "ItemCondition": "New"},
                timeout=15
            )
        except requests.RequestException as e:
            print(f"  -> [SP-API] {asin} pricing request failed ({e}), attempt {attempt}/{max_retries}")
            time.sleep(2 * attempt)
            continue

        if res.status_code == 429:
            wait = 2 * attempt
            print(f"  -> [SP-API] {asin} throttled (429). Waiting {wait}s, attempt {attempt}/{max_retries}")
            time.sleep(wait)
            continue

        if res.status_code != 200:
            print(f"  -> [SP-API] {asin} pricing HTTP {res.status_code}: {res.text[:200]}")
            return None, "", "api_error"

        payload = res.json().get("payload", {})
        offers = payload.get("Offers", [])
        summary = payload.get("Summary", {})

        # Check total offer count across all sellers
        total_offers = summary.get("TotalOfferCount", len(offers))

        if total_offers == 0 or not offers:
            if attempt == 1:
                # Amazon sometimes returns an empty offer list on the first call; check once more
                print(f"  -> [SP-API] {asin} returned no offers. Re-checking in 2s...")
                time.sleep(2)
                continue
            print(f"  -> [SP-API] {asin} no offers: status={payload.get('status')!r} "
                  f"summary={json.dumps(summary)[:300]}")
            return None, "", "out_of_stock"

        priced_offers = []
        currency_code = ""
        for o in offers:
            listing_price = o.get("ListingPrice", {}).get("Amount")
            currency_code = currency_code or o.get("ListingPrice", {}).get("CurrencyCode", "")
            shipping = o.get("Shipping", {}).get("Amount", 0) or 0
            if listing_price is not None:
                priced_offers.append(float(listing_price) + float(shipping))

        if not priced_offers:
            return None, currency_code, "out_of_stock"
        return min(priced_offers), currency_code, "in_stock"

    return None, "", "api_error"


# ==========================================
# 2. Async Scraper & Processing Logic
# ==========================================
INTERSTITIAL_SELECTORS = [
    # Global "Dismiss" for the international shipping & location popups
    "input[data-action-type='DISMISS']",
    ".glow-toaster-button-dismiss input",
    "input.a-button-input[aria-labelledby='glow-toaster-button-dismiss']",
    
    # Structural selectors for Amazon's soft-CAPTCHA
    "form[action*='validateCaptcha'] .a-button-primary input[type='submit']",
    "form[action*='validateCaptcha'] .a-button-primary button",
    
    # Substring text selectors
    "text=Continue shopping", "text=Weiter shoppen", "text=Continuer les achats",
    "text=Continuar comprando", "text=Continua a fare acquisti", "text=Verder winkelen",
    "text=Fortsätt handla", "text=Kontynuuj zakupy", "text=متابعة التسوق", "text=استمر في التسوق",
    
    # GDPR & Cookie Consent buttons
    "#sp-cc-accept", "input[name='accept']", "#glow-toaster-cell-content button",  
    "a#nav-cart-flyover-close-icon", "[data-testid='cookie-accept-all']", "[data-testid='consent-accept-all']",

    # NOTE: generic popover closers (.a-popover-close / .a-button-close) were removed from
    # this list on purpose. The JS block in dismiss_interstitials already closes *visible*
    # popovers, and a force-click on generic selectors can hit random elements on the
    # search results page.
]

COOKIE_ACCEPT_PHRASES = [
    "accept all", "accept cookies", "i accept", "accept",
    "alle akzeptieren", "akzeptieren",
    "tout accepter", "accepter",
    "aceptar todo", "aceptar",
    "accetta tutto", "accetta",
    "alles accepteren", "accepteren",
    "zaakceptuj wszystko", "zaakceptuj",
    "godkänn alla", "godkänn",
    "aceitar todos", "aceitar",
    "موافقة", "قبول",
    "متابعة التسوق", # Added regex fallback for Arabic soft-CAPTCHA
    "استمر",
]

CURRENT_PRICE_SELECTORS = [
    '#corePriceDisplay_desktop_feature_div .priceToPay .a-offscreen',
    '#corePrice_desktop .priceToPay .a-offscreen',
    '#apexPriceToPay .a-offscreen',
    '#corePriceDisplay_desktop_feature_div .a-price.a-text-normal .a-offscreen',
    '#buyNewSection .a-price .a-offscreen',
    'div[data-csa-c-buying-option-type="NEW"] .a-price .a-offscreen',
    '#desktop_buybox .a-price .a-offscreen',
    '#price_inside_buybox',
    '#newBuyBoxPrice',
    '#tp_price_block_total_price_ww .a-offscreen',
    '.a-price.priceToPay .a-offscreen',
    '#corePrice_feature_div .a-price .a-offscreen',
]

LIST_PRICE_SELECTORS = [
    '#corePriceDisplay_desktop_feature_div .basisPrice .a-offscreen',
    '#corePrice_desktop .basisPrice .a-offscreen',
    '#corePriceDisplay_desktop_feature_div span.a-text-price[data-a-strike="true"] .a-offscreen',
    '#listPrice_feature_div .a-price .a-offscreen',
    'span.a-price.a-text-price[data-a-strike="true"] .a-offscreen',
    '.basisPrice .a-price.a-text-price .a-offscreen',
    '#priceblock_ourprice_row .a-text-strike',
]


def parse_rating(raw_text):
    if not raw_text:
        return "N/A"
    match = re.search(r"(\d+[.,]\d+|\d+)", raw_text)
    if not match:
        return "N/A"
    try:
        return f"{float(match.group(1).replace(',', '.')):.1f}"
    except ValueError:
        return "N/A"


async def dismiss_interstitials(page):
    frames = [page] + list(page.frames)

    for frame in frames:
        # 1. JS POPOVER & CAPTCHA NUKER
        # Executes directly in the browser to bypass Playwright visibility/overlay traps
        try:
            captcha_cleared = await frame.evaluate('''() => {
                // A. Destroy Language/Location Modals (Clicks the 'X')
                const popoverCloseBtns = document.querySelectorAll("button[data-action='a-popover-close'], .a-popover-close");
                for (let btn of popoverCloseBtns) {
                    if (btn.offsetParent !== null) { btn.click(); }
                }
                
                // B. Bypass CAPTCHA
                const forms = document.querySelectorAll('form');
                for (let f of forms) {
                    if (f.action && f.action.includes('validateCaptcha')) {
                        const btn = f.querySelector('button, input[type="submit"], input.a-button-input');
                        if (btn) { btn.click(); return true; }
                        f.submit(); return true;
                    }
                }
                return false;
            }''')
            if captcha_cleared:
                await page.wait_for_timeout(3000)
        except Exception:
            pass

        # 2. Native Playwright Fallbacks (using force=True to ignore overlapping UI elements)
        for sel in INTERSTITIAL_SELECTORS:
            try:
                el = await frame.query_selector(sel)
                if el:
                    await el.click(timeout=1500, force=True)
                    await page.wait_for_timeout(800)
            except Exception:
                pass

        for phrase in COOKIE_ACCEPT_PHRASES:
            try:
                btn = frame.get_by_role("button", name=re.compile(re.escape(phrase), re.IGNORECASE))
                if await btn.count() > 0:
                    await btn.first.click(timeout=1500, force=True)
                    await page.wait_for_timeout(800)
                    break
            except Exception:
                pass


# ------------------------------------------------------------------
# 2a. Deterministic product navigation helpers
# ------------------------------------------------------------------
ASIN_URL_RE = re.compile(r"/(?:dp|gp/product|gp/aw/d)/([A-Z0-9]{10})", re.IGNORECASE)


def asin_from_url(url):
    m = ASIN_URL_RE.search(url or "")
    return m.group(1).upper() if m else None


async def is_expected_product_page(page, asin, timeout=5000):
    """
    True only if the current page is a product page AND it is the product
    for `asin`. Waits for #productTitle instead of sampling the DOM once,
    so we don't judge a page that is still loading.
    """
    asin = asin.upper()
    try:
        await page.wait_for_selector(PRODUCT_TITLE_SELECTOR, timeout=timeout)
    except Exception:
        return False

    if asin_from_url(page.url) == asin:
        return True

    # Fallback: Amazon's hidden input holds the ASIN currently displayed
    try:
        shown = await page.evaluate(
            """() => {
                const i = document.querySelector('input#ASIN, input[name="ASIN"]');
                return i ? i.value : '';
            }"""
        )
        return (shown or "").upper() == asin
    except Exception:
        return False


PRODUCT_TITLE_SELECTOR = "#productTitle, h1#title, #titleSection"

# Amazon's "page not found" (dog) page is served at the requested /dp/ URL with HTTP 404 and
# contains these markers. It also shows recommendation carousels, which is why it can look
# like a product page in a thumbnail.
NOT_FOUND_JS = r"""() => !!document.querySelector('a[href*="cs_404"], img[alt*="Dogs of Amazon" i]')"""


async def looks_like_404(page):
    try:
        return bool(await page.evaluate(NOT_FOUND_JS))
    except Exception:
        return False


async def get_delivery_label(page):
    """Text of the 'Deliver to ...' widget, e.g. '1000 Bruxelles'. Empty if not found."""
    try:
        el = await page.query_selector("#glow-ingress-line2")
        if el:
            return (await el.inner_text()).strip()
    except Exception:
        pass
    return ""


# Runs in the page: which key product elements exist / are visible, what hides the title,
# any alert messages, and the page text with navigation noise removed.
PAGE_REPORT_JS = r"""() => {
  const out = {};
  const ids = ['dp','dp-container','ppd','centerCol','leftCol','imgTagWrapperId','productTitle','title',
               'titleSection','desktop_buybox','buybox','add-to-cart-button','availability','outOfStock',
               'GLUXModal','glow-ingress-line2'];
  for (const id of ids) {
    const el = document.getElementById(id);
    if (!el) { out[id] = 'missing'; continue; }
    const r = el.getBoundingClientRect(), cs = getComputedStyle(el);
    out[id] = Math.round(r.width) + 'x' + Math.round(r.height) + ' display=' + cs.display + ' vis=' + cs.visibility;
  }
  let hidden = null;
  const t = document.getElementById('productTitle');
  if (t) {
    for (let n = t; n && n !== document.documentElement; n = n.parentElement) {
      const cs = getComputedStyle(n);
      if (cs.display === 'none' || cs.visibility === 'hidden') {
        hidden = (n.tagName + (n.id ? '#' + n.id : '') +
                  (n.className ? '.' + String(n.className).trim().split(/\s+/).join('.') : '')).slice(0, 140) +
                 ' display=' + cs.display + ' vis=' + cs.visibility;
        break;
      }
    }
  }
  out.hiddenAncestorOfProductTitle = hidden;
  const texts = (sel) => Array.from(document.querySelectorAll(sel))
      .map(e => (e.innerText || '').replace(/\s+/g, ' ').trim()).filter(Boolean).slice(0, 5);
  out.alerts = texts('.a-alert-content, [role="alert"], #g-unrecoverable-error');
  out.headings = texts('h1, h2').map(s => s.slice(0, 80));
  const main = document.getElementById('a-page') || document.body;
  const clone = main.cloneNode(true);
  clone.querySelectorAll('#navbar, #nav-main, header, script, style, noscript, [id^="nav-"], #a-popover-root')
       .forEach(e => e.remove());
  out.pageTextStart = (clone.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 400);
  return out;
}"""


async def dump_debug(page, asin, reason, last_doc, screenshot_bytes=None):
    """
    Print evidence for a failed ASIN and save the full-size screenshot + HTML to ./debug/.
    (The Excel thumbnail is only 150x100, which hides what the page really is.)
    """
    try:
        title = await page.title()
    except Exception:
        title = ""
    try:
        report = await page.evaluate(PAGE_REPORT_JS)
    except Exception as e:
        report = {"report_error": str(e)}
    delivery = await get_delivery_label(page)
    print(f"  -> [DEBUG] {asin} reason={reason} http={last_doc.get('status')} url={page.url}")
    print(f"  -> [DEBUG] {asin} title={title!r} delivery_to={delivery!r}")
    print(f"  -> [DEBUG] {asin} page_report={json.dumps(report, ensure_ascii=False)}")
    try:
        os.makedirs("debug", exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join("debug", f"{asin}_{stamp}_{reason}")
        with open(base + ".html", "w", encoding="utf-8") as f:
            f.write(await page.content())
        if screenshot_bytes:
            with open(base + ".png", "wb") as f:
                f.write(screenshot_bytes)
    except Exception:
        pass


async def clear_overlays(page):
    """Remove invisible shields that intercept clicks/keystrokes."""
    try:
        await page.evaluate('''() => {
            document.querySelectorAll('.a-popover-lgtbox, #nav-cover, #a-popover-root, .glow-toaster, #sp-cc').forEach(el => el.remove());
            document.body.style.overflow = 'auto';
        }''')
    except Exception:
        pass


# Runs inside the page. Finds the ORGANIC (non-sponsored) result card for the ASIN and
# tags its product link with data-scraper-target="1".
# Returns "organic" | "sponsored_only" | "none".
FIND_ORGANIC_LINK_JS = r'''(asin) => {
    document.querySelectorAll('[data-scraper-target]')
        .forEach(e => e.removeAttribute('data-scraper-target'));

    // Only real search-result cards for THIS asin (never carousels / other widgets)
    let cards = Array.from(document.querySelectorAll(
        `div[data-component-type="s-search-result"][data-asin="${asin}"]`));
    if (cards.length === 0) {
        cards = Array.from(document.querySelectorAll(`div.s-result-item[data-asin="${asin}"]`));
    }
    if (cards.length === 0) return "none";

    // Structural (language-independent) sponsored markers
    const SPONSORED = '.puis-sponsored-label-text, .s-sponsored-label-info-icon, ' +
                      '[data-component-type="sp-sponsored-result"], a[href*="/sspa/"]';
    const isSponsored = (c) => c.classList.contains('AdHolder') || !!c.querySelector(SPONSORED);
    const visible = (a) => a.getClientRects().length > 0;
    const needle = `/dp/${asin}`;

    let sawSponsored = false;
    for (const card of cards) {
        if (isSponsored(card)) { sawSponsored = true; continue; }

        const anchors = Array.from(card.querySelectorAll('a[href]')).filter(a => {
            const h = a.getAttribute('href') || '';
            return h.includes(needle) && !h.includes('/sspa/') && visible(a);
        });
        if (anchors.length === 0) continue;

        // Fixed preference order: title link > image link > anything else
        const pick = anchors.find(a => a.closest('h2') || a.querySelector('h2'))
                  || anchors.find(a => a.classList.contains('s-no-outline'))
                  || anchors[0];
        pick.setAttribute('data-scraper-target', '1');
        return "organic";
    }
    return sawSponsored ? "sponsored_only" : "none";
}'''


async def locate_organic_link(page, asin, max_wait_s=15):
    """
    Poll until the organic card for `asin` exists (results render lazily; sponsored
    cards usually appear first), scrolling to trigger lazy loading. Returns a locator
    for the tagged link, or None.
    """
    asin = asin.upper()
    deadline = time.monotonic() + max_wait_s
    state = "none"
    while True:
        try:
            state = await page.evaluate(FIND_ORGANIC_LINK_JS, asin)
        except Exception:
            state = "none"
        if state == "organic":
            return page.locator("a[data-scraper-target='1']").first
        if time.monotonic() >= deadline:
            break
        await page.mouse.wheel(0, 600)
        await page.wait_for_timeout(700)

    if state == "sponsored_only":
        print(f"  -> {asin} only appears as a SPONSORED result. Not clicking ads.")
    else:
        print(f"  -> {asin} not found on the search results page.")
    return None


async def open_via_search(page, asin, domain):
    """
    Organic-search backdoor. Search the ASIN once, click its organic (non-sponsored)
    result, and verify we landed on the right product. If the click route fails, fall
    back once to the canonical URL with the search page as referer. Returns True only
    on a verified ASIN match.
    """
    asin = asin.upper()
    await clear_overlays(page)

    box = await page.query_selector("#twotabsearchtextbox")
    if not box:
        print(f"  -> No search box on page for {asin}.")
        return False

    await box.click()
    await box.fill("")
    await box.type(asin, delay=120)  # keystroke events, like a human
    await page.wait_for_timeout(500)

    # Submit exactly once and wait for the navigation this action triggers
    try:
        async with page.expect_navigation(wait_until="domcontentloaded", timeout=20000):
            btn = await page.query_selector("#nav-search-submit-button")
            if btn:
                await btn.click()
            else:
                await box.press("Enter")
    except Exception as e:
        print(f"  -> Search submit navigation warning: {e}")

    await dismiss_interstitials(page)

    # Amazon sometimes jumps straight to the product on an exact-ASIN search
    if await is_expected_product_page(page, asin, timeout=2000):
        return True

    # Make sure a results page (not a captcha / error page) is actually there
    try:
        await page.wait_for_selector(
            "div.s-main-slot, div[data-component-type='s-search-result']", timeout=15000
        )
    except Exception:
        print(f"  -> Search results never rendered for {asin}.")

    search_url = page.url
    link = await locate_organic_link(page, asin)

    if link is not None:
        try:
            await link.scroll_into_view_if_needed(timeout=5000)
            await link.evaluate("el => el.removeAttribute('target')")  # keep same tab
            await link.click(timeout=10000)
            await page.wait_for_url(
                re.compile(rf"/dp/{asin}", re.IGNORECASE),
                timeout=30000,
                wait_until="domcontentloaded",
            )
        except Exception as e:
            print(f"  -> Click navigation warning for {asin}: {e}")

        await dismiss_interstitials(page)
        if await is_expected_product_page(page, asin, timeout=8000):
            return True
        if asin_from_url(page.url) == asin:
            print(f"  -> Landed on the right URL for {asin}, but the product block (title/gallery) is not visible.")
            return False
        print(f"  -> Click did not land on {asin}.")

    # One deterministic fallback: canonical URL, with the search page as referer
    print(f"  -> Falling back to canonical URL with search referer for {asin}...")
    try:
        await page.goto(
            f"https://www.{domain}/dp/{asin}",
            referer=search_url,
            timeout=30000,
            wait_until="domcontentloaded",
        )
    except Exception as e:
        print(f"  -> Fallback navigation warning: {e}")

    await dismiss_interstitials(page)
    return await is_expected_product_page(page, asin, timeout=8000)


# ------------------------------------------------------------------
# 2b. Front-end scraper
# ------------------------------------------------------------------
DELIVERY_POSTCODES = {
    "United States": "10001", "Canada": "M5V 2H1", "Mexico": "06000", "Brazil": "01000-000",
    "United Kingdom": "E1 6AN", "Germany": "10115", "France": "75001", "Italy": "00118",
    "Spain": "28001", "Netherlands": "1012", "Sweden": "11120", "Poland": "00-001",
    "Belgium": "1000", "Ireland": "D01 F5P2", "Australia": "2000"
}


async def apply_delivery_location(page, marketplace_name):
    """
    Set the delivery location with robust JS injection to bypass Playwright visibility 
    and interception traps, then VERIFY that the header's 'Deliver to' label changed.
    """
    before = await get_delivery_label(page)
    postcode = DELIVERY_POSTCODES.get(marketplace_name)
    steps = []
    
    try:
        # 1. Open the location modal (JavaScript click is most reliable here)
        await page.evaluate("""() => {
            const b = document.querySelector("#nav-global-location-popover-link") ||
                      document.querySelector("#nav-global-location-slot");
            if (b) b.click();
        }""")
        steps.append("opened")
        
        # Wait for modal animation to finish fetching the form
        await page.wait_for_timeout(3000)

        if postcode:
            # 2. Inject the postcode and click Apply using browser-native JS
            # This handles single-input (US/EU) and dual-input (CA/UK) layouts effortlessly
            success = await page.evaluate(f"""(zip) => {{
                const zip1 = document.querySelector("#GLUXZipUpdateInput_0");
                const zip2 = document.querySelector("#GLUXZipUpdateInput_1");
                const zipSingle = document.querySelector("#GLUXZipUpdateInput");

                // Fill the input(s)
                if (zip1 && zip2) {{
                    let parts = zip.split(" ");
                    if (parts.length < 2) parts = [zip.substring(0, 3), zip.substring(3)];
                    zip1.value = parts[0];
                    zip2.value = parts[1] || "";
                    zip1.dispatchEvent(new Event('input', {{ bubbles: true }}));
                    zip2.dispatchEvent(new Event('input', {{ bubbles: true }}));
                }} else if (zipSingle) {{
                    zipSingle.value = zip;
                    zipSingle.dispatchEvent(new Event('input', {{ bubbles: true }}));
                }} else {{
                    return false; // Inputs not found in the DOM
                }}
                
                // Click the 'Apply' button
                const applyBtn = document.querySelector("#GLUXZipUpdate input, input[aria-labelledby='GLUXZipUpdate-announce']");
                if (applyBtn) {{
                    applyBtn.click();
                    return true;
                }}
                return false;
            }}""", postcode)
            
            if success:
                steps.append("typed_and_applied")
            else:
                steps.append("input_not_found")
                
            await page.wait_for_timeout(2500)

        # 3. Click the final "Done" or "Continue" button
        await page.evaluate("""() => {
            const doneBtns = document.querySelectorAll("button[name='glowDoneButton'], #GLUXConfirmClose, .a-popover-footer input");
            for (let btn of doneBtns) {
                if (btn.offsetParent !== null) { // if visible
                    btn.click();
                    return;
                }
            }
        }""")
        steps.append("confirm_clicked")
        
        # Wait for modal to close and page to register the change
        await page.wait_for_timeout(3000)

    except Exception as e:
        steps.append(f"error:{str(e)[:80]}")

    # 4. Verify the label changed
    after = ""
    for _ in range(8):
        after = await get_delivery_label(page)
        if after and after != before:
            break
        await page.wait_for_timeout(1000)
        
    ok = bool(after) and after != before
    print(f"  -> Delivery location [{marketplace_name}] steps={steps} before={before!r} after={after!r} ok={ok}")
    return ok

async def _scrape_amazon_frontend_once(asin, browser, marketplace_name, domain, locale="en-US", timezone_id="America/New_York", storage_state=None, warm_up=False):
    url = f"https://www.{domain}/dp/{asin}"
    
    context_kwargs = {
        "user_agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "viewport": {"width": 1280, "height": 1080},
        "locale": locale,
        "timezone_id": timezone_id,
        "extra_http_headers": {"Accept-Language": f"{locale},en;q=0.8"},
    }
    if storage_state:
        context_kwargs["storage_state"] = storage_state

    context = await browser.new_context(**context_kwargs)
    
    await context.add_cookies([
        {"name": "sp-cc", "value": "ACCEPT", "domain": f".{domain}", "path": "/"}
    ])

    await context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page = await context.new_page()
    await Stealth().apply_stealth_async(page)

    # Remember the HTTP status of the last top-level document (404 = Amazon "page not found")
    last_doc = {"status": None, "url": None}

    def _track_response(response):
        try:
            if response.request.resource_type == "document" and response.frame == page.main_frame:
                last_doc["status"] = response.status
                last_doc["url"] = response.url
        except Exception:
            pass

    page.on("response", _track_response)
    updated_storage_state = storage_state

    async def check_if_product_page():
        # Product page AND the right ASIN (previously: any /dp/ page passed)
        return await is_expected_product_page(page, asin)

    async def set_local_delivery(marketplace_name):
        return await apply_delivery_location(page, marketplace_name)

    try:
        if warm_up:
            # Fresh-session retry: behave like a person. Home page -> set the delivery
            # location up front -> search -> click. No cold direct /dp/ hit.
            print(f"  -> Fresh-session retry for {asin}: warm-up on {domain}, then search path...")
            await page.goto(f"https://www.{domain}/", timeout=60000, wait_until="domcontentloaded")
            await page.wait_for_timeout(3500)
            await dismiss_interstitials(page)
            await set_local_delivery(marketplace_name)
            print(f"  -> Delivery location now: {await get_delivery_label(page)!r}")
            product_loaded = await open_via_search(page, asin, domain)
        else:
            # 1. INITIAL DIRECT LOAD
            await page.goto(url, timeout=60000, wait_until="domcontentloaded")
            await page.wait_for_timeout(3500)
            await dismiss_interstitials(page)
            
            # 2. ALWAYS enforce local delivery location (No longer waiting for a block)
            print(f"  -> Actively enforcing local delivery routing for {asin}...")
            await set_local_delivery(marketplace_name)
            print(f"  -> Delivery location now: {await get_delivery_label(page)!r}")
            
            # Reload the page to ensure the Buy Box and availability update to reflect the new postcode
            await page.goto(f"{url}?psc=1", timeout=30000, wait_until="domcontentloaded")
            await page.wait_for_timeout(3500)
            await dismiss_interstitials(page)
            
            product_loaded = await check_if_product_page()

            # 3. ESCALATION STAGE: The Organic Search Backdoor (single, verified path)
            if not product_loaded:
                print(f"  -> Initiating Organic Search Backdoor for {asin}...")
                product_loaded = await open_via_search(page, asin, domain)

        # --- UI CLEANUP ---
        try:
            await page.keyboard.press("Escape")
            await page.wait_for_timeout(500)
            close_flyover = await page.query_selector("a#nav-cart-flyover-close-icon, .a-popover-close, button[data-action='a-popover-close'], #huc-v2-order-row-close-btn")
            if close_flyover:
                await close_flyover.click(timeout=1500)
            await page.wait_for_timeout(800) 
        except Exception:
            pass

        # 4. CAPTURE DIAGNOSTIC STATE
        screenshot_bytes = await page.screenshot()

        # 5. STRICT DIAGNOSTIC FAILURE ROUTING
        if not product_loaded:
            page_title = await page.title()
            http_status = last_doc.get("status")
            landed_asin = asin_from_url(page.url)

            is_captcha = await page.evaluate("() => !!document.querySelector('form[action*=\"validateCaptcha\"]')")
            if is_captcha or "Bot Check" in page_title or "Robot Check" in page_title or "CAPTCHA" in page_title:
                reason, label, keep_shot = "captcha", "CAPTCHA Block", False
            elif http_status == 404 or await looks_like_404(page):
                reason, label, keep_shot = "not_found", "Page Not Found", False
            elif "/dp/" not in page.url and "/gp/" not in page.url:
                reason, label, keep_shot = "redirect", "Region Redirect", False
            elif landed_asin and landed_asin != asin.upper():
                reason, label, keep_shot = "asin_mismatch", "ASIN Mismatch", False
            elif landed_asin and landed_asin == asin.upper():
                # Change keep_shot to True to retain the screenshot of the empty shell page
                reason, label, keep_shot = "block_missing", "Product Block Missing", True
            else:
                reason, label, keep_shot = "blocked", "Visual Scrape Blocked", True

            await dump_debug(page, asin, reason, last_doc, screenshot_bytes)
            return True, label, "N/A", "N/A", "N/A", "N/A", (screenshot_bytes if keep_shot else None), reason, storage_state

        # 6. VERIFY OUT OF STOCK EDGE CASES
        is_out_of_stock = False
        out_of_stock_elem = await page.query_selector("#outOfStock")
        if out_of_stock_elem:
            is_out_of_stock = True
        else:
            availability_elem = await page.query_selector("#availability")
            if availability_elem:
                text = await availability_elem.inner_text()
                if "unavailable" in text.lower() or "we don't know when" in text.lower() or "out of stock" in text.lower():
                    is_out_of_stock = True

        # 7. SUCCESS PATH METRICS EXTRACTION
        rating, reviews, coupon = "N/A", "N/A", "No Coupon"
        rating_elem = await page.query_selector("#acrPopover")
        if rating_elem:
            title_attr = await rating_elem.get_attribute("title")
            if title_attr: rating = parse_rating(title_attr)

        review_elem = await page.query_selector("#acrCustomerReviewText")
        if review_elem:
            rev_text = await review_elem.inner_text()
            if rev_text: reviews = rev_text.split(" ")[0].strip()

        coupon = await page.evaluate('''() => {
            const selectors = [
                'label[for*="coupon"]', 
                'label[id^="couponBadge"]', 
                '#vpcButton', 
                'label[for="vpcButton"]',
                '.vpc-title', 
                '[id*="voucherBadge"]', 
                '.promoPriceBlockMessage',
                'label:has(input[type="checkbox"])'
            ];
            
            for (let sel of selectors) {
                const els = document.querySelectorAll(sel);
                for (let el of els) {
                    const clone = el.cloneNode(true);
                    
                    const prices = clone.querySelectorAll('.a-price');
                    prices.forEach(p => {
                        const offscreen = p.querySelector('.a-offscreen');
                        if (offscreen) p.textContent = ' ' + offscreen.textContent + ' ';
                    });
                    
                    let text = clone.textContent.replace(/\\s+/g, ' ').trim();
                    
                    // 1. STRICT EXCLUSIONS
                    if (text.match(/geschenk|gift/i)) continue;
                    
                    // 2. REQUIRED KEYWORDS: It MUST explicitly say coupon or voucher
                    if (!text.match(/coupon|voucher/i)) continue;
                    
                    // 3. STRICT PERCENTAGE SCENARIO
                    // Requires the % to be anchored directly to the word "coupon" (e.g., "10 %-Coupon")
                    let pctStrict = text.match(/(\\d+(?:[.,]\\d+)?)\\s*%\\s*-?\\s*(?:coupon|voucher)/i);
                    if (pctStrict) return pctStrict[1] + "%";
                    
                    // Fallback for "Coupon: 10%" or "Voucher: Save 10%"
                    let pctPrefix = text.match(/(?:coupon|voucher).*?(\\d+(?:[.,]\\d+)?)\\s*%/i);
                    if (pctPrefix) return pctPrefix[1] + "%";
                    
                    // 4. STRICT FIXED PRICE SCENARIO
                    // Requires currency symbols to be anchored directly to the coupon (e.g., "$5 Coupon" or "15€ Coupon")
                    let curStrict = text.match(/(?:coupon|voucher).*?([$£€]|USD|EUR)\\s*(\\d+(?:[.,]\\d+)?)/i) ||
                                    text.match(/([$£€]|USD|EUR)\\s*(\\d+(?:[.,]\\d+)?)\\s*-?\\s*(?:coupon|voucher)/i) ||
                                    text.match(/(\\d+(?:[.,]\\d+)?)\\s*([$£€]|USD|EUR)\\s*-?\\s*(?:coupon|voucher)/i);
                    if (curStrict) return (curStrict[1] + curStrict[2]).replace(/\\s+/g, '');
                    
                    let curPrefix = text.match(/(?:coupon|voucher).*?(\\d+(?:[.,]\\d+)?)\\s*([$£€]|USD|EUR)/i);
                    if (curPrefix) return (curPrefix[1] + curPrefix[2]).replace(/\\s+/g, '');
                }
            }
            return "No Coupon";
        }''')

        price_result = await page.evaluate(
            '''(sels) => {
                const clean = (t) => t ? t.replace(/\\s+/g, ' ').trim() : "N/A";
                let cp = "N/A", cpSel = null, lp = "N/A", lpSel = null;
                for (const sel of sels.cp) {
                    const el = document.querySelector(sel);
                    if (el && el.textContent.match(/[0-9]/)) { cp = clean(el.textContent); cpSel = sel; break; }
                }
                for (const sel of sels.lp) {
                    const el = document.querySelector(sel);
                    if (el && el.textContent.match(/[0-9]/)) { lp = clean(el.textContent); lpSel = sel; break; }
                }
                if (cp === lp) { lp = "N/A"; lpSel = null; }
                if (cp === "N/A" && lp !== "N/A") { cp = lp; cpSel = lpSel; lp = "N/A"; lpSel = null; }
                return { current: cp, list: lp, current_selector: cpSel, list_selector: lpSel };
            }''',
            {"cp": CURRENT_PRICE_SELECTORS, "lp": LIST_PRICE_SELECTORS}
        )

        try:
            updated_storage_state = await context.storage_state()
        except Exception:
            pass

        selector_used = "out_of_stock" if is_out_of_stock else (price_result.get("current_selector") or "none")
        return True, rating, reviews, coupon, price_result.get("current", "N/A"), price_result.get("list", "N/A"), screenshot_bytes, selector_used, updated_storage_state

    except Exception as e:
        print(f"  -> PIPELINE ERROR for {asin}: {e}")
        return True, "Scrape Error", "Scrape Error", "Scrape Error", "Error", "Error", None, "exception", storage_state
        
    finally:
        await page.close()
        await context.close()
        
RETRYABLE_FAILURES = {"Page Not Found", "CAPTCHA Block", "Region Redirect", "ASIN Mismatch",
                      "Visual Scrape Blocked", "Scrape Error", "Product Block Missing"}


async def scrape_amazon_frontend(asin, browser, marketplace_name, domain, locale="en-US", timezone_id="America/New_York", storage_state=None):
    """
    Attempt 1: normal flow (direct load -> delivery bypass -> search backdoor), reusing the
    shared session. If that fails, attempt 2 runs ONCE in a brand-new session (no stored
    cookies/history) using the human-like path (home -> location -> search -> click).
    """
    first = await _scrape_amazon_frontend_once(
        asin, browser, marketplace_name, domain, locale, timezone_id, storage_state, warm_up=False)
    if first[1] not in RETRYABLE_FAILURES:
        return first

    print(f"  -> {asin}: attempt 1 failed ({first[1]}). Retrying once in a fresh session...")
    second = await _scrape_amazon_frontend_once(
        asin, browser, marketplace_name, domain, locale, timezone_id, None, warm_up=True)
    if second[1] in RETRYABLE_FAILURES:
        # Still failing: report it, but keep the previous shared session for the next ASIN
        return second[:8] + (storage_state,)
    return second


async def process_asin(asin, marketplace_id, marketplace_name, endpoint, access_token, browser, storage_state=None):
    domain = get_amazon_domain(marketplace_name)
    locale = get_locale(marketplace_name)
    timezone_id = get_timezone(marketplace_name)
    currency = get_currency_symbol(marketplace_name)

    row = {
        "ID": asin, "Region": marketplace_name, "Brand": "N/A", "Product Name": "N/A",
        "Product URL": f"https://www.{domain}/dp/{asin}", "Listing Price": "N/A", "Current Price": "N/A",
        "Savings": "0%", "Coupon Info": "No Coupon", "Rating": "N/A", "Review Count": "N/A",
        "Status": "Unavailable", "Product Screenshot": "Unavailable", "Ranking": "N/A",
        "Screenshot_Bytes": None, "Screenshot_Path": ""
    }

    headers = {"x-amz-access-token": access_token}
    cat_url = f"{endpoint}/catalog/2022-04-01/items/{asin}"
    try:
        cat_res = requests.get(cat_url, headers=headers, params={"marketplaceIds": marketplace_id, "includedData": "summaries,salesRanks"}, timeout=15)
        if cat_res.status_code == 200:
            item_data = cat_res.json()
            if item_data.get("summaries"):
                row["Brand"] = item_data["summaries"][0].get("brand", "")
                row["Product Name"] = item_data["summaries"][0].get("itemName", "")
            
            ranks_data = item_data.get("salesRanks", [])
            if ranks_data:
                rank_list = []
                # 1. Try to get specific sub-category ranks first (e.g., "Baby Carrier")
                if "classificationRanks" in ranks_data[0] and ranks_data[0]["classificationRanks"]:
                    rank_list = [f"#{r.get('rank')} in {r.get('title')}" for r in ranks_data[0]["classificationRanks"]]
                
                # 2. Fall back to broad display group ranks if classification is missing (e.g., "Baby")
                elif "displayGroupRanks" in ranks_data[0] and ranks_data[0]["displayGroupRanks"]:
                    rank_list = [f"#{r.get('rank')} in {r.get('title')}" for r in ranks_data[0]["displayGroupRanks"]]
                
                if rank_list:
                    row["Ranking"] = ", ".join(rank_list)
    except requests.RequestException:
        pass

    sp_current, sp_currency_code, stock_state = get_sp_api_prices(endpoint, asin, marketplace_id, access_token)

    if stock_state == "out_of_stock":
        print(f"  -> [SP-API] {asin} is OUT OF STOCK in {marketplace_name}. Skipping browser scrape.")
        row["Status"] = "Unavailable"
        return row, storage_state

    if stock_state == "api_error":
        # Don't trust a failed API call as "out of stock"; check the live page instead
        print(f"  -> [SP-API] {asin}: pricing check failed. Checking the live page instead...")
    else:
        row["Current Price"] = f"{currency}{sp_current:.2f}"
        row["Listing Price"] = f"{currency}{sp_current:.2f}"
        print(f"  -> [SP-API] {asin} is ACTIVE ({currency}{sp_current:.2f}). Launching browser scrape...")

    is_match, r_rating, r_reviews, r_coupon, r_current, r_list, screen_bytes, selector_used, storage_state = \
        await scrape_amazon_frontend(asin, browser, marketplace_name, domain, locale, timezone_id, storage_state)

    if screen_bytes:
        os.makedirs("screenshots", exist_ok=True)
        filename = f"{asin}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        screenshot_path = os.path.join("screenshots", filename)
        with open(screenshot_path, "wb") as f:
            f.write(screen_bytes)
        row["Screenshot_Path"] = screenshot_path
        row["Screenshot_Bytes"] = screen_bytes

    # Handle Diagnostic Failures (Prevents errors from leaking into the Rating column)
    failure_states = ["CAPTCHA Block", "Region Redirect", "Visual Scrape Blocked",
                      "Consent Wall Block", "Scrape Error", "ASIN Mismatch", "Page Not Found", "Product Block Missing"]
    if not is_match or r_rating in failure_states:
        row["Status"] = f"Active ({r_rating})" if r_rating in failure_states else "Active (Visual Scrape Failed)"
        row["Rating"] = "N/A"
        return row, storage_state

    # Override the Status column if the frontend scrape found it to be out of stock
    if selector_used == "out_of_stock":
        row["Status"] = "Unavailable"
    else:
        row["Status"] = "Active"

    # Write the successfully extracted information to the row
    row["Rating"] = r_rating
    row["Review Count"] = r_reviews
    row["Coupon Info"] = r_coupon

    scraped_current = parse_price_to_float(r_current)
    scraped_list = parse_price_to_float(r_list)

    final_current = scraped_current if scraped_current is not None else sp_current
    row["Current Price"] = f"{currency}{final_current:.2f}" if final_current is not None else "N/A"

    effective_list = scraped_list if scraped_list is not None else final_current
    row["Listing Price"] = f"{currency}{effective_list:.2f}" if effective_list is not None else "N/A"

    savings_pct = calculate_savings_percent(final_current, effective_list)
    row["Savings"] = f"{savings_pct}%"

    return row, storage_state

async def async_workflow(asins, marketplace_id, marketplace_name, endpoint, access_token):
    processed_rows = []
    domain_storage_state = None

    virtual_display = None
    try:
        virtual_display = Display(visible=0, size=(1280, 1080))
        virtual_display.start()
    except FileNotFoundError:
        print("  -> Local environment detected: Running without Xvfb virtual display...")

    try:
        async with async_playwright() as p:
            # Force "chrome" channel so Playwright uses standard Branded Chrome, not Chromium
            browser = await p.chromium.launch(
                channel="chrome", 
                headless=False, 
                args=["--disable-blink-features=AutomationControlled"]
            )

            init_db()

            for asin in asins:
                row_data, domain_storage_state = await process_asin(
                    asin, marketplace_id, marketplace_name, endpoint, access_token, browser, domain_storage_state
                )
                processed_rows.append(row_data)

                log_to_database(row_data)
                print(f"  -> Successfully logged {asin} to SQLite database.")

                await asyncio.sleep(0.3)

            await browser.close()
    finally:
        if virtual_display:
            virtual_display.stop()

    return processed_rows

# ==========================================
# 3. Flask API Route (File Upload Endpoint)
# ==========================================
@app.route('/', methods=['GET'])
def index():
    dynamic_mapping = {}
    for config in CONFIG_JSON["sp_api_configurations"]:
        acc = config["account_name"]
        if acc not in dynamic_mapping:
            dynamic_mapping[acc] = []
        for mp in config["marketplaces"].keys():
            if mp not in dynamic_mapping[acc]:
                dynamic_mapping[acc].append(mp)

    mapping_json = json.dumps(dynamic_mapping)

    return f'''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Amazon ASIN Scraper Tester</title>
        <script>
            const accountMap = {mapping_json};

            function updateMarketplaces() {{
                const accountSelect = document.getElementById("account_name");
                const marketplaceSelect = document.getElementById("marketplace_name");
                const selectedAccount = accountSelect.value;

                marketplaceSelect.innerHTML = "";

                if (accountMap[selectedAccount]) {{
                    accountMap[selectedAccount].forEach(function(mp) {{
                        const option = document.createElement("option");
                        option.value = mp;
                        option.text = mp;
                        marketplaceSelect.appendChild(option);
                    }});
                }}
            }}

            window.onload = function() {{
                const accountSelect = document.getElementById("account_name");

                for (const acc in accountMap) {{
                    const option = document.createElement("option");
                    option.value = acc;
                    option.text = acc;
                    accountSelect.appendChild(option);
                }}

                updateMarketplaces();
            }};
        </script>
        <style>
            body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 50px; background-color: #f9f9fb; }}
            .container {{ max-width: 450px; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
            select, input[type="file"] {{ width: 100%; padding: 8px; margin-top: 5px; margin-bottom: 20px; border: 1px solid #ccc; border-radius: 4px; font-size: 14px; }}
            label {{ font-weight: 600; color: #333; }}
            button {{ width: 100%; padding: 12px; font-weight: bold; cursor: pointer; background-color: #007bff; color: white; border: none; border-radius: 4px; font-size: 16px; transition: background 0.2s; }}
            button:hover {{ background-color: #0056b3; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>Upload ASIN List</h2>
            <a href="/dashboard" style="display:block; text-align:center; padding:10px; background:#28a745; color:white; text-decoration:none; border-radius:4px; margin-bottom:20px; font-weight:bold;">View Live Dashboard</a>
            <form action="/process-excel" method="POST" enctype="multipart/form-data">
                <label>Excel File:</label><br>
                <input type="file" name="file" accept=".xlsx" required>

                <label>Account Name:</label><br>
                <select id="account_name" name="account_name" onchange="updateMarketplaces()" required>
                </select>

                <label>Marketplace:</label><br>
                <select id="marketplace_name" name="marketplace_name" required>
                </select>

                <button type="submit">Run Scraper</button>
            </form>
        </div>
    </body>
    </html>
    '''

from flask import send_from_directory

@app.route('/screenshots/<path:filename>')
def serve_screenshot(filename):
    return send_from_directory('screenshots', filename)

@app.route('/process-excel', methods=['POST'])
def process_excel():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files['file']
    account_name = request.form.get("account_name")
    marketplace_name = request.form.get("marketplace_name")

    config, marketplace_id = get_target_config(account_name, marketplace_name)
    if not config:
        return jsonify({"error": "Invalid account or marketplace selection"}), 400

    try:
        # 1. Read the file normally first
        input_df = pd.read_excel(file)
        
        # 2. Check if the first row was accidentally consumed as a header
        first_col_name = str(input_df.columns[0]).strip().upper()
        if len(first_col_name) == 10 and (first_col_name.startswith('B') or first_col_name[0].isdigit()):
            # The header is actually an ASIN. Reload the file without a header row.
            file.seek(0)
            input_df = pd.read_excel(file, header=None)
        
        # 3. Find the ASIN column (defaults to the first column if no named header matches)
        asin_col = next((col for col in input_df.columns if str(col).strip().upper() in ["ASIN", "ID", "PRODUCT ID"]), input_df.columns[0])
        
        # 4. Extract the ASINs
        asins = [str(val).strip().upper() for val in input_df[asin_col].dropna().tolist() if str(val).strip()]
        
        # 5. Clean out header strings if they accidentally got mixed into the data rows
        asins = [a for a in asins if a not in ["ASIN", "ID", "PRODUCT ID"]]

        access_token = get_access_token(config["refresh_token"])

        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        processed_rows = loop.run_until_complete(
            async_workflow(asins, marketplace_id, marketplace_name, config["endpoint"], access_token)
        )

        # 1. Create a lookup dictionary of our scraped results
        scraped_lookup = {r["ID"]: r for r in processed_rows}

        # 2. Define the new columns we want to append
        final_columns = [
            "ID", "Marketplace", "Product URL", "Image", "Product Name", 
            "Brand", "Ranking", "Rating", "No. of Review", "List Price", 
            "Current Price", "Savings", "Coupon", "Status"
        ]

        # Map the new Excel column names to the internal keys your scraper uses
        # (This prevents breaking your SQLite database which expects the old names)
        key_mapping = {
            "ID": "ID", "Marketplace": "Region", "Product URL": "Product URL", 
            "Image": "Product Screenshot", "Product Name": "Product Name", 
            "Brand": "Brand", "Ranking": "Ranking", "Rating": "Rating", 
            "No. of Review": "Review Count", "List Price": "Listing Price", 
            "Current Price": "Current Price", "Savings": "Savings", 
            "Coupon": "Coupon Info", "Status": "Status"
        }

        # 3. Ensure these columns exist in the original dataframe
        for col in final_columns:
            if col not in input_df.columns:
                input_df[col] = ""

        # 4. Map the scraped data directly into the row structure
        screenshot_list = []
        for idx in input_df.index:
            val = input_df.at[idx, asin_col]
            asin = str(val).strip().upper() if pd.notna(val) else ""
            
            if asin in scraped_lookup:
                row_data = scraped_lookup[asin]
                for col in final_columns:
                    internal_key = key_mapping[col]
                    input_df.at[idx, col] = row_data.get(internal_key, "N/A")
                screenshot_list.append(row_data.get("Screenshot_Bytes"))
            else:
                screenshot_list.append(None)

        # 4b. Reorder the dataframe to match the requested column order
        # This brings the specific columns to the front, and keeps any extra original columns at the end
        other_cols = [c for c in input_df.columns if c not in final_columns and c != asin_col]
        input_df = input_df[final_columns + other_cols]

        # 5. Export back to Excel keeping original structure
        output_buffer = BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            input_df.to_excel(writer, index=False, sheet_name="Results")

        output_buffer.seek(0)
        wb = load_workbook(output_buffer)
        ws = wb.active

        # 6. Apply formatting
        center_aligned = Alignment(horizontal="center", vertical="center", wrap_text=True)
        hyperlink_font = Font(color="0563C1", underline="single") 
        
        url_col_idx = next((c for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value == "Product URL"), None)

        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            for cell in row:
                cell.alignment = center_aligned
                if row_idx > 1 and url_col_idx and cell.column == url_col_idx and cell.value and cell.value != "N/A":
                    cell.hyperlink = str(cell.value)
                    cell.font = hyperlink_font

        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.row_dimensions[1].height = 30

        for col_idx in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = 25

        # 7. Embed images based on the new list mapping
        # Note: We now search for "Image" instead of "Product Screenshot"
        img_col_idx = next((c for c in range(1, ws.max_column + 1) if ws.cell(row=1, column=c).value == "Image"), None)

        if img_col_idx:
            ws.column_dimensions[get_column_letter(img_col_idx)].width = 49

        for row_idx in range(2, ws.max_row + 1):
            ws.row_dimensions[row_idx].height = 260
            
            if img_col_idx:
                cell = ws.cell(row=row_idx, column=img_col_idx)
                screen_bytes = screenshot_list[row_idx - 2] 
                if screen_bytes:
                    try:
                        excel_img = ExcelImage(BytesIO(screen_bytes))
                        _from = AnchorMarker(col=img_col_idx - 1, colOff=50000, row=row_idx - 1, rowOff=50000)
                        _to = AnchorMarker(col=img_col_idx, colOff=-50000, row=row_idx, rowOff=-50000)
                        excel_img.anchor = TwoCellAnchor(_from=_from, to=_to, editAs='twoCell')
                        ws.add_image(excel_img)
                        cell.value = "" 
                    except Exception:
                        cell.value = "Render Error"
                else:
                    if not cell.value or cell.value == "N/A":
                        cell.value = "Unavailable"

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        filename = f"Processed_{marketplace_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(
            final_buffer,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/dashboard', methods=['GET'])
def dashboard():
    return '''
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Amazon ASIN Dashboard</title>
        <style>
            body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 30px; background-color: #f9f9fb; margin: 0; }
            .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
            .btn { padding: 10px 15px; background-color: #007bff; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; border: none; cursor: pointer; }
            .btn:hover { background-color: #0056b3; }
            
            /* Filter Bar */
            .filters { display: flex; flex-wrap: wrap; gap: 15px; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); margin-bottom: 20px; }
            .filter-group { display: flex; flex-direction: column; }
            .filter-group label { font-size: 12px; font-weight: bold; margin-bottom: 5px; color: #555; }
            .filter-group input { padding: 8px; border: 1px solid #ccc; border-radius: 4px; width: 130px; font-size: 14px; }
            
            /* Table */
            .table-container { background: white; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow-x: auto; }
            table { width: 100%; border-collapse: collapse; min-width: 1200px; }
            th, td { padding: 12px 15px; text-align: left; border-bottom: 1px solid #eee; font-size: 13px; vertical-align: middle; }
            th { background-color: #f1f3f5; color: #333; font-weight: bold; position: sticky; top: 0; }
            tr:hover { background-color: #f8f9fa; }
            .img-preview { max-width: 80px; max-height: 60px; border-radius: 4px; object-fit: cover; }
            a { color: #007bff; text-decoration: none; }
            a:hover { text-decoration: underline; }
            
            .status-active { color: #0f8b8d; font-weight: bold; }
            .status-out { color: #d90429; font-weight: bold; }
            .coupon-badge { background: #28a745; color: white; padding: 3px 6px; border-radius: 3px; font-size: 11px; font-weight: bold; }
        </style>
    </head>
    <body>
        <div class="header">
            <h2>Scraper Dashboard</h2>
            <a href="/" class="btn">← Back to Scraper</a>
        </div>

        <div class="filters">
            <div class="filter-group">
                <label>Min Price</label>
                <input type="number" id="minPrice" placeholder="e.g. 10" oninput="applyFilters()">
            </div>
            <div class="filter-group">
                <label>Max Price</label>
                <input type="number" id="maxPrice" placeholder="e.g. 100" oninput="applyFilters()">
            </div>
            <div class="filter-group">
                <label>Min Savings (%)</label>
                <input type="number" id="minSavings" placeholder="e.g. 15" oninput="applyFilters()">
            </div>
            <div class="filter-group">
                <label>Min Rating</label>
                <input type="number" id="minRating" step="0.1" placeholder="e.g. 4.0" oninput="applyFilters()">
            </div>
            <div class="filter-group">
                <label>Min Reviews</label>
                <input type="number" id="minReviews" placeholder="e.g. 100" oninput="applyFilters()">
            </div>
            <div class="filter-group">
                <label>Search Keyword</label>
                <input type="text" id="searchRank" placeholder="ASIN, Brand, Rank..." oninput="applyFilters()" style="width: 200px;">
            </div>
            <div class="filter-group" style="justify-content: flex-end;">
                <button class="btn" onclick="resetFilters()">Reset Filters</button>
            </div>
        </div>

        <div class="table-container">
            <table>
                <thead>
                    <tr>
                        <th>ASIN / Region</th>
                        <th>Image</th>
                        <th>Product Details</th>
                        <th>Price & Savings</th>
                        <th>Ranking</th>
                        <th>Rating / Reviews</th>
                        <th>Coupon Price</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    <tr><td colspan="8" style="text-align:center;">Loading data...</td></tr>
                </tbody>
            </table>
        </div>

        <script>
            let allData = [];

            // Fetch data from API on load
            window.onload = async function() {
                try {
                    const res = await fetch('/api/data');
                    allData = await res.json();
                    renderTable(allData);
                } catch (error) {
                    document.getElementById('tableBody').innerHTML = `<tr><td colspan="8" style="color:red; text-align:center;">Error loading data: ${error.message}</td></tr>`;
                }
            };

            // Utility to extract raw numbers from formatted text (e.g. "$1,200.50" -> 1200.5)
            function extractNumber(val) {
                if (!val || val === "N/A" || val === "Error") return null;
                let clean = String(val);
                
                // Convert European comma decimals to dots
                if (clean.includes(',') && clean.includes('.')) {
                    clean = clean.lastIndexOf(',') > clean.lastIndexOf('.') 
                        ? clean.replace(/\./g, '').replace(',', '.') 
                        : clean.replace(/,/g, '');
                } else if (clean.includes(',')) {
                    const parts = clean.split(',');
                    clean = parts[parts.length - 1].length <= 2 
                        ? clean.replace(/,/g, '.') 
                        : clean.replace(/,/g, '');
                }
                
                clean = clean.replace(/[^0-9.-]+/g, "");
                const parsed = parseFloat(clean);
                return isNaN(parsed) ? null : parsed;
            }

            // Filtering logic
            function applyFilters() {
                const minPrice = parseFloat(document.getElementById('minPrice').value);
                const maxPrice = parseFloat(document.getElementById('maxPrice').value);
                const minSavings = parseFloat(document.getElementById('minSavings').value);
                const minRating = parseFloat(document.getElementById('minRating').value);
                const minReviews = parseFloat(document.getElementById('minReviews').value);
                const searchTxt = document.getElementById('searchRank').value.toLowerCase();

                const filtered = allData.filter(row => {
                    const price = extractNumber(row.current_price);
                    const savings = extractNumber(row.savings);
                    const rating = extractNumber(row.rating);
                    const reviews = extractNumber(row.review_count);

                    if (!isNaN(minPrice) && (price === null || price < minPrice)) return false;
                    if (!isNaN(maxPrice) && (price === null || price > maxPrice)) return false;
                    if (!isNaN(minSavings) && (savings === null || savings < minSavings)) return false;
                    if (!isNaN(minRating) && (rating === null || rating < minRating)) return false;
                    if (!isNaN(minReviews) && (reviews === null || reviews < minReviews)) return false;

                    if (searchTxt) {
                        const searchTarget = `${row.asin} ${row.brand} ${row.product_name} ${row.ranking}`.toLowerCase();
                        if (!searchTarget.includes(searchTxt)) return false;
                    }

                    return true;
                });

                renderTable(filtered);
            }

            function resetFilters() {
                document.querySelectorAll('.filters input').forEach(inp => inp.value = '');
                renderTable(allData);
            }

            // Render DOM Table
            function renderTable(data) {
                const tbody = document.getElementById('tableBody');
                tbody.innerHTML = '';

                if (data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;">No matching records found.</td></tr>';
                    return;
                }

                data.forEach(row => {
                    const tr = document.createElement('tr');
                    
                    // Image Handling
                    let imgHtml = "N/A";
                    if (row.screenshot_path) {
                        const safePath = row.screenshot_path.replace(/\\\\/g, '/');
                        imgHtml = `<img src="/${safePath}" class="img-preview" onerror="this.style.display='none'">`;
                    }

                    // Coupon Price Logic: Calculate final numerical price
                    let finalCouponPrice = "N/A";
                    let originalCouponText = "No Coupon";
                    
                    if (row.coupon_info && row.coupon_info !== "No Coupon") {
                        originalCouponText = row.coupon_info;
                        const currentPriceNum = extractNumber(row.current_price);
                        const couponNum = extractNumber(row.coupon_info);
                        
                        if (couponNum !== null && currentPriceNum !== null) {
                            // If the text contains "%", treat it as a fractional voucher
                            if (String(row.coupon_info).includes("%")) {
                                finalCouponPrice = (currentPriceNum * (1 - (couponNum / 100))).toFixed(2);
                            } else {
                                // Keep the standard US Coupon logic (Flat Discount or Target Price)
                                if (couponNum > (currentPriceNum * 0.5)) {
                                    finalCouponPrice = couponNum.toFixed(2); // It's a target price
                                } else {
                                    finalCouponPrice = Math.max(0, currentPriceNum - couponNum).toFixed(2); // It's a flat discount
                                }
                            }
                        } else if (couponNum !== null) {
                            finalCouponPrice = couponNum.toFixed(2);
                        }
                    }

                    // Status styling
                    const statusClass = String(row.status).includes("Active") ? "status-active" : "status-out";

                    tr.innerHTML = `
                        <td>
                            <strong><a href="${row.product_url}" target="_blank">${row.asin}</a></strong><br>
                            <span style="color:#777">${row.region}</span>
                        </td>
                        <td>${imgHtml}</td>
                        <td>
                            <strong>${row.brand || 'N/A'}</strong><br>
                            <span style="font-size:11px; color:#555; display:-webkit-box; -webkit-line-clamp:3; -webkit-box-orient:vertical; overflow:hidden;">
                                ${row.product_name || 'N/A'}
                            </span>
                        </td>
                        <td>
                            <strong>${row.current_price}</strong> 
                            <span style="text-decoration:line-through; color:#999; font-size:11px; margin-left:4px;">${row.listing_price}</span>
                            <br><span style="color:#28a745; font-weight:bold;">${row.savings || '0%'} off</span>
                        </td>
                        <td style="max-width:150px; font-size:11px; line-height:1.4;">${row.ranking || 'N/A'}</td>
                        <td>
                            ⭐ ${row.rating}<br>
                            <span style="color:#777; font-size:12px;">(${row.review_count} revs)</span>
                        </td>
                        <td>
                            <strong style="color:#d90429;">${finalCouponPrice}</strong><br>${originalCouponText !== "No Coupon" ? `<span class="coupon-badge">${originalCouponText}</span>` : ''}
                        </td>
                        <td class="${statusClass}">${row.status}</td>
                    `;
                    tbody.appendChild(tr);
                });
            }
        </script>
    </body>
    </html>
    '''
    
@app.route('/api/data', methods=['GET'])
def get_data():
    try:
        conn = sqlite3.connect("amazon_tracker.db")
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM products ORDER BY last_updated DESC")
        rows = cursor.fetchall()
        conn.close()
        return jsonify([dict(row) for row in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
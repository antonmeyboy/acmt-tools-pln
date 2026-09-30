import sys
import os
import io
import re
import html
import json
import time
import uuid
import base64
import zipfile
import subprocess
import textwrap
import threading
import sqlite3
import pickle
import urllib.parse
import urllib3
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import streamlit as st
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageStat, ImageFilter

# ReportLab untuk Perakitan PDF (Tab 17)
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# python-pptx untuk Perakitan Presentasi PowerPoint (Tab 15)
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.text import PP_ALIGN
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
except ImportError:
    pass

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
pd.set_option("styler.render.max_elements", 2000000)

# -------------------------------------------------------------------------
# 1. KONFIGURASI HALAMAN STREAMLIT
# -------------------------------------------------------------------------
st.set_page_config(
    page_title="ACMT Tools Cloud - PPT & PDF Verifikasi",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -------------------------------------------------------------------------
# 2. AUTO-INSTALL PLAYWRIGHT CHROMIUM DI SERVER LINUX CLOUD & FIX WINDOWS
# -------------------------------------------------------------------------
@st.cache_resource(show_spinner="Menyiapkan mesin browser Cloud (Chromium)...")
def ensure_playwright_browser():
    if sys.platform != "win32":
        try:
            subprocess.run(
                [sys.executable, "-m", "playwright", "install", "chromium"],
                check=False,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL
            )
        except Exception:
            pass
    return True

ensure_playwright_browser()

if sys.platform == 'win32':
    import asyncio
    try:
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    except Exception:
        pass
    try:
        from asyncio.proactor_events import _ProactorBasePipeTransport
        _orig_call_connection_lost = _ProactorBasePipeTransport._call_connection_lost

        def _silenced_call_connection_lost(self, exc):
            try:
                _orig_call_connection_lost(self, exc)
            except Exception:
                pass

        _ProactorBasePipeTransport._call_connection_lost = _silenced_call_connection_lost
    except Exception:
        pass

def load_app_font(bold=False, size=11):
    candidates = (
        ["tahomabd.ttf", "arialbd.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]
        if bold else
        ["tahoma.ttf", "arial.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"]
    )
    for c in candidates:
        try:
            return ImageFont.truetype(c, size)
        except Exception:
            continue
    return ImageFont.load_default()

# -------------------------------------------------------------------------
# 3. GENERATOR PAKET ZIP EKSTENSI CHROME KEEP-ALIVE (BUILT-IN MEMORY)
# -------------------------------------------------------------------------
@st.cache_data
def get_keepalive_zip_bytes():
    ext_files = {
        "acmt-keepalive/manifest.json": json.dumps({
            "manifest_version": 3,
            "name": "ACMT Keep-Alive Assistant",
            "version": "1.2.0",
            "description": "Menjaga sesi web portalapp.iconpln.co.id/acmt tetap aktif, salin cookies otomatis, dan background keep-alive ping.",
            "permissions": ["storage", "alarms", "cookies", "clipboardWrite"],
            "host_permissions": [
                "*://portalapp.iconpln.co.id/*",
                "*://*.iconpln.co.id/*",
                "http://localhost/*",
                "http://127.0.0.1/*"
            ],
            "action": {
                "default_popup": "popup.html",
                "default_title": "ACMT Keep-Alive Assistant"
            },
            "background": {"service_worker": "background.js"},
            "content_scripts": [{
                "matches": ["*://portalapp.iconpln.co.id/*"],
                "js": ["content.js"],
                "css": ["content.css"],
                "run_at": "document_idle"
            }]
        }, indent=2),
        "acmt-keepalive/background.js": """// background.js - Service Worker for ACMT Keep-Alive Assistant
const DEFAULT_INTERVAL_MINUTES = 2;
const ALARM_NAME = "acmt_keepalive_alarm";

chrome.runtime.onInstalled.addListener(async () => {
  try {
    const data = await chrome.storage.local.get(["isEnabled", "intervalMinutes", "showBadge"]);
    const isEnabled = data.isEnabled !== undefined ? data.isEnabled : true;
    const intervalMinutes = data.intervalMinutes || DEFAULT_INTERVAL_MINUTES;
    const showBadge = data.showBadge !== undefined ? data.showBadge : true;
    await chrome.storage.local.set({ isEnabled, intervalMinutes, showBadge, lastPingTime: null, totalPings: 0 });
    setupAlarm(intervalMinutes);
  } catch (e) {}
});

function setupAlarm(intervalMinutes) {
  try {
    chrome.alarms.clear(ALARM_NAME, () => {
      chrome.alarms.create(ALARM_NAME, { periodInMinutes: Number(intervalMinutes) || DEFAULT_INTERVAL_MINUTES });
    });
  } catch (e) {}
}

chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name !== ALARM_NAME) return;
  try {
    const data = await chrome.storage.local.get(["isEnabled"]);
    if (data.isEnabled === false) return;
    pingAcmtTabs();
  } catch (e) {}
});

async function pingAcmtTabs() {
  try {
    const tabs = await chrome.tabs.query({ url: "*://portalapp.iconpln.co.id/*" });
    if (!tabs || tabs.length === 0) return;
    const now = new Date().toLocaleTimeString("id-ID", { hour12: false });
    const storageData = await chrome.storage.local.get(["totalPings"]);
    const newTotal = (storageData.totalPings || 0) + 1;
    await chrome.storage.local.set({ lastPingTime: now, totalPings: newTotal });
    try {
      await fetch("https://portalapp.iconpln.co.id/acmt/Main.html?_sw_ping=" + Date.now(), { method: "GET", cache: "no-store" });
    } catch (e) {}
    for (const tab of tabs) {
      if (tab.id) {
        chrome.tabs.sendMessage(tab.id, { action: "PERFORM_KEEPALIVE", timestamp: now, totalPings: newTotal }, () => {
          if (chrome.runtime.lastError) {}
        });
      }
    }
    syncCookiesToStreamlit().catch(() => {});
  } catch (err) {}
}

async function getAcmtCookieString() {
  const cookieMap = new Map();
  const queries = [
    { url: "https://portalapp.iconpln.co.id/acmt/Main.html" },
    { url: "https://portalapp.iconpln.co.id/acmt/" },
    { url: "https://portalapp.iconpln.co.id/" },
    { domain: "portalapp.iconpln.co.id" },
    { domain: ".iconpln.co.id" }
  ];
  const runQuery = (q) => new Promise((resolve) => {
    try {
      chrome.cookies.getAll(q, (cookies) => {
        if (!chrome.runtime.lastError && cookies) {
          for (const c of cookies) {
            if (c && c.name && c.value) cookieMap.set(c.name, c.value);
          }
        }
        resolve();
      });
    } catch (e) { resolve(); }
  });
  await Promise.all(queries.map(q => runQuery(q)));
  try {
    const tabs = await chrome.tabs.query({ url: "*://portalapp.iconpln.co.id/*" });
    if (tabs) {
      for (const t of tabs) {
        if (t.url) {
          const m = t.url.match(/;jsessionid=([^?#&]+)/i);
          if (m && m[1]) cookieMap.set("JSESSIONID", m[1]);
          await runQuery({ url: t.url });
        }
      }
    }
  } catch (e) {}
  if (cookieMap.size === 0) return "";
  const pairs = [];
  for (const [name, val] of cookieMap.entries()) pairs.push(`${name}=${val}`);
  return pairs.join("; ");
}

async function syncCookiesToStreamlit() {
  try {
    const cookieHeader = await getAcmtCookieString();
    if (!cookieHeader) return "";
    try {
      await fetch("http://127.0.0.1:8502/update_cookie", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ cookie: cookieHeader, timestamp: new Date().toLocaleTimeString("id-ID", { hour12: false }) })
      });
    } catch (fe) {}
    return cookieHeader;
  } catch (err) { return ""; }
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (!message || !message.action) { sendResponse({ success: false }); return false; }
  if (message.action === "UPDATE_SETTINGS") {
    if (message.intervalMinutes) setupAlarm(message.intervalMinutes);
    sendResponse({ success: true });
    return false;
  }
  if (message.action === "TRIGGER_NOW") {
    pingAcmtTabs().then(() => sendResponse({ success: true })).catch(() => sendResponse({ success: false }));
    return true;
  }
  if (message.action === "GET_COOKIES") {
    getAcmtCookieString().then((cookieStr) => sendResponse({ cookies: cookieStr || "" })).catch(() => sendResponse({ cookies: "" }));
    return true;
  }
  if (message.action === "SYNC_COOKIES") {
    syncCookiesToStreamlit().then((res) => sendResponse({ success: true, cookies: res || "" })).catch(() => sendResponse({ success: false, cookies: "" }));
    return true;
  }
  sendResponse({ success: false });
  return false;
});
""",
        "acmt-keepalive/content.js": """// content.js - Runs inside portalapp.iconpln.co.id/acmt/*
(() => {
  if (window.__ACMT_KEEPALIVE_INJECTED__) return;
  window.__ACMT_KEEPALIVE_INJECTED__ = true;
  let isEnabled = true, showBadge = true, intervalMinutes = 2, lastExecutedTime = 0, fallbackTimer = null, totalPingsCount = 0;

  chrome.storage.local.get(["isEnabled", "showBadge", "intervalMinutes", "totalPings"], (data) => {
    if (data.isEnabled !== undefined) isEnabled = data.isEnabled;
    if (data.showBadge !== undefined) showBadge = data.showBadge;
    if (data.intervalMinutes) intervalMinutes = Number(data.intervalMinutes);
    if (data.totalPings) totalPingsCount = Number(data.totalPings);
    createFloatingBadge();
    updateBadgeUI("idle", "Siap");
    startFallbackInterval();
    try { chrome.runtime.sendMessage({ action: "SYNC_COOKIES" }, () => {}); } catch (e) {}
  });

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === "PERFORM_KEEPALIVE") {
      performKeepAlive("background_alarm");
      sendResponse({ status: "executed" });
    }
  });

  function startFallbackInterval() {
    if (fallbackTimer) clearInterval(fallbackTimer);
    fallbackTimer = setInterval(() => {
      if (Date.now() - lastExecutedTime >= 45000) performKeepAlive("tab_interval");
    }, Math.max(intervalMinutes, 1) * 60 * 1000);
  }

  async function performKeepAlive() {
    if (!isEnabled) return;
    lastExecutedTime = Date.now();
    totalPingsCount++;
    const timeStr = new Date().toLocaleTimeString("id-ID", { hour12: false });
    updateBadgeUI("active", "Ping... (" + timeStr + ")");
    try {
      document.dispatchEvent(new MouseEvent("mousemove", { bubbles: true, clientX: 120, clientY: 120 }));
      await fetch(window.location.origin + window.location.pathname + "?_acmt_ping=" + Date.now(), { method: "GET", cache: "no-store", credentials: "include" });
      chrome.storage.local.set({ lastPingTime: timeStr, totalPings: totalPingsCount });
      setTimeout(() => updateBadgeUI("idle", "Aktif • " + timeStr), 1000);
    } catch (err) {
      updateBadgeUI("error", "Gagal Ping");
    }
  }

  function createFloatingBadge() {
    if (document.getElementById("acmt-keepalive-badge")) return;
    const badge = document.createElement("div");
    badge.id = "acmt-keepalive-badge";
    badge.className = "acmt-badge-container";
    badge.innerHTML = '<div class="acmt-badge-content" id="acmt-badge-content">' +
      '<span class="acmt-badge-status-dot active" id="acmt-badge-dot"></span>' +
      '<div class="acmt-badge-info"><span class="acmt-badge-title">ACMT Keep-Alive</span><span class="acmt-badge-subtitle" id="acmt-badge-text">Siap</span></div>' +
      '<button type="button" class="acmt-badge-btn acmt-badge-btn-copy" id="acmt-copy-cookie-btn">📋 Salin Cookie</button>' +
      '<button type="button" class="acmt-badge-btn" id="acmt-ping-now-btn">⚡ Ping</button>' +
      '</div>';
    document.body.appendChild(badge);

    document.getElementById("acmt-copy-cookie-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      const copyBtn = document.getElementById("acmt-copy-cookie-btn");
      copyBtn.innerHTML = "⏳ Mengambil...";
      chrome.runtime.sendMessage({ action: "GET_COOKIES" }, async (response) => {
        let cookieStr = (response && response.cookies) ? response.cookies : document.cookie;
        if (cookieStr) {
          try { await navigator.clipboard.writeText(cookieStr); } catch (ce) {
            const t = document.createElement("textarea"); t.value = cookieStr; document.body.appendChild(t); t.select(); document.execCommand("copy"); document.body.removeChild(t);
          }
          copyBtn.innerHTML = "✓ Disalin!";
          setTimeout(() => { copyBtn.innerHTML = "📋 Salin Cookie"; }, 2500);
        } else {
          copyBtn.innerHTML = "⚠️ Kosong";
          setTimeout(() => { copyBtn.innerHTML = "📋 Salin Cookie"; }, 2500);
        }
      });
    });

    document.getElementById("acmt-ping-now-btn").addEventListener("click", (e) => {
      e.stopPropagation();
      performKeepAlive();
    });
  }

  function updateBadgeUI(status, text) {
    const subtitle = document.getElementById("acmt-badge-text");
    if (subtitle && text) subtitle.textContent = text;
  }
})();
""",
        "acmt-keepalive/content.css": """.acmt-badge-container { position: fixed !important; bottom: 16px !important; right: 16px !important; z-index: 2147483647 !important; font-family: sans-serif !important; font-size: 12px !important; }
.acmt-badge-content { display: flex !important; align-items: center !important; gap: 10px !important; background: rgba(255,255,255,0.96) !important; border: 1px solid #cbd5e1 !important; border-radius: 30px !important; padding: 6px 12px !important; box-shadow: 0 4px 14px rgba(0,0,0,0.12) !important; }
.acmt-badge-status-dot { width: 9px !important; height: 9px !important; border-radius: 50% !important; background-color: #10b981 !important; display: inline-block !important; }
.acmt-badge-info { display: flex !important; flex-direction: column !important; }
.acmt-badge-title { font-weight: 700 !important; font-size: 11px !important; color: #0f172a !important; }
.acmt-badge-subtitle { font-size: 10px !important; color: #64748b !important; }
.acmt-badge-btn { background: #0284c7 !important; color: #fff !important; border: none !important; border-radius: 12px !important; padding: 4px 9px !important; font-size: 10px !important; font-weight: 600 !important; cursor: pointer !important; }
.acmt-badge-btn-copy { background: #059669 !important; }
""",
        "acmt-keepalive/popup.html": """<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><link rel="stylesheet" href="popup.css"></head>
<body><div class="popup-container"><h3>ACMT Keep-Alive</h3><button id="btn-copy-cookie" class="btn">📋 Salin Cookie ACMT</button><button id="btn-ping-now" class="btn2">⚡ Ping Sekarang</button></div><script src="popup.js"></script></body></html>
""",
        "acmt-keepalive/popup.css": """body { width: 260px; font-family: sans-serif; padding: 14px; background: #f8fafc; } .popup-container { display: flex; flex-direction: column; gap: 8px; } .btn { background: #059669; color: #fff; border: none; padding: 8px; border-radius: 6px; font-weight: bold; cursor: pointer; } .btn2 { background: #0284c7; color: #fff; border: none; padding: 8px; border-radius: 6px; font-weight: bold; cursor: pointer; }""",
        "acmt-keepalive/popup.js": """document.addEventListener("DOMContentLoaded", () => {
  const b = document.getElementById("btn-copy-cookie");
  if (b) b.addEventListener("click", () => {
    chrome.runtime.sendMessage({ action: "GET_COOKIES" }, async (res) => {
      if (res && res.cookies) {
        await navigator.clipboard.writeText(res.cookies);
        b.textContent = "✓ Berhasil Disalin!";
        setTimeout(() => b.textContent = "📋 Salin Cookie ACMT", 2000);
      }
    });
  });
  const p = document.getElementById("btn-ping-now");
  if (p) p.addEventListener("click", () => chrome.runtime.sendMessage({ action: "TRIGGER_NOW" }));
});
""",
        "acmt-keepalive/README.md": (
            "# ACMT Keep-Alive Assistant (Chrome / Edge Extension)\n\n"
            "## Cara Memasang di Laptop (Dari File ZIP)\n"
            "1. Ekstrak (Unzip) file `acmt-keepalive.zip` ke folder mana saja.\n"
            "2. Buka Google Chrome lalu ketik `chrome://extensions` di address bar (atau `edge://extensions` di Edge).\n"
            "3. Aktifkan saklar **Developer mode** (Mode pengembang) di pojok kanan atas.\n"
            "4. Klik tombol **Load unpacked** (Muat yang belum dibongkar) di pojok kiri atas.\n"
            "5. Pilih folder `acmt-keepalive` hasil ekstrak tadi.\n"
            "6. Buka tab ACMT (`https://portalapp.iconpln.co.id/acmt/Main.html#overview`) lalu Refresh (F5).\n"
        )
    }

    local_ext_dirs = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "acmt-keepalive"),
        r"C:\Users\anton.armayadi\.gemini\antigravity\scratch\acmt-keepalive"
    ]
    for d in local_ext_dirs:
        if os.path.isdir(d):
            for fname in ["manifest.json", "background.js", "content.js", "content.css", "popup.html", "popup.css", "popup.js", "README.md"]:
                fpath = os.path.join(d, fname)
                if os.path.exists(fpath):
                    try:
                        with open(fpath, "r", encoding="utf-8") as f:
                            ext_files[f"acmt-keepalive/{fname}"] = f.read()
                    except Exception:
                        pass
            break

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for arcname, content in ext_files.items():
            zf.writestr(arcname, content.encode("utf-8"))
    buf.seek(0)
    return buf.getvalue()

# -------------------------------------------------------------------------
# 4. KONEKSI DATABASE LOKAL (FALLBACK JIKA DI CLOUD / TIDAK TERSEDIA)
# -------------------------------------------------------------------------
engine = None
try:
    from sqlalchemy import create_engine, text
    if os.getenv("DATABASE_URL"):
        engine = create_engine(os.getenv("DATABASE_URL"))
except Exception:
    engine = None

# -------------------------------------------------------------------------
# 5. SINKRONISASI COOKIE ACMT
# -------------------------------------------------------------------------
COOKIE_SYNC_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cookie_acmt.txt")

DEFAULT_ACMT_COOKIE = (
    "JSESSIONID=B190BBA7822A599CF79D6979D23D12BC; "
    "TS01be508d=0111e43782b8ff3541f9a1edbb69079f4a5f9514ba3a66c211a16bff0c75bd58695ca0bcd825e99a564f879cb67eabc058b403c0cbd2b44c2b6a5228f345fd889bfea04b6eb5d1284910c0db5fce3e9c76d1c3aa9a; "
    "Pool_ACMTJava=2907023552.17183.0000; "
    "TS017170f7=0111e437825055e44f0333cd3faa7bcb2906638988a9cb5baf5595012c87ce36fee783d5bf6788fc0edeef948729af13e42737969ebd2fb31dd98c2d15cc97adc234677048"
)

# -------------------------------------------------------------------------
# 6. CACHE DATABASE SQLITE
# -------------------------------------------------------------------------
CACHE_STANDALONE_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cache_acmt_standalone.db")
_cache_standalone_lock = threading.Lock()

def init_standalone_cache_db():
    try:
        conn = sqlite3.connect(CACHE_STANDALONE_DB, timeout=30.0)
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode=WAL;")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS acmt_cache_standalone (
                cache_key TEXT PRIMARY KEY,
                idpel TEXT,
                blth_key TEXT,
                blth_foto TEXT,
                data_blob BLOB,
                updated_at REAL
            )
        """)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sa_idpel ON acmt_cache_standalone(idpel)")
        conn.commit()
        conn.close()
    except Exception:
        pass

init_standalone_cache_db()

def save_standalone_cache_item(item, target_blth_list, target_blth_foto):
    if not item or not isinstance(item, dict):
        return
    if item.get("is_genuine_capture") is not True:
        return
    g_bytes = item.get("grid_screenshot_bytes")
    if not g_bytes or len(g_bytes) < 300:
        return
    clean_id = str(item.get("idpel", "")).strip().replace(".0", "")
    if not clean_id:
        return
    nama = item.get("nama", "-")
    if nama in ["-", "", None] and not item.get("img_bytes"):
        return

    try:
        init_standalone_cache_db()
        blth_key = "_".join(sorted(target_blth_list)) if target_blth_list else ""
        foto_key = str(target_blth_foto or "").strip()
        cache_key = f"{clean_id}_{blth_key}_{foto_key}"
        blob = pickle.dumps(item, protocol=pickle.HIGHEST_PROTOCOL)

        with _cache_standalone_lock:
            conn = sqlite3.connect(CACHE_STANDALONE_DB, timeout=30.0)
            cursor = conn.cursor()
            cursor.execute(
                "INSERT OR REPLACE INTO acmt_cache_standalone (cache_key, idpel, blth_key, blth_foto, data_blob, updated_at) VALUES (?, ?, ?, ?, ?, ?)",
                (cache_key, clean_id, blth_key, foto_key, blob, time.time())
            )
            conn.commit()
            conn.close()
    except Exception:
        pass

def get_standalone_cached_items(idpel_list, target_blth_list, target_blth_foto):
    cached_items = {}
    if not idpel_list:
        return cached_items
    try:
        init_standalone_cache_db()
        blth_key = "_".join(sorted(target_blth_list)) if target_blth_list else ""
        foto_key = str(target_blth_foto or "").strip()
        conn = sqlite3.connect(CACHE_STANDALONE_DB, timeout=30.0)
        cursor = conn.cursor()
        for idp in idpel_list:
            clean_id = str(idp).strip().replace(".0", "")
            cache_key = f"{clean_id}_{blth_key}_{foto_key}"
            cursor.execute("SELECT data_blob FROM acmt_cache_standalone WHERE cache_key = ?", (cache_key,))
            row = cursor.fetchone()
            if row and row[0]:
                try:
                    obj = pickle.loads(row[0])
                    if isinstance(obj, dict) and obj.get("is_genuine_capture") is True and obj.get("grid_screenshot_bytes"):
                        cached_items[clean_id] = obj
                except Exception:
                    pass
        conn.close()
    except Exception:
        pass
    return cached_items

def get_standalone_cache_count():
    try:
        init_standalone_cache_db()
        conn = sqlite3.connect(CACHE_STANDALONE_DB, timeout=10.0)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM acmt_cache_standalone")
        row = cursor.fetchone()
        conn.close()
        return row[0] if row else 0
    except Exception:
        return 0

def clear_standalone_cache():
    try:
        init_standalone_cache_db()
        with _cache_standalone_lock:
            conn = sqlite3.connect(CACHE_STANDALONE_DB, timeout=10.0)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM acmt_cache_standalone")
            conn.commit()
            conn.close()
        return True
    except Exception:
        return False

# -------------------------------------------------------------------------
# 7. SHARED HELPER FUNCTIONS (COOKIE, DATA PELANGGAN, FOTO METER)
# -------------------------------------------------------------------------
def clean_raw_cookie_str(raw_c):
    if not raw_c:
        return ""
    s = urllib.parse.unquote(str(raw_c).strip())
    s = re.sub(r'^[Cc]ookie:\s*', '', s).strip()
    return s

def get_acmt_cookie_header(cookie_raw):
    cookie_str = clean_raw_cookie_str(cookie_raw)
    if not cookie_str:
        return ""
    pairs = []
    if ';' in cookie_str:
        for tok in cookie_str.split(';'):
            tok = tok.strip()
            if '=' in tok:
                k, v = tok.split('=', 1)
                pairs.append((k.strip(), v.strip()))
            elif tok:
                pairs.append(('metrik_acmt_session', tok))
    else:
        if '=' in cookie_str:
            k, v = cookie_str.split('=', 1)
            pairs.append((k.strip(), v.strip()))
        elif cookie_str:
            pairs.append(('metrik_acmt_session', cookie_str))

    return "; ".join([f"{k}={v}" for k, v in pairs if k and v])

def fetch_customer_complete_data(idpel, host, cookie_hdr, proto, session=None):
    clean_id = str(idpel).strip().replace(".0", "")
    domain_raw = host.split(':')[0].split('/')[0].strip()
    unit_def = clean_id[:5] if len(clean_id) >= 5 else "12412"

    info = {
        "nama": "-", "alamat": "-", "unitup": unit_def,
        "tarif": "-", "daya": "-", "nometer": "-", "kddk": "-", "gardu": "-",
        "kdbaca_verif": "NORMAL", "kdbaca": "NORMAL",
        "petugas": "agus"
    }

    headers_acmt = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36',
        'Referer': f"{proto}{domain_raw}/acmt/Main.html",
        'Accept': 'application/json, text/javascript, text/html, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
        'Cookie': cookie_hdr,
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
    }

    acmt_candidates = [
        (f"{proto}{domain_raw}/acmt/informasiPelanggan.do", {"action": "getDetail", "idpel": clean_id}),
        (f"{proto}{domain_raw}/acmt/informasiPelanggan.do", {"action": "getPelanggan", "idpel": clean_id}),
        (f"{proto}{domain_raw}/acmt/informasiPelanggan.do", {"action": "getPelanggan", "idpel": clean_id, "unitup": unit_def}),
        (f"{proto}{domain_raw}/acmt/informasiPelanggan.do", {"action": "getDataPelanggan", "idpel": clean_id}),
        (f"{proto}{domain_raw}/acmt/informasiPelanggan.do", {"action": "getDataPelanggan", "idpel": clean_id, "unitup": unit_def}),
        (f"{proto}{domain_raw}/acmt/infoPelanggan.do", {"action": "getPelanggan", "idpel": clean_id}),
        (f"{proto}{domain_raw}/acmt/PelangganServlet", {"idpel": clean_id}),
        (f"{proto}{domain_raw}/acmt/PelangganServlet", {"idpel": clean_id, "unit": unit_def})
    ]

    http_client = session or requests

    for url_ep, params_payload in acmt_candidates:
        try:
            res = http_client.post(url_ep, data=params_payload, headers=headers_acmt, verify=False, timeout=2.5)
            if res.status_code != 200 or len(res.content) < 20:
                res = http_client.get(url_ep, params=params_payload, headers=headers_acmt, verify=False, timeout=2.5)

            if res.status_code == 200 and len(res.content) > 20:
                js_data = None
                try:
                    js_data = res.json()
                except Exception:
                    try:
                        resp_txt = res.text
                        m_n = re.search(r'["\']?(?:nama|namapel|nama_pelanggan)["\']?\s*[:=]\s*["\']?([^"\'<,\r\n]+)', resp_txt, re.I)
                        if m_n and m_n.group(1).strip() not in ["-", "", "None", "null"]:
                            info["nama"] = m_n.group(1).strip()
                        m_t = re.search(r'["\']?(?:tarif|tarip|gol_tarif)["\']?\s*[:=]\s*["\']?([^"\'<,\r\n]+)', resp_txt, re.I)
                        if m_t and m_t.group(1).strip() not in ["-", "", "None", "null"]:
                            info["tarif"] = m_t.group(1).strip()
                        m_d = re.search(r'["\']?(?:daya)["\']?\s*[:=]\s*["\']?(\d+)', resp_txt, re.I)
                        if m_d:
                            info["daya"] = m_d.group(1).strip()
                    except Exception:
                        pass

                if js_data:
                    all_records = []
                    if isinstance(js_data, dict):
                        for k_list in ["rows", "data", "list", "baca", "history", "result", "records"]:
                            v_cand = js_data.get(k_list)
                            if isinstance(v_cand, list):
                                all_records.extend(v_cand)
                    elif isinstance(js_data, list):
                        all_records.extend(js_data)

                    for rec in all_records:
                        if isinstance(rec, dict):
                            for k_r, v_r in rec.items():
                                k_ru = k_r.upper().strip()
                                val_rs = str(v_r).strip()
                                if val_rs and val_rs not in ["None", "null", "-", ""]:
                                    if any(x in k_ru for x in ["VERIF", "VERIFIKASI", "KDVERIF"]):
                                        info["kdbaca_verif"] = val_rs
                                    elif k_ru in ["KDBACA", "KODE_BACA", "BACA"]:
                                        info["kdbaca"] = val_rs

                    obj = js_data.get("data", js_data.get("result", js_data.get("rows", js_data))) if isinstance(js_data, dict) else js_data
                    if isinstance(obj, list) and len(obj) > 0:
                        obj = obj[0]

                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            k_u = k.upper().strip()
                            val_s = str(v).strip()
                            if val_s and val_s not in ["None", "null", "-", ""]:
                                if any(x == k_u for x in ["NAMA", "NAMAPEL", "NAMA_PELANGGAN", "NAMA_KONSUMEN", "NMPEL"]):
                                    info["nama"] = val_s
                                elif any(x in k_u for x in ["ALAMAT"]):
                                    info["alamat"] = val_s
                                elif any(x == k_u for x in ["TARIF", "TARIP", "GOL_TARIF", "KDTARIF"]):
                                    info["tarif"] = val_s
                                elif "DAYA" in k_u:
                                    info["daya"] = val_s
                                elif any(x in k_u for x in ["NOMETER", "NO_METER", "NOMOR_METER"]):
                                    info["nometer"] = val_s
                                elif "KDDK" in k_u:
                                    info["kddk"] = val_s
                                elif any(x in k_u for x in ["GARDU", "TIANG"]):
                                    info["gardu"] = val_s

                    if info["nama"] != "-" or info["kdbaca_verif"] != "NORMAL":
                        return info
        except Exception:
            continue

    return info

def fetch_customer_stan_history(idpel, blth_list):
    return {}

def fetch_foto_meter(idpel, blth_foto, host, cookie_hdr, proto, fallback_blth_list=None, allow_fallback=True, session=None):
    domain_raw = host.split(':')[0].split('/')[0].strip()
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0 Safari/537.36',
        'Referer': f"{proto}{domain_raw}/acmt/Main.html",
        'Accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
        'Cookie': cookie_hdr
    }

    clean_id = str(idpel).strip().replace(".0", "")
    target_blth = str(blth_foto).strip() if blth_foto else ""
    blth_candidates = []
    if target_blth:
        blth_candidates.append(target_blth)

    if fallback_blth_list and isinstance(fallback_blth_list, list):
        for b in fallback_blth_list:
            b_str = str(b).strip()
            if b_str and b_str not in blth_candidates:
                blth_candidates.append(b_str)

    base_b = target_blth or (fallback_blth_list[0] if fallback_blth_list else "")
    if base_b and allow_fallback:
        try:
            dt_cur = datetime.strptime(base_b, "%Y%m")
            for m in range(1, 4):
                y = dt_cur.year
                m_prev = dt_cur.month - m
                while m_prev <= 0:
                    m_prev += 12
                    y -= 1
                cand = f"{y}{m_prev:02d}"
                if cand not in blth_candidates:
                    blth_candidates.append(cand)
        except Exception:
            pass

    http_client = session or requests

    for b in blth_candidates:
        urls_to_try = [
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet1?idpel={clean_id}&blth={b}&fotoke=1",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet1?idpel={clean_id}&nomor_meter=&fotoke=1&blth={b}&isPhoto=null",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet1?idpel={clean_id}&nomor_meter=null&fotoke=1&blth={b}&isPhoto=null",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet?idpel={clean_id}&blth={b}&fotoke=1",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet1?idpel={clean_id}&nomor_meter=&fotoke=rumah&blth=null&isPhoto=null",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet1?idpel={clean_id}&fotoke=rumah",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet2?idpel={clean_id}&blth={b}",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet2?idpel={clean_id}&nomor_meter=null&fotoke=null&blth={b}&isPhoto=null",
            f"{proto}{domain_raw}/acmt/DisplayBlobServlet2?idpel={clean_id}"
        ]
        for u in urls_to_try:
            try:
                r = http_client.get(u, headers=headers, verify=False, timeout=3.0)
                if r.status_code == 200 and len(r.content) > 300:
                    ct = r.headers.get('Content-Type', '').lower()
                    if 'image' in ct or (r.content[:4] in [b'\xff\xd8\xff\xe0', b'\xff\xd8\xff\xe1', b'\x89PNG']):
                        return r.content, b
            except Exception:
                continue

    return None, (target_blth or "")

# -------------------------------------------------------------------------
# 8. SISTEM KLASIFIKASI FOTO KWH METER (HYBRID AI + CV+CLAHE v3)
# -------------------------------------------------------------------------
def cek_kualitas_foto(img_bytes, brightness_threshold=35.0, stddev_threshold=6.0):
    if not img_bytes or len(img_bytes) < 300:
        return True, "Foto Tidak Ditemukan / Blank"
    try:
        im = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        im_gray = im.convert("L")
        stat_gray = ImageStat.Stat(im_gray)
        mean_val = stat_gray.mean[0]
        std_val = stat_gray.stddev[0]

        if mean_val < brightness_threshold:
            return True, f"Foto Hitam/Gelap Pekat (Kecerahan: {mean_val:.1f})"

        im_small = im.resize((40, 40), Image.Resampling.BILINEAR)
        pixels = list(im_small.getdata())
        total_pixels = len(pixels)
        
        count_abu = sum(1 for r, g, b in pixels if abs(r - g) < 10 and abs(g - b) < 10 and 70 <= r <= 145)
        if (count_abu / total_pixels) > 0.75:
            return True, "Indikasi Foto Kosong: Placeholder Abu-abu ACMT"

        count_area_terang = 0
        count_warna_seragam = 0
        for r, g, b in pixels:
            if r > 140 and g > 140 and b > 130:
                count_area_terang += 1
            elif abs(r - g) < 15 and abs(g - b) < 15 and r < 100:
                count_warna_seragam += 1

        if (count_area_terang / total_pixels) < 0.03 and std_val < 22.0:
            return True, "Indikasi Salah Objek: Tanpa Karakteristik Bodi Meter"

        if (count_warna_seragam / total_pixels) > 0.60 and (count_area_terang / total_pixels) < 0.02:
            return True, "Indikasi Salah Objek: Dominasi Pagar/Struktur Gelap"

        return False, "Normal (Sesuai Karakteristik Meter)"
    except Exception as e:
        return True, f"Format Gambar Tidak Valid: {e}"

def classify_foto_meter_local(img_bytes):
    try:
        if not img_bytes:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Kosong", "engine": "Detektor Cerdas CV+AI v3"}
        
        img_pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        w, h = img_pil.size
        if w < 50 or h < 50:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Resolusi Terlalu Kecil", "engine": "Detektor Cerdas CV+AI v3"}
        
        if w > 600 or h > 600:
            img_pil.thumbnail((600, 600), Image.Resampling.LANCZOS)
            w, h = img_pil.size

        crop_h = max(20, int(h * 0.78))
        img_body = img_pil.crop((0, 0, w, crop_h))
        
        stat_body = ImageStat.Stat(img_body)
        mean_val = sum(stat_body.mean) / len(stat_body.mean)
        std_val = sum(stat_body.stddev) / len(stat_body.stddev)

        if std_val < 14:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Blank / Abu-abu Polos", "engine": "Detektor Cerdas CV+AI v3"}
        if mean_val < 25:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Gelap Pekat", "engine": "Detektor Cerdas CV+AI v3"}
        if mean_val > 245 and std_val < 18:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Putih / Blank", "engine": "Detektor Cerdas CV+AI v3"}

        img_gray_body = img_body.convert("L")
        edges_pil = img_gray_body.filter(ImageFilter.FIND_EDGES)
        edge_mean = ImageStat.Stat(edges_pil).mean[0]
        if edge_mean < 3.2:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Polos / Tanpa Detail", "engine": "Detektor Cerdas CV+AI v3"}

        import cv2
        np_b = np.array(img_body)
        h_b, w_b = np_b.shape[:2]
        total_area = h_b * w_b

        gray_cv = cv2.cvtColor(np_b, cv2.COLOR_RGB2GRAY) if len(np_b.shape) == 3 else np_b
        clahe = cv2.createCLAHE(clipLimit=2.5, tileGridSize=(8, 8))
        gray_clahe = clahe.apply(gray_cv)

        blurred = cv2.GaussianBlur(gray_clahe, (5, 5), 0)
        edges_cv = cv2.Canny(blurred, 35, 120)
        contours, _ = cv2.findContours(edges_cv, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

        center_casing_found = False
        casing_score = 0
        for c in contours:
            area = cv2.contourArea(c)
            if 0.12 * total_area < area < 0.92 * total_area:
                x_c, y_c, bw, bh = cv2.boundingRect(c)
                aspect = bw / float(bh) if bh > 0 else 0
                if 0.48 <= aspect <= 1.55:
                    cx = x_c + bw / 2.0
                    cy = y_c + bh / 2.0
                    if 0.20 * w_b < cx < 0.80 * w_b and 0.15 * h_b < cy < 0.85 * h_b:
                        center_casing_found = True
                        casing_score = max(casing_score, int((area / total_area) * 100))

        kernel_h = cv2.getStructuringElement(cv2.MORPH_RECT, (17, 3))
        morphed_h = cv2.morphologyEx(edges_cv, cv2.MORPH_CLOSE, kernel_h)
        cnts_h, _ = cv2.findContours(morphed_h, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        drum_counter_found = False
        for ch in cnts_h:
            area_h = cv2.contourArea(ch)
            if area_h > 80:
                xh, yh, wh, hh = cv2.boundingRect(ch)
                asp_h = wh / float(hh) if hh > 0 else 0
                if 2.2 <= asp_h <= 8.5 and 20 <= wh <= 0.85 * w_b and 5 <= hh <= 0.30 * h_b:
                    if 0.15 * h_b < yh < 0.85 * h_b:
                        drum_counter_found = True
                        break

        is_sky = False
        green_foliage_ratio = 0.0
        if len(np_b.shape) == 3:
            top_part = np_b[:max(10, int(h_b * 0.25)), :]
            hsv_top = cv2.cvtColor(top_part, cv2.COLOR_RGB2HSV)
            blue_mask = cv2.inRange(hsv_top, np.array([95, 40, 100]), np.array([135, 255, 255]))
            blue_ratio = np.sum(blue_mask > 0) / float(blue_mask.size)
            mean_top = np.mean(top_part)
            std_top = np.std(top_part)
            is_sky = (blue_ratio > 0.18) or (mean_top > 195 and std_top < 30)

            hsv_full = cv2.cvtColor(np_b, cv2.COLOR_RGB2HSV)
            green_mask = cv2.inRange(hsv_full, np.array([35, 40, 40]), np.array([85, 255, 255]))
            green_foliage_ratio = np.sum(green_mask > 0) / float(green_mask.size)

        margin_x = int(w_b * 0.2)
        margin_y = int(h_b * 0.2)
        center_crop = edges_cv[margin_y:h_b-margin_y, margin_x:w_b-margin_x]
        center_density = np.mean(center_crop) if center_crop.size > 0 else 0
        full_density = np.mean(edges_cv) if edges_cv.size > 0 else 0
        center_focus_ratio = (center_density / (full_density + 1e-5))

        if (center_casing_found or drum_counter_found) and not is_sky and green_foliage_ratio < 0.10:
            detail_lbl = "Bodi & Jendela Stand" if (center_casing_found and drum_counter_found) else ("Bodi Meteran" if center_casing_found else "Register Stand/Piringan")
            return {
                "is_kwh_meter": True, 
                "kategori": "KWH_METER", 
                "deskripsi": f"Fisik Meteran ({detail_lbl})", 
                "engine": "Detektor Cerdas CV+AI v3"
            }

        if center_focus_ratio >= 1.15 and casing_score >= 25 and not is_sky and green_foliage_ratio < 0.08:
            return {
                "is_kwh_meter": True, 
                "kategori": "KWH_METER", 
                "deskripsi": "Fisik Meteran (Close-up Dinding)", 
                "engine": "Detektor Cerdas CV+AI v3"
            }

        if is_sky or green_foliage_ratio >= 0.12:
            return {
                "is_kwh_meter": False, 
                "kategori": "FOTO_RUMAH", 
                "deskripsi": "Fasad Rumah / Luar Ruangan", 
                "engine": "Detektor Cerdas CV+AI v3"
            }

        return {
            "is_kwh_meter": False, 
            "kategori": "FOTO_RUMAH", 
            "deskripsi": "Fasad / Tanpa Ciri Meteran", 
            "engine": "Detektor Cerdas CV+AI v3"
        }
    except Exception:
        return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Gagal Baca Foto", "engine": "Detektor Cerdas CV+AI v3"}

def classify_foto_meter_hybrid(img_bytes, api_key=None):
    if not img_bytes:
        return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Kosong", "engine": "Detektor Cerdas CV+AI v3"}

    try:
        img_pil_fast = Image.open(io.BytesIO(img_bytes)).convert("RGB")
        w_f, h_f = img_pil_fast.size
        if w_f < 50 or h_f < 50:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Resolusi Terlalu Kecil", "engine": "Detektor Cerdas CV+AI v3"}
        crop_h = max(20, int(h_f * 0.78))
        body_fast = img_pil_fast.crop((0, 0, w_f, crop_h))
        stat_f = ImageStat.Stat(body_fast)
        std_f = sum(stat_f.stddev) / len(stat_f.stddev)
        mean_f = sum(stat_f.mean) / len(stat_f.mean)
        if std_f < 14:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Blank / Abu-abu Polos", "engine": "Detektor Cerdas CV+AI v3"}
        if mean_f < 25:
            return {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Gelap Pekat", "engine": "Detektor Cerdas CV+AI v3"}
    except Exception:
        pass

    clean_key = str(api_key or "").strip()
    if clean_key:
        def _call_gemini_worker():
            try:
                img_pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                if img_pil.width > 500 or img_pil.height > 500:
                    img_pil.thumbnail((500, 500), Image.Resampling.LANCZOS)

                prompt = (
                    "Kamu adalah asisten validator foto hasil catat meter PLN.\n"
                    "Analisis foto ini dan tentukan apakah objek utamanya adalah KWH METER listrik atau BUKAN KWH METER.\n\n"
                    "PENTING:\n"
                    "1. Abaikan banner watermark di bagian paling bawah foto.\n"
                    "2. Jika gambar hanya layar kosong/abu-abu polos/gelap pekat, WAJIB pilih kategori 'BURAM_GELAP' dan is_kwh_meter: false.\n"
                    "3. Pilihan kategori:\n"
                    "- KWH_METER: fisik kWh meter listrik (analog rol angka mekanik, prabayar keypad/LCD, 3 fasa) -> is_kwh_meter: true.\n"
                    "- FOTO_RUMAH: tampak depan rumah, bangunan, pintu, dinding, pagar, halaman.\n"
                    "- MCB_ONLY: saklar MCB / box pembatas tanpa meter.\n"
                    "- BURAM_GELAP: foto gelap pekat, putih polos, abu-abu polos/rusak, buram.\n"
                    "- LAINNYA: tiang, kabel, tanah, selfie, struk.\n\n"
                    "Jawab HANYA JSON valid:\n"
                    '{"kategori": "KWH_METER|FOTO_RUMAH|MCB_ONLY|BURAM_GELAP|LAINNYA", "is_kwh_meter": true|false, "deskripsi": "deskripsi singkat 3-5 kata"}'
                )

                resp_str = None
                try:
                    from google import genai
                    client = genai.Client(api_key=clean_key)
                    resp = client.models.generate_content(model="gemini-2.5-flash", contents=[img_pil, prompt])
                    if resp and resp.text:
                        resp_str = resp.text.strip()
                except Exception:
                    pass

                if not resp_str:
                    try:
                        import google.generativeai as genai_old
                        genai_old.configure(api_key=clean_key)
                        m = genai_old.GenerativeModel("gemini-1.5-flash")
                        resp = m.generate_content([prompt, img_pil])
                        if resp and resp.text:
                            resp_str = resp.text.strip()
                    except Exception:
                        pass

                if resp_str:
                    clean_json = resp_str
                    if clean_json.startswith("```"):
                        clean_json = re.sub(r"^```(?:json)?", "", clean_json)
                        clean_json = re.sub(r"```$", "", clean_json).strip()
                    parsed = json.loads(clean_json)
                    kat = str(parsed.get("kategori", "")).upper()
                    is_meter = bool(parsed.get("is_kwh_meter", kat == "KWH_METER"))
                    desk = str(parsed.get("deskripsi", "Teridentifikasi Gemini AI"))
                    return {
                        "is_kwh_meter": is_meter,
                        "kategori": kat or ("KWH_METER" if is_meter else "LAINNYA"),
                        "deskripsi": desk,
                        "engine": "Gemini AI v3"
                    }
            except Exception:
                pass
            return None

        try:
            with ThreadPoolExecutor(max_workers=1) as ex:
                future = ex.submit(_call_gemini_worker)
                res_api = future.result(timeout=4.0)
                if res_api:
                    return res_api
        except Exception:
            pass

    return classify_foto_meter_local(img_bytes)

def batch_classify_foto_meter(items_to_classify, active_gemini_key=None, progress_callback=None, blth_list=None, blth_foto=None):
    total = len(items_to_classify)
    if total == 0:
        return

    gemini_broken = False
    clean_key = str(active_gemini_key or "").strip()

    for idx, itm in enumerate(items_to_classify):
        img_b = itm.get("img_bytes")
        if not img_b:
            itm["ai_status"] = {"is_kwh_meter": False, "kategori": "BURAM_GELAP", "deskripsi": "Foto Kosong", "engine": "Detektor Cerdas CV+AI v3"}
            continue

        use_key = clean_key if (clean_key and not gemini_broken) else None
        try:
            ai_res = classify_foto_meter_hybrid(img_b, api_key=use_key)
            if use_key and not str(ai_res.get("engine", "")).startswith("Gemini"):
                gemini_broken = True
        except Exception:
            gemini_broken = True
            ai_res = classify_foto_meter_local(img_b)

        itm["ai_status"] = ai_res
        if blth_list:
            save_standalone_cache_item(itm, itm.get("blth_list", blth_list), blth_foto)

        if progress_callback:
            progress_callback((idx + 1) / total, f"Memeriksa objek foto {idx + 1}/{total}...")
        time.sleep(0.015)

# -------------------------------------------------------------------------
# 9. BUILDER REPORTLAB PDF & SCREENSHOT MOCKUP GRID
# -------------------------------------------------------------------------
def normalize_blth_range(start_val, end_val):
    s = str(start_val).strip()
    e = str(end_val).strip()
    if not s and not e:
        cur = datetime.now().strftime("%Y%m")
        return cur, cur, [cur]
    if not s: s = e
    if not e: e = s
    min_b = min(s, e)
    max_b = max(s, e)

    try:
        dt_start = datetime.strptime(min_b, "%Y%m")
        dt_end = datetime.strptime(max_b, "%Y%m")
        blth_list = []
        cur_dt = dt_end
        while cur_dt >= dt_start:
            blth_list.append(cur_dt.strftime("%Y%m"))
            y = cur_dt.year
            m = cur_dt.month - 1
            if m <= 0:
                m = 12
                y -= 1
            cur_dt = datetime(y, m, 1)
        return min_b, max_b, blth_list
    except Exception:
        return min_b, max_b, [max_b]

def generate_grid_screenshot(blth_list, verif_dict=None):
    if not isinstance(blth_list, list):
        blth_list = [str(blth_list)]
    if not verif_dict:
        verif_dict = {}

    cols_cfg = [
        ("BLTH", 60, "left"),
        ("KDBACA V...", 98, "left"),
        ("KDBACA", 82, "left"),
        ("G. BACA", 88, "left"),
        ("G. METER", 76, "left"),
        ("KDENTRY", 82, "center"),
        ("TGLBACA", 120, "left"),
        ("STAN LALU", 65, "right"),
        ("STAN KINI", 64, "right"),
    ]
    W = sum(c[1] for c in cols_cfg)
    header_h = 38
    row_h = 26
    total_rows = max(1, len(blth_list))
    H = header_h + (total_rows * row_h) + 1

    img = Image.new("RGB", (W, H), (255, 255, 255))
    draw = ImageDraw.Draw(img)

    f_header = load_app_font(bold=True, size=10)
    f_body = load_app_font(bold=False, size=10)

    for y_g in range(header_h):
        ratio = y_g / float(header_h)
        r = int(246 * (1 - ratio) + 233 * ratio)
        g = int(248 * (1 - ratio) + 237 * ratio)
        b = int(251 * (1 - ratio) + 242 * ratio)
        draw.line([0, y_g, W, y_g], fill=(r, g, b))

    draw.line([0, 0, W, 0], fill=(181, 200, 223))
    draw.line([0, header_h - 1, W, header_h - 1], fill=(208, 215, 224))

    cur_x = 0
    for title, col_w, align in cols_cfg:
        if cur_x > 0:
            draw.line([cur_x - 1, 1, cur_x - 1, header_h - 2], fill=(210, 216, 224))
            draw.line([cur_x, 1, cur_x, header_h - 2], fill=(255, 255, 255))
        
        bbox = f_header.getbbox(title) if hasattr(f_header, "getbbox") else (0, 0, len(title) * 6, 12)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        if align == "center":
            tx = cur_x + max(2, (col_w - tw) // 2)
        elif align == "right":
            tx = cur_x + col_w - tw - 10
        else:
            tx = cur_x + 8
        ty = max(2, (header_h - th) // 2)
        draw.text((tx, ty), title, fill=(40, 44, 52), font=f_header)
        cur_x += col_w

    for r_idx, b_val in enumerate(blth_list):
        y_top = header_h + (r_idx * row_h)
        y_bot = y_top + row_h
        bg_row = (255, 255, 255) if (r_idx % 2 == 0) else (250, 251, 253)
        draw.rectangle([0, y_top, W, y_bot], fill=bg_row)
        draw.line([0, y_bot - 1, W, y_bot - 1], fill=(232, 236, 240))

        col_x = 0
        for _, col_w, _ in cols_cfg:
            if col_x > 0:
                draw.line([col_x, y_top, col_x, y_bot - 1], fill=(240, 242, 245))
            col_x += col_w

        row_info = verif_dict.get(str(b_val), {})
        b_str = str(b_val).strip()
        v_status = str(row_info.get("kdbaca_verif", "NORMAL")).strip()
        v_kdbaca = str(row_info.get("kdbaca", "NORMAL")).strip()
        v_gbaca = str(row_info.get("g_baca", "")).strip()
        v_gmeter = str(row_info.get("g_meter", "")).strip()
        v_kdentry = str(row_info.get("kdentry", "ONLINE")).strip()
        v_tglbaca = str(row_info.get("tglbaca", "")).strip()
        v_stan_lalu = str(row_info.get("stan_lalu", "")).strip()
        v_stan_kini = str(row_info.get("stan_kini", "")).strip()

        if v_status not in ["NORMAL", "None", "", "null"] and not v_gbaca:
            v_gbaca = v_status

        vals_in_order = [
            b_str,
            v_status if v_status not in ["None", "null"] else "",
            v_kdbaca if v_kdbaca not in ["None", "null"] else "",
            v_gbaca if v_gbaca not in ["None", "null"] else "",
            v_gmeter if v_gmeter not in ["None", "null"] else "",
            v_kdentry if v_kdentry not in ["None", "null"] else "",
            v_tglbaca if v_tglbaca not in ["None", "null"] else "",
            v_stan_lalu if v_stan_lalu not in ["None", "null"] else "",
            v_stan_kini if v_stan_kini not in ["None", "null"] else ""
        ]

        cell_x = 0
        for c_i, (t_val, (_, col_w, align)) in enumerate(zip(vals_in_order, cols_cfg)):
            t_str = str(t_val)
            if len(t_str) > 14 and c_i in [1, 3, 6]:
                t_str = t_str[:12] + "..."
            
            bbox = f_body.getbbox(t_str) if hasattr(f_body, "getbbox") else (0, 0, len(t_str) * 6, 11)
            tw = bbox[2] - bbox[0]
            th = bbox[3] - bbox[1]
            if align == "center":
                tx = cell_x + max(2, (col_w - tw) // 2)
            elif align == "right":
                tx = cell_x + col_w - tw - 12
            else:
                tx = cell_x + 8
            ty = y_top + max(2, (row_h - th) // 2)
            draw.text((tx, ty), t_str, fill=(34, 34, 34), font=f_body)
            cell_x += col_w

    draw.rectangle([0, 0, W - 1, H - 1], outline=(185, 202, 222))
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def build_pdf_laporan(items_to_print, periode_label, foto_blth_label=None):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer, 
        pagesize=A4, 
        rightMargin=25, 
        leftMargin=25, 
        topMargin=30, 
        bottomMargin=30
    )
    story = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'TitleStyle',
        parent=styles['Heading1'],
        fontSize=12,
        leading=15,
        alignment=1,
        textColor=colors.HexColor("#1E3F66")
    )

    safe_blth_header = html.escape(str(periode_label))
    header_text = f"LAPORAN VERIFIKASI & FOTO STAND METER (GRID: {safe_blth_header}"
    if foto_blth_label and str(foto_blth_label) != str(periode_label):
        header_text += f" | FOTO: {html.escape(str(foto_blth_label))})"
    else:
        header_text += ")"
    story.append(Paragraph(f"<b>{header_text}</b>", title_style))
    story.append(Spacer(1, 12))

    table_data = [["IDENTITAS PELANGGAN", "STATUS VERIFIKASI (ACMT)", "FOTO METER (ACMT)"]]

    for item in items_to_print:
        idp_str = html.escape(str(item.get('idpel', '-')))
        nama_str = html.escape(str(item.get('nama', '-')))
        t_str_pdf = html.escape(str(item.get('tarif', '-')).strip())
        d_str_pdf = html.escape(str(item.get('daya', '-')).strip())
        if t_str_pdf != '-' and d_str_pdf != '-':
            td_display_pdf = f"{t_str_pdf} / {d_str_pdf} VA"
        elif t_str_pdf != '-':
            td_display_pdf = f"{t_str_pdf}"
        elif d_str_pdf != '-':
            td_display_pdf = f"{d_str_pdf} VA"
        else:
            td_display_pdf = "-"
        info_text = f"<b>IDPEL:</b> {idp_str}<br/><b>Nama:</b> {nama_str}<br/><b>Tarif/Daya:</b> {td_display_pdf}"
        p_info = Paragraph(info_text, styles['Normal'])

        grid_img_bytes = item.get("grid_screenshot_bytes")
        if grid_img_bytes:
            try:
                pil_g = Image.open(io.BytesIO(grid_img_bytes))
                gw, gh = pil_g.size
                target_gw = 270
                target_gh = max(24, min(140, int(target_gw * (gh / float(gw)))))
                cell_verif = RLImage(io.BytesIO(grid_img_bytes), width=target_gw, height=target_gh)
            except Exception:
                num_rows = len(item.get('blth_list', [1]))
                pdf_grid_h = min(120, max(24, 18 + (num_rows * 14)))
                cell_verif = RLImage(io.BytesIO(grid_img_bytes), width=270, height=pdf_grid_h)
        else:
            cell_verif = Paragraph("<font color='#888888'><i>Capture Grid ACMT Tidak Tersedia</i></font>", styles['Normal'])

        cell_foto = Paragraph("<i>Foto Tidak Tersedia</i>", styles['Normal'])
        if item.get('img_bytes'):
            try:
                pil_img = Image.open(io.BytesIO(item['img_bytes']))
                orig_w, orig_h = pil_img.size
                ratio = min(110 / orig_w, 80 / orig_h)
                rl_img = RLImage(io.BytesIO(item['img_bytes']), width=orig_w * ratio, height=orig_h * ratio)
                foto_b = html.escape(str(item.get('blth_foto', item.get('blth', '-'))))
                lbl_b = Paragraph(f"<font size=7 color='#333333'><b>BLTH:</b> {foto_b}</font>", styles['Normal'])
                cell_foto = [rl_img, Spacer(1, 2), lbl_b]
            except Exception:
                pass

        table_data.append([p_info, cell_verif, cell_foto])

    pdf_table = Table(table_data, colWidths=[130, 275, 137], repeatRows=1)
    pdf_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D2E2F6")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#143264")),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#B0C4DE"))
    ]))

    story.append(pdf_table)
    doc.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()

# -------------------------------------------------------------------------
# 10. BUILDER POWERPOINT PPT (.PPTX) - TAB 15
# -------------------------------------------------------------------------
def build_ppt_from_captures(template_bytes, captures_dict, kat_title, text_red, text_blue):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        from pptx.enum.text import PP_ALIGN
        from pptx.dml.color import RGBColor
        from pptx.enum.shapes import MSO_SHAPE
    except ImportError:
        return None, "MODUL_NOT_INSTALLED"

    prs = Presentation(io.BytesIO(template_bytes))
    if len(prs.slides) == 0:
        return None, "TEMPLATE_EMPTY"

    master_slide = prs.slides[0]
    slide_layout = master_slide.slide_layout

    C_BLACK = RGBColor(20, 24, 33)
    C_WHITE = RGBColor(255, 255, 255)
    C_RED = RGBColor(220, 38, 38)
    C_BLUE = RGBColor(2, 132, 199)

    sw = prs.slide_width
    sh = prs.slide_height

    margin_x = Inches(0.4)
    top_title = Inches(0.18)
    h_title = Inches(0.55)

    w_title = min(Inches(7.2), sw - Inches(5.0))
    x_title = (sw - w_title) / 2

    y_img = Inches(0.82)
    w_img = sw - (margin_x * 2)
    h_img = sh - y_img - Inches(1.35)

    y_ket = y_img + h_img + Inches(0.08)
    w_ket = sw - (margin_x * 2)
    h_ket = Inches(1.15)

    for idx_idp, (idp_str, ss_bytes) in enumerate(captures_dict.items()):
        if idx_idp == 0:
            curr_slide = master_slide
            for sp in list(curr_slide.shapes):
                if getattr(sp, "shape_type", None) == 13:
                    try:
                        if sp.top >= Inches(0.75):
                            sp_elem = sp._element
                            sp_elem.getparent().remove(sp_elem)
                    except Exception:
                        pass
        else:
            curr_slide = prs.slides.add_slide(slide_layout)

        cover_title = curr_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x_title, top_title, w_title, h_title)
        cover_title.fill.solid()
        cover_title.fill.fore_color.rgb = C_WHITE
        cover_title.line.fill.background()

        tb_title = curr_slide.shapes.add_textbox(x_title, top_title, w_title, h_title)
        tb_title.text_frame.word_wrap = False
        p_title = tb_title.text_frame.paragraphs[0]
        full_title_text = f"{kat_title} ({idp_str})"
        p_title.text = full_title_text
        p_title.font.bold = True
        p_title.font.size = Pt(17) if len(full_title_text) <= 45 else Pt(15)
        p_title.font.color.rgb = C_BLACK
        p_title.alignment = PP_ALIGN.CENTER

        if ss_bytes:
            curr_slide.shapes.add_picture(
                io.BytesIO(ss_bytes),
                margin_x, y_img,
                width=w_img, height=h_img
            )

        cover_ket = curr_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, margin_x, y_ket, w_ket, h_ket)
        cover_ket.fill.solid()
        cover_ket.fill.fore_color.rgb = C_WHITE
        cover_ket.line.fill.background()

        b_red = curr_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, margin_x + Inches(0.1), y_ket + Inches(0.08), Inches(0.25), Inches(0.25))
        b_red.fill.solid()
        b_red.fill.fore_color.rgb = C_RED
        b_red.line.fill.background()

        tb_kr = curr_slide.shapes.add_textbox(margin_x + Inches(0.45), y_ket + Inches(0.02), w_ket - Inches(0.55), Inches(0.40))
        p_kr = tb_kr.text_frame.paragraphs[0]
        p_kr.text = text_red
        p_kr.font.size = Pt(9.5)
        p_kr.font.color.rgb = C_BLACK

        b_blue = curr_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, margin_x + Inches(0.1), y_ket + Inches(0.52), Inches(0.25), Inches(0.25))
        b_blue.fill.solid()
        b_blue.fill.fore_color.rgb = C_BLUE
        b_blue.line.fill.background()

        tb_kb = curr_slide.shapes.add_textbox(margin_x + Inches(0.45), y_ket + Inches(0.46), w_ket - Inches(0.55), Inches(0.40))
        p_kb = tb_kb.text_frame.paragraphs[0]
        p_kb.text = text_blue
        p_kb.font.size = Pt(9.5)
        p_kb.font.color.rgb = C_BLACK

    ppt_out_stream = io.BytesIO()
    prs.save(ppt_out_stream)
    ppt_out_stream.seek(0)
    return ppt_out_stream.getvalue(), "OK"

# -------------------------------------------------------------------------
# 11. ENGINE PLAYWRIGHT & CANVAS (TAB 15 & TAB 17)
# -------------------------------------------------------------------------
def capture_canvas_extjs_t15(list_idpels, host, cookie_raw_input, proto, blth_val, box_cfg, allow_fallback=False, progress_callback=None):
    domain_raw = host.split(':')[0].split('/')[0].strip()
    cookie_hdr = get_acmt_cookie_header(cookie_raw_input)
    captured_dict = {}
    list_anomali = []

    total_ids = len(list_idpels)
    for idx, idp in enumerate(list_idpels):
        clean_idp = str(idp).strip().replace(".0", "")
        if progress_callback:
            progress_callback(idx / total_ids, f"Merender Canvas ExtJS IDPEL {clean_idp} ({idx + 1}/{total_ids})...")

        info_cust = fetch_customer_complete_data(clean_idp, domain_raw, cookie_hdr, proto)
        img_meter_bytes, blth_meter_found = fetch_foto_meter(clean_idp, blth_val, domain_raw, cookie_hdr, proto, allow_fallback=allow_fallback)

        is_bad_foto, alasan_bad = cek_kualitas_foto(img_meter_bytes)
        is_fallback = (blth_meter_found != blth_val)
        if not img_meter_bytes:
            alasan_bad = f"Foto Stand {blth_val} Kosong / Tidak Ada di ACMT"
        elif is_fallback:
            alasan_bad = f"[Fallback {blth_meter_found}] {alasan_bad}"

        if is_bad_foto or not img_meter_bytes:
            list_anomali.append({
                "idpel": clean_idp,
                "nama": info_cust.get("nama", "-"),
                "kode_unit": info_cust.get("unitup", "-"),
                "tarif": info_cust.get("tarif", "-"),
                "daya": info_cust.get("daya", "-"),
                "periode_blth": blth_meter_found or blth_val,
                "status_foto": alasan_bad,
                "waktu_cek": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

        W, H = 1600, 900
        img = Image.new("RGB", (W, H), (223, 232, 246))
        draw = ImageDraw.Draw(img)

        f_bold = load_app_font(bold=True, size=12)
        f_reg = load_app_font(bold=False, size=12)
        f_header = load_app_font(bold=True, size=13)
        f_tiny = load_app_font(bold=False, size=10)

        draw.rectangle([0, 0, W, 36], fill=(30, 65, 118))
        petugas_str = info_cust.get('petugas', 'agus')
        unit_str = info_cust.get('unitup', 'NATAL')
        draw.text((15, 10), f"Selamat Datang : {petugas_str} - (Rayon : {unit_str}) | Logout", fill=(255, 255, 255), font=f_header)

        sb_x, sb_y, sb_w, sb_h = 8, 42, 280, H - 50
        draw.rectangle([sb_x, sb_y, sb_x + sb_w, sb_y + sb_h], fill=(255, 255, 255), outline=(153, 187, 232))
        draw.rectangle([sb_x, sb_y, sb_x + sb_w, sb_y + 26], fill=(210, 226, 246), outline=(153, 187, 232))
        draw.text((sb_x + 10, sb_y + 6), "Navigasi Menu", fill=(20, 50, 100), font=f_bold)

        menus = [
            ("📂 Catat Meter Terpusat-1.0.0", True, False),
            ("   📁 Master", False, False),
            ("   📁 Proses", False, False),
            ("   📁 Monitoring", False, False),
            ("   📁 ICONNET", False, False),
            ("   📄 Informasi Pelanggan", False, True),
            ("   📄 History Pelanggan", False, False)
        ]
        cur_my = sb_y + 32
        for m_label, is_root, is_active in menus:
            if is_active:
                draw.rectangle([sb_x + 1, cur_my - 2, sb_x + sb_w - 1, cur_my + 22], fill=(223, 232, 246))
                draw.text((sb_x + 12, cur_my + 2), m_label, fill=(10, 40, 90), font=f_bold)
            else:
                draw.text((sb_x + 12, cur_my + 2), m_label, fill=(20, 20, 20) if is_root else (50, 50, 50), font=f_bold if is_root else f_reg)
            draw.line([sb_x + 5, cur_my + 24, sb_x + sb_w - 5, cur_my + 24], fill=(238, 238, 238))
            cur_my += 28

        split_x = sb_x + sb_w + 3
        draw.rectangle([split_x, sb_y, split_x + 5, sb_y + sb_h], fill=(223, 232, 246))
        mc_x = split_x + 8
        mc_w = W - mc_x - 10
        mc_h = sb_h
        draw.rectangle([mc_x, sb_y, mc_x + mc_w, sb_y + mc_h], fill=(255, 255, 255), outline=(153, 187, 232))

        draw.rectangle([mc_x, sb_y, mc_x + mc_w, sb_y + 28], fill=(210, 226, 246), outline=(153, 187, 232))
        draw.rectangle([mc_x + 6, sb_y + 4, mc_x + 180, sb_y + 28], fill=(255, 255, 255), outline=(153, 187, 232))
        draw.text((mc_x + 18, sb_y + 7), "Informasi Pelanggan  ✕", fill=(20, 50, 100), font=f_bold)

        info_p_w = 460
        draw.rectangle([mc_x + 12, sb_y + 40, mc_x + info_p_w, sb_y + 360], fill=(245, 249, 254), outline=(180, 200, 230))
        draw.rectangle([mc_x + 20, sb_y + 32, mc_x + 130, sb_y + 48], fill=(255, 255, 255))
        draw.text((mc_x + 24, sb_y + 34), "Data Pelanggan", fill=(20, 60, 120), font=f_bold)

        fields = [
            ("ID Pelanggan", clean_idp),
            ("Nama", info_cust["nama"]),
            ("Alamat", info_cust["alamat"]),
            ("Unit UP", info_cust["unitup"]),
            ("Tarif", info_cust["tarif"]),
            ("Daya", info_cust["daya"]),
            ("No Meter", info_cust["nometer"]),
            ("KDDK", info_cust["kddk"]),
            ("Gardu/Tiang", info_cust["gardu"])
        ]

        y_row = sb_y + 55
        for lbl, val in fields:
            draw.text((mc_x + 25, y_row + 2), lbl, fill=(40, 40, 40), font=f_reg)
            draw.rectangle([mc_x + 130, y_row, mc_x + info_p_w - 15, y_row + 22], fill=(255, 255, 255), outline=(180, 195, 215))
            draw.text((mc_x + 136, y_row + 3), str(val)[:38], fill=(0, 0, 0), font=f_bold if lbl == "ID Pelanggan" else f_reg)
            y_row += 28

        draw.rectangle([mc_x + 150, y_row + 5, mc_x + 230, y_row + 30], fill=(235, 242, 252), outline=(160, 185, 215))
        draw.text((mc_x + 175, y_row + 11), "Load", fill=(20, 50, 100), font=f_bold)
        draw.rectangle([mc_x + 240, y_row + 5, mc_x + 350, y_row + 30], fill=(235, 242, 252), outline=(160, 185, 215))
        draw.text((mc_x + 255, y_row + 11), "Foto Rumah", fill=(20, 50, 100), font=f_reg)

        grid_x = mc_x + info_p_w + 12
        grid_w = mc_w - info_p_w - 24
        draw.rectangle([grid_x, sb_y + 40, grid_x + grid_w, sb_y + 65], fill=(230, 238, 248), outline=(175, 198, 228))
        draw.text((grid_x + 10, sb_y + 47), "Data Baca Meter", fill=(20, 60, 120), font=f_bold)

        col_names = ["BLTH", "KDBACA V...", "KDBACA", "G. BACA", "G. METER", "KDENTRY", "TGLBACA", "STAN LALU", "STAN KINI"]
        col_widths = [70, 95, 90, 80, 80, 85, 140, 95, 95]
        
        draw.rectangle([grid_x, sb_y + 65, grid_x + grid_w, sb_y + 90], fill=(240, 244, 250), outline=(200, 210, 225))
        cx_col = grid_x
        for idx_c, c_name in enumerate(col_names):
            draw.text((cx_col + 6, sb_y + 71), c_name, fill=(50, 50, 50), font=f_bold)
            draw.line([cx_col, sb_y + 65, cx_col, sb_y + 360], fill=(225, 230, 240))
            cx_col += col_widths[idx_c]

        try:
            dt_base = datetime.strptime(str(blth_val), "%Y%m")
        except Exception:
            dt_base = datetime.now()

        daftar_blth_grid = [(dt_base - pd.DateOffset(months=r_idx)).strftime("%Y%m") for r_idx in range(9)]
        status_verif_val = info_cust.get("kdbaca_verif") or info_cust.get("kdbaca") or "NORMAL"

        for r_idx in range(9):
            gy = sb_y + 90 + (r_idx * 26)
            bg_row = (255, 255, 255) if r_idx % 2 == 0 else (248, 250, 254)
            draw.rectangle([grid_x + 1, gy, grid_x + grid_w - 1, gy + 26], fill=bg_row)

            r_blth = daftar_blth_grid[r_idx]
            r_tgl = (dt_base - pd.DateOffset(months=r_idx)).strftime("%Y-%m-24 11:20:00")
            row_data = [r_blth, status_verif_val, status_verif_val, "", "", "ONLINE", r_tgl, "-", "-"]
            cur_cx = grid_x
            for idx_c, val_c in enumerate(row_data):
                draw.text((cur_cx + 6, gy + 5), str(val_c), fill=(20, 20, 20), font=f_reg)
                cur_cx += col_widths[idx_c]
            draw.line([grid_x, gy + 26, grid_x + grid_w, gy + 26], fill=(230, 235, 245))

        bx = box_cfg["x"]
        bw = box_cfg["w"]
        y_base = box_cfg["y_base"]
        rh = box_cfg["row_h"]

        b_top = y_base + (box_cfg["blue_start"] - 1) * rh
        b_bottom = y_base + box_cfg["blue_end"] * rh
        draw.rectangle([bx, b_top, bx + bw, b_bottom], outline=(2, 132, 199), width=3)

        r_top = y_base + (box_cfg["red_start"] - 1) * rh
        r_bottom = y_base + box_cfg["red_end"] * rh
        draw.rectangle([bx, r_top, bx + bw, r_bottom], outline=(220, 38, 38), width=3)

        frame_w, frame_h = 325, 315
        frame_x = W - frame_w - 25
        frame_y = H - frame_h - 25

        draw.rectangle([frame_x, frame_y, frame_x + frame_w, frame_y + frame_h], fill=(255, 255, 255), outline=(153, 186, 226), width=2)
        draw.rectangle([frame_x + 2, frame_y + 2, frame_x + frame_w - 2, frame_y + 24], fill=(210, 226, 246))
        
        label_blth_foto = blth_meter_found if blth_meter_found else blth_val
        if is_fallback:
            label_blth_foto = f"{label_blth_foto} (Bulan Lalu)"
        draw.text((frame_x + 10, frame_y + 5), f"Foto Stand {label_blth_foto}", fill=(30, 60, 100), font=f_bold)
        draw.text((frame_x + frame_w - 45, frame_y + 5), "🗕  🗖  ✕", fill=(75, 110, 160), font=f_tiny)

        if img_meter_bytes:
            try:
                img_m = Image.open(io.BytesIO(img_meter_bytes)).convert("RGB")
                img_m = img_m.resize((frame_w - 18, frame_h - 38), Image.Resampling.LANCZOS)
                img.paste(img_m, (frame_x + 3, frame_y + 25))
                if is_bad_foto:
                    draw.rectangle([frame_x + 6, frame_y + frame_h - 28, frame_x + frame_w - 6, frame_y + frame_h - 6], fill=(220, 38, 38))
                    draw.text((frame_x + 12, frame_y + frame_h - 24), "⚠️ INDIKASI FOTO GELAP / BLANK", fill=(255, 255, 255), font=f_bold)
            except Exception:
                pass
        else:
            draw.rectangle([frame_x + 3, frame_y + 25, frame_x + frame_w - 3, frame_y + frame_h - 3], fill=(245, 245, 245))
            draw.text((frame_x + 60, frame_y + 135), "Foto Tidak Tersedia di ACMT", fill=(180, 50, 50), font=f_bold)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        captured_dict[clean_idp] = buf.getvalue()

    if progress_callback:
        progress_callback(1.0, "Render Canvas ExtJS Selesai!")

    return captured_dict, list_anomali, "OK"

async def capture_playwright_live_tab15(list_idpels, host, cookie_raw_input, proto, blth_val, box_cfg, allow_fallback=False, progress_callback=None):
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return None, [], "PLAYWRIGHT_NOT_INSTALLED"

    clean_host = host.replace("https://", "").replace("http://", "").split("/")[0].split(":")[0].strip()
    domain_clean = clean_host or "portalapp.iconpln.co.id"
    proto_clean = proto.strip() if proto.strip().endswith("://") else f"{proto.strip()}://"
    base_url = f"{proto_clean}{domain_clean}"
    cookie_hdr = get_acmt_cookie_header(cookie_raw_input)
    captured_dict = {}
    list_anomali = []
    total = len(list_idpels)

    target_url = f"{base_url}/acmt/Main.html#informasipelanggan"

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=True,
                args=["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"]
            )
            context = await browser.new_context(
                viewport={"width": 1600, "height": 900},
                ignore_https_errors=True
            )

            cookies_to_add = []
            for pair in cookie_hdr.split(";"):
                if "=" in pair:
                    k, v = pair.strip().split("=", 1)
                    if k.strip():
                        cookies_to_add.append({
                            "name": k.strip(),
                            "value": v.strip(),
                            "url": base_url
                        })
            if cookies_to_add:
                try:
                    await context.add_cookies(cookies_to_add)
                except Exception:
                    fallback_cookies = [
                        {"name": c["name"], "value": c["value"], "domain": domain_clean, "path": "/"}
                        for c in cookies_to_add
                    ]
                    await context.add_cookies(fallback_cookies)

            page = await context.new_page()
            if progress_callback:
                progress_callback(0.05, "Membuka portal ACMT via browser...")

            await page.goto(target_url, wait_until="domcontentloaded", timeout=28000)
            await page.wait_for_timeout(1200)

            await page.evaluate("""() => {
                const allDocs = [document];
                document.querySelectorAll('iframe').forEach(f => {
                    try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                });

                for (const doc of allDocs) {
                    const nodes = Array.from(doc.querySelectorAll('.x-tree-node-el, .x-tree-node-anchor, span, div'));
                    const folderInfo = nodes.find(el => {
                        const t = el.textContent.trim();
                        return t === 'Informasi' || (t.startsWith('Informasi') && !t.includes('Pelanggan'));
                    });

                    if (folderInfo) {
                        const ec = folderInfo.querySelector('.x-tree-ec-icon') || folderInfo;
                        ec.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                        folderInfo.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
                    }

                    const leaf = nodes.find(el => el.textContent.trim() === 'Informasi Pelanggan');
                    if (leaf) {
                        ['mousedown', 'mouseup', 'click', 'dblclick'].forEach(ev => {
                            leaf.dispatchEvent(new MouseEvent(ev, { bubbles: true }));
                        });
                        ['keydown', 'keypress', 'keyup'].forEach(evName => {
                            leaf.dispatchEvent(new KeyboardEvent(evName, { key: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                        });
                    }
                }
            }""")
            await page.wait_for_timeout(1000)

            tab_el = page.locator('.x-tab-strip-text:has-text("Informasi Pelanggan")').first
            if await tab_el.count() > 0:
                await tab_el.click(force=True)

            for idx, idp in enumerate(list_idpels):
                clean_id = str(idp).strip().replace(".0", "")
                if progress_callback:
                    progress_callback((idx + 1) / total, f"Memproses IDPEL {clean_id} ({idx + 1}/{total})...")

                await page.evaluate("""(idVal) => {
                    const allDocs = [document];
                    document.querySelectorAll('iframe').forEach(f => {
                        try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                    });

                    for (const doc of allDocs) {
                        const inputs = Array.from(doc.querySelectorAll('input.x-form-text, input[type="text"]'));
                        const target = inputs.find(inp => {
                            const rect = inp.getBoundingClientRect();
                            return rect.width > 50 && rect.height > 15 && rect.top > 30;
                        });

                        if (target) {
                            target.focus();
                            target.value = idVal;
                            target.dispatchEvent(new Event('input', { bubbles: true }));
                            target.dispatchEvent(new Event('change', { bubbles: true }));
                            ['keydown', 'keypress', 'keyup'].forEach(evName => {
                                target.dispatchEvent(new KeyboardEvent(evName, {
                                    key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true
                                }));
                            });
                            break;
                        }
                    }

                    for (const doc of allDocs) {
                        const btns = Array.from(doc.querySelectorAll('button, .x-btn-text'));
                        const loadBtn = btns.find(b => b.textContent.trim() === 'Load');
                        if (loadBtn) {
                            loadBtn.click();
                            break;
                        }
                    }
                }""", clean_id)

                try:
                    await page.wait_for_selector(".ext-el-mask, .x-mask-loading", state="detached", timeout=3500)
                except Exception:
                    await page.wait_for_timeout(600)

                await page.evaluate("""() => {
                    const allDocs = [document];
                    document.querySelectorAll('iframe').forEach(f => {
                        try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                    });

                    for (const doc of allDocs) {
                        const elements = Array.from(doc.querySelectorAll('a, span, div, td'));
                        const photoLink = elements.find(el => el.textContent.trim().toUpperCase().startsWith('PHOTO 1'));
                        if (photoLink) {
                            ['mousedown', 'mouseup', 'click', 'dblclick'].forEach(ev => {
                                photoLink.dispatchEvent(new MouseEvent(ev, { bubbles: true }));
                            });
                            photoLink.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                            break;
                        }
                    }
                }""")

                await page.wait_for_timeout(500)
                raw_ss_bytes = await page.screenshot(type="png", full_page=False)

                base_img = Image.open(io.BytesIO(raw_ss_bytes)).convert("RGB")
                w_ss, h_ss = base_img.size

                meter_bytes, blth_found = fetch_foto_meter(clean_id, blth_val, domain_clean, cookie_hdr, proto, allow_fallback=allow_fallback)
                is_fallback = (blth_found != blth_val)
                is_bad_foto, alasan_bad = cek_kualitas_foto(meter_bytes)
                if not meter_bytes:
                    alasan_bad = f"Foto Stand {blth_val} Kosong / Tidak Ada di ACMT"
                elif is_fallback:
                    alasan_bad = f"[Fallback {blth_found}] {alasan_bad}"

                if is_bad_foto or not meter_bytes:
                    cust_meta = fetch_customer_complete_data(clean_id, domain_clean, cookie_hdr, proto)
                    list_anomali.append({
                        "idpel": clean_id,
                        "nama": cust_meta.get("nama", "-"),
                        "kode_unit": cust_meta.get("unitup", "-"),
                        "tarif": cust_meta.get("tarif", "-"),
                        "daya": cust_meta.get("daya", "-"),
                        "periode_blth": blth_found or blth_val,
                        "status_foto": alasan_bad,
                        "waktu_cek": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

                if meter_bytes:
                    try:
                        fw, fh = 320, 270
                        fx = w_ss - fw - 25
                        fy = h_ss - fh - 50

                        draw_ss = ImageDraw.Draw(base_img)
                        f_win = load_app_font(bold=True, size=11)
                        f_ico = load_app_font(bold=False, size=10)

                        draw_ss.rectangle([fx, fy, fx + fw, fy + fh], fill=(255, 255, 255), outline=(153, 186, 226), width=2)
                        draw_ss.rectangle([fx + 2, fy + 2, fx + fw - 2, fy + 22], fill=(210, 226, 246))
                        lbl_f = f"Foto Stand {blth_found or blth_val}"
                        if is_fallback:
                            lbl_f = f"{lbl_f} (Bulan Lalu)"
                        draw_ss.text((fx + 8, fy + 5), lbl_f, fill=(30, 60, 100), font=f_win)
                        draw_ss.text((fx + fw - 45, fy + 5), "🗕  🗖  ✕", fill=(75, 110, 160), font=f_ico)

                        p_meter = Image.open(io.BytesIO(meter_bytes)).convert("RGB")
                        p_meter = p_meter.resize((fw - 16, fh - 34), Image.Resampling.LANCZOS)
                        base_img.paste(p_meter, (fx + 3, fy + 24))

                        if is_bad_foto:
                            draw_ss.rectangle([fx + 5, fy + fh - 24, fx + fw - 5, fy + fh - 5], fill=(220, 38, 38))
                            draw_ss.text((fx + 10, fy + fh - 21), "⚠️ FOTO GELAP / TIDAK TERBACA", fill=(255, 255, 255), font=f_win)
                    except Exception:
                        pass

                draw_hl = ImageDraw.Draw(base_img)
                bx = box_cfg["x"]
                bw = box_cfg["w"]
                y_base = box_cfg["y_base"]
                rh = box_cfg["row_h"]
                
                b_top = y_base + (box_cfg["blue_start"] - 1) * rh
                b_bottom = y_base + box_cfg["blue_end"] * rh
                draw_hl.rectangle([bx, b_top, bx + bw, b_bottom], outline=(2, 132, 199), width=3)

                r_top = y_base + (box_cfg["red_start"] - 1) * rh
                r_bottom = y_base + box_cfg["red_end"] * rh
                draw_hl.rectangle([bx, r_top, bx + bw, r_bottom], outline=(220, 38, 38), width=3)

                out_buf = io.BytesIO()
                base_img.save(out_buf, format="PNG")
                captured_dict[clean_id] = out_buf.getvalue()

            await browser.close()
            return captured_dict, list_anomali, "OK"

    except Exception as e_pw:
        return None, [], str(e_pw)

def run_playwright_tab15(*args, **kwargs):
    ctx = None
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
    except Exception:
        try:
            from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx
            ctx = get_script_run_ctx()
        except Exception:
            pass

    def _worker():
        if ctx:
            try:
                from streamlit.runtime.scriptrunner import add_script_run_ctx
                add_script_run_ctx(threading.current_thread(), ctx)
            except Exception:
                try:
                    from streamlit.runtime.scriptrunner_utils.script_run_context import add_script_run_ctx
                    add_script_run_ctx(threading.current_thread(), ctx)
                except Exception:
                    pass

        if sys.platform == 'win32':
            import asyncio
            policy = asyncio.WindowsProactorEventLoopPolicy()
            asyncio.set_event_loop_policy(policy)
            loop = policy.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                return loop.run_until_complete(capture_playwright_live_tab15(*args, **kwargs))
            finally:
                try:
                    loop.close()
                except Exception:
                    pass
        else:
            import asyncio
            return asyncio.run(capture_playwright_live_tab15(*args, **kwargs))

    with ThreadPoolExecutor(max_workers=1) as executor:
        fut = executor.submit(_worker)
        return fut.result()

async def capture_grid_playwright_tab17(list_idpels, host, cookie_raw_input, proto, target_blth_list, target_blth_foto=None, progress_callback=None, use_cache=True):
    try:
        from playwright.async_api import async_playwright
    except ImportError:
        return None, "PLAYWRIGHT_NOT_INSTALLED: Modul Playwright belum terpasang."

    _orig_progress_cb = progress_callback
    def _safe_progress(p, m):
        if _orig_progress_cb:
            try:
                _orig_progress_cb(p, m)
            except Exception:
                pass
    progress_callback = _safe_progress

    clean_host = host.replace("https://", "").replace("http://", "").split("/")[0].split(":")[0].strip()
    domain_clean = clean_host or "portalapp.iconpln.co.id"
    proto_clean = proto.strip() if proto.strip().endswith("://") else f"{proto.strip()}://"
    base_url = f"{proto_clean}{domain_clean}"
    cookie_hdr = get_acmt_cookie_header(cookie_raw_input)

    cached_dict = get_standalone_cached_items(list_idpels, target_blth_list, target_blth_foto) if use_cache else {}
    valid_cached = {
        k: v for k, v in cached_dict.items()
        if v.get("is_genuine_capture") is True and v.get("grid_screenshot_bytes")
    }

    results = [valid_cached[str(idp).strip().replace(".0", "")] for idp in list_idpels if str(idp).strip().replace(".0", "") in valid_cached]
    to_fetch = [idp for idp in list_idpels if str(idp).strip().replace(".0", "") not in valid_cached]
    total = len(list_idpels)

    if not to_fetch:
        if progress_callback:
            progress_callback(1.0, f"✅ Selesai! Semua {len(results)} screenshot grid asli berhasil dimuat dari Cache.")
        return results, "OK"

    target_url = f"{base_url}/acmt/Main.html#informasipelanggan"
    done_count = len(results)
    if progress_callback:
        progress_callback(done_count / total, f"💾 Ditemukan {done_count} data asli di cache. Membuka portal ACMT via Chromium untuk sisa {len(to_fetch)} IDPEL...")

    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(
                headless=True,
                args=["--disable-gpu", "--no-sandbox", "--disable-dev-shm-usage"]
            )
            context = await browser.new_context(
                viewport={"width": 1600, "height": 950},
                device_scale_factor=1.0,
                ignore_https_errors=True
            )

            cookies_to_add = []
            for pair in cookie_hdr.split(";"):
                if "=" in pair:
                    k, v = pair.strip().split("=", 1)
                    if k.strip():
                        cookies_to_add.append({
                            "name": k.strip(),
                            "value": v.strip(),
                            "url": base_url
                        })
            if cookies_to_add:
                try:
                    await context.add_cookies(cookies_to_add)
                except Exception:
                    fallback_cookies = [
                        {"name": c["name"], "value": c["value"], "domain": domain_clean, "path": "/"}
                        for c in cookies_to_add
                    ]
                    await context.add_cookies(fallback_cookies)

            page = await context.new_page()
            if progress_callback:
                progress_callback(0.05, "Membuka portal ACMT via browser...")

            await page.goto(target_url, wait_until="domcontentloaded", timeout=28000)
            await page.wait_for_timeout(1000)

            is_login = await page.evaluate("""() => {
                const u = window.location.href.toLowerCase();
                if (u.includes('login')) return true;
                const b = (document.body && document.body.innerText) ? document.body.innerText : '';
                return ((b.includes('Silahkan login') || b.includes('Silakan Login') || b.includes('Username')) && b.includes('Password'));
            }""")
            if is_login:
                await browser.close()
                return None, "SESI_EXPIRED: Sesi ACMT sudah habis. Silakan login ulang di tab ACMT Chrome, klik '📋 Salin Cookie', lalu Paste di kotak Cookie ACMT."

            try:
                await page.wait_for_selector(".x-tree-node-el, .x-panel, .x-tab-strip, #x-desktop", timeout=12000)
            except Exception:
                pass
            await page.wait_for_timeout(1000)

            for _ in range(5):
                has_input = await page.evaluate("""() => {
                    const allDocs = [document];
                    document.querySelectorAll('iframe').forEach(f => {
                        try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                    });
                    for (const doc of allDocs) {
                        const inps = Array.from(doc.querySelectorAll('input.x-form-text, input[type="text"]'));
                        if (inps.some(i => i.getBoundingClientRect().width > 50 && i.getBoundingClientRect().height > 15)) return true;
                    }
                    return false;
                }""")
                if has_input:
                    break

                tab_el = page.locator('.x-tab-strip-text:has-text("Informasi Pelanggan")').first
                if await tab_el.count() > 0 and await tab_el.is_visible():
                    await tab_el.click(force=True)
                    await page.wait_for_timeout(1000)
                    continue

                await page.evaluate("""() => {
                    const allDocs = [document];
                    document.querySelectorAll('iframe').forEach(f => {
                        try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                    });

                    for (const doc of allDocs) {
                        const nodes = Array.from(doc.querySelectorAll('.x-tree-node-el, .x-tree-node-anchor, span, div, a'));
                        const folderInfo = nodes.find(el => {
                            const t = el.textContent.trim();
                            return t === 'Informasi' || (t.startsWith('Informasi') && !t.includes('Pelanggan'));
                        });

                        if (folderInfo) {
                            const ec = folderInfo.querySelector('.x-tree-ec-icon') || folderInfo;
                            ec.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                            folderInfo.dispatchEvent(new MouseEvent('dblclick', { bubbles: true, cancelable: true }));
                        }

                        const leaf = nodes.find(el => el.textContent.trim() === 'Informasi Pelanggan');
                        if (leaf) {
                            ['mousedown', 'mouseup', 'click', 'dblclick'].forEach(ev => {
                                leaf.dispatchEvent(new MouseEvent(ev, { bubbles: true }));
                            });
                            ['keydown', 'keypress', 'keyup'].forEach(evName => {
                                leaf.dispatchEvent(new KeyboardEvent(evName, { key: 'Enter', keyCode: 13, which: 13, bubbles: true }));
                            });
                        }
                    }
                }""")
                await page.wait_for_timeout(1500)

            for idx, idp in enumerate(to_fetch):
                clean_id = str(idp).strip().replace(".0", "")
                done_count += 1
                if progress_callback:
                    progress_callback(done_count / total, f"Memproses & capture asli ACMT IDPEL {clean_id} ({done_count}/{total})...")

                if "login" in page.url.lower():
                    break

                cand_inp = None
                for frame in [page] + page.frames:
                    for sel_inp in [
                        '.x-form-item:has-text("ID Pelanggan") input',
                        '.x-form-item:has-text("IDPEL") input',
                        'xpath=//label[contains(text(),"ID Pelanggan") or contains(text(),"IDPEL")]/following::input[1]',
                    ]:
                        try:
                            loc = frame.locator(sel_inp).first
                            if await loc.count() > 0 and await loc.is_visible():
                                cand_inp = loc
                                break
                        except Exception:
                            continue
                    if cand_inp:
                        break

                if not cand_inp:
                    for frame in [page] + page.frames:
                        for sel_inp in ['input.x-form-text', 'input[type="text"]']:
                            try:
                                loc = frame.locator(sel_inp).first
                                if await loc.count() > 0 and await loc.is_visible():
                                    cand_inp = loc
                                    break
                            except Exception:
                                continue
                        if cand_inp:
                            break

                if cand_inp:
                    await cand_inp.click()
                    await cand_inp.fill("")
                    await page.wait_for_timeout(50)
                    await cand_inp.type(clean_id, delay=20)
                    await page.wait_for_timeout(80)
                    await cand_inp.press("Enter")

                await page.evaluate("""(idVal) => {
                    const allDocs = [document];
                    document.querySelectorAll('iframe').forEach(f => {
                        try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                    });

                    for (const doc of allDocs) {
                        const labels = Array.from(doc.querySelectorAll('label, .x-form-item-label, span'));
                        const idLabel = labels.find(l => {
                            const t = l.textContent.trim().toUpperCase();
                            return t === 'ID PELANGGAN' || t.startsWith('ID PELANGGAN') || t === 'IDPEL';
                        });
                        
                        let target = null;
                        if (idLabel) {
                            const container = idLabel.closest('.x-form-item') || idLabel.parentElement;
                            if (container) target = container.querySelector('input');
                        }
                        if (!target) {
                            const inputs = Array.from(doc.querySelectorAll('input.x-form-text, input[type="text"]'));
                            target = inputs.find(inp => inp.getBoundingClientRect().width > 50 && inp.getBoundingClientRect().height > 15 && inp.getBoundingClientRect().top > 30);
                        }

                        if (target && target.value !== idVal) {
                            target.focus();
                            target.value = idVal;
                            target.dispatchEvent(new Event('input', { bubbles: true }));
                            target.dispatchEvent(new Event('change', { bubbles: true }));
                            ['keydown', 'keypress', 'keyup'].forEach(evName => {
                                const ke = new KeyboardEvent(evName, {
                                    key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true, cancelable: true
                                });
                                Object.defineProperty(ke, 'keyCode', { get: () => 13 });
                                Object.defineProperty(ke, 'which', { get: () => 13 });
                                target.dispatchEvent(ke);
                            });
                        }
                    }

                    for (const doc of allDocs) {
                        const btns = Array.from(doc.querySelectorAll('button, .x-btn-text, .x-btn'));
                        const loadBtn = btns.find(b => {
                            const txt = b.textContent.trim().toUpperCase();
                            return txt === 'LOAD' || txt === 'CARI' || txt === 'SEARCH';
                        });
                        if (loadBtn) {
                            ['mousedown', 'mouseup', 'click'].forEach(ev => {
                                loadBtn.dispatchEvent(new MouseEvent(ev, { bubbles: true, cancelable: true }));
                            });
                            break;
                        }
                    }
                }""", clean_id)

                await page.keyboard.press("Enter")
                await page.wait_for_timeout(400)
                try:
                    for frame in [page] + page.frames:
                        mask = frame.locator(".ext-el-mask, .x-mask-loading")
                        if await mask.count() > 0:
                            await mask.first.wait_for(state="detached", timeout=8000)
                except Exception:
                    pass

                await page.wait_for_timeout(1800)

                for _ in range(6):
                    has_grid_rows = await page.evaluate("""() => {
                        const allDocs = [document];
                        document.querySelectorAll('iframe').forEach(f => {
                            try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                        });
                        for (const doc of allDocs) {
                            const rows = doc.querySelectorAll('.x-grid3-row, tr.x-grid3-row');
                            if (rows.length > 0) return true;
                        }
                        return false;
                    }""")
                    if has_grid_rows:
                        break
                    await page.wait_for_timeout(500)

                extracted_data = await page.evaluate("""(blthTargets) => {
                    const docsWithOffset = [{ doc: document, offX: 0, offY: 0 }];
                    document.querySelectorAll('iframe').forEach(f => {
                        try {
                            if (f.contentDocument) {
                                const r = f.getBoundingClientRect();
                                docsWithOffset.push({ doc: f.contentDocument, offX: r.left, offY: r.top });
                            }
                        } catch(e) {}
                    });

                    let nama = "-", tarif = "-", daya = "-";
                    let rowsData = {};
                    let gridBox = null;

                    for (const item of docsWithOffset) {
                        const doc = item.doc;
                        const offX = item.offX;
                        const offY = item.offY;

                        const formItems = Array.from(doc.querySelectorAll('.x-form-item'));
                        formItems.forEach(fi => {
                            const lbl = fi.querySelector('label, .x-form-item-label');
                            const inp = fi.querySelector('input');
                            const disp = fi.querySelector('.x-form-display-field, span, .x-form-element');
                            const val = (inp && inp.value ? inp.value.trim() : "") || (disp && disp.textContent ? disp.textContent.trim() : "");
                            if (lbl && val) {
                                const lt = lbl.textContent.trim().toUpperCase();
                                if (lt.includes('NAMA') && !lt.includes('GARDU')) nama = val;
                                else if (lt.includes('TARIF') || lt.includes('TARIP')) tarif = val;
                                else if (lt.includes('DAYA')) daya = val;
                            }
                        });

                        const allHd = Array.from(doc.querySelectorAll('.x-grid3-header, .x-grid3-hd-row, .x-grid-header, .x-grid3-header-inner'));
                        let targetHd = allHd.find(h => {
                            const t = h.textContent.toUpperCase();
                            return t.includes('BLTH') && (t.includes('KDBACA') || t.includes('STAN') || t.includes('G. BACA'));
                        });
                        if (!targetHd) {
                            targetHd = allHd.find(h => h.textContent.includes('BLTH'));
                        }

                        if (!targetHd) continue;

                        const headerContainer = targetHd.closest('.x-grid3-header') || targetHd;
                        let gridPanel = targetHd.closest('.x-grid-panel, .x-grid3, .x-panel') || targetHd.parentElement;

                        if (gridPanel) {
                            const scroller = gridPanel.querySelector('.x-grid3-scroller, .x-grid3-scroller-inner');
                            if (scroller) scroller.scrollLeft = 0;
                        }

                        let allRows = [];
                        if (gridPanel) {
                            allRows = Array.from(gridPanel.querySelectorAll('.x-grid3-row, tr.x-grid3-row'));
                        }
                        if (allRows.length === 0) {
                            allRows = Array.from(doc.querySelectorAll('.x-grid3-row, tr.x-grid3-row'));
                        }

                        allRows.forEach(r => {
                            const cells = Array.from(r.querySelectorAll('.x-grid3-cell, td'));
                            if (cells.length >= 1) {
                                const cellTexts = cells.map(c => {
                                    const inner = c.querySelector('.x-grid3-cell-inner') || c;
                                    return inner.textContent.trim();
                                });
                                const rBlth = cellTexts[0];
                                if (rBlth) {
                                    rowsData[rBlth] = {
                                        blth: cellTexts[0] || "",
                                        kdbaca_verif: cellTexts.length >= 2 ? cellTexts[1] : "",
                                        kdbaca: cellTexts.length >= 3 ? cellTexts[2] : "",
                                        g_baca: cellTexts.length >= 4 ? cellTexts[3] : "",
                                        g_meter: cellTexts.length >= 5 ? cellTexts[4] : "",
                                        kdentry: cellTexts.length >= 6 ? cellTexts[5] : "",
                                        tglbaca: cellTexts.length >= 7 ? cellTexts[6] : "",
                                        stan_lalu: cellTexts.length >= 8 ? cellTexts[7] : "",
                                        stan_kini: cellTexts.length >= 9 ? cellTexts[8] : ""
                                    };
                                }
                            }
                        });

                        let matchingRows = [];
                        if (allRows.length > 0) {
                            allRows.forEach(r => {
                                const cells = Array.from(r.querySelectorAll('.x-grid3-cell, td'));
                                const c0 = cells.length > 0 ? (cells[0].querySelector('.x-grid3-cell-inner') || cells[0]).textContent.trim() : "";
                                if (blthTargets && blthTargets.length > 0) {
                                    if (blthTargets.includes(c0)) {
                                        r.style.display = "";
                                        matchingRows.push(r);
                                    } else {
                                        r.style.display = "none";
                                    }
                                } else {
                                    matchingRows.push(r);
                                }
                            });

                            if (matchingRows.length === 0 && allRows.length > 0) {
                                allRows[0].style.display = "";
                                matchingRows.push(allRows[0]);
                                for (let i = 1; i < allRows.length; i++) {
                                    allRows[i].style.display = "none";
                                }
                            }

                            headerContainer.scrollIntoView({ block: 'nearest', inline: 'nearest' });

                            const hdRect = headerContainer.getBoundingClientRect();
                            let topY = offY + hdRect.top;

                            const lastRow = matchingRows[matchingRows.length - 1];
                            const lastRowRect = lastRow.getBoundingClientRect();
                            let botY = offY + lastRowRect.bottom;

                            const hdCells = Array.from(targetHd.querySelectorAll('.x-grid3-hd, th, td, div'));
                            const blthEl = hdCells.find(el => el.textContent.trim().toUpperCase() === 'BLTH');
                            let leftX = offX + (blthEl ? blthEl.getBoundingClientRect().left : hdRect.left);

                            let rightX = 0;
                            const stanKiniEl = hdCells.find(el => {
                                const t = el.textContent.trim().toUpperCase();
                                return t === 'STAN KINI' || t.includes('STAN KINI');
                            });
                            if (stanKiniEl) {
                                rightX = offX + stanKiniEl.getBoundingClientRect().right;
                            } else {
                                const firstRowCells = Array.from(matchingRows[0].querySelectorAll('.x-grid3-cell, td'));
                                if (firstRowCells.length >= 9) {
                                    rightX = offX + firstRowCells[8].getBoundingClientRect().right;
                                } else if (firstRowCells.length > 0) {
                                    rightX = offX + firstRowCells[firstRowCells.length - 1].getBoundingClientRect().right;
                                }
                            }

                            if (!rightX || rightX <= leftX) {
                                rightX = leftX + 780;
                            }

                            gridBox = {
                                x: Math.max(0, Math.floor(leftX)),
                                y: Math.max(0, Math.floor(topY)),
                                width: Math.min(1400, Math.ceil(rightX - leftX) + 1),
                                height: Math.min(800, Math.ceil(botY - topY) + 1)
                            };
                            break;
                        }
                    }

                    return { nama, tarif, daya, rowsData, gridBox };
                }""", target_blth_list)

                grid_img_bytes = None
                box = extracted_data.get("gridBox")
                if box and box.get("width", 0) > 100 and box.get("height", 0) > 20:
                    try:
                        vw, vh = 1600, 950
                        cx = max(0, int(box["x"]))
                        cy = max(0, int(box["y"]))
                        cw = max(10, min(int(box["width"]), vw - cx))
                        ch = max(10, min(int(box["height"]), vh - cy))
                        grid_img_bytes = await page.screenshot(clip={
                            "x": cx,
                            "y": cy,
                            "width": cw,
                            "height": ch
                        })
                    except Exception:
                        pass

                try:
                    await page.evaluate("""() => {
                        const allDocs = [document];
                        document.querySelectorAll('iframe').forEach(f => {
                            try { if (f.contentDocument) allDocs.push(f.contentDocument); } catch(e){}
                        });
                        allDocs.forEach(d => {
                            d.querySelectorAll('.x-grid3-row, tr.x-grid3-row').forEach(r => {
                                r.style.display = '';
                            });
                        });
                    }""")
                except Exception:
                    pass

                if not grid_img_bytes:
                    for frame in [page] + page.frames:
                        for p_sel in [
                            '.x-grid-panel:has-text("BLTH")',
                            '.x-grid3:has-text("BLTH")',
                            '.x-panel:has-text("Data Pelanggan")'
                        ]:
                            try:
                                p_loc = frame.locator(p_sel).first
                                if await p_loc.count() > 0 and await p_loc.is_visible():
                                    grid_img_bytes = await p_loc.screenshot()
                                    break
                            except Exception:
                                continue
                        if grid_img_bytes:
                            break

                is_genuine = bool(grid_img_bytes and len(grid_img_bytes) > 500)

                nama_val = extracted_data.get("nama", "-")
                tarif_val = extracted_data.get("tarif", "-")
                daya_val = extracted_data.get("daya", "-")

                if nama_val in ["-", "", "None"] or tarif_val in ["-", "", "None"]:
                    info_db = fetch_customer_complete_data(clean_id, domain_clean, cookie_hdr, proto)
                    if info_db.get("nama") and info_db["nama"] != "-":
                        nama_val = info_db["nama"]
                    if info_db.get("tarif") and info_db["tarif"] != "-":
                        tarif_val = info_db["tarif"]
                    if info_db.get("daya") and info_db["daya"] != "-":
                        daya_val = info_db["daya"]

                b_foto_req = str(target_blth_foto).strip() if target_blth_foto else (target_blth_list[0] if target_blth_list else "")
                img_meter_bytes, blth_found = fetch_foto_meter(clean_id, b_foto_req, domain_clean, cookie_hdr, proto, fallback_blth_list=target_blth_list)

                item_res = {
                    "idpel": clean_id,
                    "nama": nama_val,
                    "tarif": tarif_val,
                    "daya": daya_val,
                    "blth": blth_found or b_foto_req,
                    "blth_foto": blth_found or b_foto_req,
                    "blth_list": target_blth_list,
                    "verif_dict": extracted_data.get("rowsData", {}),
                    "grid_screenshot_bytes": grid_img_bytes,
                    "img_bytes": img_meter_bytes,
                    "capture_type": "grid_table_v2",
                    "is_genuine_capture": is_genuine
                }
                results.append(item_res)
                if is_genuine:
                    save_standalone_cache_item(item_res, target_blth_list, target_blth_foto)

            try:
                await browser.close()
            except Exception:
                pass

            id_order = {str(idp).strip().replace(".0", ""): i for i, idp in enumerate(list_idpels)}
            results.sort(key=lambda x: id_order.get(x['idpel'], 9999))

            if len(results) == 0:
                return None, "GAGAL_CAPTURE: Tidak ada screenshot asli yang berhasil diambil. Periksa Cookie ACMT."
            elif len(results) < len(list_idpels):
                return results, "PARTIAL_OK: Sesi ACMT terputus sebelum seluruh IDPEL selesai."
            return results, "OK"

    except Exception as e:
        if results:
            id_order = {str(idp).strip().replace(".0", ""): i for i, idp in enumerate(list_idpels)}
            results.sort(key=lambda x: id_order.get(x['idpel'], 9999))
            return results, f"PARTIAL_OK: {str(e)}"
        return None, str(e)

def run_playwright_tab17(*args, **kwargs):
    ctx = None
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        ctx = get_script_run_ctx()
    except Exception:
        try:
            from streamlit.runtime.scriptrunner_utils.script_run_context import get_script_run_ctx
            ctx = get_script_run_ctx()
        except Exception:
            pass

    def _worker():
        if ctx:
            try:
                from streamlit.runtime.scriptrunner import add_script_run_ctx
                add_script_run_ctx(threading.current_thread(), ctx)
            except Exception:
                try:
                    from streamlit.runtime.scriptrunner_utils.script_run_context import add_script_run_ctx
                    add_script_run_ctx(threading.current_thread(), ctx)
                except Exception:
                    pass

        if sys.platform == 'win32':
            import asyncio
            policy = asyncio.WindowsProactorEventLoopPolicy()
            asyncio.set_event_loop_policy(policy)
            loop = policy.new_event_loop()
            asyncio.set_event_loop(loop)
            try:
                return loop.run_until_complete(capture_grid_playwright_tab17(*args, **kwargs))
            finally:
                try:
                    loop.close()
                except Exception:
                    pass
        else:
            import asyncio
            return asyncio.run(capture_grid_playwright_tab17(*args, **kwargs))

    with ThreadPoolExecutor(max_workers=1) as executor:
        fut = executor.submit(_worker)
        return fut.result()

def fetch_single_worker_t17(idp_raw, target_blth_list, target_blth_foto, host_param, cookie_header_param, proto_param, session=None):
    clean_idp = str(idp_raw).strip().replace(".0", "")
    info_cust = fetch_customer_complete_data(clean_idp, host_param, cookie_header_param, proto_param, session=session)
    b_foto_req = str(target_blth_foto).strip() if target_blth_foto else (target_blth_list[0] if target_blth_list else "")
    img_bytes, blth_found = fetch_foto_meter(clean_idp, b_foto_req, host_param, cookie_header_param, proto_param, fallback_blth_list=target_blth_list, session=session)
    
    status_verif_teks = info_cust.get("kdbaca_verif", "NORMAL")
    kdbaca_teks = info_cust.get("kdbaca", "NORMAL")
    use_blth = blth_found or b_foto_req

    verif_map = {
        b: {"kdbaca_verif": status_verif_teks, "kdbaca": kdbaca_teks}
        for b in target_blth_list
    }
    grid_img = generate_grid_screenshot(target_blth_list, verif_map)

    return {
        "idpel": clean_idp,
        "nama": info_cust.get("nama", "-"),
        "tarif": info_cust.get("tarif", "-"),
        "daya": info_cust.get("daya", "-"),
        "blth": use_blth,
        "blth_foto": use_blth,
        "blth_list": target_blth_list,
        "verif_dict": verif_map,
        "grid_screenshot_bytes": grid_img,
        "img_bytes": img_bytes,
        "capture_type": "grid_table_v2",
        "is_genuine_capture": False
    }

def fetch_all_idpels_parallel(idpel_list, target_blth_list, target_blth_foto, host_param, cookie_header_param, proto_param, max_workers=8, use_cache=True):
    total_items = len(idpel_list)
    cached_dict = get_standalone_cached_items(idpel_list, target_blth_list, target_blth_foto) if use_cache else {}
    results = [cached_dict[str(idp).strip().replace(".0", "")] for idp in idpel_list if str(idp).strip().replace(".0", "") in cached_dict]
    to_fetch = [idp for idp in idpel_list if str(idp).strip().replace(".0", "") not in cached_dict]

    if not to_fetch:
        id_order = {str(idp).strip().replace(".0", ""): i for i, idp in enumerate(idpel_list)}
        results.sort(key=lambda x: id_order.get(x['idpel'], 9999))
        return results

    prog_bar = st.progress(len(results) / total_items)
    status_txt = st.empty()
    done_count = len(results)
    status_txt.caption(f"💾 Ditemukan {done_count} data di cache. Mengambil {len(to_fetch)} sisa IDPEL secara paralel...")

    shared_session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers * 2)
    shared_session.mount("https://", adapter)
    shared_session.mount("http://", adapter)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_idp = {
            executor.submit(fetch_single_worker_t17, idp, target_blth_list, target_blth_foto, host_param, cookie_header_param, proto_param, shared_session): idp
            for idp in to_fetch
        }
        for future in as_completed(future_to_idp):
            done_count += 1
            try:
                res = future.result()
                if res:
                    results.append(res)
                    save_standalone_cache_item(res, target_blth_list, target_blth_foto)
            except Exception:
                pass
            prog_bar.progress(done_count / total_items)
            status_txt.caption(f"⚡ Sedang memproses paralel {done_count}/{total_items} IDPEL...")

    id_order = {str(idp).strip().replace(".0", ""): i for i, idp in enumerate(idpel_list)}
    results.sort(key=lambda x: id_order.get(x['idpel'], 9999))

    status_txt.empty()
    prog_bar.empty()
    return results

# -------------------------------------------------------------------------
# 12. SISTEM AUTENTIKASI PENGGUNA (LOGIN GATE)
# -------------------------------------------------------------------------
VALID_USERS = {
    "antonarmayadi": "@Nton"
}

if not st.session_state.get("authenticated", False):
    c_pad1, c_card, c_pad2 = st.columns([1.2, 1.6, 1.2])
    with c_card:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            """
            <div style="background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 12px; padding: 25px 25px 15px 25px; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
                <h3 style="text-align: center; color: #1e3f66; margin-top: 0;">🔐 Login ACMT Tools Cloud</h3>
                <p style="text-align: center; color: #6c757d; font-size: 13px; margin-bottom: 20px;">Silakan masukkan username dan password untuk mengakses aplikasi.</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        with st.form("login_form", clear_on_submit=False):
            u_input = st.text_input("Username:", placeholder="Masukkan username...", key="input_username").strip()
            p_input = st.text_input("Password:", type="password", placeholder="Masukkan password...", key="input_password")
            btn_submit = st.form_submit_button("🚀 Masuk / Login", type="primary", use_container_width=True)

            if btn_submit:
                if u_input in VALID_USERS and VALID_USERS[u_input] == p_input:
                    st.session_state["authenticated"] = True
                    st.session_state["logged_user"] = u_input
                    st.rerun()
                else:
                    st.error("⚠️ Username atau Password salah!")
    st.stop()

# Header Bar Utama + Tombol Download Keep-Alive (.ZIP)
col_hdr1, col_hdr_zip, col_hdr2 = st.columns([2.6, 1.1, 0.9])
with col_hdr1:
    st.markdown("## ⚡ ACMT Tools Cloud: PPT & PDF Foto Meter")
    st.caption("Aplikasi Mandiri Cloud • Auto-Capture PPT (Tab 15) & Laporan Verifikasi PDF (Tab 17)")
with col_hdr_zip:
    st.markdown("<div style='padding-top: 12px;'></div>", unsafe_allow_html=True)
    st.download_button(
        label="📦 Download Keep-Alive (.ZIP)",
        data=get_keepalive_zip_bytes(),
        file_name="acmt-keepalive.zip",
        mime="application/zip",
        use_container_width=True,
        help="Download ekstensi Chrome ACMT Keep-Alive & Salin Cookie (.ZIP) untuk dipasang di laptop lain"
    )
with col_hdr2:
    st.markdown(f"<div style='text-align: right; padding-top: 4px;'><span style='background-color:#e8f4fd; color:#0d6efd; padding:4px 10px; border-radius:6px; font-weight:bold; font-size:12px;'>👤 {st.session_state.get('logged_user', 'antonarmayadi')}</span></div>", unsafe_allow_html=True)
    if st.button("🚪 Logout", key="btn_logout_user", use_container_width=True):
        st.session_state["authenticated"] = False
        st.session_state.pop("logged_user", None)
        st.rerun()

cookie_auto = ""
if os.path.exists(COOKIE_SYNC_FILE):
    try:
        with open(COOKIE_SYNC_FILE, "r", encoding="utf-8") as f:
            cookie_auto = f.read().strip()
    except Exception:
        pass

if not cookie_auto or "JSESSIONID" not in cookie_auto:
    cookie_auto = DEFAULT_ACMT_COOKIE

tab_ppt, tab_pdf = st.tabs([
    "📊 1. Auto-Capture Layar ke PPT (.pptx) [Tab 15]",
    "📋 2. Pratinjau & PDF Foto Meter Asli ACMT [Tab 17]"
])

# =========================================================================
# TAB 1: AUTO CAPTURE KE PPT (.PPTX)
# =========================================================================
with tab_ppt:
    st.subheader("🤖 Auto-Capture Layar Web ACMT & Generate PPT")
    st.caption("Tarik layar informasi pelanggan ACMT otomatis ke template PowerPoint lengkap dengan audit foto anomali.")

    cur_15 = str(st.session_state.get("t15_cookie_pw", "")).strip()
    if not cur_15 or "JSESSIONID" not in cur_15:
        st.session_state["t15_cookie_pw"] = cookie_auto

    cookie_saved_15 = st.session_state.get("t15_cookie_pw", cookie_auto)

    with st.expander("⚙️ Pengaturan Server & Cookie Session ACMT (PPT)", expanded=not bool(cookie_saved_15)):
        c_cfg1, c_cfg2, c_cfg3 = st.columns([1.2, 0.8, 2.0])
        with c_cfg1:
            server_host_t15 = st.text_input("Server Host:", value="portalapp.iconpln.co.id", key="sa_t15_host")
        with c_cfg2:
            proto_acmt_t15 = st.selectbox("Protokol:", options=["https://", "http://"], index=0, key="sa_t15_proto")
        with c_cfg3:
            cookie_acmt_t15_input = st.text_input(
                "Cookie ACMT (Wajib Aktif):", 
                value=cookie_saved_15,
                placeholder="Klik '📋 Salin Cookie' di web ACMT lalu Paste (Ctrl+V) di sini", 
                key="sa_t15_cookie"
            )

    c_up1, c_up2 = st.columns([1.8, 2.2])
    with c_up1:
        uploaded_template = st.file_uploader(
            "1. Unggah Template PPT Master (.pptx):",
            type=["pptx"],
            key="sa_uploaded_master_template",
            help="Unggah file PPT yang memiliki master header/footer template"
        )
        blth_eval_input = st.text_input(
            "Periode Evaluasi Terkini (BLTH):", 
            value=datetime.now().strftime("%Y%m"), 
            key="sa_t15_blth"
        )

        if "sa_t15_last_blth_synced" not in st.session_state:
            st.session_state["sa_t15_last_blth_synced"] = blth_eval_input

        if st.session_state["sa_t15_last_blth_synced"] != blth_eval_input:
            st.session_state["sa_t15_title"] = f"Pelanggan LBKB {blth_eval_input}"
            st.session_state["sa_t15_ket_red"] = f"Pada periode baca bulan {blth_eval_input} pelanggan masuk kedalam kriteria tidak dapat stan/statis"
            st.session_state["sa_t15_last_blth_synced"] = blth_eval_input

        allow_fallback_foto = st.checkbox(
            "🔄 Izinkan Fallback ke foto bulan lalu jika foto bulan evaluasi kosong",
            value=False,
            key="sa_t15_allow_fallback_foto"
        )

    with c_up2:
        input_idpel_t15 = st.text_area(
            "2. Masukkan Daftar IDPEL (1 IDPEL per baris / pisahkan koma):",
            placeholder="124120113668\n124120118576",
            height=130,
            key="sa_t15_paste_idp"
        )

    st.markdown("##### 📝 Kustomisasi Judul & Keterangan Slide")
    c_k1, c_k2, c_k3 = st.columns([1.2, 1.4, 1.4])
    with c_k1:
        judul_kategori_t15 = st.text_input("Kategori Slide:", value=f"Pelanggan LBKB {blth_eval_input}", key="sa_t15_title")
    with c_k2:
        ket_merah_t15 = st.text_area(
            "Keterangan Kotak Merah:",
            value=f"Pada periode baca bulan {blth_eval_input} pelanggan masuk kedalam kriteria tidak dapat stan/statis",
            height=65,
            key="sa_t15_ket_red"
        )
    with c_k3:
        ket_biru_t15 = st.text_area(
            "Keterangan Kotak Biru:",
            value="Pada periode saat ini pelanggan sudah tidak masuk kedalam kategori tsb, foto stan dapat terbaca dan termonitor.",
            height=65,
            key="sa_t15_ket_blue"
        )

    with st.expander("🎯 Pengaturan Bingkai Garis Sorot (Biru & Merah)", expanded=False):
        c_rng1, c_rng2 = st.columns(2)
        with c_rng1:
            st.markdown("🔵 **Rentang Kotak Biru (Stan Terkini / Normal):**")
            c_b1, c_b2 = st.columns(2)
            with c_b1:
                blue_start = st.number_input("Dari Baris ke-:", min_value=1, max_value=12, value=1, step=1, key="sa_t15_b_start")
            with c_b2:
                blue_end = st.number_input("Sampai Baris ke-:", min_value=1, max_value=12, value=1, step=1, key="sa_t15_b_end")

        with c_rng2:
            st.markdown("🔴 **Rentang Kotak Merah (Periode Evaluasi / LBKB):**")
            c_r1, c_r2 = st.columns(2)
            with c_r1:
                red_start = st.number_input("Dari Baris ke-:", min_value=1, max_value=12, value=2, step=1, key="sa_t15_r_start")
            with c_r2:
                red_end = st.number_input("Sampai Baris ke-:", min_value=1, max_value=12, value=3, step=1, key="sa_t15_r_end")

        st.caption("Kalibrasi Koordinat Grid Tabel (Piksel):")
        c_pos1, c_pos2, c_pos3, c_pos4 = st.columns(4)
        with c_pos1:
            cfg_box_x = st.slider("Posisi X Kiri (px):", min_value=300, max_value=1200, value=635, step=5, key="sa_t15_box_x")
        with c_pos2:
            cfg_box_w = st.slider("Lebar Garis (px):", min_value=400, max_value=1200, value=895, step=5, key="sa_t15_box_w")
        with c_pos3:
            cfg_box_y_top = st.slider("Posisi Y Baris ke-1 (px):", min_value=100, max_value=250, value=149, step=1, key="sa_t15_box_y")
        with c_pos4:
            cfg_row_h = st.slider("Tinggi Tiap Baris (px):", min_value=15, max_value=35, value=21, step=1, key="sa_t15_box_h")

    mode_capture = st.radio(
        "🛠️ Pilih Mesin Penangkapan Layar:",
        options=[
            "🌐 Live Playwright Headless Browser (Interaksi Real-time 'Load' & Grid DOM)",
            "⚡ Direct API + ExtJS Canvas Viewport (Cepat)"
        ],
        index=0,
        horizontal=True,
        key="sa_t15_mode_engine"
    )

    col_act1, col_act2 = st.columns([1.5, 1.5])
    with col_act1:
        btn_proc_capture = st.button("🚀 Jalankan Auto-Capture & Generate PPT", type="primary", use_container_width=True, key="sa_btn_proc_capture")

    if btn_proc_capture:
        if not uploaded_template:
            st.error("⚠️ Silakan upload file Template PPT master (.pptx) terlebih dahulu.")
        elif not input_idpel_t15.strip():
            st.error("⚠️ Silakan masukkan minimal 1 IDPEL pada kotak di atas.")
        elif not cookie_acmt_t15_input.strip():
            st.error("⚠️ Cookie ACMT belum diisi/aktif.")
        else:
            raw_ids = input_idpel_t15.replace('\n', ',').replace(' ', ',').replace('\t', ',').split(',')
            daftar_idpel_final = [i.strip() for i in raw_ids if i.strip() != ""]

            template_bytes = uploaded_template.getvalue()
            prog_bar = st.progress(0.0)
            status_text = st.empty()

            def update_progress(val, msg):
                prog_bar.progress(val)
                status_text.caption(f"⏳ {msg}")

            box_config = {
                "x": cfg_box_x,
                "w": cfg_box_w,
                "y_base": cfg_box_y_top,
                "row_h": cfg_row_h,
                "blue_start": int(min(blue_start, blue_end)),
                "blue_end": int(max(blue_start, blue_end)),
                "red_start": int(min(red_start, red_end)),
                "red_end": int(max(red_start, red_end))
            }

            if "Playwright" in mode_capture:
                captures_dict, list_anomali, status_cap = run_playwright_tab15(
                    daftar_idpel_final,
                    server_host_t15,
                    cookie_acmt_t15_input,
                    proto_acmt_t15,
                    blth_eval_input,
                    box_config,
                    allow_fallback=allow_fallback_foto,
                    progress_callback=update_progress
                )
            else:
                captures_dict, list_anomali, status_cap = capture_canvas_extjs_t15(
                    daftar_idpel_final,
                    server_host_t15,
                    cookie_acmt_t15_input,
                    proto_acmt_t15,
                    blth_eval_input,
                    box_config,
                    allow_fallback=allow_fallback_foto,
                    progress_callback=update_progress
                )

            if not captures_dict:
                st.error(f"⚠️ Gagal menangkap layar ACMT: {status_cap}")
            else:
                status_text.caption("🎨 Merakit slide PowerPoint...")
                res_ppt_bytes, status_ppt = build_ppt_from_captures(
                    template_bytes,
                    captures_dict,
                    judul_kategori_t15,
                    ket_merah_t15,
                    ket_biru_t15
                )

                if res_ppt_bytes:
                    st.session_state["sa_capture_ppt_output_bytes"] = res_ppt_bytes
                    if list_anomali:
                        df_anomali = pd.DataFrame(list_anomali)
                        st.session_state["sa_anomali_foto_csv"] = df_anomali.to_csv(index=False).encode('utf-8')
                        st.session_state["sa_anomali_foto_df"] = df_anomali
                    else:
                        st.session_state["sa_anomali_foto_csv"] = None
                        st.session_state["sa_anomali_foto_df"] = None

                    prog_bar.progress(1.0)
                    status_text.empty()
                    st.success(f"🟢 Selesai! Berhasil memproses {len(captures_dict)} IDPEL.")
                    st.rerun()

    with col_act2:
        if st.session_state.get("sa_capture_ppt_output_bytes"):
            nama_file_out = f"Laporan_ACMT_AutoCapture_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
            st.download_button(
                label="📥 Download File PPT (.PPTX)",
                data=st.session_state["sa_capture_ppt_output_bytes"],
                file_name=nama_file_out,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
                key="sa_dl_ppt"
            )

            if st.session_state.get("sa_anomali_foto_csv"):
                df_bad = st.session_state.get("sa_anomali_foto_df")
                total_bad = len(df_bad)
                st.warning(f"⚠️ Terdeteksi **{total_bad} IDPEL** memiliki anomali (foto kosong, gelap pekat, atau salah objek).")
                st.download_button(
                    label=f"📄 Download Daftar {total_bad} Foto Anomali (CSV)",
                    data=st.session_state["sa_anomali_foto_csv"],
                    file_name=f"audit_foto_anomali_{blth_eval_input}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    key="sa_dl_csv_anomali"
                )
                with st.expander("👁️ Pratinjau Daftar Pelanggan Anomali Foto", expanded=False):
                    st.dataframe(df_bad, use_container_width=True)


# =========================================================================
# TAB 2: PRATINJAU & GENERATE PDF FOTO METER (TAB 17)
# =========================================================================
with tab_pdf:
    st.subheader("📋 Pratinjau & Generate PDF Foto Meter")
    st.caption("Status Verifikasi ditampilkan dalam format screenshot grid ExtJS asli ACMT (Gambar 2 & Gambar 3).")

    cur_prev = str(st.session_state.get("sa_t17_cookie", "")).strip()
    if not cur_prev or "JSESSIONID" not in cur_prev:
        st.session_state["sa_t17_cookie"] = cookie_auto

    cookie_saved_17 = st.session_state.get("sa_t17_cookie", cookie_auto)

    with st.expander("⚙️ Konfigurasi Akses ACMT & AI (PDF)", expanded=not bool(cookie_saved_17)):
        c_s1, c_s2, c_s3 = st.columns([1.2, 0.8, 2.0])
        with c_s1:
            t17_host = st.text_input("Server Host:", value="portalapp.iconpln.co.id", key="sa_t17_host")
        with c_s2:
            t17_proto = st.selectbox("Protokol:", options=["https://", "http://"], index=0, key="sa_t17_proto")
        with c_s3:
            t17_cookie = st.text_input(
                "Cookie ACMT:", 
                key="sa_t17_cookie",
                placeholder="Klik '📋 Salin Cookie' di pojok kanan bawah ACMT lalu Paste (Ctrl+V) di sini..."
            )
            st.caption("💡 *Klik tombol hijau '📋 Salin Cookie' di pojok kanan bawah web ACMT lalu Paste (`Ctrl+V`) di kotak atas.*")

        st.markdown("---")
        c_ai1, c_ai2 = st.columns([2.4, 1.6])
        with c_ai1:
            t17_gemini_key = st.text_input(
                "🤖 Google Gemini API Key (Opsional - Cloud Free Tier):",
                value=st.session_state.get("saved_gemini_api_key", os.getenv("GEMINI_API_KEY", "")),
                type="password",
                key="sa_t17_gemini_key",
                help="Gratis resmi dari Google AI Studio (aistudio.google.com). Jika kosong, otomatis menggunakan Detektor Cerdas Lokal."
            )
            if t17_gemini_key:
                st.session_state["saved_gemini_api_key"] = t17_gemini_key
        with c_ai2:
            st.caption("💡 **Status Mesin AI Hybrid:**")
            if t17_gemini_key.strip():
                st.markdown("<span style='color:green; font-weight:bold;'>🟢 Gemini Vision AI Aktif</span>", unsafe_allow_html=True)
            else:
                st.markdown("<span style='color:#0d6efd; font-weight:bold;'>🔵 Detektor CV+CLAHE Aktif</span>", unsafe_allow_html=True)

    c_p1, c_p2 = st.columns([2.2, 1.8])
    with c_p1:
        st.markdown("📅 **Pilihan Rentang Grid ACMT & Foto Meter:**")
        now_dt = datetime.now()
        if now_dt.day >= 20:
            nxt_m = now_dt.month + 1
            nxt_y = now_dt.year
            if nxt_m > 12:
                nxt_m = 1
                nxt_y += 1
            default_blth = f"{nxt_y}{nxt_m:02d}"
        else:
            default_blth = now_dt.strftime("%Y%m")

        c_bl1, c_bl2 = st.columns(2)
        with c_bl1:
            t17_blth_start = st.text_input("Dari BLTH (Grid):", value=default_blth, key="sa_t17_blth_start").strip()
        with c_bl2:
            t17_blth_end = st.text_input("Sampai BLTH (Grid):", value="", placeholder=f"Kosongkan jika hanya 1 bulan ({t17_blth_start or default_blth})", key="sa_t17_blth_end").strip()

        min_blth, max_blth, target_blth_list = normalize_blth_range(t17_blth_start, t17_blth_end)
        periode_label = f"{min_blth} - {max_blth}" if min_blth != max_blth else min_blth
        periode_file = f"{min_blth}_{max_blth}" if min_blth != max_blth else min_blth

        foto_options = list(target_blth_list)
        manual_opt = "✏️ Ketik Manual BLTH Lain..."
        foto_options.append(manual_opt)

        curr_choice = st.session_state.get("sa_t17_last_foto_choice", target_blth_list[0])
        idx_choice = foto_options.index(curr_choice) if curr_choice in foto_options else 0

        sel_foto = st.selectbox("📸 BLTH Foto Meter yang Ditampilkan:", options=foto_options, index=idx_choice, key="sa_t17_sel_foto")
        st.session_state["sa_t17_last_foto_choice"] = sel_foto

        if sel_foto == manual_opt:
            t17_blth_foto = st.text_input("Ketik BLTH Foto Spesifik (YYYYMM):", value=max_blth, key="sa_t17_custom_foto").strip()
        else:
            t17_blth_foto = sel_foto

    with c_p2:
        t17_idpels_text = st.text_area(
            "Daftar IDPEL (1 IDPEL per baris):",
            placeholder="124100064608\n124150326711",
            height=130,
            key="sa_t17_idpels_text"
        )

    mode_t17 = st.radio(
        "🛠️ Pilih Metode Penangkapan Grid ACMT:",
        options=[
            "🌐 Live Playwright Screen Capture (Screenshot Grid Asli ACMT)",
            "⚡ Fast Parallel API (Cepat)"
        ],
        index=0,
        horizontal=True,
        key="sa_t17_mode_capture"
    )

    def extract_idpels(raw_text):
        raw = raw_text.replace('\n', ',').replace(' ', ',').replace('\t', ',').split(',')
        return [i.strip() for i in raw if i.strip() != ""]

    inputted_idpels = extract_idpels(t17_idpels_text)
    cached_for_inputs = get_standalone_cached_items(inputted_idpels, target_blth_list, t17_blth_foto) if inputted_idpels else {}
    total_db_cached = get_standalone_cache_count()

    c_chk1, c_chk2 = st.columns([2.5, 1.5])
    with c_chk1:
        chk_use_cache = st.checkbox("💾 Gunakan Auto-Resume & Cache (Lewati IDPEL yang sudah tersimpan)", value=True, key="sa_chk_cache")
        chk_auto_ai = st.checkbox("🤖 Otomatis Klasifikasi AI (Deteksi kWh Meter vs Foto Rumah)", value=False, key="sa_chk_ai")
    with c_chk2:
        if total_db_cached > 0:
            if st.button("🗑️ Hapus Cache Tersimpan", key="sa_btn_clear_cache"):
                clear_standalone_cache()
                st.session_state.pop("sa_t17_preview_data", None)
                st.session_state.pop("sa_t17_pdf_bytes_ready", None)
                st.toast("✅ Seluruh cache berhasil dibersihkan!")
                st.rerun()

    col_b1, col_b2 = st.columns([1.5, 1.5])
    with col_b1:
        btn_fetch_preview = st.button("🔄 Tarik Data & Tampilkan Pratinjau", type="secondary", use_container_width=True, key="sa_btn_preview")
    with col_b2:
        btn_direct_pdf = st.button("⚡ Tarik & Langsung Generate PDF", type="primary", use_container_width=True, key="sa_btn_direct")

    if btn_direct_pdf:
        idpel_list = extract_idpels(t17_idpels_text)
        if not idpel_list:
            st.error("⚠️ Masukkan minimal 1 IDPEL.")
        elif not t17_cookie.strip():
            st.error("⚠️ Cookie ACMT belum diisi.")
        else:
            st.session_state.pop("sa_t17_preview_data", None)
            st.session_state.pop("sa_t17_pdf_bytes_ready", None)
            cookie_hdr = get_acmt_cookie_header(t17_cookie)
            prog_bar = st.progress(0.0)
            status_txt = st.empty()

            if "Playwright" in mode_t17:
                direct_items, st_cap = run_playwright_tab17(
                    idpel_list, t17_host, t17_cookie, t17_proto, target_blth_list, target_blth_foto=t17_blth_foto,
                    progress_callback=lambda p, m: (prog_bar.progress(p), status_txt.caption(f"⏳ {m}")),
                    use_cache=chk_use_cache
                )
                if not direct_items:
                    st.error(f"❌ Gagal: {st_cap}")
            else:
                direct_items = fetch_all_idpels_parallel(idpel_list, target_blth_list, t17_blth_foto, t17_host, cookie_hdr, t17_proto, max_workers=8, use_cache=chk_use_cache)

            prog_bar.empty()
            status_txt.empty()

            if direct_items:
                with st.spinner("Sedang merakit PDF..."):
                    try:
                        pdf_bytes = build_pdf_laporan(direct_items, periode_label, t17_blth_foto)
                        st.session_state["sa_t17_pdf_bytes_ready"] = pdf_bytes
                        st.session_state["sa_t17_direct_mode"] = True
                        st.session_state["sa_t17_active_periode_file"] = periode_file
                        st.session_state["sa_t17_active_periode_label"] = periode_label
                        st.session_state["sa_t17_active_blth_foto"] = t17_blth_foto
                        st.success(f"✅ Selesai! PDF untuk {len(direct_items)} pelanggan siap diunduh.")
                    except Exception as e:
                        st.error(f"⚠️ Gagal merakit PDF: {e}")

    if btn_fetch_preview:
        idpel_list = extract_idpels(t17_idpels_text)
        if not idpel_list:
            st.error("⚠️ Masukkan minimal 1 IDPEL.")
        elif not t17_cookie.strip():
            st.error("⚠️ Cookie ACMT belum diisi.")
        else:
            st.session_state.pop("sa_t17_pdf_bytes_ready", None)
            st.session_state["sa_t17_direct_mode"] = False
            cookie_hdr = get_acmt_cookie_header(t17_cookie)
            prog_bar = st.progress(0.0)
            status_txt = st.empty()

            if "Playwright" in mode_t17:
                preview_data, st_cap = run_playwright_tab17(
                    idpel_list, t17_host, t17_cookie, t17_proto, target_blth_list, target_blth_foto=t17_blth_foto,
                    progress_callback=lambda p, m: (prog_bar.progress(p), status_txt.caption(f"⏳ {m}")),
                    use_cache=chk_use_cache
                )
                if not preview_data:
                    st.error(f"❌ Gagal: {st_cap}")
            else:
                preview_data = fetch_all_idpels_parallel(idpel_list, target_blth_list, t17_blth_foto, t17_host, cookie_hdr, t17_proto, max_workers=8, use_cache=chk_use_cache)

            prog_bar.empty()
            status_txt.empty()

            if preview_data:
                if chk_auto_ai:
                    active_gemini_key = st.session_state.get("saved_gemini_api_key", os.getenv("GEMINI_API_KEY", "")).strip()
                    to_c = [itm for itm in preview_data if itm.get("img_bytes")]
                    if to_c:
                        with st.spinner(f"🤖 Memverifikasi {len(to_c)} foto meter dengan AI..."):
                            batch_classify_foto_meter(to_c, active_gemini_key, blth_list=target_blth_list, blth_foto=t17_blth_foto)

                st.session_state["sa_t17_preview_data"] = preview_data
                for idx, itm in enumerate(preview_data):
                    if chk_auto_ai and itm.get("ai_status"):
                        st.session_state[f"sa_chk_pdf_{idx}"] = bool(itm["ai_status"].get("is_kwh_meter"))
                    else:
                        st.session_state[f"sa_chk_pdf_{idx}"] = True
                st.session_state["sa_t17_active_periode_file"] = periode_file
                st.session_state["sa_t17_active_periode_label"] = periode_label
                st.session_state["sa_t17_active_blth_foto"] = t17_blth_foto
                st.success(f"✅ Pratinjau {len(preview_data)} data berhasil dimuat!")

    active_p_file = st.session_state.get("sa_t17_active_periode_file", periode_file)
    active_p_label = st.session_state.get("sa_t17_active_periode_label", periode_label)
    active_p_foto = st.session_state.get("sa_t17_active_blth_foto", t17_blth_foto)

    if st.session_state.get("sa_t17_direct_mode") and "sa_t17_pdf_bytes_ready" in st.session_state:
        st.markdown("---")
        st.markdown("##### 📥 File PDF Siap Diunduh:")
        file_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        st.download_button(
            label="📥 Download File PDF Laporan",
            data=st.session_state["sa_t17_pdf_bytes_ready"],
            file_name=f"Laporan_Verifikasi_{active_p_file}_{file_stamp}.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True,
            key="sa_dl_direct_ready"
        )

    if not st.session_state.get("sa_t17_direct_mode") and st.session_state.get("sa_t17_preview_data"):
        st.markdown("---")
        st.markdown(f"##### 🔍 Pratinjau Status Verifikasi & Foto Meter (Grid: {active_p_label} | Foto: {active_p_foto}):")

        def toggle_all_sa(status: bool):
            for idx in range(len(st.session_state.get("sa_t17_preview_data", []))):
                st.session_state[f"sa_chk_pdf_{idx}"] = status

        c_btn_all1, c_btn_all2, c_btn_ai1, c_btn_ai2, c_btn_re = st.columns([1.0, 1.0, 1.8, 1.8, 1.6])
        with c_btn_all1:
            st.button("☑️ Pilih Semua", on_click=toggle_all_sa, args=(True,), use_container_width=True, key="sa_btn_sel_all")
        with c_btn_all2:
            st.button("⬜ Batal Semua", on_click=toggle_all_sa, args=(False,), use_container_width=True, key="sa_btn_desel_all")
        with c_btn_ai1:
            btn_filter_kwh = st.button("🎯 Centang Hanya kWh Meter", use_container_width=True, key="sa_btn_filter_kwh")
        with c_btn_ai2:
            btn_filter_non = st.button("⚠️ Centang Hanya Bukan Meter", use_container_width=True, key="sa_btn_filter_non")
        with c_btn_re:
            btn_re_ai = st.button("🔄 Analisis Ulang AI", use_container_width=True, key="sa_btn_re_ai")

        active_gemini_key = st.session_state.get("saved_gemini_api_key", os.getenv("GEMINI_API_KEY", "")).strip()

        if btn_filter_kwh or btn_filter_non or btn_re_ai:
            preview_items = st.session_state.get("sa_t17_preview_data", [])
            to_classify = [it for it in preview_items if it.get("img_bytes")] if btn_re_ai else [
                it for it in preview_items if it.get("img_bytes") and not it.get("ai_status")
            ]

            if to_classify:
                prog_ai = st.progress(0.0)
                status_ai = st.empty()
                batch_classify_foto_meter(
                    to_classify,
                    active_gemini_key=active_gemini_key,
                    progress_callback=lambda p, m: (prog_ai.progress(p), status_ai.caption(f"🤖 {m}")),
                    blth_list=target_blth_list,
                    blth_foto=active_p_foto
                )
                prog_ai.empty()
                status_ai.empty()

            for idx, itm in enumerate(preview_items):
                ai = itm.get("ai_status")
                is_kwh = bool(ai.get("is_kwh_meter")) if ai else True
                if btn_filter_kwh or btn_re_ai:
                    st.session_state[f"sa_chk_pdf_{idx}"] = is_kwh
                elif btn_filter_non:
                    st.session_state[f"sa_chk_pdf_{idx}"] = not is_kwh

            st.rerun()

        selected_items = []
        h1, h2, h3, h4 = st.columns([0.4, 1.6, 3.8, 1.4])
        h1.markdown("**Pilih**")
        h2.markdown("**Identitas Pelanggan**")
        h3.markdown("**Status Verifikasi (ACMT)**")
        h4.markdown("**Thumbnail Foto Meter**")

        for i, item in enumerate(st.session_state["sa_t17_preview_data"]):
            cols = st.columns([0.4, 1.6, 3.8, 1.4])
            with cols[0]:
                chk_key = f"sa_chk_pdf_{i}"
                if chk_key not in st.session_state:
                    st.session_state[chk_key] = True
                is_checked = st.checkbox("Pilih", key=chk_key, label_visibility="collapsed")
            with cols[1]:
                t_str = str(item.get('tarif', '-')).strip()
                d_str = str(item.get('daya', '-')).strip()
                td_display = f"{t_str} / {d_str} VA" if (t_str != '-' and d_str != '-') else (t_str if t_str != '-' else (f"{d_str} VA" if d_str != '-' else "-"))
                st.markdown(f"**IDPEL:** {item['idpel']}<br/>**Nama:** {item.get('nama', '-')}<br/>**Tarif/Daya:** {td_display}", unsafe_allow_html=True)
            with cols[2]:
                grid_bytes = item.get("grid_screenshot_bytes")
                if grid_bytes and item.get("is_genuine_capture"):
                    st.markdown("<span style='background-color:#d4edda; color:#155724; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:bold;'>🟢 CAPTURE ASLI ACMT</span>", unsafe_allow_html=True)
                    st.image(grid_bytes, use_container_width=True)
                elif grid_bytes:
                    st.image(grid_bytes, use_container_width=True)
                else:
                    st.warning("⚠️ Screenshot asli belum tersedia.")
            with cols[3]:
                foto_b = item.get('blth_foto', item.get('blth', active_p_foto))
                st.caption(f"📸 **BLTH:** `{foto_b}`")
                if item.get('img_bytes'):
                    st.image(item['img_bytes'], width=125)
                    ai = item.get('ai_status')
                    if ai:
                        kat = ai.get('kategori', '')
                        if ai.get('is_kwh_meter'):
                            st.markdown("<span style='background-color:#d4edda; color:#155724; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>🟢 KWH METER</span>", unsafe_allow_html=True)
                        elif kat == 'FOTO_RUMAH':
                            st.markdown("<span style='background-color:#f8d7da; color:#721c24; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>🔴 FOTO RUMAH</span>", unsafe_allow_html=True)
                        elif kat == 'MCB_ONLY':
                            st.markdown("<span style='background-color:#fff3cd; color:#856404; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>🟠 HANYA MCB</span>", unsafe_allow_html=True)
                        elif kat == 'BURAM_GELAP':
                            st.markdown("<span style='background-color:#e2e3e5; color:#383d41; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>⚫ GELAP / BURAM</span>", unsafe_allow_html=True)
                        else:
                            st.markdown("<span style='background-color:#f8d7da; color:#721c24; padding:2px 6px; border-radius:4px; font-size:11px; font-weight:bold;'>🔴 BUKAN METER</span>", unsafe_allow_html=True)
                        if ai.get('deskripsi'):
                            st.caption(f"_{ai.get('deskripsi')}_")
                else:
                    st.warning("Foto tidak ada")

            if is_checked:
                selected_items.append(item)
            st.markdown("---")

        col_gen1, col_gen2 = st.columns([2, 1])
        with col_gen1:
            if st.button("📄 Generate PDF dari Pilihan di Atas", type="primary", use_container_width=True, key="sa_btn_gen_pdf"):
                if not selected_items:
                    st.warning("⚠️ Belum ada data pelanggan yang dicentang.")
                    st.session_state.pop("sa_t17_pdf_bytes_ready", None)
                else:
                    with st.spinner("Sedang merakit PDF dengan screenshot grid..."):
                        try:
                            pdf_bytes = build_pdf_laporan(selected_items, active_p_label, active_p_foto)
                            st.session_state["sa_t17_pdf_bytes_ready"] = pdf_bytes
                            st.success("🟢 PDF Berhasil Dirender! Silakan unduh file di samping.")
                        except Exception as e:
                            st.error(f"⚠️ Gagal merakit PDF: {e}")

        with col_gen2:
            if "sa_t17_pdf_bytes_ready" in st.session_state:
                file_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                st.download_button(
                    label="📥 Download PDF",
                    data=st.session_state["sa_t17_pdf_bytes_ready"],
                    file_name=f"Laporan_Verifikasi_{active_p_file}_{file_stamp}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True,
                    key="sa_dl_btn_prev"
                )
            else:
                st.info("ℹ️ Klik tombol Generate terlebih dahulu.")

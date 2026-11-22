# -*- coding: utf-8 -*-
"""
POS UMKM BULO JAYA (Sembako & Kebutuhan Harian)
Jalankan:  streamlit run app.py
Fitur: Kasir (scan barcode / foto AI / pilih gambar), harga eceran-grosir otomatis,
stok otomatis berkurang, tambah-kurang stok (tombol cepat di galeri), pemantauan omset-modal-untung,
transaksi keuangan, struk, role Owner & Kasir.
Tambahan (gabungan app v2): keranjang tersimpan di database (aman mati lampu), SQLite WAL, barcode EAN-13 otomatis
untuk barang >= Rp5.000 + cetak barcode, 3 akun kasir awal.
Katalog: 100 produk toko kelontong (sesuai file Database_100_produk_toko_kelontong.docx), foto kemasan asli diunduh otomatis dari internet.
"""
import os, io, re, json, html, base64, hashlib, hmac, secrets, sqlite3, time, shutil, threading, textwrap
import http.client, urllib.request, urllib.parse, urllib.error
import datetime as dt
from contextlib import closing
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

DB = os.environ.get("POS_DB", "bulojaya_pos.db")
IMG_DIR = "gambar_produk"
os.makedirs(IMG_DIR, exist_ok=True)
TZ = dt.timezone(dt.timedelta(hours=8))  # WITA (Makassar)
HARGA_BARCODE_MIN = 5000  # barang dengan harga eceran >= ini otomatis dibuatkan barcode EAN-13

st.set_page_config(page_title="POS Bulo Jaya", layout="wide", initial_sidebar_state="expanded")
# ── Ikon navigasi (Lucide, stroke tunggal) disuntik lewat CSS: tidak butuh library/file tambahan ──
_LUCIDE = {
    "keranjang": '<circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/>',
    "stok": '<path d="m7.5 4.27 9 5.15"/><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/>',
    "dashboard": '<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>',
    "keuangan": '<path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1"/><path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4"/>',
    "pengaturan": '<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>',
    "keluar": '<path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" x2="9" y1="12" y2="12"/>',
}

def _icon_uri(name, color="#DCEAE1"):
    svg = (f"<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='{color}' stroke-width='2' "
           f"stroke-linecap='round' stroke-linejoin='round'>{_LUCIDE[name]}</svg>")
    return 'url("data:image/svg+xml,' + urllib.parse.quote(svg, safe="") + '")'

def _nav_css():
    """Urutan ikon = urutan menu di main(): Kasir, Stok Barang, Dashboard, Keuangan, Pengaturan."""
    out = [f'section[data-testid="stSidebar"] [role="radiogroup"]>label:nth-of-type({i})::before{{background-image:{_icon_uri(n)}}}'
           for i, n in enumerate(("keranjang", "stok", "dashboard", "keuangan", "pengaturan"), 1)]
    out.append(f'.st-key-sb_bottom button::before{{background-image:{_icon_uri("keluar", "#FFFFFF")}}}')
    return "\n".join(out)

def tulis_tema():
    """Warna utama Streamlit (radio, checkbox, fokus input) dibuat hijau lewat .streamlit/config.toml.
    Hanya dibuat bila belum ada; berlaku setelah aplikasi dijalankan ulang satu kali."""
    try:
        p = os.path.join(".streamlit", "config.toml")
        if not os.path.exists(p):
            os.makedirs(".streamlit", exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                f.write('[theme]\nbase = "light"\nprimaryColor = "#16834A"\nbackgroundColor = "#F5F6F8"\n'
                        'secondaryBackgroundColor = "#FFFFFF"\ntextColor = "#1E293B"\nfont = "sans serif"\n\n'
                        '[browser]\ngatherUsageStats = false\n')
    except OSError:
        pass

tulis_tema()
CSS = r"""
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
:root{--bg:#F5F6F8;--side:#173D2B;--pri:#16834A;--pri-d:#12703F;--card:#FFFFFF;--tx:#1E293B;--mu:#64748B;--bd:#E5E7EB;--warn:#B45309;--bad:#DC2626}
.stApp h1,.stApp h2,.stApp h3,.stApp h4,.stApp p,.stApp label,.stApp button,.stApp input,.stApp textarea,.stApp [data-baseweb="tab"],.stApp [data-testid="stMetricValue"],.stApp [data-testid="stMetricLabel"],.stApp .stCaption,.stApp li{font-family:'Inter',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif}
.stApp{background:var(--bg);color:var(--tx)}
[data-testid="stHeader"]{background:transparent}
.block-container{padding:1.4rem 2rem 2.5rem;max-width:1500px}
footer,.stDeployButton{visibility:hidden}
.stApp h3{font-size:1rem;font-weight:600;color:var(--tx);margin:.4rem 0 .6rem}
.stApp [data-testid="stCaptionContainer"],.stApp .stCaption{color:var(--mu)}
/* ---------- Sidebar ---------- */
section[data-testid="stSidebar"]{background:var(--side);border-right:0}
section[data-testid="stSidebar"] *{color:#DCEAE1}
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"]>div[data-testid="stVerticalBlock"]{min-height:calc(100vh - 5rem)}
.sbrand{display:flex;gap:11px;align-items:center;padding:2px 4px 16px;margin-bottom:14px;border-bottom:1px solid rgba(255,255,255,.12)}
.sbrand .mk{width:36px;height:36px;border-radius:8px;background:var(--pri);display:grid;place-items:center;font-weight:700;font-size:.85rem;letter-spacing:.4px;flex:none}
section[data-testid="stSidebar"] .sbrand .mk{color:#fff}
.sbrand .bn{font-weight:600;font-size:.95rem;line-height:1.2}
section[data-testid="stSidebar"] .sbrand .bn{color:#fff}
.sbrand .bs{font-size:.72rem;color:#9DB8A8}
.userbox{display:flex;justify-content:space-between;align-items:center;background:rgba(255,255,255,.07);border-radius:6px;padding:8px 12px;font-size:.82rem;margin-bottom:16px}
.userbox em{font-style:normal;font-size:.7rem;text-transform:uppercase;letter-spacing:.06em;color:#9DB8A8}
.navlabel{font-size:.68rem;text-transform:uppercase;letter-spacing:.09em;color:#8FAF9D;padding:0 12px;margin-bottom:4px}
section[data-testid="stSidebar"] [role="radiogroup"]{gap:2px;width:100%}
section[data-testid="stSidebar"] [role="radiogroup"]>label{display:flex;align-items:center;gap:11px;width:100%;margin:0;padding:9px 12px;border-radius:6px;cursor:pointer;transition:background .12s}
section[data-testid="stSidebar"] [role="radiogroup"]>label::before{content:"";width:18px;height:18px;flex:none;background-repeat:no-repeat;background-position:center;background-size:contain;opacity:.8}
section[data-testid="stSidebar"] [role="radiogroup"]>label>div:nth-of-type(1){display:none}
section[data-testid="stSidebar"] [role="radiogroup"]>label p{font-size:.9rem;font-weight:500}
section[data-testid="stSidebar"] [role="radiogroup"]>label:hover{background:rgba(255,255,255,.08)}
section[data-testid="stSidebar"] [role="radiogroup"]>label:has(input:checked){background:rgba(255,255,255,.14);box-shadow:inset 3px 0 0 #4FBF85}
section[data-testid="stSidebar"] [role="radiogroup"]>label:has(input:checked) *{color:#fff;font-weight:600}
section[data-testid="stSidebar"] [role="radiogroup"]>label:has(input:checked)::before{opacity:1}
.st-key-sb_bottom{margin-top:auto}
.st-key-sb_bottom button{display:flex;align-items:center;justify-content:center;gap:8px;background:transparent;border:1px solid rgba(255,255,255,.28);color:#fff;border-radius:6px}
.st-key-sb_bottom button:hover{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.5);color:#fff;box-shadow:none}
.st-key-sb_bottom button::before{content:"";width:16px;height:16px;background-repeat:no-repeat;background-size:contain}
/* ---------- Header halaman ---------- */
.phead{display:flex;align-items:baseline;justify-content:space-between;gap:16px;padding:0 0 12px;margin-bottom:16px;border-bottom:1px solid var(--bd)}
.phead .tt{font-weight:700;font-size:1.4rem;color:var(--tx);letter-spacing:-.01em;line-height:1.2}
.phead .ss{color:var(--mu);font-size:.85rem;text-align:right}
.statbar{display:flex;background:var(--card);border:1px solid var(--bd);border-radius:8px;margin:0 0 12px;overflow:hidden}
.statbar .it{flex:1;display:flex;justify-content:space-between;align-items:baseline;gap:10px;padding:9px 16px;border-right:1px solid var(--bd)}
.statbar .it:last-child{border-right:0}
.statbar span{color:var(--mu);font-size:.78rem}
.statbar b{font-size:.98rem;font-weight:700;color:var(--tx)}
.statbar b.warn{color:var(--warn)}
.promo{font-size:.78rem;color:var(--mu);margin:0 0 14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.promo b{color:var(--tx);font-weight:600}
.sech{font-weight:600;color:var(--tx);font-size:1rem;margin:6px 0 10px}
.alertbar{background:#FFFBEB;border:1px solid #FCE7A8;border-left:3px solid #D97706;color:#7C4A03;border-radius:6px;padding:9px 14px;margin:4px 0 12px;font-size:.85rem}
/* ---------- Kartu produk ---------- */
[class*="st-key-pc_"],[class*="st-key-pg_"]{background:var(--card);border:1px solid var(--bd);border-radius:8px;overflow:hidden;gap:0!important;transition:border-color .12s,box-shadow .12s}
[class*="st-key-pc_"]:hover,[class*="st-key-pg_"]:hover{border-color:#9CCBB1;box-shadow:0 2px 8px rgba(15,23,42,.08)}
[class*="st-key-pc_"] [data-testid="stButton"]{padding:0 10px 10px}
[class*="st-key-pc_"] .stButton>button{min-height:34px;background:#fff;color:var(--pri);border:1px solid var(--pri)}
[class*="st-key-pc_"] .stButton>button:hover:not(:disabled){background:var(--pri);color:#fff}
[class*="st-key-pc_"] .stButton>button:disabled{color:#94A3B8;border-color:var(--bd);background:#F8FAFC}
[class*="st-key-pg_"] [data-testid="stHorizontalBlock"]{padding:0 10px 10px;gap:8px}
.pimg{position:relative;aspect-ratio:4/3;background:#fff;overflow:hidden;border-bottom:1px solid #F1F5F9}
.pimg img{width:100%;height:100%;object-fit:cover;display:block}
.pimg.out img{filter:grayscale(1);opacity:.5}
.tg{position:absolute;font-size:.66rem;font-weight:600;padding:2px 7px;border-radius:4px;color:#fff;letter-spacing:.02em}
.tg.tl{top:8px;left:8px}.tg.tr{top:8px;right:8px}
.tg.low{background:#D97706}.tg.out{background:var(--bad)}.tg.cart{background:var(--pri)}
.pbody{padding:10px 12px 8px}
.pbody .nm{font-weight:600;color:var(--tx);font-size:.86rem;line-height:1.3;min-height:2.6em;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
.pbody .pr{font-weight:700;color:var(--tx);font-size:1rem;margin:4px 0 1px}
.pbody .pr small{font-weight:400;color:var(--mu);font-size:.72rem}
.pbody .gr{font-size:.72rem;color:var(--mu);margin-bottom:8px}
.pbody .gr.none{visibility:hidden}
.sbar{height:4px;border-radius:99px;background:#EEF0F3;overflow:hidden}
.sbar i{display:block;height:100%;border-radius:99px}
.sinfo{display:flex;justify-content:space-between;font-size:.72rem;color:var(--mu);margin-top:4px}
/* ---------- Filter kategori ---------- */
.st-key-cat_k [role="radiogroup"]{gap:6px;flex-wrap:wrap}
.st-key-cat_k label[data-baseweb="radio"]{background:var(--card);border:1px solid var(--bd);border-radius:6px;padding:5px 12px;margin:0;transition:.12s}
.st-key-cat_k label[data-baseweb="radio"] p{font-size:.82rem;font-weight:500}
.st-key-cat_k label[data-baseweb="radio"]:hover{border-color:var(--pri)}
.st-key-cat_k label[data-baseweb="radio"]>div:nth-of-type(1){display:none}
.st-key-cat_k label[data-baseweb="radio"]:has(input:checked){background:var(--pri);border-color:var(--pri)}
.st-key-cat_k label[data-baseweb="radio"]:has(input:checked) *{color:#fff;font-weight:600}
/* ---------- Keranjang ---------- */
.st-key-cartpanel{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:16px 16px 12px}
div[data-testid="stHorizontalBlock"]:has(>div[data-testid="stColumn"] .st-key-cartpanel){align-items:flex-start}
div[data-testid="stColumn"]:has(>div .st-key-cartpanel){position:sticky;top:3.2rem;max-height:calc(100vh - 4.5rem);overflow-y:auto}
.carttitle{font-weight:700;font-size:1.05rem;color:var(--tx);display:flex;justify-content:space-between;align-items:center;padding-bottom:10px;margin-bottom:4px;border-bottom:1px solid var(--bd)}
.carttitle .pill{background:#EEF2F6;color:var(--mu);border-radius:4px;font-size:.74rem;font-weight:600;padding:2px 8px}
[class*="st-key-ci_"]{border-bottom:1px solid #F1F5F9;padding:10px 0 8px}
.cihead{display:flex;gap:10px;align-items:center}
.cihead img{width:44px;height:44px;border-radius:6px;object-fit:cover;border:1px solid var(--bd);background:#fff}
.cihead .nm{font-weight:600;color:var(--tx);font-size:.88rem;line-height:1.25}
.cihead .sb{font-size:.74rem;color:var(--mu)}
.empty{text-align:center;padding:30px 12px;border:1px dashed #CBD5E1;border-radius:8px;background:#FAFBFC;color:var(--tx);font-size:.9rem}
.empty small{display:block;margin-top:4px;color:var(--mu);font-size:.8rem}
.badge-g{background:#E7F5EC;color:var(--pri);border:1px solid #BFE3CD;padding:1px 8px;border-radius:4px;font-weight:600;font-size:.72rem}
.badge-e{background:#F1F5F9;color:var(--mu);padding:1px 8px;border-radius:4px;font-size:.72rem;font-weight:500}
.total{display:flex;justify-content:space-between;align-items:flex-end;padding:12px 0 6px;margin-top:6px;border-top:1px solid var(--bd)}
.total small{display:block;font-size:.72rem;font-weight:600;color:var(--mu);text-transform:uppercase;letter-spacing:.07em}
.total .amt{font-size:1.7rem;font-weight:700;color:var(--tx);line-height:1.2;letter-spacing:-.01em}
.total b{color:var(--mu);font-size:.8rem;font-weight:500}
/* ---------- Peringkat terlaris ---------- */
.rank{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:10px;position:relative}
.rank img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:6px;background:#fff}
.rank .md{position:absolute;top:16px;left:16px;background:var(--side);color:#fff;font-size:.7rem;font-weight:700;border-radius:4px;padding:1px 7px}
.rank .nm{font-weight:600;font-size:.84rem;color:var(--tx);margin-top:8px;line-height:1.25}
.rank .sv{color:var(--mu);font-weight:500;font-size:.78rem}
/* ---------- Komponen bawaan ---------- */
[data-testid="stMetric"]{background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:12px 16px}
[data-testid="stMetricLabel"]{color:var(--mu)}
[data-testid="stMetricLabel"] p{font-size:.78rem;font-weight:500}
[data-testid="stMetricValue"]{font-weight:700;color:var(--tx);font-size:1.4rem}
.stTabs [data-baseweb="tab-list"]{gap:4px;background:transparent}
.stTabs [data-baseweb="tab"]{padding:8px 14px;height:auto;background:transparent;color:var(--mu);font-weight:500}
.stTabs [aria-selected="true"]{color:var(--pri);font-weight:600}
.stTabs [data-baseweb="tab-highlight"]{background:var(--pri);height:2px}
.stTabs [data-baseweb="tab-border"]{background:var(--bd);height:1px}
.stButton>button,.stDownloadButton>button,.stFormSubmitButton>button{border-radius:6px;font-weight:500;min-height:38px;border:1px solid #D1D5DB;background:#fff;color:var(--tx);transition:.12s}
.stButton>button:hover,.stDownloadButton>button:hover,.stFormSubmitButton>button:hover{border-color:var(--pri);color:var(--pri)}
button[kind="primary"],button[data-testid="stBaseButton-primary"],button[data-testid="stBaseButton-primaryFormSubmit"]{background:var(--pri);color:#fff;border:1px solid var(--pri);font-weight:600}
button[kind="primary"]:hover,button[data-testid="stBaseButton-primary"]:hover,button[data-testid="stBaseButton-primaryFormSubmit"]:hover{background:var(--pri-d);border-color:var(--pri-d);color:#fff}
.st-key-cartpanel button[data-testid="stBaseButton-primary"]{min-height:46px;font-size:.98rem}
[data-baseweb="input"],[data-baseweb="select"]>div,[data-baseweb="textarea"]{border-radius:6px!important;background:#fff}
[data-baseweb="input"]:focus-within,[data-baseweb="select"]>div:focus-within,[data-baseweb="textarea"]:focus-within{border-color:var(--pri)!important}
[data-testid="stDataFrame"]{border:1px solid var(--bd);border-radius:8px;overflow:hidden}
[data-testid="stAlert"]{border-radius:6px}
[data-testid="stExpander"]{border:1px solid var(--bd);border-radius:8px;background:var(--card)}
[data-testid="stExpander"] summary p{font-weight:500}
[data-testid="stForm"]{border:1px solid var(--bd);border-radius:8px;background:var(--card)}
hr{border-color:var(--bd)!important}
/* ---------- Login ---------- */
.st-key-loginbox{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:28px 30px 20px;box-shadow:0 12px 32px -16px rgba(15,23,42,.25);margin-top:6vh}
.login-hero{text-align:center;margin-bottom:10px}
.login-hero .mk{width:48px;height:48px;border-radius:10px;background:var(--side);color:#fff;display:grid;place-items:center;margin:0 auto 12px;font-weight:700;letter-spacing:.5px}
.login-hero .brand{font-weight:700;font-size:1.35rem;color:var(--tx)}
.login-hero .tag{color:var(--mu);font-size:.86rem;margin-top:2px}
""" + _nav_css()
st.markdown("<style>" + CSS + "</style>", unsafe_allow_html=True)

# ───────────────────────── Util ─────────────────────────
def rp(n): return "Rp " + f"{int(round(n or 0)):,}".replace(",", ".")
def now(): return dt.datetime.now(TZ).replace(tzinfo=None).isoformat(" ", "seconds")
def today(): return dt.datetime.now(TZ).date()

def hash_pw(pw, salt=None):
    salt = salt or secrets.token_hex(8)
    h = hashlib.pbkdf2_hmac("sha256", pw.encode(), salt.encode(), 100_000).hex()
    return f"{salt}${h}"

def check_pw(pw, stored):
    salt, _ = stored.split("$", 1)
    return hmac.compare_digest(hash_pw(pw, salt), stored)

def conn():
    c = sqlite3.connect(DB, timeout=30)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA journal_mode=WAL")    # tahan mati lampu
    c.execute("PRAGMA synchronous=FULL")    # data langsung ditulis ke disk
    return c

def run(sql, p=()):
    with closing(conn()) as c:
        cur = c.execute(sql, p); c.commit(); return cur.lastrowid

def one(sql, p=()):
    with closing(conn()) as c:
        return c.execute(sql, p).fetchone()

def qdf(sql, p=()):
    with closing(conn()) as c:
        return pd.read_sql_query(sql, c, params=p)

def setting(k, d=""):
    r = one("SELECT v FROM settings WHERE k=?", (k,))
    return r["v"] if r else d

# ───────────────────────── Database ─────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY, username TEXT UNIQUE, nama TEXT, role TEXT, pw TEXT);
CREATE TABLE IF NOT EXISTS products(id INTEGER PRIMARY KEY, kode TEXT UNIQUE, nama TEXT, kategori TEXT, emoji TEXT, gambar TEXT,
  satuan TEXT, harga_beli REAL, harga_eceran REAL, harga_grosir REAL, min_grosir INTEGER, stok INTEGER, stok_min INTEGER, aktif INTEGER DEFAULT 1,
  gambar_url TEXT DEFAULT '', kata_cari TEXT DEFAULT '', sumber TEXT DEFAULT 'f', foto_coba INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS stock_moves(id INTEGER PRIMARY KEY, waktu TEXT, product_id INTEGER, tipe TEXT, qty INTEGER, stok_akhir INTEGER, ket TEXT, user TEXT);
CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY, nota TEXT UNIQUE, waktu TEXT, kasir TEXT, total REAL, modal REAL, untung REAL,
  metode TEXT, bayar REAL, kembali REAL, status TEXT DEFAULT 'OK');
CREATE TABLE IF NOT EXISTS sale_items(id INTEGER PRIMARY KEY, sale_id INTEGER, product_id INTEGER, nama TEXT, kategori TEXT, qty INTEGER,
  harga REAL, mode TEXT, subtotal REAL, modal REAL);
CREATE TABLE IF NOT EXISTS cash(id INTEGER PRIMARY KEY, waktu TEXT, tipe TEXT, kategori TEXT, nominal REAL, ket TEXT, user TEXT, sale_id INTEGER);
CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
CREATE TABLE IF NOT EXISTS draft(username TEXT, product_id INTEGER, qty INTEGER, PRIMARY KEY(username, product_id));
"""

# Katalog 100 produk toko kelontong — nama, satuan & HARGA JUAL ECERAN diambil persis dari
# Database_100_produk_toko_kelontong.docx (nomor di akhir baris = ID di file itu, kode barang = BJ + ID).
# Kolom lain TIDAK ada di file tersebut, jadi diisi perkiraan yang bisa kamu ubah di menu Stok > Kelola Barang:
#   harga_beli (modal) ≈ 70-88% harga jual · harga_grosir ≈ 5-10% lebih murah (0 = tanpa grosir) · stok awal contoh.
# (nama, kategori, satuan, harga_beli, harga_eceran, harga_grosir, min_grosir, stok, stok_min, kata_cari, sumber)
#  kata_cari : kata kunci pencarian foto. Pisahkan alternatif dengan "|".
#              Bagian setelah "||" dipakai khusus untuk Wikimedia Commons (bahasa Inggris).
#  sumber    : f = kemasan makanan/minuman (Open Food Facts) · b = perawatan/kebersihan (Open Beauty/Products Facts)
#              w = barang umum (Wikimedia Commons) · x = tanpa foto
SEED = [
 ("Beras Premium","Sembako & Dapur","kg",14100,16000,15200,5,40,10,"white rice grains","w"),   # 1
 ("Beras Medium","Sembako & Dapur","kg",12300,14000,13300,5,40,10,"uncooked rice","w"),   # 2
 ("Beras 5 Kg","Sembako & Dapur","sak",68650,78000,74100,5,40,10,"beras 5 kg|beras premium 5 kg||rice sack 5 kg","f"),   # 3
 ("Minyak Goreng 1 Liter","Sembako & Dapur","pcs",19350,22000,20900,6,24,6,"minyak goreng 1 liter|bimoli 1 liter||cooking oil pouch","f"),   # 4
 ("Minyak Goreng 2 Liter","Sembako & Dapur","pcs",39600,45000,42750,6,4,6,"minyak goreng 2 liter|bimoli 2 liter||cooking oil bottle","f"),   # 5
 ("Gula Pasir 1 Kg","Sembako & Dapur","kg",16700,19000,18050,5,40,10,"gula pasir 1 kg|gulaku||white sugar package","f"),   # 6
 ("Gula Pasir 1/2 Kg","Sembako & Dapur","pcs",8500,10000,9300,12,48,12,"gula pasir 500 g|gulaku 500 g||white sugar","f"),   # 7
 ("Telur Ayam","Sembako & Dapur","kg",26400,30000,28500,5,40,10,"chicken eggs","w"),   # 8
 ("Tepung Terigu 1 Kg","Sembako & Dapur","kg",9700,11000,10450,5,40,10,"tepung terigu segitiga biru|tepung terigu 1 kg||wheat flour package","f"),   # 9
 ("Tepung Terigu 1/2 Kg","Sembako & Dapur","pcs",5100,6000,5600,12,48,12,"tepung terigu 500 g|tepung terigu||wheat flour","f"),   # 10
 ("Tepung Tapioka","Sembako & Dapur","kg",8500,10000,9300,5,40,10,"tepung tapioka|tapioka||tapioca flour","f"),   # 11
 ("Garam","Sembako & Dapur","pcs",1550,2000,1800,20,120,30,"garam beryodium 250 g|garam dapur||table salt","f"),   # 12
 ("Kecap Manis Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"kecap manis sachet|kecap bango sachet","f"),   # 13
 ("Kecap Manis Botol","Sembako & Dapur","botol",10550,12000,11400,6,4,6,"kecap manis abc|kecap manis bango","f"),   # 14
 ("Saus Sambal Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"saus sambal sachet|saus sambal abc","f"),   # 15
 ("Saus Tomat Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"saus tomat sachet|saus tomat abc","f"),   # 16
 ("Penyedap Rasa Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"masako ayam|royco ayam","f"),   # 17
 ("Micin Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"ajinomoto|vetsin","f"),   # 18
 ("Merica Bubuk Sachet","Sembako & Dapur","pcs",700,1000,900,20,200,50,"ladaku merica bubuk|merica bubuk||ground pepper","f"),   # 19
 ("Kaldu Bubuk","Sembako & Dapur","pcs",700,1000,900,20,200,50,"royco kaldu sapi|kaldu bubuk","f"),   # 20
 ("Mi Goreng Instan","Instan & Sarapan","pcs",3000,3500,3250,12,120,30,"indomie mi goreng|mi goreng instan","f"),   # 21
 ("Mi Kuah Ayam Bawang","Instan & Sarapan","pcs",3000,3500,3250,12,120,30,"indomie ayam bawang|mie sedaap ayam bawang","f"),   # 22
 ("Mi Kuah Soto","Instan & Sarapan","pcs",3000,3500,3250,12,120,30,"indomie soto mie|mie sedaap soto","f"),   # 23
 ("Mi Kuah Kari","Instan & Sarapan","pcs",3000,3500,3250,12,120,30,"indomie kari ayam|mi kari ayam","f"),   # 24
 ("Mi Goreng Pedas","Instan & Sarapan","pcs",3000,3500,3250,12,120,30,"indomie mi goreng pedas|mi goreng pedas","f"),   # 25
 ("Mi Instan Premium","Instan & Sarapan","pcs",4250,5000,4650,12,48,12,"indomie jumbo|mi instan premium","f"),   # 26
 ("Bihun Instan","Instan & Sarapan","pcs",3400,4000,3700,12,10,12,"bihun kuah instan|bihun instan","f"),   # 27
 ("Bubur Instan","Instan & Sarapan","pcs",4250,5000,4650,12,48,12,"bubur instan|bubur ayam instan","f"),   # 28
 ("Susu Kental Manis Sachet","Instan & Sarapan","pcs",1550,2000,1800,20,120,30,"indomilk susu kental manis sachet|susu kental manis sachet","f"),   # 29
 ("Susu Kental Manis Kaleng","Instan & Sarapan","kaleng",11450,13000,12350,6,24,6,"frisian flag susu kental manis|susu kental manis kaleng","f"),   # 30
 ("Kopi Hitam Sachet","Instan & Sarapan","pcs",1150,1500,1350,20,120,30,"kapal api kopi bubuk|kopi hitam","f"),   # 31
 ("Kopi Susu Sachet","Instan & Sarapan","pcs",1550,2000,1800,20,120,30,"good day cappuccino|kopi susu sachet","f"),   # 32
 ("Kopi 3-in-1","Instan & Sarapan","pcs",1550,2000,1800,20,120,30,"kapal api special mix|kopi 3 in 1","f"),   # 33
 ("Teh Sachet","Instan & Sarapan","pcs",700,1000,900,20,200,50,"teh tubruk sachet|teh bubuk","f"),   # 34
 ("Teh Celup","Instan & Sarapan","kotak",5950,7000,6500,12,48,12,"sariwangi teh celup|teh celup","f"),   # 35
 ("Minuman Cokelat Sachet","Instan & Sarapan","pcs",1550,2000,1800,20,120,30,"milo sachet|energen coklat","f"),   # 36
 ("Susu Bubuk Sachet","Instan & Sarapan","pcs",1550,2000,1800,20,120,30,"dancow sachet|susu bubuk sachet","f"),   # 37
 ("Sereal Sachet","Instan & Sarapan","pcs",2350,3000,2700,20,0,30,"energen sereal|sereal sachet","f"),   # 38
 ("Roti Tawar","Instan & Sarapan","bungkus",13200,15000,14250,6,24,6,"sari roti tawar|roti tawar||sliced bread loaf","f"),   # 39
 ("Roti Manis","Instan & Sarapan","pcs",2350,3000,2700,20,120,30,"sari roti manis|roti manis||sweet bun bread","f"),   # 40
 ("Air Mineral 600 ml","Minuman","botol",3400,4000,3700,12,48,12,"aqua 600 ml|le minerale 600 ml","f"),   # 41
 ("Air Mineral 1,5 Liter","Minuman","botol",5950,7000,6500,12,48,12,"aqua 1.5 l|le minerale 1.5 l","f"),   # 42
 ("Air Mineral Gelas","Minuman","gelas",700,1000,900,20,200,50,"aqua gelas 220 ml|air mineral gelas","f"),   # 43
 ("Teh Botol","Minuman","botol",4250,5000,4650,12,48,12,"teh botol sosro","f"),   # 44
 ("Teh Kotak","Minuman","kotak",4250,5000,4650,12,48,12,"teh kotak sosro|ultra teh kotak","f"),   # 45
 ("Teh Pucuk","Minuman","botol",4250,5000,4650,12,48,12,"teh pucuk harum","f"),   # 46
 ("Minuman Isotonik","Minuman","botol",5950,7000,6500,12,10,12,"pocari sweat|mizone","f"),   # 47
 ("Minuman Soda","Minuman","botol",5950,7000,6500,12,48,12,"coca-cola|sprite","f"),   # 48
 ("Minuman Susu UHT","Minuman","kotak",4250,5000,4650,12,48,12,"ultra milk putih|susu uht","f"),   # 49
 ("Minuman Cokelat UHT","Minuman","kotak",4250,5000,4650,12,48,12,"ultra milk coklat|susu uht coklat","f"),   # 50
 ("Minuman Kopi Botol","Minuman","botol",5950,7000,6500,12,48,12,"kopi susu botol|good day freeze","f"),   # 51
 ("Minuman Energi","Minuman","botol",5950,7000,6500,12,48,12,"kratingdaeng|minuman energi","f"),   # 52
 ("Minuman Sirup","Minuman","botol",13200,15000,14250,6,24,6,"sirup marjan|sirup abc","f"),   # 53
 ("Minuman Serbuk Jeruk","Minuman","sachet",1150,1500,1350,20,120,30,"nutrisari jeruk|marimas jeruk","f"),   # 54
 ("Minuman Serbuk Mangga","Minuman","sachet",1150,1500,1350,20,120,30,"nutrisari mangga|marimas mangga","f"),   # 55
 ("Minuman Serbuk Jambu","Minuman","sachet",1150,1500,1350,20,120,30,"nutrisari jambu|marimas jambu","f"),   # 56
 ("Minuman Serbuk Sirsak","Minuman","sachet",1150,1500,1350,20,120,30,"nutrisari sirsak|marimas sirsak","f"),   # 57
 ("Minuman Serbuk Anggur","Minuman","sachet",1150,1500,1350,20,120,30,"marimas anggur|nutrisari anggur","f"),   # 58
 ("Minuman Serbuk Melon","Minuman","sachet",1150,1500,1350,20,120,30,"marimas melon|nutrisari melon","f"),   # 59
 ("Minuman Serbuk Rasa Teh","Minuman","sachet",1550,2000,1800,20,120,30,"marimas lemon tea|marimas teh","f"),   # 60
 ("Keripik Singkong","Snack & Jajanan","bungkus",1550,2000,1800,20,120,30,"keripik singkong||cassava chips","f"),   # 61
 ("Keripik Pisang","Snack & Jajanan","bungkus",1550,2000,1800,20,120,30,"keripik pisang||banana chips","f"),   # 62
 ("Kerupuk","Snack & Jajanan","bungkus",1550,2000,1800,20,120,30,"kerupuk udang|kerupuk||shrimp crackers","f"),   # 63
 ("Snack Jagung","Snack & Jajanan","bungkus",1550,2000,1800,20,120,30,"chiki twist|snack jagung||corn snack","f"),   # 64
 ("Snack Kentang","Snack & Jajanan","bungkus",1550,2000,1800,20,28,30,"chitato|snack kentang||potato chips","f"),   # 65
 ("Wafer Cokelat","Snack & Jajanan","pcs",1550,2000,1800,20,120,30,"tango wafer coklat|wafer coklat||chocolate wafer","f"),   # 66
 ("Wafer Vanila","Snack & Jajanan","pcs",1550,2000,1800,20,120,30,"tango wafer vanilla|wafer vanilla||vanilla wafer","f"),   # 67
 ("Biskuit Sandwich","Snack & Jajanan","bungkus",1550,2000,1800,20,120,30,"oreo original|biskuit sandwich","f"),   # 68
 ("Biskuit Marie","Snack & Jajanan","bungkus",2350,3000,2700,20,120,30,"marie regal biskuit|biskuit marie","f"),   # 69
 ("Biskuit Crackers","Snack & Jajanan","bungkus",2350,3000,2700,20,120,30,"khong guan crackers|biskuit crackers","f"),   # 70
 ("Permen Mint","Snack & Jajanan","pcs",350,500,450,20,200,50,"permen mint|relaxa||mint candy","f"),   # 71
 ("Permen Buah","Snack & Jajanan","pcs",350,500,450,20,200,50,"yupi permen|permen buah||fruit candy","f"),   # 72
 ("Permen Kopi","Snack & Jajanan","pcs",350,500,450,20,200,50,"kopiko coffee candy|permen kopi||coffee candy","f"),   # 73
 ("Cokelat Batangan Mini","Snack & Jajanan","pcs",1550,2000,1800,20,120,30,"silverqueen|cokelat batang||chocolate bar","f"),   # 74
 ("Kacang Kulit","Snack & Jajanan","bungkus",2350,3000,2700,20,120,30,"kacang garuda|kacang kulit||peanuts in shell","f"),   # 75
 ("Sabun Mandi Batang","Mandi & Kebersihan","pcs",3400,4000,3700,12,48,12,"lifebuoy sabun batang|lux sabun mandi","b"),   # 76
 ("Sabun Mandi Cair Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"lifebuoy body wash sachet|sabun cair sachet","b"),   # 77
 ("Shampoo Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"sunsilk shampoo sachet|clear shampoo sachet","b"),   # 78
 ("Shampoo Botol","Mandi & Kebersihan","botol",13200,15000,14250,6,24,6,"sunsilk shampoo|pantene shampoo","b"),   # 79
 ("Pasta Gigi","Mandi & Kebersihan","pcs",4250,5000,4650,12,10,12,"pepsodent pasta gigi|close up pasta gigi","b"),   # 80
 ("Sikat Gigi","Mandi & Kebersihan","pcs",4250,5000,4650,12,48,12,"toothbrush","w"),   # 81
 ("Deterjen Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"rinso sachet|so klin sachet","b"),   # 82
 ("Deterjen 800 Gram","Mandi & Kebersihan","bungkus",13200,15000,14250,6,24,6,"rinso detergent|daia deterjen","b"),   # 83
 ("Pewangi Pakaian Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"molto pewangi|downy sachet","b"),   # 84
 ("Sabun Cuci Piring Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"sunlight sachet|mama lemon sachet","b"),   # 85
 ("Sabun Cuci Piring Botol","Mandi & Kebersihan","botol",6800,8000,7450,12,48,12,"sunlight jeruk nipis|mama lemon","b"),   # 86
 ("Pembersih Lantai Sachet","Mandi & Kebersihan","sachet",700,1000,900,20,200,50,"wipol sachet|super pell sachet","b"),   # 87
 ("Pembersih Lantai Botol","Mandi & Kebersihan","botol",8500,10000,9300,12,48,12,"wipol karbol|super pell","b"),   # 88
 ("Tisu","Mandi & Kebersihan","bungkus",2350,3000,2700,20,120,30,"paseo facial tissue|tessa tissue","b"),   # 89
 ("Pembalut Wanita","Mandi & Kebersihan","bungkus",4250,5000,4650,12,48,12,"charm pembalut|laurier pembalut","b"),   # 90
 ("Lilin","Rumah Tangga","pcs",1550,2000,1800,20,120,30,"white candles","w"),   # 91
 ("Korek Api Gas","Rumah Tangga","pcs",2350,3000,2700,20,120,30,"gas lighter","w"),   # 92
 ("Korek Api Kayu","Rumah Tangga","kotak",700,1000,900,20,48,50,"matchbox safety matches","w"),   # 93
 ("Obat Nyamuk Bakar","Rumah Tangga","bungkus",2350,3000,2700,20,120,30,"fumakilla obat nyamuk bakar||mosquito coil","b"),   # 94
 ("Obat Nyamuk Sachet","Rumah Tangga","sachet",1550,2000,1800,20,120,30,"vape obat nyamuk||mosquito repellent","b"),   # 95
 ("Baterai AA","Rumah Tangga","pasang",4250,5000,4650,5,10,12,"abc baterai alkaline aa||AA battery","b"),   # 96
 ("Baterai AAA","Rumah Tangga","pasang",4250,5000,4650,5,0,12,"abc baterai alkaline aaa||AAA battery","b"),   # 97
 ("Kantong Plastik","Rumah Tangga","bungkus",4250,5000,4650,12,48,12,"plastic carrier bags","w"),   # 98
 ("Gas LPG 3 Kg","Rumah Tangga","tabung",19350,22000,0,0,30,8,"LPG cylinder 3 kg","w"),   # 99
 ("Air Galon Isi Ulang","Rumah Tangga","galon",5950,7000,0,0,30,8,"water gallon dispenser","w"),   # 100
]
KATALOG_VER = "v3-100-produk-kelontong"
KATEGORI = ["Sembako & Dapur", "Instan & Sarapan", "Minuman", "Snack & Jajanan", "Mandi & Kebersihan", "Rumah Tangga"]

# 1 owner + 3 kasir. Ubah di sini kalau komposisinya beda (hanya dipakai saat database masih kosong).
SEED_USERS = [("owner", "Pemilik Bulo Jaya", "owner", "owner123"), ("kasir1", "Kasir 1", "kasir", "kasir123"),
              ("kasir2", "Kasir 2", "kasir", "kasir123"), ("kasir3", "Kasir 3", "kasir", "kasir123")]

def migrate_schema():
    """Tambah kolom baru (foto & kata kunci) pada database lama tanpa menghapus data."""
    with closing(conn()) as c:
        have = {r["name"] for r in c.execute("PRAGMA table_info(products)")}
        for col, ddl in (("gambar_url", "TEXT DEFAULT ''"), ("kata_cari", "TEXT DEFAULT ''"),
                         ("sumber", "TEXT DEFAULT 'f'"), ("foto_coba", "INTEGER DEFAULT 0"),
                         ("barcode", "TEXT DEFAULT ''")):
            if col not in have: c.execute(f"ALTER TABLE products ADD COLUMN {col} {ddl}")
        c.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_products_barcode ON products(barcode) WHERE barcode<>''")
        c.commit()

def _kunci_foto(kata_cari):
    """Kata kunci utama (alternatif pertama, huruf kecil) untuk mencocokkan foto lama dengan barang baru."""
    return (kata_cari or "").split("||")[0].split("|")[0].strip().lower()

def install_katalog():
    """Ganti katalog lama dengan katalog 100 produk kelontong (sekali saja).
    Barang lama dinonaktifkan, bukan dihapus, sehingga riwayat penjualan, kas & mutasi tetap utuh."""
    if setting("katalog") == KATALOG_VER: return
    if os.path.exists(DB) and one("SELECT 1 FROM products"):
        bak = f"{DB}.backup_sebelum_katalog_baru"
        if not os.path.exists(bak):
            try: shutil.copy2(DB, bak)
            except OSError: pass
    with closing(conn()) as c:
        # foto lama yang kata kunci utamanya sama dipakai ulang (hemat unduhan, tidak perlu internet)
        foto_lama = {}
        for r in c.execute("SELECT kata_cari,sumber,gambar FROM products WHERE aktif=1 AND COALESCE(gambar,'')<>''").fetchall():
            if os.path.exists(r["gambar"]):
                foto_lama.setdefault((_kunci_foto(r["kata_cari"]), r["sumber"]), r["gambar"])
        c.execute("UPDATE products SET aktif=0, kode=kode||'#lama'||id WHERE aktif=1")
        for i, s in enumerate(SEED, 1):
            nama, kat, sat, hb, he, hg, mg, stok, smin, cari, src = s
            pid = c.execute("INSERT INTO products(kode,nama,kategori,emoji,gambar,satuan,harga_beli,harga_eceran,harga_grosir,min_grosir,stok,stok_min,"
                            "gambar_url,kata_cari,sumber,foto_coba) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)",
                            (f"BJ{i:03d}", nama, kat, "", foto_lama.get((_kunci_foto(cari), src), ""), sat, hb, he, hg, mg, stok, smin, "", cari, src)).lastrowid
            c.execute("INSERT INTO stock_moves(waktu,product_id,tipe,qty,stok_akhir,ket,user) VALUES(?,?,?,?,?,?,?)",
                      (now(), pid, "AWAL", stok, stok, "Stok awal katalog 100 produk kelontong", "sistem"))
        c.execute("INSERT OR REPLACE INTO settings VALUES('katalog',?)", (KATALOG_VER,))
        c.commit()

def init_db():
    with closing(conn()) as c:
        c.executescript(SCHEMA); c.commit()
    migrate_schema()
    if not one("SELECT 1 FROM users"):
        for un, nm, rl, pw in SEED_USERS:
            run("INSERT INTO users(username,nama,role,pw) VALUES(?,?,?,?)", (un, nm, rl, hash_pw(pw)))
    if not one("SELECT 1 FROM settings WHERE k='toko'"):
        for k, v in {"toko": "UMKM Bulo Jaya", "alamat": "Makassar, Sulawesi Selatan", "footer": "Terima kasih, selamat berbelanja kembali!"}.items():
            run("INSERT OR REPLACE INTO settings VALUES(?,?)", (k, v))
    install_katalog()
    backfill_barcode()

# ───────────────────────── Logika bisnis ─────────────────────────
def get_product(pid): return one("SELECT * FROM products WHERE id=? AND aktif=1", (int(pid),))
def by_kode(kode): return one("SELECT * FROM products WHERE (kode=? OR barcode=? OR lower(nama)=lower(?)) AND aktif=1", (kode.strip(), kode.strip(), kode.strip()))

# ── Barcode EAN-13 (barang >= HARGA_BARCODE_MIN) ──
def ean13(pid):
    d = f"899{int(pid):09d}"
    t = sum(int(x) * (3 if i % 2 else 1) for i, x in enumerate(d))
    return d + str((10 - t % 10) % 10)

def ensure_barcode(pid):
    p = one("SELECT harga_eceran, barcode FROM products WHERE id=?", (int(pid),))
    if p and (p["harga_eceran"] or 0) >= HARGA_BARCODE_MIN and not p["barcode"]:
        run("UPDATE products SET barcode=? WHERE id=?", (ean13(pid), int(pid)))

def backfill_barcode():
    with closing(conn()) as c:
        for r in c.execute("SELECT id FROM products WHERE aktif=1 AND harga_eceran>=? AND COALESCE(barcode,'')=''", (HARGA_BARCODE_MIN,)).fetchall():
            c.execute("UPDATE products SET barcode=? WHERE id=?", (ean13(r["id"]), r["id"]))
        c.commit()

def barcode_png(code):
    from barcode import EAN13
    from barcode.writer import ImageWriter
    buf = io.BytesIO(); EAN13(code[:12], writer=ImageWriter()).write(buf)
    return buf.getvalue()

# ── Keranjang disimpan di database (aman kalau mati lampu / browser tertutup) ──
def load_cart(username):
    with closing(conn()) as c:
        return {int(r["product_id"]): int(r["qty"]) for r in c.execute("SELECT product_id,qty FROM draft WHERE username=?", (username,))}

def save_cart():
    u = st.session_state.get("user")
    if not u: return
    cart = {int(k): int(v) for k, v in st.session_state.get("cart", {}).items()}
    if st.session_state.get("cart_saved") == cart: return
    with closing(conn()) as c:
        c.execute("DELETE FROM draft WHERE username=?", (u["username"],))
        c.executemany("INSERT INTO draft(username,product_id,qty) VALUES(?,?,?)", [(u["username"], k, v) for k, v in cart.items()])
        c.commit()
    st.session_state.cart_saved = dict(cart)

def price_for(p, qty):
    """Harga bertingkat: grosir aktif bila qty >= batas grosir."""
    if (p["harga_grosir"] or 0) > 0 and (p["min_grosir"] or 0) > 0 and qty >= p["min_grosir"]:
        return p["harga_grosir"], "GROSIR"
    return p["harga_eceran"], "ECER"

def move_stock(c, pid, delta, tipe, ket, user):
    stok = c.execute("SELECT stok FROM products WHERE id=?", (pid,)).fetchone()["stok"] + delta
    if stok < 0: raise ValueError("Stok tidak boleh negatif")
    c.execute("UPDATE products SET stok=? WHERE id=?", (stok, pid))
    c.execute("INSERT INTO stock_moves(waktu,product_id,tipe,qty,stok_akhir,ket,user) VALUES(?,?,?,?,?,?,?)", (now(), pid, tipe, delta, stok, ket, user))

def add_cash(c, tipe, kategori, nominal, ket, user, sale_id=None):
    c.execute("INSERT INTO cash(waktu,tipe,kategori,nominal,ket,user,sale_id) VALUES(?,?,?,?,?,?,?)", (now(), tipe, kategori, nominal, ket, user, sale_id))

def checkout(cart, metode, bayar, kasir):
    with closing(conn()) as c:
        try:
            items, total, modal = [], 0, 0
            for pid, qty in cart.items():
                p = c.execute("SELECT * FROM products WHERE id=?", (int(pid),)).fetchone()
                if not p or p["stok"] < qty: raise ValueError(f"Stok Tidak Mencukupi: {p['nama'] if p else pid}")
                harga, mode = price_for(p, qty)
                items.append((p, qty, harga, mode, harga * qty, p["harga_beli"] * qty))
                total += harga * qty; modal += p["harga_beli"] * qty
            if metode == "Tunai" and bayar < total: raise ValueError("Uang diterima kurang")
            bayar = bayar if metode == "Tunai" else total
            n = c.execute("SELECT COUNT(*) n FROM sales WHERE date(waktu)=date(?)", (now(),)).fetchone()["n"] + 1
            nota = f"BJ{today():%Y%m%d}-{n:04d}"
            sid = c.execute("INSERT INTO sales(nota,waktu,kasir,total,modal,untung,metode,bayar,kembali) VALUES(?,?,?,?,?,?,?,?,?)",
                            (nota, now(), kasir, total, modal, total - modal, metode, bayar, bayar - total)).lastrowid
            for p, qty, harga, mode, sub, mdl in items:
                c.execute("INSERT INTO sale_items(sale_id,product_id,nama,kategori,qty,harga,mode,subtotal,modal) VALUES(?,?,?,?,?,?,?,?,?)",
                          (sid, p["id"], p["nama"], p["kategori"], qty, harga, mode, sub, mdl))
                move_stock(c, p["id"], -qty, "PENJUALAN", nota, kasir)
            add_cash(c, "MASUK", "Penjualan", total, nota, kasir, sid)
            c.commit()
        except Exception:
            c.rollback(); raise
    return sid

def void_sale(sid, user):
    with closing(conn()) as c:
        try:
            s = c.execute("SELECT * FROM sales WHERE id=?", (sid,)).fetchone()
            if s["status"] != "OK": raise ValueError("Transaksi sudah dibatalkan")
            for it in c.execute("SELECT * FROM sale_items WHERE sale_id=?", (sid,)).fetchall():
                move_stock(c, it["product_id"], it["qty"], "BATAL", f"Batal {s['nota']}", user)
            c.execute("UPDATE sales SET status='BATAL' WHERE id=?", (sid,))
            add_cash(c, "KELUAR", "Pembatalan Penjualan", s["total"], s["nota"], user, sid)
            c.commit()
        except Exception:
            c.rollback(); raise

def receipt_text(sid):
    s = one("SELECT * FROM sales WHERE id=?", (sid,))
    its = qdf("SELECT * FROM sale_items WHERE sale_id=?", (sid,))
    W = 34
    L = [setting("toko").center(W), setting("alamat").center(W), "=" * W, f"No   : {s['nota']}", f"Waktu: {s['waktu']}", f"Kasir: {s['kasir']}", "-" * W]
    for r in its.itertuples():
        L.append(r.nama[:W]); tag = " [GROSIR]" if r.mode == "GROSIR" else ""
        L.append(f"  {r.qty} x {int(r.harga):,}".replace(",", ".") + tag)
        L.append(f"{rp(r.subtotal):>{W}}")
    L += ["-" * W, f"{'TOTAL':<12}{rp(s['total']):>{W-12}}", f"{s['metode']:<12}{rp(s['bayar']):>{W-12}}", f"{'Kembali':<12}{rp(s['kembali']):>{W-12}}", "=" * W, setting("footer").center(W)]
    return "\n".join(L)

# ───────────────────────── Foto produk asli dari internet ─────────────────────────
# Sumber terbuka: Open Food Facts / Open Beauty Facts / Open Products Facts (foto kemasan, CC BY-SA)
# dan Wikimedia Commons (barang umum). Diunduh pelan-pelan di latar belakang & disimpan di folder gambar_produk,
# jadi hanya perlu internet sekali. Batas Open Food Facts: 10 pencarian/menit/IP, karena itu ada jeda antar permintaan.
UA = f"BuloJayaPOS/2.0 ({os.environ.get('POS_CONTACT', 'kasir@bulojaya.local')})"

@st.cache_resource
def _net():
    """Status bersama lintas rerun Streamlit (kunci, jeda antar permintaan, cache hasil pencarian)."""
    return {"tl": threading.Lock(), "sl": threading.Lock(), "last": {}, "urls": {}}
_N = _net()

class NetError(Exception):
    """Internet mati / server sibuk (bukan berarti fotonya tidak ada)."""

def _throttle(key, gap):
    with _N["tl"]:
        wait = _N["last"].get(key, 0) + gap - time.time()
        if wait > 0: time.sleep(wait)
        _N["last"][key] = time.time()

def http_get(url, timeout=15, limit=8_000_000):
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = r.read(limit + 1)
    except urllib.error.HTTPError as e:
        if e.code in (429, 502, 503, 504): raise NetError(f"server sibuk (HTTP {e.code})") from e
        raise ValueError(f"HTTP {e.code}") from e
    except (urllib.error.URLError, OSError, http.client.HTTPException) as e:
        raise NetError(str(e)) from e
    if len(data) > limit: raise ValueError("file terlalu besar")
    return data

def _jget(url): return json.loads(http_get(url).decode("utf-8", "replace"))

def _teks(v):
    if isinstance(v, str): return v
    if isinstance(v, (list, tuple)): return " ".join(_teks(x) for x in v)
    if isinstance(v, dict): return " ".join(_teks(x) for x in v.values())
    return ""

def _skor(q, teks):
    """Seberapa cocok teks dengan kata kunci. 0 = tidak cocok (kata utama/brand harus ada)."""
    toks = re.findall(r"[a-z0-9]+", q.lower())
    kata = [t for t in toks if not t.isdigit() and len(t) >= 3]
    if not kata: return 1
    tl = teks.lower()
    if kata[0] not in tl: return 0
    return sum(t in tl for t in toks if len(t) >= 2 or t.isdigit())

def _gambar_dari(h):
    for k in ("image_front_url", "image_url", "image_front_small_url", "image_small_url"):
        u = h.get(k)
        if isinstance(u, str) and u.startswith("http"): return u
    try:
        for u in (h.get("selected_images") or {}).get("front", {}).get("display", {}).values():
            if isinstance(u, str) and u.startswith("http"): return u
    except AttributeError: pass
    return None

def _off_produk(kode):
    """Ambil URL foto depan dari barcode (batas 15 permintaan/menit)."""
    _throttle("off_produk", 4.5)
    d = _jget(f"https://world.openfoodfacts.org/api/v2/product/{urllib.parse.quote(str(kode))}.json?fields=image_front_url,image_url")
    return _gambar_dari(d.get("product") or {})

def _pilih(q, hits, teks_fn, resolve):
    ranked = sorted(((_skor(q, teks_fn(h)), -i, h) for i, h in enumerate(hits)), key=lambda x: (x[0], x[1]), reverse=True)
    for sk, _, h in ranked[:3]:
        if sk <= 0: break
        try: u = resolve(h)
        except ValueError: u = None
        if u: return u
    return None

def off_sal(q):
    """Open Food Facts via Search-a-licious (pencarian teks resmi). Utamakan produk yang dijual di Indonesia."""
    def cari(query):
        _throttle("sal", 1.5)
        return _jget("https://search.openfoodfacts.org/search?" + urllib.parse.urlencode({"q": query, "page_size": 10})).get("hits") or []
    def resolve(h):
        return _gambar_dari(h) or (_off_produk(h["code"]) if h.get("code") else None)
    teks = lambda h: _teks([h.get("product_name"), h.get("brands"), h.get("generic_name")])
    for query in (f'{q} countries_tags:"en:indonesia"', q):
        try: hits = cari(query)
        except ValueError: continue          # sintaks filter ditolak server → coba tanpa filter
        u = _pilih(q, hits, teks, resolve)
        if u: return u
    return None

def legacy_cari(host, q):
    """Open Beauty Facts / Open Products Facts (API pencarian lama)."""
    _throttle(host, 6.5)
    d = _jget(f"https://{host}/cgi/search.pl?" + urllib.parse.urlencode(
        {"search_terms": q, "search_simple": 1, "action": "process", "json": 1, "page_size": 10}))
    return _pilih(q, d.get("products") or [], lambda h: _teks([h.get("product_name"), h.get("brands"), h.get("generic_name")]), _gambar_dari)

def commons_cari(q):
    """Wikimedia Commons — foto barang umum (galon, gelas, tabung gas, dsb.)."""
    _throttle("commons", 1.0)
    d = _jget("https://commons.wikimedia.org/w/api.php?" + urllib.parse.urlencode({
        "action": "query", "format": "json", "generator": "search", "gsrsearch": q + " filetype:bitmap", "gsrnamespace": 6,
        "gsrlimit": 10, "prop": "imageinfo", "iiprop": "url|mime|size", "iiurlwidth": 700}))
    pages = sorted(((d.get("query") or {}).get("pages") or {}).values(), key=lambda x: x.get("index", 99))
    cands = []
    for pg in pages:
        ii = (pg.get("imageinfo") or [{}])[0]
        if ii.get("mime") not in ("image/jpeg", "image/png"): continue
        if min(ii.get("width", 0), ii.get("height", 0)) < 250: continue
        u = ii.get("thumburl") or ii.get("url")
        if u: cands.append((_skor(q, pg.get("title", "")), u))
    if not cands: return None
    best = max(cands, key=lambda x: x[0])          # urutan relevansi Commons dipakai bila skor sama
    return best[1] if best[0] > 0 else cands[0][1]

def cari_url_foto(p):
    """Tentukan URL foto untuk satu barang (baris produk). None = tidak ketemu. NetError = masalah jaringan."""
    if (p["gambar_url"] or "").strip(): return p["gambar_url"].strip()
    src = p["sumber"] or "f"
    if src == "x": return None
    db_part, _, w_part = (p["kata_cari"] or p["nama"]).partition("||")
    q_db = [q.strip() for q in db_part.split("|") if q.strip()][:2]
    q_w = [q.strip() for q in (w_part or db_part).split("|") if q.strip()][:2]
    plan = []
    if src == "f": plan += [("sal", q) for q in q_db]
    if src == "b" and q_db: plan += [("obf", q_db[0]), ("opf", q_db[0])]
    plan += [("wiki", q) for q in q_w]
    fn = {"sal": off_sal, "obf": lambda q: legacy_cari("world.openbeautyfacts.org", q),
          "opf": lambda q: legacy_cari("world.openproductsfacts.org", q), "wiki": commons_cari}
    net = False
    for kind, q in plan:
        key = (kind, q.lower())
        if key not in _N["urls"]:
            try: _N["urls"][key] = fn[kind](q)
            except NetError: net = True; continue
            except (ValueError, KeyError, TypeError, AttributeError): _N["urls"][key] = None
        if _N["urls"][key]: return _N["urls"][key]
    if net: raise NetError("jaringan bermasalah")
    return None

def simpan_foto(tag, data):
    """Validasi & simpan foto sebagai JPEG (latar transparan → putih). Mengembalikan path file."""
    from PIL import Image, ImageOps
    im = Image.open(io.BytesIO(data)); im.load()
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA"); bg = Image.new("RGB", im.size, (255, 255, 255)); bg.paste(im, mask=im.getchannel("A")); im = bg
    else: im = im.convert("RGB")
    if min(im.size) < 100: raise ValueError("gambar terlalu kecil")
    im.thumbnail((900, 900))
    path = os.path.join(IMG_DIR, f"{tag}_{secrets.token_hex(3)}.jpg")
    im.save(path, "JPEG", quality=88)
    return path

def ambil_foto(pid):
    """Cari & unduh foto satu barang. Return (ok, pesan). NetError bila internet bermasalah."""
    p = one("SELECT * FROM products WHERE id=?", (int(pid),))
    if not p: return False, "Barang tidak ditemukan"
    try:
        url = cari_url_foto(p)
        if not url: raise LookupError("foto belum ketemu di sumber terbuka")
        path = simpan_foto(f"p{pid}", http_get(url))
    except NetError: raise
    except Exception as e:
        run("UPDATE products SET foto_coba=foto_coba+1 WHERE id=?", (int(pid),))
        return False, str(e) or e.__class__.__name__
    run("UPDATE products SET gambar=?, foto_coba=0 WHERE id=?", (path, int(pid)))
    return True, "OK"

@st.cache_resource
def photo_state(): return {"running": False, "offline": False, "last": ""}

def _foto_loop(S):
    gagal_net = 0
    try:
        while True:
            r = one("SELECT id,nama FROM products WHERE aktif=1 AND COALESCE(sumber,'f')<>'x' AND COALESCE(gambar,'')='' AND COALESCE(foto_coba,0)=0 ORDER BY id LIMIT 1")
            if not r: break
            S["last"] = r["nama"]
            try: ambil_foto(r["id"]); gagal_net = 0
            except NetError:
                gagal_net += 1
                if gagal_net >= 3: S["offline"] = True; break
                time.sleep(5)
    except Exception: pass
    finally: S["running"] = False

def mulai_unduh_foto(ulang=False):
    """Mulai pekerja latar belakang (aman dipanggil berulang)."""
    S = photo_state()
    if ulang:
        run("UPDATE products SET foto_coba=0 WHERE aktif=1 AND COALESCE(gambar,'')=''"); _N["urls"].clear(); S["offline"] = False
    with _N["sl"]:
        if S["running"] or S["offline"]: return
        if not one("SELECT 1 FROM products WHERE aktif=1 AND COALESCE(sumber,'f')<>'x' AND COALESCE(gambar,'')='' AND COALESCE(foto_coba,0)=0 LIMIT 1"): return
        S["running"] = True
    threading.Thread(target=_foto_loop, args=(S,), daemon=True, name="foto-worker").start()

def foto_stats():
    r = one("SELECT COUNT(*) n, SUM(CASE WHEN COALESCE(gambar,'')<>'' THEN 1 ELSE 0 END) ada, "
            "SUM(CASE WHEN COALESCE(gambar,'')='' AND COALESCE(foto_coba,0)>0 THEN 1 ELSE 0 END) gagal "
            "FROM products WHERE aktif=1 AND COALESCE(sumber,'f')<>'x'")
    return int(r["n"] or 0), int(r["ada"] or 0), int(r["gagal"] or 0)

def _foto_body(compact):
    S = photo_state(); n, ada, gagal = foto_stats()
    if not n: return
    if S["running"] and compact:
        st.caption(f"Mengunduh foto asli produk di latar belakang… {ada}/{n} siap — aplikasi tetap bisa dipakai.")
    elif S["running"]:
        st.info(f"Mengunduh foto asli produk dari internet… **{ada}/{n}** siap (sekarang: {S['last']}). "
                "Berjalan di latar belakang — aplikasi tetap bisa dipakai. Barang tanpa foto sementara memakai gambar nama.")
    elif S["offline"]:
        st.warning("Tidak bisa terhubung ke server foto (internet mati / server sibuk). Barang tanpa foto sementara memakai gambar nama.")
        if st.button("Coba unduh lagi", key=f"foto_retry_{int(compact)}"): mulai_unduh_foto(ulang=True); st.rerun()
    elif gagal and not compact:
        st.caption(f"{ada}/{n} barang sudah berfoto asli. {gagal} barang belum ketemu fotonya otomatis — "
                   "unggah foto / tempel URL foto di tab **Kelola Barang**, atau coba cari ulang.")
        if st.button("Cari ulang foto yang belum ada", key="foto_retry_0"): mulai_unduh_foto(ulang=True); st.rerun()
    baru = ada - st.session_state.get("foto_rendered", ada)
    if baru > 0 and st.button(f"{baru} foto baru siap — klik untuk menampilkan", key=f"foto_new_{int(compact)}", type="primary"): st.rerun()

def foto_status(compact=False):
    """Status unduhan foto; memperbarui diri tiap 5 detik selama unduhan berjalan."""
    S = photo_state()
    st.session_state.foto_rendered = foto_stats()[1]
    st.fragment(run_every=5 if S["running"] else None)(_foto_body)(compact)

# ───────────────────────── Tampilan: gambar produk & kartu ─────────────────────────
esc = html.escape
HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
CAT_PAL = {  # (warna terang, warna tengah, warna teks)
    "Sembako & Dapur": ("#fde68a", "#f59e0b", "#92400e"), "Instan & Sarapan": ("#fed7aa", "#fb923c", "#9a3412"),
    "Minuman": ("#bae6fd", "#38bdf8", "#075985"), "Snack & Jajanan": ("#fef08a", "#facc15", "#854d0e"),
    "Mandi & Kebersihan": ("#ddd6fe", "#a78bfa", "#5b21b6"), "Rumah Tangga": ("#a7f3d0", "#34d399", "#065f46"),
}

def daftar_kategori():
    ada = qdf("SELECT DISTINCT kategori FROM products WHERE aktif=1")["kategori"].tolist()
    return [k for k in KATEGORI if k in ada] + sorted(k for k in ada if k not in KATEGORI)

@st.cache_data(show_spinner=False)
def product_art(nama, kategori, emoji, satuan):
    """Gambar sementara (tanpa emoji) bila foto asli belum ada: kartu nama barang di atas warna kategori."""
    a, b, c = CAT_PAL.get(kategori, ("#bbf7d0", "#4ade80", "#166534"))
    n = int(hashlib.md5(nama.encode()).hexdigest()[:8], 16)
    ang, c1x, c2x = 20 + n % 60, 30 + n % 110, 330 + (n >> 3) % 120
    wr = textwrap.wrap(nama, 17) or [nama[:17]]
    lines = wr[:3]
    if len(wr) > 3: lines[-1] = lines[-1][:15] + "…"
    h = 20 + 40 * len(lines) + 44
    y0 = (360 - h) // 2 - 6
    txt = "".join(f'<text x="240" y="{y0 + 50 + i * 40}" font-size="30" font-weight="700" fill="{c}" text-anchor="middle" '
                  f'font-family="Arial,sans-serif">{esc(l)}</text>' for i, l in enumerate(lines))
    lab = esc(f"foto belum tersedia · per {satuan}" if satuan else "foto belum tersedia")
    svg = (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 480 360">'
           f'<defs><linearGradient id="g" gradientTransform="rotate({ang} .5 .5)"><stop offset="0" stop-color="{a}"/><stop offset="1" stop-color="{b}"/></linearGradient></defs>'
           f'<rect width="480" height="360" fill="url(#g)"/>'
           f'<circle cx="{c1x}" cy="40" r="70" fill="#fff" opacity=".2"/><circle cx="{c2x}" cy="320" r="90" fill="#fff" opacity=".16"/>'
           f'<rect x="40" y="{y0}" width="400" height="{h}" rx="26" fill="#fff" opacity=".94"/>{txt}'
           f'<text x="240" y="{y0 + h - 16}" font-size="15" fill="#94a3b8" text-anchor="middle" font-family="Arial,sans-serif">{lab}</text></svg>')
    return "data:image/svg+xml;base64," + base64.b64encode(svg.encode()).decode()

@st.cache_data(show_spinner=False)
def photo_uri(path, mtime, size=(360, 270), q=76):
    """Foto asli → dimuat utuh di atas latar putih (kemasan tidak terpotong) & dikecilkan supaya halaman tetap ringan."""
    try:
        from PIL import Image, ImageOps
        im = ImageOps.pad(ImageOps.exif_transpose(Image.open(path)).convert("RGB"), size, color=(255, 255, 255))
        buf = io.BytesIO(); im.save(buf, "JPEG", quality=q)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""

def prod_img(nama, kategori, emoji, satuan, gambar):
    if isinstance(gambar, str) and gambar and os.path.exists(gambar):
        u = photo_uri(gambar, os.path.getmtime(gambar))
        if u: return u
    return product_art(str(nama), str(kategori or ""), "", str(satuan or ""))

def thumb_uri(gambar):
    """Thumbnail kecil untuk tabel; None bila belum ada foto."""
    if isinstance(gambar, str) and gambar and os.path.exists(gambar):
        return photo_uri(gambar, os.path.getmtime(gambar), (96, 72), 70) or None
    return None

def stock_pct(stok, smin): return 0 if stok <= 0 else max(8, min(100, round(stok / (max(smin, 1) * 4) * 100)))

def card_html(r, incart=0):
    """Kartu produk (gambar + harga + bar stok). r = baris dari df.itertuples()."""
    stok, smin = int(r.stok), int(r.stok_min or 0)
    src = prod_img(r.nama, r.kategori, r.emoji, r.satuan, r.gambar)
    if stok <= 0: tl, col, out = "<span class='tg tl out'>HABIS</span>", "#DC2626", "out"
    elif stok <= smin: tl, col, out = "<span class='tg tl low'>MENIPIS</span>", "#D97706", ""
    else: tl, col, out = "", "#16834A", ""
    tr = f"<span class='tg tr cart'>Di keranjang: {incart}</span>" if incart else ""
    hg = r.harga_grosir if pd.notna(r.harga_grosir) else 0
    mg = r.min_grosir if pd.notna(r.min_grosir) else 0
    gr = f"<div class='gr'>Grosir {rp(hg)} · min {int(mg)}</div>" if hg > 0 and mg > 0 else "<div class='gr none'>&nbsp;</div>"
    sat = esc(str(r.satuan or ""))
    return (f"<div class='pimg {out}'><img src='{src}' alt=''>{tl}{tr}</div>"
            f"<div class='pbody'><div class='nm'>{esc(str(r.nama))}</div><div class='pr'>{rp(r.harga_eceran)} <small>/ {sat}</small></div>{gr}"
            f"<div class='sbar'><i style='width:{stock_pct(stok, smin)}%;background:{col}'></i></div>"
            f"<div class='sinfo'><span>Stok</span><b style='color:{col}'>{stok} {sat}</b></div></div>")

def page_head(title, sub=""):
    st.markdown(f"<div class='phead'><div class='tt'>{esc(title)}</div><div class='ss'>{esc(sub)}</div></div>", unsafe_allow_html=True)

def kasir_header(u):
    t = today()
    s = one("SELECT COUNT(*) n, COALESCE(SUM(total),0) t FROM sales WHERE status='OK' AND date(waktu)=?", (str(t),))
    low = one("SELECT COUNT(*) n FROM products WHERE aktif=1 AND stok<=stok_min")["n"]
    page_head(f"{HARI[t.weekday()]}, {t.day} {BULAN[t.month - 1]} {t.year}  ·  {u['nama']}")
    items = [("Transaksi hari ini", str(s["n"]), "")]
    if u["role"] == "owner": items.append(("Omset hari ini", rp(s["t"]), ""))
    items.append(("Stok menipis", f"{low} barang", "warn" if low else ""))
    st.markdown("<div class='statbar'>" + "".join(f"<div class='it'><span>{k}</span><b class='{c}'>{esc(v)}</b></div>" for k, v, c in items) + "</div>",
                unsafe_allow_html=True)

def promo_ticker():
    d = qdf("SELECT nama,satuan,harga_eceran,harga_grosir,min_grosir FROM products WHERE aktif=1 AND stok>0 AND harga_grosir>0 AND min_grosir>0 "
            "ORDER BY (harga_eceran-harga_grosir)*1.0/harga_eceran DESC LIMIT 4")
    if d.empty: return
    tx = "  ·  ".join(f"<b>{esc(r.nama)}</b> ≥{int(r.min_grosir)} {esc(str(r.satuan))} = {rp(r.harga_grosir)}" for r in d.itertuples())
    st.markdown(f"<div class='promo'>Harga grosir: {tx}</div>", unsafe_allow_html=True)

# ───────────────────────── Scan: barcode & AI ─────────────────────────
def _strs(x):
    if isinstance(x, str): return [x] if x else []
    if isinstance(x, (list, tuple)): return [s for i in x for s in _strs(i)]
    return []

def decode_barcodes(data):
    out = []
    try:
        import numpy as np, cv2
        img = cv2.imdecode(np.frombuffer(data, np.uint8), cv2.IMREAD_COLOR)
        try:
            det = cv2.barcode.BarcodeDetector()
            if hasattr(det, "detectAndDecodeWithType"): out += _strs(det.detectAndDecodeWithType(img)[1])
            else: out += _strs(det.detectAndDecode(img)[1])
        except Exception: pass
        try: out += _strs(cv2.QRCodeDetector().detectAndDecode(img)[0])
        except Exception: pass
    except ImportError:
        st.warning("Library opencv belum terpasang: pip install opencv-python-headless")
    return out

def ai_identify(data, mime):
    """Kenali produk dari foto memakai Claude Vision (opsional, butuh ANTHROPIC_API_KEY)."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key: return None, "Isi dulu variabel lingkungan ANTHROPIC_API_KEY untuk memakai fitur AI."
    try:
        import anthropic
        cat = qdf("SELECT id,nama FROM products WHERE aktif=1")
        daftar = "\n".join(f"{r.id}|{r.nama}" for r in cat.itertuples())
        msg = anthropic.Anthropic(api_key=key).messages.create(
            model=os.environ.get("ANTHROPIC_MODEL", "claude-sonnet-5-5"), max_tokens=100,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": mime or "image/jpeg", "data": base64.b64encode(data).decode()}},
                {"type": "text", "text": f"Foto ini adalah barang di toko sembako. Pilih barang yang paling cocok dari daftar (format id|nama):\n{daftar}\n\n"
                                         'Balas HANYA JSON: {"id": <angka atau null>, "alasan": "<singkat>"}'}]}])
        m = re.search(r"\{.*\}", msg.content[0].text, re.S)
        j = json.loads(m.group(0)) if m else {}
        return j.get("id"), j.get("alasan", "")
    except Exception as e:
        return None, f"AI gagal: {e}"

# ───────────────────────── Keranjang ─────────────────────────
def flash(kind, msg): st.session_state.flash = (kind, msg)

def cart_add(pid, n=1):
    cart, k = st.session_state.cart, f"q_{pid}"
    cur = int(st.session_state.get(k, cart.get(pid, 0)))
    st.session_state[k] = cur + n; cart[pid] = cur + n

def cart_step(pid, d):
    k = f"q_{pid}"; v = int(st.session_state.get(k, 1)) + d
    if v < 1: cart_del(pid)
    else: st.session_state[k] = v; st.session_state.cart[pid] = v

def cart_del(pid):
    st.session_state.cart.pop(pid, None); st.session_state.pop(f"q_{pid}", None)

def cart_clear():
    for pid in list(st.session_state.cart): cart_del(pid)

def add_by_code(kode):
    p = by_kode(kode)
    if not p: flash("error", f"Kode '{kode}' tidak ditemukan"); return False
    if p["stok"] <= 0: flash("error", f"{p['nama']} stok habis"); return False
    cart_add(int(p["id"])); flash("success", f"{p['nama']} ditambahkan"); return True

def scan_manual():
    v = st.session_state.get("kode_in", "").strip()
    if v: add_by_code(v)
    st.session_state.kode_in = ""

# ───────────────────────── Halaman: Kasir ─────────────────────────
def scan_input():
    st.text_input("Scan barcode / kode barang", key="kode_in", on_change=scan_manual, placeholder="Arahkan scanner USB ke barcode, atau ketik kode (mis. BJ001) lalu Enter")

def scan_tab():
    mode = st.radio("Kamera", ["Mati", "Scan Barcode", "Kenali dengan AI"], horizontal=True)
    if mode == "Mati": return
    img = st.camera_input("Arahkan barcode / foto barang", key=f"cam_{mode}")
    if not img: return
    data = img.getvalue(); h = hashlib.md5(data).hexdigest()
    if st.session_state.get("last_img") == h: return
    st.session_state.last_img = h
    if mode == "Scan Barcode":
        codes = decode_barcodes(data)
        if not codes: flash("warning", "Barcode tidak terbaca, coba lebih dekat / ketik manual.")
        else:
            for cd in dict.fromkeys(codes):
                if add_by_code(cd): break
    else:
        with st.spinner("AI mengenali barang..."):
            pid, ket = ai_identify(data, img.type)
        p = get_product(pid) if pid else None
        if p and p["stok"] > 0: cart_add(int(p["id"])); flash("success", f"AI mengenali: {p['nama']} — {ket}")
        else: flash("warning", f"AI tidak yakin. {ket or ''}")
    st.rerun()

SORTS = {
    "Nama A–Z": "p.nama", "Terlaris": "COALESCE(t.q,0) DESC, p.nama", "Harga termurah": "p.harga_eceran, p.nama",
    "Harga termahal": "p.harga_eceran DESC, p.nama", "Stok terbanyak": "p.stok DESC, p.nama", "Stok menipis dulu": "(p.stok<=p.stok_min) DESC, p.stok, p.nama",
}

def catalog_tab():
    s = st.session_state.get("cat_s", "")      # kolom cari ada di atas halaman kasir
    cats = ["Semua"] + daftar_kategori()
    c1, c2 = st.columns([3, 1])
    with c1: k = st.radio("Kategori", cats, horizontal=True, key="cat_k", label_visibility="collapsed")
    srt = c2.selectbox("Urutkan", list(SORTS), key="cat_o", label_visibility="collapsed", format_func=lambda x: f"Urutkan: {x}")
    sig = (s, srt, k)
    if st.session_state.get("cat_sig") != sig: st.session_state.cat_sig = sig; st.session_state.cat_lim = 45
    lim = st.session_state.get("cat_lim", 45)
    where, p = " WHERE p.aktif=1", []
    if k != "Semua": where += " AND p.kategori=?"; p.append(k)
    if s: where += " AND p.nama LIKE ?"; p.append(f"%{s}%")
    total = int(one("SELECT COUNT(*) n FROM products p" + where, p)["n"])
    df = qdf("SELECT p.* FROM products p LEFT JOIN (SELECT i.product_id pid, SUM(i.qty) q FROM sale_items i JOIN sales s ON s.id=i.sale_id "
             "WHERE s.status='OK' GROUP BY 1) t ON t.pid=p.id" + where + f" ORDER BY {SORTS[srt]} LIMIT ?", p + [lim])
    st.caption(f"Menampilkan {len(df)} dari {total} barang · klik **Tambah** untuk memasukkan ke keranjang")
    if df.empty: st.info("Barang tidak ditemukan. Coba kata kunci lain."); return
    cart, cols = st.session_state.cart, st.columns(3)
    for i, r in enumerate(df.itertuples()):
        with cols[i % 3].container(key=f"pc_{r.id}"):
            st.markdown(card_html(r, cart.get(int(r.id), 0)), unsafe_allow_html=True)
            st.button("Tambah", key=f"add_{r.id}", on_click=cart_add, args=(int(r.id),), disabled=r.stok <= 0, use_container_width=True)
    if total > len(df):
        st.button(f"Tampilkan {min(45, total - len(df))} barang lagi", key="cat_more", use_container_width=True,
                  on_click=lambda: st.session_state.update(cat_lim=lim + 45))

def cart_panel():
    cart = st.session_state.cart
    st.markdown(f"<div class='carttitle'><span>Keranjang</span><span class='pill'>{sum(cart.values())} item</span></div>", unsafe_allow_html=True)
    if st.session_state.pop("celebrate", False): st.toast("Transaksi berhasil disimpan")
    if st.session_state.get("last_sale"):
        sid = st.session_state.last_sale; txt = receipt_text(sid)
        st.success("Transaksi berhasil disimpan — stok & kas sudah diperbarui.")
        st.code(txt, language=None)
        a, b = st.columns(2)
        a.download_button("Simpan struk", txt, file_name="struk.txt", use_container_width=True)
        if b.button("Tutup struk", use_container_width=True): st.session_state.last_sale = None; st.rerun()
        components.html(f"<style>@media print{{button{{display:none}}}}</style><pre style='font:13px monospace'>{html.escape(txt)}</pre>"
                        "<button onclick='window.print()' style='padding:8px 16px'>Cetak struk</button>", height=120)
        st.divider()
    if not cart:
        st.markdown("<div class='empty'><b>Keranjang kosong</b><small>Scan barcode atau pilih barang dari katalog</small></div>", unsafe_allow_html=True)
        return
    total, ok = 0, True
    for pid in list(cart):
        p = get_product(pid)
        if not p: cart_del(pid); continue
        k = f"q_{pid}"
        if k not in st.session_state: st.session_state[k] = cart[pid]
        with st.container(key=f"ci_{pid}"):
            src = prod_img(p["nama"], p["kategori"], p["emoji"], p["satuan"], p["gambar"])
            st.markdown(f"<div class='cihead'><img src='{src}' alt=''><div><div class='nm'>{esc(p['nama'])}</div>"
                        f"<div class='sb'>{rp(p['harga_eceran'])} / {esc(p['satuan'] or '')} · sisa stok {p['stok']}</div></div></div>", unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns([1, 2, 1, 1.6])
            c1.button("−", key=f"m{pid}", on_click=cart_step, args=(pid, -1), use_container_width=True)
            qty = int(c2.number_input("Jumlah", min_value=1, step=1, key=k, label_visibility="collapsed"))
            c3.button("+", key=f"p{pid}", on_click=cart_step, args=(pid, 1), use_container_width=True)
            c4.button("Hapus", key=f"d{pid}", on_click=cart_del, args=(pid,), use_container_width=True)
            cart[pid] = qty
            harga, mode = price_for(p, qty); sub = harga * qty; total += sub
            badge = "<span class='badge-g'>Harga grosir</span>" if mode == "GROSIR" else "<span class='badge-e'>Eceran</span>"
            st.markdown(f"{badge} &nbsp; {qty} × {rp(harga)} = **{rp(sub)}**", unsafe_allow_html=True)
            if mode == "ECER" and p["harga_grosir"] and p["min_grosir"]:
                st.caption(f"Beli ≥ {p['min_grosir']} {p['satuan']} → harga grosir {rp(p['harga_grosir'])}")
            if qty > p["stok"]: ok = False; st.error(f"Stok Tidak Mencukupi (sisa {p['stok']} {p['satuan']})")
    if not cart: return
    st.markdown(f"<div class='total'><div><small>Total bayar</small><div class='amt'>{rp(total)}</div></div><b>{sum(cart.values())} item</b></div>", unsafe_allow_html=True)
    metode = st.radio("Metode bayar", ["Tunai", "QRIS", "Transfer"], horizontal=True)
    bayar = total
    if metode == "Tunai":
        bayar = st.number_input("Uang diterima (Rp)", min_value=0, step=1000, key="bayar")
        st.button("Uang pas", on_click=lambda: st.session_state.update(bayar=int(total)))
        st.metric("Kembalian", rp(max(bayar - total, 0)) if bayar >= total else "Kurang " + rp(total - bayar))
    disabled = (not ok) or (metode == "Tunai" and bayar < total)
    if st.button("Bayar & Cetak Struk", type="primary", use_container_width=True, disabled=disabled):
        try:
            sid = checkout(dict(cart), metode, float(bayar), st.session_state.user["nama"])
            cart_clear(); st.session_state.last_sale = sid; st.session_state.celebrate = True; st.session_state.pop("bayar", None); st.rerun()
        except Exception as e: st.error(str(e))
    st.button("Kosongkan keranjang", on_click=cart_clear, use_container_width=True)

def page_kasir():
    kasir_header(st.session_state.user)
    promo_ticker()
    foto_status(compact=True)
    f = st.session_state.pop("flash", None)
    if f: getattr(st, f[0])(f[1])
    L, R = st.columns([5, 3], gap="large")
    with L:
        q1, q2 = st.columns(2)
        with q1: scan_input()
        q2.text_input("Cari barang", key="cat_s", placeholder="Nama barang, mis. beras, indomie, aqua, gas")
        with st.expander("Scan dengan kamera / kenali barang dengan AI"): scan_tab()
        catalog_tab()
    with R:
        with st.container(key="cartpanel"): cart_panel()

# ───────────────────────── Halaman: Stok ─────────────────────────
def stok_status(r): return "HABIS" if r.stok <= 0 else ("MENIPIS" if r.stok <= r.stok_min else "Aman")

def quick_stock(pid, delta):
    """Tombol ➕/➖ cepat di galeri stok (dipanggil sebagai on_click)."""
    p = get_product(pid)
    if not p: return
    try:
        with closing(conn()) as cn:
            move_stock(cn, pid, delta, "MASUK" if delta > 0 else "KELUAR", "Penyesuaian cepat (galeri stok)", st.session_state.user["nama"])
            cn.commit()
        st.session_state.flash_stok = ("success", f"{p['nama']}: {'+' if delta > 0 else '-'}{abs(delta)} {p['satuan'] or ''} → stok {p['stok'] + delta}")
    except Exception as e:
        st.session_state.flash_stok = ("error", f"{p['nama']}: {e}")

def page_stok():
    page_head("Stok Barang", "Pantau, tambah, dan kurangi persediaan toko")
    foto_status()
    f = st.session_state.pop("flash_stok", None)
    if f: st.toast(f[1])
    owner = st.session_state.user["role"] == "owner"
    df = qdf("SELECT * FROM products WHERE aktif=1 ORDER BY kategori, nama")
    df["status"] = [stok_status(r) for r in df.itertuples()]
    c = st.columns(4)
    c[0].metric("Jenis barang", len(df)); c[1].metric("Menipis", int((df.status == "MENIPIS").sum()))
    c[2].metric("Habis", int((df.status == "HABIS").sum()))
    if owner: c[3].metric("Nilai persediaan (modal)", rp((df.stok * df.harga_beli).sum()))
    krit = df[df.status != "Aman"].sort_values("stok")
    if len(krit):
        nm = ", ".join(f"{esc(r.nama)} ({r.stok})" for r in krit.head(6).itertuples())
        st.markdown(f"<div class='alertbar'><b>{len(krit)} barang perlu restok:</b> {nm}{' …' if len(krit) > 6 else ''}</div>", unsafe_allow_html=True)
    kats = ["Semua"] + daftar_kategori()
    tabs = st.tabs(["Daftar Stok", "Galeri Stok", "Tambah / Kurang Stok", "Kelola Barang", "Riwayat Mutasi"])
    with tabs[0]:
        f1, f2, f3 = st.columns([2, 1, 1])
        s = f1.text_input("Cari", key="stk_s"); stt = f2.selectbox("Status", ["Semua", "MENIPIS", "HABIS", "Aman"]); kk = f3.selectbox("Kategori", kats, key="stk_k")
        v = df[df.nama.str.contains(s, case=False, regex=False)] if s else df
        if stt != "Semua": v = v[v.status == stt]
        if kk != "Semua": v = v[v.kategori == kk]
        v = v.copy()
        v["foto"] = [thumb_uri(g) for g in v.gambar]
        v["level"] = [stock_pct(int(a), int(b or 0)) for a, b in zip(v.stok, v.stok_min)]
        v["kondisi"] = v.status.map({"HABIS": "Habis", "MENIPIS": "Menipis", "Aman": "Aman"})
        cols = ["foto", "kode", "barcode", "nama", "kategori", "stok", "level", "satuan", "stok_min", "harga_eceran", "harga_grosir", "min_grosir", "kondisi"]
        if owner: cols.insert(cols.index("harga_eceran"), "harga_beli")
        cc = st.column_config
        cfg = {"foto": cc.ImageColumn("Foto", width="small"), "level": cc.ProgressColumn("Level stok", min_value=0, max_value=100, format="%d%%"),
               "harga_eceran": cc.NumberColumn("Harga eceran", format="Rp %d"), "harga_grosir": cc.NumberColumn("Harga grosir", format="Rp %d"),
               "harga_beli": cc.NumberColumn("Harga beli", format="Rp %d")}
        st.dataframe(v[cols], use_container_width=True, hide_index=True, column_config=cfg)
    with tabs[1]:
        st.caption("Tekan **−** / **+** di bawah foto barang untuk mengurangi / menambah stok langsung. Atur jumlah per klik di bawah. "
                   "Untuk mencatat pembelian ke kas, pakai tab **Tambah / Kurang Stok**.")
        g1, g2, g3, g4 = st.columns([2, 1, 1, 1])
        gs = g1.text_input("Cari di galeri", key="gal_s", placeholder="ketik nama barang…")
        gk = g2.selectbox("Kategori", kats, key="gal_k"); gt = g3.selectbox("Tampilkan", ["Semua", "MENIPIS", "HABIS", "Aman"], key="gal_t")
        step = int(g4.number_input("Jumlah per klik", min_value=1, max_value=10000, value=1, step=1, key="gal_step"))
        gv = df[df.nama.str.contains(gs, case=False, regex=False)] if gs else df
        if gk != "Semua": gv = gv[gv.kategori == gk]
        if gt != "Semua": gv = gv[gv.status == gt]
        sig = (gs, gk, gt)
        if st.session_state.get("gal_sig") != sig: st.session_state.gal_sig = sig; st.session_state.gal_lim = 40
        lim = st.session_state.get("gal_lim", 40)
        if gv.empty: st.info("Tidak ada barang yang cocok.")
        gc = st.columns(4)
        for i, r in enumerate(gv.head(lim).itertuples()):
            with gc[i % 4].container(key=f"pg_{r.id}"):
                st.markdown(card_html(r), unsafe_allow_html=True)
                b1, b2 = st.columns(2)
                b1.button(f"− {step}", key=f"gm_{r.id}", on_click=quick_stock, args=(int(r.id), -step), disabled=int(r.stok) < step, use_container_width=True)
                b2.button(f"+ {step}", key=f"gp_{r.id}", on_click=quick_stock, args=(int(r.id), step), use_container_width=True)
        if len(gv) > lim:
            st.button(f"Tampilkan {min(40, len(gv) - lim)} barang lagi (sisa {len(gv) - lim})", key="gal_more", use_container_width=True,
                      on_click=lambda: st.session_state.update(gal_lim=lim + 40))
    with tabs[2]:
        t1, t2 = st.columns([2, 1])
        ts = t1.text_input("Cari barang", key="tk_s", placeholder="ketik nama barang untuk mempersempit pilihan…")
        tk = t2.selectbox("Kategori", kats, key="tk_k")
        dv = df[df.nama.str.contains(ts, case=False, regex=False)] if ts else df
        if tk != "Semua": dv = dv[dv.kategori == tk]
        if dv.empty: st.info("Barang tidak ditemukan.")
        else:
            opt = {int(r.id): f"{r.nama} ({r.kode}) — stok {r.stok} {r.satuan}" for r in dv.itertuples()}
            pid = st.selectbox("Barang", list(opt), format_func=opt.get, key="tk_pid")
            p = get_product(pid)
            src = prod_img(p["nama"], p["kategori"], "", p["satuan"], p["gambar"])
            warna = "#ef4444" if p["stok"] <= 0 else ("#f59e0b" if p["stok"] <= p["stok_min"] else "#16a34a")
            st.markdown(f"<div style='display:flex;gap:14px;align-items:center;margin:4px 0 10px'><img src='{src}' style='width:120px;border-radius:14px;box-shadow:0 4px 12px rgba(0,0,0,.18)'>"
                        f"<div><div style='font-weight:700;font-size:1.05rem'>{esc(p['nama'])}</div>"
                        f"<div style='color:#64748b;font-size:.85rem'>Stok sekarang: <b style='color:{warna};font-size:1.1rem'>{p['stok']} {esc(p['satuan'] or '')}</b> · batas kritis {p['stok_min']}</div></div></div>", unsafe_allow_html=True)
            aksi = st.radio("Aksi", ["Tambah (barang masuk)", "Kurangi (rusak/hilang/kedaluwarsa)", "Koreksi stok (stok opname)"], horizontal=True, key="tk_aksi")
            koreksi = aksi.startswith("Koreksi")
            qty = st.number_input("Stok fisik sebenarnya" if koreksi else "Jumlah", min_value=0, step=1, value=int(p["stok"]) if koreksi else 0, key=f"tk_qty_{pid}_{int(koreksi)}")
            ket = st.text_input("Keterangan", placeholder="mis. restok dari supplier / dus rusak", key="tk_ket")
            catat = False
            if aksi.startswith("Tambah") and owner: catat = st.checkbox("Catat sebagai pengeluaran kas (harga beli × jumlah)", value=True, key="tk_kas")
            if st.button("Simpan perubahan stok", type="primary", key="tk_save"):
                if qty <= 0 and not koreksi: st.warning("Isi jumlah lebih dari 0.")
                else:
                    try:
                        with closing(conn()) as cn:
                            user = st.session_state.user["nama"]
                            if aksi.startswith("Tambah"):
                                move_stock(cn, pid, int(qty), "MASUK", ket or "Barang masuk", user)
                                if catat and qty: add_cash(cn, "KELUAR", "Pembelian Stok", p["harga_beli"] * qty, f"{p['nama']} x{qty}", user)
                            elif aksi.startswith("Kurangi"): move_stock(cn, pid, -int(qty), "KELUAR", ket or "Barang keluar", user)
                            else: move_stock(cn, pid, int(qty) - p["stok"], "KOREKSI", ket or "Stok opname", user)
                            cn.commit()
                        st.session_state.flash_stok = ("success", f"Stok {p['nama']} diperbarui.")
                        st.rerun()
                    except Exception as e: st.error(str(e))
    with tabs[3]:
        if not owner: st.info("Hanya Owner yang dapat menambah/mengubah/menghapus data barang.")
        else: form_barang(df)
    with tabs[4]:
        st.dataframe(qdf("SELECT m.waktu,p.nama,m.tipe,m.qty,m.stok_akhir,m.ket,m.user FROM stock_moves m JOIN products p ON p.id=m.product_id "
                         "ORDER BY m.id DESC LIMIT 300"), use_container_width=True, hide_index=True)

SUMBER_LBL = {"f": "Makanan/minuman kemasan (Open Food Facts)", "b": "Perawatan/kebersihan (Open Beauty/Products Facts)",
              "w": "Barang umum (Wikimedia Commons)", "x": "Tanpa foto internet (unggah sendiri)"}

def form_barang(df):
    opt = {0: "Barang baru"} | {int(r.id): f"{r.nama} ({r.kode})" for r in df.itertuples()}
    pid = st.selectbox("Pilih", list(opt), format_func=opt.get, key="kb_sel")
    p = get_product(pid) if pid else None
    if p:
        pv = prod_img(p["nama"], p["kategori"], "", p["satuan"], p["gambar"])
        st.markdown(f"<div style='display:flex;gap:14px;align-items:center;margin:6px 0 10px'><img src='{pv}' style='width:150px;border-radius:14px;box-shadow:0 4px 12px rgba(0,0,0,.18)'>"
                    "<div style='color:#64748b;font-size:.85rem'>Foto saat ini.<br>Foto asli diunduh otomatis dari internet. "
                    "Bisa diganti dengan <b>unggah foto</b> sendiri atau <b>tempel URL gambar</b> di bawah.</div></div>", unsafe_allow_html=True)
        b1, b2, _ = st.columns([1.3, 1, 2])
        if b1.button("Cari foto otomatis untuk barang ini", key=f"kb_cari_{pid}", use_container_width=True):
            with st.spinner("Mencari foto di internet… (bisa sampai ±20 detik)"):
                try: ok, msg = ambil_foto(pid)
                except NetError: ok, msg = False, "Tidak ada koneksi internet / server sedang sibuk."
            if ok: st.success("Foto ditemukan & disimpan."); st.rerun()
            else: st.warning(f"Belum berhasil: {msg}. Coba tempel URL gambar atau unggah foto sendiri.")
        if p["gambar"] and b2.button("Hapus foto", key=f"kb_hf_{pid}", use_container_width=True):
            run("UPDATE products SET gambar='', foto_coba=1 WHERE id=?", (pid,)); st.rerun()
    g = lambda f, d: (p[f] if p and p[f] is not None else d)
    cats = list(dict.fromkeys(KATEGORI + sorted(set(df.kategori.tolist()))))
    with st.form(f"fb_{pid}"):
        a, b = st.columns(2)
        kode = a.text_input("Kode / Barcode", g("kode", "")); nama = b.text_input("Nama barang", g("nama", ""))
        kat = a.selectbox("Kategori", cats, index=cats.index(g("kategori", "Sembako")) if g("kategori", "Sembako") in cats else 0); kat2 = b.text_input("…atau kategori baru")
        sat = a.text_input("Satuan", g("satuan", "pcs"))
        hb = a.number_input("Harga beli / modal", 0, step=500, value=int(g("harga_beli", 0)))
        he = b.number_input("Harga eceran", 0, step=500, value=int(g("harga_eceran", 0)))
        hg = a.number_input("Harga grosir (0 = tanpa grosir)", 0, step=500, value=int(g("harga_grosir", 0)))
        mg = b.number_input("Batas qty grosir", 0, step=1, value=int(g("min_grosir", 0)))
        stok = a.number_input("Stok awal" if not p else "Stok (ubah lewat tab Tambah/Kurang)", 0, step=1, value=int(g("stok", 0)), disabled=bool(p))
        smin = b.number_input("Batas stok kritis", 0, step=1, value=int(g("stok_min", 5)))
        st.markdown("**Foto barang**")
        up = st.file_uploader("Unggah foto sendiri (opsional)", type=["png", "jpg", "jpeg", "webp"])
        url = st.text_input("…atau tempel URL gambar (opsional)", g("gambar_url", ""), placeholder="https://…/foto.jpg")
        d1, d2 = st.columns(2)
        cari = d1.text_input("Kata kunci cari foto otomatis", g("kata_cari", ""), help="Pisahkan alternatif dengan |  ·  Tambahkan || lalu kata kunci bahasa Inggris untuk Wikimedia. Contoh: indomie mi goreng")
        skeys = list(SUMBER_LBL)
        src = d2.selectbox("Sumber foto otomatis", skeys, index=skeys.index(g("sumber", "f")) if g("sumber", "f") in skeys else 0, format_func=SUMBER_LBL.get)
        save = st.form_submit_button("Simpan", type="primary")
    if save:
        if not kode or not nama: st.error("Kode dan nama wajib diisi."); return
        if hg and hg >= he: st.warning("Harga grosir sebaiknya lebih rendah dari eceran.")
        img = g("gambar", "")
        if up:
            try: img = simpan_foto("upload", up.getvalue())
            except Exception as e: st.error(f"Foto tidak bisa dibaca: {e}"); return
        elif url.strip() and url.strip() != g("gambar_url", ""):
            try: img = simpan_foto("url", http_get(url.strip()))
            except NetError: st.warning("Tidak ada koneksi internet — foto dari URL belum diunduh.")
            except Exception as e: st.warning(f"Foto dari URL gagal dipakai: {e}")
        kat = kat2.strip() or kat
        try:
            if p:
                run("UPDATE products SET kode=?,nama=?,kategori=?,gambar=?,satuan=?,harga_beli=?,harga_eceran=?,harga_grosir=?,min_grosir=?,stok_min=?,"
                    "gambar_url=?,kata_cari=?,sumber=? WHERE id=?", (kode, nama, kat, img, sat, hb, he, hg, mg, smin, url.strip(), cari.strip(), src, pid))
                if not img: run("UPDATE products SET foto_coba=0 WHERE id=?", (pid,))
            else:
                new = run("INSERT INTO products(kode,nama,kategori,emoji,gambar,satuan,harga_beli,harga_eceran,harga_grosir,min_grosir,stok,stok_min,gambar_url,kata_cari,sumber,foto_coba) "
                          "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)", (kode, nama, kat, "", img, sat, hb, he, hg, mg, stok, smin, url.strip(), cari.strip() or nama, src))
                run("INSERT INTO stock_moves(waktu,product_id,tipe,qty,stok_akhir,ket,user) VALUES(?,?,?,?,?,?,?)", (now(), new, "AWAL", stok, stok, "Barang baru", st.session_state.user["nama"]))
            ensure_barcode(pid if p else new)
            mulai_unduh_foto()
            st.success("Tersimpan."); st.rerun()
        except sqlite3.IntegrityError: st.error("Kode barang sudah dipakai.")
    if p:
        with st.expander("Barcode & cetak"):
            if p["barcode"]:
                st.caption(f"Barcode EAN-13: `{p['barcode']}` — tempel di kemasan, lalu scan di kasir.")
                try:
                    png = barcode_png(p["barcode"]); st.image(png, width=300)
                    st.download_button("Unduh PNG", png, file_name=f"barcode_{p['barcode']}.png", mime="image/png", key=f"bc_{pid}")
                except ImportError: st.error("Jalankan: pip install python-barcode pillow")
            else: st.info(f"Barcode otomatis dibuat untuk barang dengan harga eceran ≥ {rp(HARGA_BARCODE_MIN)}.")
    if p and st.button("Hapus barang ini (nonaktifkan)"):
        run("UPDATE products SET aktif=0, kode=kode||'#del'||id WHERE id=?", (pid,)); st.rerun()

# ───────────────────────── Periode & Dashboard ─────────────────────────
def pilih_periode(key):
    o = st.selectbox("Periode", ["Hari ini", "7 hari terakhir", "30 hari terakhir", "Bulan ini", "Semua", "Pilih tanggal"], key=key)
    t = today()
    if o == "Hari ini": return t, t
    if o == "7 hari terakhir": return t - dt.timedelta(days=6), t
    if o == "30 hari terakhir": return t - dt.timedelta(days=29), t
    if o == "Bulan ini": return t.replace(day=1), t
    if o == "Semua": return dt.date(2000, 1, 1), t
    r = st.date_input("Rentang", (t - dt.timedelta(days=6), t), key=key + "d")
    return (r[0], r[1]) if len(r) == 2 else (t, t)

def page_dashboard():
    page_head("Dashboard & Laporan", "Omset, modal, untung, dan barang terlaris")
    d1, d2 = pilih_periode("per_d"); P = (str(d1), str(d2))
    s = qdf("SELECT * FROM sales WHERE status='OK' AND date(waktu) BETWEEN ? AND ?", P)
    omset, modal = s.total.sum(), s.modal.sum()
    low_n = int(one("SELECT COUNT(*) n FROM products WHERE aktif=1 AND stok<=stok_min")["n"])
    r1 = st.columns(3)
    r1[0].metric("Omset", rp(omset)); r1[1].metric("Untung", rp(omset - modal)); r1[2].metric("Jumlah transaksi", len(s))
    r2 = st.columns(3)
    r2[0].metric("Modal (HPP)", rp(modal)); r2[1].metric("Margin", f"{(omset - modal) / omset * 100:.1f}%" if omset else "0%")
    r2[2].metric("Barang stok menipis", low_n)
    t1, t2, t3 = st.tabs(["Grafik", "Stok Kritis", "Riwayat Penjualan"])
    with t1:
        if s.empty: st.info("Belum ada penjualan pada periode ini.")
        else:
            s["tgl"] = s.waktu.str[:10]; s["jam"] = s.waktu.str[11:13]
            it = qdf("SELECT i.* FROM sale_items i JOIN sales s ON s.id=i.sale_id WHERE s.status='OK' AND date(s.waktu) BETWEEN ? AND ?", P)
            top = it.groupby("nama").qty.sum().nlargest(5)
            if len(top):
                st.markdown("<div class='sech'>Barang terlaris</div>", unsafe_allow_html=True)
                pm = qdf("SELECT nama,kategori,emoji,satuan,gambar FROM products ORDER BY aktif DESC, id DESC").drop_duplicates("nama").set_index("nama")
                rc = st.columns(5)
                for i, (nm, q) in enumerate(top.items()):
                    if nm in pm.index: m = pm.loc[nm]; src_img, sat = prod_img(nm, m.kategori, m.emoji, m.satuan, m.gambar), m.satuan
                    else: src_img, sat = product_art(nm, "", "", ""), ""
                    rc[i].markdown(f"<div class='rank'><span class='md'>{i + 1}</span><img src='{src_img}' alt=''>"
                                   f"<div class='nm'>{esc(nm)}</div><div class='sv'>{int(q)} {esc(str(sat or ''))} terjual</div></div>", unsafe_allow_html=True)
            a, b = st.columns(2)
            a.subheader("Omset, modal & untung harian")
            a.bar_chart(s.groupby("tgl")[["total", "modal", "untung"]].sum().rename(columns={"total": "Omset", "modal": "Modal", "untung": "Untung"}), color=["#16834A", "#94A3B8", "#173D2B"])
            b.subheader("Jam sibuk (jumlah transaksi)")
            b.bar_chart(s.groupby("jam").size().rename("Transaksi"), color="#16834A")
            a, b = st.columns(2)
            a.subheader("10 barang terlaris")
            a.bar_chart(it.groupby("nama").qty.sum().nlargest(10), color="#16834A")
            b.subheader("Untung per kategori")
            it["untung"] = it.subtotal - it.modal
            b.bar_chart(it.groupby("kategori").untung.sum(), color="#173D2B")
    with t2:
        k = qdf("SELECT kode,nama,kategori,stok,satuan,stok_min FROM products WHERE aktif=1 AND stok<=stok_min ORDER BY stok")
        st.dataframe(k, use_container_width=True, hide_index=True) if len(k) else st.success("Semua stok aman.")
    with t3:
        h = qdf("SELECT id,nota,waktu,kasir,total,modal,untung,metode,status FROM sales WHERE date(waktu) BETWEEN ? AND ? ORDER BY id DESC", P)
        st.dataframe(h.drop(columns="id"), use_container_width=True, hide_index=True)
        st.download_button("Unduh CSV", h.to_csv(index=False), file_name="penjualan.csv")
        if len(h):
            sel = st.selectbox("Lihat detail nota", h.id.tolist(), format_func=lambda i: h.set_index("id").loc[i, "nota"])
            st.code(receipt_text(int(sel)), language=None)
            if h.set_index("id").loc[sel, "status"] == "OK" and st.button("Batalkan transaksi ini (stok dikembalikan)"):
                try: void_sale(int(sel), st.session_state.user["nama"]); st.success("Dibatalkan."); st.rerun()
                except Exception as e: st.error(str(e))

# ───────────────────────── Halaman: Keuangan ─────────────────────────
def page_keuangan():
    page_head("Pemantauan Keuangan", "Pemasukan, pengeluaran, dan saldo kas")
    d1, d2 = pilih_periode("per_k"); P = (str(d1), str(d2))
    df = qdf("SELECT * FROM cash WHERE date(waktu) BETWEEN ? AND ? ORDER BY id DESC", P)
    masuk = df[df.tipe == "MASUK"].nominal.sum(); keluar = df[df.tipe == "KELUAR"].nominal.sum()
    saldo = qdf("SELECT COALESCE(SUM(CASE WHEN tipe='MASUK' THEN nominal ELSE -nominal END),0) s FROM cash")["s"][0]
    c = st.columns(4)
    c[0].metric("Pemasukan", rp(masuk)); c[1].metric("Pengeluaran", rp(keluar)); c[2].metric("Selisih periode", rp(masuk - keluar)); c[3].metric("Saldo kas total", rp(saldo))
    with st.expander("Tambah / kurangi catatan kas", expanded=True):
        a, b, cc = st.columns([1, 1, 1])
        tipe = a.radio("Jenis", ["Pemasukan", "Pengeluaran"], horizontal=True)
        kat = b.selectbox("Kategori", ["Modal Awal", "Setoran Modal", "Pendapatan Lain"] if tipe.startswith("Pemasukan")
                          else ["Pembelian Stok", "Gaji", "Listrik/Air", "Sewa", "Transportasi", "Prive", "Lain-lain"])
        nom = cc.number_input("Nominal (Rp)", min_value=0, step=5000)
        ket = st.text_input("Keterangan", key="kas_ket")
        if st.button("Simpan catatan kas", type="primary") and nom > 0:
            with closing(conn()) as cn:
                add_cash(cn, "MASUK" if tipe.startswith("Pemasukan") else "KELUAR", kat, nom, ket, st.session_state.user["nama"]); cn.commit()
            st.rerun()
    st.subheader("Transaksi keuangan")
    st.dataframe(df.drop(columns=["sale_id"]), use_container_width=True, hide_index=True)
    st.download_button("Unduh CSV", df.to_csv(index=False), file_name="kas.csv")
    man = df[df.sale_id.isna()]
    if len(man):
        hid = st.selectbox("Hapus catatan manual", man.id.tolist(), format_func=lambda i: f"#{i} {man.set_index('id').loc[i, 'kategori']} {rp(man.set_index('id').loc[i, 'nominal'])}")
        if st.button("Hapus catatan"): run("DELETE FROM cash WHERE id=?", (int(hid),)); st.rerun()

# ───────────────────────── Halaman: Pengaturan ─────────────────────────
def page_pengaturan():
    page_head("Pengaturan", "Toko, pengguna, dan data")
    t1, t2, t3 = st.tabs(["Toko & Struk", "Pengguna", "Data"])
    with t1:
        with st.form("toko"):
            tk = st.text_input("Nama toko", setting("toko")); al = st.text_input("Alamat", setting("alamat")); ft = st.text_input("Footer struk", setting("footer"))
            if st.form_submit_button("Simpan"):
                for k, v in (("toko", tk), ("alamat", al), ("footer", ft)): run("INSERT OR REPLACE INTO settings VALUES(?,?)", (k, v))
                st.success("Tersimpan.")
    with t2:
        u = qdf("SELECT id,username,nama,role FROM users"); st.dataframe(u, hide_index=True, use_container_width=True)
        with st.form("usr"):
            a, b, c, d = st.columns(4)
            un = a.text_input("Username"); nm = b.text_input("Nama"); rl = c.selectbox("Peran", ["kasir", "owner"]); pw = d.text_input("Password", type="password")
            if st.form_submit_button("Tambah / reset password"):
                if un and pw and nm:
                    ex = one("SELECT id FROM users WHERE username=?", (un,))
                    if ex: run("UPDATE users SET nama=?,role=?,pw=? WHERE id=?", (nm, rl, hash_pw(pw), ex["id"]))
                    else: run("INSERT INTO users(username,nama,role,pw) VALUES(?,?,?,?)", (un, nm, rl, hash_pw(pw)))
                    st.success("Tersimpan."); st.rerun()
                else: st.error("Lengkapi semua kolom.")
        dl = st.selectbox("Hapus pengguna", u.username.tolist())
        if st.button("Hapus pengguna"):
            if dl == st.session_state.user["username"]: st.error("Tidak bisa menghapus akun sendiri.")
            else: run("DELETE FROM users WHERE username=?", (dl,)); st.rerun()
    with t3:
        if os.path.exists(DB): st.download_button("Backup database", open(DB, "rb").read(), file_name="backup_bulojaya.db")
        st.warning("Hapus semua riwayat transaksi, kas & mutasi (data barang & stok tetap). Gunakan setelah uji coba, sebelum dipakai sungguhan.")
        if st.text_input("Ketik HAPUS untuk konfirmasi") == "HAPUS" and st.button("Hapus riwayat transaksi"):
            for t in ("sales", "sale_items", "cash", "stock_moves"): run(f"DELETE FROM {t}")
            st.success("Riwayat dihapus.")

# ───────────────────────── Login & Navigasi ─────────────────────────
def login():
    _, m, _ = st.columns([1, 1.1, 1])
    with m:
        with st.container(key="loginbox"):
            st.markdown("<div class='login-hero'><div class='mk'>BJ</div><div class='brand'>POS Bulo Jaya</div><div class='tag'>Sistem kasir &amp; persediaan UMKM</div></div>", unsafe_allow_html=True)
            with st.form("login"):
                u = st.text_input("Username"); p = st.text_input("Password", type="password")
                if st.form_submit_button("Masuk", type="primary", use_container_width=True):
                    r = one("SELECT * FROM users WHERE username=?", (u.strip(),))
                    if r and check_pw(p, r["pw"]):
                        st.session_state.user = dict(r)
                        st.session_state.cart = load_cart(r["username"]); st.session_state.cart_saved = dict(st.session_state.cart)
                        st.rerun()
                    else: st.error("Username atau password salah.")
            st.caption("Akun awal → owner / owner123  ·  kasir1, kasir2, kasir3 / kasir123  (ganti di menu Pengaturan)")

def main():
    init_db()
    st.session_state.setdefault("cart", {})
    if "user" not in st.session_state: return login()
    mulai_unduh_foto()
    u = st.session_state.user
    pages = {"Kasir": page_kasir, "Stok Barang": page_stok}
    if u["role"] == "owner": pages |= {"Dashboard & Laporan": page_dashboard, "Keuangan": page_keuangan, "Pengaturan": page_pengaturan}
    with st.sidebar:
        st.markdown(f"<div class='sbrand'><div class='mk'>BJ</div><div><div class='bn'>{esc(setting('toko'))}</div><div class='bs'>Kasir &amp; Persediaan</div></div></div>"
                    f"<div class='userbox'><span>{esc(u['nama'])}</span><em>{esc(u['role'])}</em></div><div class='navlabel'>Menu</div>", unsafe_allow_html=True)
        pg = st.radio("Menu", list(pages), label_visibility="collapsed")
        with st.container(key="sb_bottom"):
            if st.button("Keluar", use_container_width=True): st.session_state.clear(); st.rerun()
    try: pages[pg]()
    finally: save_cart()  # simpan keranjang ke database setiap ada perubahan

if __name__ == "__main__":
    main()

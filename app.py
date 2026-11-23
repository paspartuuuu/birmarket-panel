import os
import sqlite3
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, Response

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("DB_PATH", os.path.join(BASE_DIR, "birmarket_mobile.db"))
EXCEL_PATH = os.environ.get("EXCEL_PATH", os.path.join(BASE_DIR, "Birmarket Tarifləri.xlsx"))
PANEL_PIN = os.environ.get("PANEL_PIN", "1234")  # İstəsən serverdə dəyişə bilərsən

PAYMENT_RATES = {
    "Kart / Apple Pay / Google Pay": 1.5,
    "Nağd ödəniş": 0.0,
}

SVG_ICON = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512">
  <rect width="512" height="512" rx="110" fill="#0f172a"/>
  <rect x="44" y="44" width="424" height="424" rx="82" fill="#1e293b" stroke="#334155" stroke-width="8"/>
  <text x="256" y="215" text-anchor="middle" fill="#38bdf8" font-family="Arial, sans-serif" font-size="112" font-weight="900">BM</text>
  <text x="256" y="335" text-anchor="middle" fill="#22c55e" font-family="Arial, sans-serif" font-size="82" font-weight="800">+₼</text>
</svg>"""


def init_db():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            order_no TEXT,
            product_name TEXT NOT NULL,
            qty INTEGER NOT NULL DEFAULT 1,
            category_name TEXT,
            cat_percent REAL NOT NULL DEFAULT 0.0,
            buy_price REAL NOT NULL DEFAULT 0.0,
            sell_price REAL NOT NULL DEFAULT 0.0,
            logistics_class TEXT,
            logistics_fee REAL NOT NULL DEFAULT 0.0,
            payment_method TEXT NOT NULL,
            pay_percent REAL NOT NULL DEFAULT 1.5,
            extra_cost REAL NOT NULL DEFAULT 0.0,
            cat_fee_azn REAL NOT NULL DEFAULT 0.0,
            pay_fee_azn REAL NOT NULL DEFAULT 0.0,
            net_profit REAL NOT NULL DEFAULT 0.0,
            status TEXT NOT NULL DEFAULT 'Gözləyir'
        )
    """)
    conn.commit()
    conn.close()


def load_categories():
    categories = []
    if not os.path.exists(EXCEL_PATH):
        return categories
    try:
        import openpyxl
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb.active
        for idx, row in enumerate(ws.iter_rows(values_only=True)):
            if idx == 0 or not row or len(row) < 4:
                continue
            vertikal = str(row[0] or "").strip()
            grup = str(row[1] or "").strip()
            alt = str(row[2] or "").strip()
            try:
                tarif = float(str(row[3]).replace(",", ".").strip())
            except (TypeError, ValueError):
                continue
            if alt:
                categories.append({
                    "vertikal": vertikal,
                    "grup": grup,
                    "alt": alt,
                    "tarif": round(tarif, 2),
                })
    except Exception:
        pass
    return categories


CATEGORIES_CACHE = load_categories()


def compute_order(data):
    qty = max(1, int(float(str(data.get("qty", 1)).replace(",", ".") or 1)))
    cat_pct = max(0.0, float(str(data.get("cat_percent", 0)).replace(",", ".") or 0))
    buy_unit = max(0.0, float(str(data.get("buy_price", 0)).replace(",", ".") or 0))
    sell_unit = max(0.0, float(str(data.get("sell_price", 0)).replace(",", ".") or 0))
    log_fee = max(0.0, float(str(data.get("logistics_fee", 0)).replace(",", ".") or 0))
    extra_fee = max(0.0, float(str(data.get("extra_cost", 0)).replace(",", ".") or 0))

    pay_method = data.get("payment_method") or "Kart / Apple Pay / Google Pay"
    pay_pct = PAYMENT_RATES.get(pay_method, 1.5)

    total_buy = buy_unit * qty
    total_sell = sell_unit * qty
    cat_fee_azn = total_sell * (cat_pct / 100.0)
    pay_fee_azn = total_sell * (pay_pct / 100.0)
    net_profit = total_sell - total_buy - cat_fee_azn - pay_fee_azn - log_fee - extra_fee

    return {
        "qty": qty,
        "cat_pct": round(cat_pct, 2),
        "total_buy": round(total_buy, 4),
        "total_sell": round(total_sell, 4),
        "cat_fee_azn": round(cat_fee_azn, 4),
        "pay_pct": round(pay_pct, 2),
        "pay_fee_azn": round(pay_fee_azn, 4),
        "log_fee": round(log_fee, 4),
        "extra_fee": round(extra_fee, 4),
        "net_profit": round(net_profit, 4),
    }


HTML_PAGE = """<!DOCTYPE html>
<html lang="az">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <meta name="theme-color" content="#0f172a" />
  <meta name="mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-capable" content="yes" />
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <meta name="apple-mobile-web-app-title" content="Birmarket Panel" />
  <link rel="manifest" href="/manifest.json" />
  <link rel="icon" type="image/svg+xml" href="/icon.svg" />
  <link rel="apple-touch-icon" href="/icon.svg" />
  <title>Birmarket Mobil Mənfəət Paneli</title>
  <style>
    :root {
      --bg: #f1f5f9;
      --card: #ffffff;
      --text: #0f172a;
      --muted: #64748b;
      --primary: #1d4ed8;
      --success: #15803d;
      --success-bg: #dcfce7;
      --danger: #b91c1c;
      --danger-bg: #fee2e2;
      --warn: #b45309;
      --border: #e2e8f0;
    }
    * { box-sizing: border-box; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    body { margin: 0; background: var(--bg); color: var(--text); }
    header {
      background: #0f172a; color: #fff; padding: 14px 16px;
      position: sticky; top: 0; z-index: 20; display: flex; justify-content: space-between; align-items: center;
    }
    header h1 { margin: 0; font-size: 16px; font-weight: 700; }
    header span { font-size: 12px; color: #94a3b8; }
    .container { max-width: 1120px; margin: 0 auto; padding: 14px; }
    .stats { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px; margin-bottom: 14px; }
    @media (min-width: 768px) {
      .stats { grid-template-columns: repeat(4, 1fr); }
      .grid-main { display: grid; grid-template-columns: 400px 1fr; gap: 14px; align-items: start; }
    }
    .stat-card {
      background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 12px;
    }
    .stat-title { font-size: 11px; color: var(--muted); font-weight: 700; text-transform: uppercase; }
    .stat-val { font-size: 18px; font-weight: 800; margin-top: 4px; }
    .card {
      background: var(--card); border: 1px solid var(--border); border-radius: 14px; padding: 14px; margin-bottom: 14px;
    }
    .card h2 { margin: 0 0 12px 0; font-size: 15px; color: #1e3a8a; }
    .field { margin-bottom: 10px; }
    label { display: block; font-size: 12px; font-weight: 600; color: #334155; margin-bottom: 4px; }
    input, select {
      width: 100%; padding: 10px 11px; border: 1px solid #cbd5e1; border-radius: 9px;
      font-size: 14px; background: #fff; color: var(--text);
    }
    .row-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .preview-box {
      background: #f8fafc; border: 1px solid var(--border); border-radius: 10px; padding: 10px; margin: 12px 0; font-size: 13px;
    }
    .preview-line { display: flex; justify-content: space-between; margin: 3px 0; }
    .net-banner {
      margin-top: 8px; padding: 10px; border-radius: 9px; text-align: center; font-weight: 800; font-size: 15px;
    }
    .net-pos { background: var(--success-bg); color: var(--success); }
    .net-neg { background: var(--danger-bg); color: var(--danger); }
    .btn {
      width: 100%; border: none; border-radius: 10px; padding: 11px 14px; font-size: 14px;
      font-weight: 700; cursor: pointer;
    }
    .btn-save { background: var(--success); color: #fff; }
    .btn-sm {
      width: auto; padding: 6px 10px; font-size: 12px; border-radius: 7px; margin-right: 4px;
    }
    .btn-accept { background: #16a34a; color: #fff; }
    .btn-cancel { background: #f59e0b; color: #fff; }
    .btn-del { background: #dc2626; color: #fff; }
    .order-item {
      border: 1px solid var(--border); border-radius: 11px; padding: 11px; margin-bottom: 10px; background: #fcfcfd;
    }
    .order-top { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
    .order-title { font-weight: 700; font-size: 14px; }
    .badge {
      font-size: 11px; font-weight: 700; padding: 3px 8px; border-radius: 999px;
      background: #e2e8f0; color: #1e293b;
    }
    .badge-ok { background: #dcfce7; color: #166534; }
    .badge-no { background: #fee2e2; color: #991b1b; }
    .order-meta { font-size: 12px; color: var(--muted); margin-top: 4px; }
    .order-profit { font-weight: 800; font-size: 14px; margin-top: 6px; }
    .order-actions { margin-top: 8px; display: flex; flex-wrap: wrap; gap: 4px; }
  </style>
</head>
<body>
  <header>
    <div>
      <h1>BIRMARKET MOBİL PANEL</h1>
      <span>Sifariş qəbulu və xalis mənfəət hesabı</span>
    </div>
    <span id="catCount"></span>
  </header>

  <div class="container">
    <div class="stats">
      <div class="stat-card">
        <div class="stat-title">Aktiv Sifariş</div>
        <div class="stat-val" id="stOrders">0</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Satış Dövriyyəsi</div>
        <div class="stat-val" style="color:#0369a1" id="stSales">0.00 ₼</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Komissiya + Logistika</div>
        <div class="stat-val" style="color:#b45309" id="stFees">0.00 ₼</div>
      </div>
      <div class="stat-card">
        <div class="stat-title">Xalis Mənfəət</div>
        <div class="stat-val" style="color:#15803d" id="stProfit">0.00 ₼</div>
      </div>
    </div>

    <div class="grid-main">
      <div class="card">
        <h2>Yeni Sifariş Qeyd Et</h2>
        <div class="row-2">
          <div class="field">
            <label>Sifariş №</label>
            <input id="order_no" placeholder="məs: 10452" />
          </div>
          <div class="field">
            <label>Say (ədəd)</label>
            <input id="qty" type="number" min="1" value="1" oninput="recalc()" />
          </div>
        </div>

        <div class="field">
          <label>Məhsulun adı</label>
          <input id="product_name" placeholder="məs: My Lady / Head & Shoulders" />
        </div>

        <div class="field">
          <label>Kateqoriya axtar</label>
          <input id="cat_search" placeholder="Axtar: bez, şampun, geyim..." oninput="filterCats()" />
        </div>

        <div class="field">
          <label>Kateqoriya seçimi</label>
          <select id="cat_select" onchange="onCatChange()"></select>
        </div>

        <div class="row-2">
          <div class="field">
            <label>Komissiya (%)</label>
            <input id="cat_percent" type="number" step="0.01" value="12.00" oninput="recalc()" />
          </div>
          <div class="field">
            <label>Ödəniş üsulu</label>
            <select id="payment_method" onchange="recalc()">
              <option value="Kart / Apple Pay / Google Pay">Kart / Apple / Google Pay (1.5%)</option>
              <option value="Nağd ödəniş">Nağd ödəniş (0.0%)</option>
            </select>
          </div>
        </div>

        <div class="row-2">
          <div class="field">
            <label>Alış / Maya (1 əd, ₼)</label>
            <input id="buy_price" type="number" step="0.01" value="1.00" oninput="recalc()" />
          </div>
          <div class="field">
            <label>Satış qiyməti (1 əd, ₼)</label>
            <input id="sell_price" type="number" step="0.01" value="1.50" oninput="recalc()" />
          </div>
        </div>

        <div class="row-2">
          <div class="field">
            <label>Logistika sinfi</label>
            <select id="logistics_class">
              <option>Paket (PACKAGE)</option>
              <option>Standart (STANDARD)</option>
              <option>KQT (OVERSIZED)</option>
              <option>KQT+ (EXTRA)</option>
            </select>
          </div>
          <div class="field">
            <label>Logistika haqqı (₼)</label>
            <input id="logistics_fee" type="number" step="0.01" value="0.00" oninput="recalc()" />
          </div>
        </div>

        <div class="field">
          <label>Əlavə / Paketləmə xərci (₼)</label>
          <input id="extra_cost" type="number" step="0.01" value="0.00" oninput="recalc()" />
        </div>

        <div class="preview-box">
          <div class="preview-line"><span>Ümumi satış:</span><strong id="p_sell">0.00 ₼</strong></div>
          <div class="preview-line"><span>Maya dəyəri:</span><span id="p_buy">0.00 ₼</span></div>
          <div class="preview-line"><span>Kateqoriya komissiyası:</span><span id="p_cat">0.00 ₼</span></div>
          <div class="preview-line"><span>Ödəniş komissiyası:</span><span id="p_pay">0.00 ₼</span></div>
          <div class="preview-line"><span>Logistika + əlavə:</span><span id="p_log">0.00 ₼</span></div>
          <div id="p_net" class="net-banner net-pos">XALİS MƏNFƏƏT: 0.00 ₼</div>
        </div>

        <button class="btn btn-save" onclick="saveOrder()">SİFARİŞİ YADDA SAXLA</button>
      </div>

      <div class="card">
        <h2>Sifarişlər və Qəbul Paneli</h2>
        <div id="ordersList">Yüklənir...</div>
      </div>
    </div>
  </div>

  <script>
    if ('serviceWorker' in navigator) {
      window.addEventListener('load', () => {
        navigator.serviceWorker.register('/sw.js').catch(() => {});
      });
    }

    let allCats = [];

    async function loadCategories() {
      const res = await fetch('/api/categories');
      allCats = await res.json();
      document.getElementById('catCount').textContent = allCats.length + ' kateqoriya';
      renderCatOptions(allCats);
    }

    function renderCatOptions(list) {
      const sel = document.getElementById('cat_select');
      sel.innerHTML = '';
      list.slice(0, 250).forEach(c => {
        const opt = document.createElement('option');
        opt.value = c.tarif;
        opt.dataset.alt = c.alt;
        opt.textContent = `${c.alt} (${c.tarif.toFixed(2)}%)`;
        sel.appendChild(opt);
      });
    }

    function filterCats() {
      const q = document.getElementById('cat_search').value.trim().toLowerCase();
      if (!q) return renderCatOptions(allCats);
      const filtered = allCats.filter(c =>
        c.alt.toLowerCase().includes(q) || c.grup.toLowerCase().includes(q)
      );
      renderCatOptions(filtered);
    }

    function onCatChange() {
      const sel = document.getElementById('cat_select');
      if (sel.value !== '') {
        document.getElementById('cat_percent').value = parseFloat(sel.value).toFixed(2);
        recalc();
      }
    }

    function num(id) {
      return parseFloat(document.getElementById(id).value) || 0;
    }

    function recalc() {
      const qty = Math.max(1, parseInt(document.getElementById('qty').value) || 1);
      const catPct = Math.max(0, num('cat_percent'));
      const buy = Math.max(0, num('buy_price')) * qty;
      const sell = Math.max(0, num('sell_price')) * qty;
      const logFee = Math.max(0, num('logistics_fee'));
      const extraFee = Math.max(0, num('extra_cost'));
      const payMethod = document.getElementById('payment_method').value;
      const payPct = payMethod.includes('Nağd') ? 0 : 1.5;

      const catFee = sell * (catPct / 100);
      const payFee = sell * (payPct / 100);
      const net = sell - buy - catFee - payFee - logFee - extraFee;

      document.getElementById('p_sell').textContent = sell.toFixed(2) + ' ₼';
      document.getElementById('p_buy').textContent = '-' + buy.toFixed(2) + ' ₼';
      document.getElementById('p_cat').textContent = '-' + catFee.toFixed(4) + ' ₼';
      document.getElementById('p_pay').textContent = '-' + payFee.toFixed(4) + ' ₼';
      document.getElementById('p_log').textContent = '-' + (logFee + extraFee).toFixed(2) + ' ₼';

      const banner = document.getElementById('p_net');
      if (net >= 0) {
        banner.className = 'net-banner net-pos';
        banner.textContent = `XALİS MƏNFƏƏT: +${net.toFixed(2)} ₼`;
      } else {
        banner.className = 'net-banner net-neg';
        banner.textContent = `ZƏRƏR: ${net.toFixed(2)} ₼`;
      }
    }

    async function saveOrder() {
      const prod = document.getElementById('product_name').value.trim();
      if (!prod) return alert('Məhsulun adını yazın.');

      const sel = document.getElementById('cat_select');
      const catName = sel.options[sel.selectedIndex]?.dataset?.alt || 'Digər';

      const payload = {
        order_no: document.getElementById('order_no').value.trim(),
        product_name: prod,
        qty: document.getElementById('qty').value,
        category_name: catName,
        cat_percent: document.getElementById('cat_percent').value,
        buy_price: document.getElementById('buy_price').value,
        sell_price: document.getElementById('sell_price').value,
        logistics_class: document.getElementById('logistics_class').value,
        logistics_fee: document.getElementById('logistics_fee').value,
        payment_method: document.getElementById('payment_method').value,
        extra_cost: document.getElementById('extra_cost').value
      };

      const res = await fetch('/api/orders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        document.getElementById('order_no').value = '';
        document.getElementById('product_name').value = '';
        loadOrders();
      }
    }

    async function setStatus(id, status) {
      await fetch(`/api/orders/${id}/status`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status })
      });
      loadOrders();
    }

    async function deleteOrder(id) {
      if (!confirm('Bu sifarişi silmək istəyirsən?')) return;
      await fetch(`/api/orders/${id}`, { method: 'DELETE' });
      loadOrders();
    }

    async function loadOrders() {
      const res = await fetch('/api/orders');
      const data = await res.json();

      document.getElementById('stOrders').textContent = data.summary.active_count;
      document.getElementById('stSales').textContent = data.summary.total_sales.toFixed(2) + ' ₼';
      document.getElementById('stFees').textContent = data.summary.total_fees.toFixed(2) + ' ₼';
      document.getElementById('stProfit').textContent = data.summary.net_profit.toFixed(2) + ' ₼';

      const wrap = document.getElementById('ordersList');
      if (!data.orders.length) {
        wrap.innerHTML = '<p style="color:#64748b">Hələ sifariş qeyd edilməyib.</p>';
        return;
      }

      wrap.innerHTML = data.orders.map(o => {
        const badgeClass = o.status === 'Qəbul edildi' ? 'badge badge-ok'
          : (o.status === 'Ləğv / Qaytarıldı' ? 'badge badge-no' : 'badge');
        return `
          <div class="order-item">
            <div class="order-top">
              <div class="order-title">#${o.id} ${o.order_no ? '(' + o.order_no + ')' : ''} — ${o.product_name} ×${o.qty}</div>
              <span class="${badgeClass}">${o.status}</span>
            </div>
            <div class="order-meta">
              ${o.created_at} • ${o.category_name} (${o.cat_percent}%) • Satış: ${o.sell_price.toFixed(2)} ₼ • Alış: ${o.buy_price.toFixed(2)} ₼
            </div>
            <div class="order-profit" style="color:${o.net_profit >= 0 ? '#15803d' : '#b91c1c'}">
              Xalis Mənfəət: ${o.net_profit >= 0 ? '+' : ''}${o.net_profit.toFixed(2)} ₼
            </div>
            <div class="order-actions">
              <button class="btn btn-sm btn-accept" onclick="setStatus(${o.id}, 'Qəbul edildi')">Qəbul Et</button>
              <button class="btn btn-sm btn-cancel" onclick="setStatus(${o.id}, 'Ləğv / Qaytarıldı')">Ləğv / Qaytar</button>
              <button class="btn btn-sm btn-del" onclick="deleteOrder(${o.id})">Sil</button>
            </div>
          </div>
        `;
      }).join('');
    }

    loadCategories().then(() => { onCatChange(); recalc(); });
    loadOrders();
  </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_PAGE)


@app.route("/manifest.json")
def manifest():
    data = """{
      "name": "Birmarket Mobil Panel",
      "short_name": "Birmarket Panel",
      "description": "Birmarket sifariş qəbulu və xalis mənfəət hesabı",
      "start_url": "/",
      "scope": "/",
      "display": "standalone",
      "orientation": "portrait",
      "background_color": "#0f172a",
      "theme_color": "#0f172a",
      "icons": [
        {
          "src": "/icon.svg",
          "sizes": "any",
          "type": "image/svg+xml",
          "purpose": "any maskable"
        }
      ]
    }"""
    return Response(data, mimetype="application/manifest+json")


@app.route("/sw.js")
def service_worker():
    js = """
    self.addEventListener('install', event => {
      self.skipWaiting();
    });
    self.addEventListener('activate', event => {
      event.waitUntil(self.clients.claim());
    });
    self.addEventListener('fetch', event => {
      event.respondWith(fetch(event.request));
    });
    """
    return Response(js, mimetype="application/javascript")


@app.route("/icon.svg")
def icon_svg():
    return Response(SVG_ICON, mimetype="image/svg+xml")


@app.route("/api/categories")
def api_categories():
    return jsonify(CATEGORIES_CACHE)


@app.route("/api/orders", methods=["GET"])
def get_orders():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM orders ORDER BY id DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()

    active_rows = [r for r in rows if r.get("status") != "Ləğv / Qaytarıldı"]
    active_count = len(active_rows)
    total_sales = sum(r["sell_price"] for r in active_rows)
    total_fees = sum(r["cat_fee_azn"] + r["pay_fee_azn"] + r["logistics_fee"] + r["extra_cost"] for r in active_rows)
    net_profit = sum(r["net_profit"] for r in active_rows)

    return jsonify({
        "orders": rows,
        "summary": {
            "active_count": active_count,
            "total_sales": round(total_sales, 2),
            "total_fees": round(total_fees, 2),
            "net_profit": round(net_profit, 2),
        }
    })


@app.route("/api/orders", methods=["POST"])
def create_order():
    data = request.get_json(force=True) or {}
    prod = str(data.get("product_name", "")).strip()
    if not prod:
        return jsonify({"error": "Məhsulun adı boş ola bilməz"}), 400

    calc = compute_order(data)
    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO orders (
            created_at, order_no, product_name, qty, category_name, cat_percent,
            buy_price, sell_price, logistics_class, logistics_fee,
            payment_method, pay_percent, extra_cost, cat_fee_azn, pay_fee_azn,
            net_profit, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Gözləyir')
    """, (
        now_str,
        str(data.get("order_no", "")).strip(),
        prod,
        calc["qty"],
        str(data.get("category_name", "Digər")).strip() or "Digər",
        calc["cat_pct"],
        calc["total_buy"],
        calc["total_sell"],
        str(data.get("logistics_class", "Paket (PACKAGE)")),
        calc["log_fee"],
        str(data.get("payment_method", "Kart / Apple Pay / Google Pay")),
        calc["pay_pct"],
        calc["extra_fee"],
        calc["cat_fee_azn"],
        calc["pay_fee_azn"],
        calc["net_profit"],
    ))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return jsonify({"id": new_id, "status": "ok"}), 201


@app.route("/api/orders/<int:order_id>/status", methods=["PATCH"])
def update_status(order_id):
    data = request.get_json(force=True) or {}
    status = str(data.get("status", "Gözləyir")).strip()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("UPDATE orders SET status = ? WHERE id = ?", (status, order_id))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})


@app.route("/api/orders/<int:order_id>", methods=["DELETE"])
def delete_order(order_id):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})


init_db()

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)

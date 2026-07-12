// ── Google Apps Script Web App URL ───────────────────────────
const GAS_URL = 'https://script.google.com/macros/s/AKfycby-CppfB0E1OkV3n38RT4V88srnzV178_fXRWIlEzlDROc7AxeZYFYY4X9HSzzTpp1h/exec';

// ── CORS Proxy（透過自家 GAS，避開 allorigins 等第三方）─────
function p(url) {
  return GAS_URL + '?proxy=' + encodeURIComponent(url);
}

async function gasGet() {
  const resp = await fetch(GAS_URL, { redirect: 'follow' });
  return resp.json();
}

async function gasPost(body) {
  const resp = await fetch(GAS_URL, {
    method: 'POST',
    redirect: 'follow',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify(body),
  });
  return resp.json();
}

// ── localStorage 封裝（取代 chrome.storage）──────────────────
function storageGet(keys) {
  const result = {};
  for (const key of keys) {
    const val = localStorage.getItem(key);
    try { result[key] = val !== null ? JSON.parse(val) : undefined; }
    catch { result[key] = val; }
  }
  return result;
}

function storageSet(obj) {
  for (const [key, val] of Object.entries(obj)) {
    localStorage.setItem(key, JSON.stringify(val));
  }
}

// ── Utilities ────────────────────────────────────────────────

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function parseNum(s) {
  return parseFloat(String(s).replace(/,/g, ''));
}

// ── API ──────────────────────────────────────────────────────

async function getRealtimeInfo(stockNo) {
  // 用 Yahoo Finance：TWSE MIS API 會擋 Google IP，透過 GAS proxy 拿不到
  // 中文簡稱從本機 stock list 補上（Yahoo 的 longName/shortName 是英文）
  for (const suffix of ['TW', 'TWO']) {
    try {
      const url = `https://query1.finance.yahoo.com/v8/finance/chart/${stockNo}.${suffix}?interval=1d&range=5d`;
      const resp = await fetch(p(url));
      const data = await resp.json();
      const result = data?.chart?.result?.[0];
      if (!result || !result.meta) continue;

      const m  = result.meta;
      const ts = result.timestamp || [];
      const q  = result.indicators?.quote?.[0] || {};
      const i  = ts.length - 1;
      if (i < 0) continue;

      const fmt = v => (v == null || isNaN(v)) ? '--' : Number(v).toFixed(2);

      let chineseName = null;
      try {
        const list = await loadStockList();
        const found = list.find(s => s.code === stockNo);
        if (found) chineseName = found.name;
      } catch { /* 拿不到中文名沒關係 */ }

      return {
        n: chineseName || m.longName || m.shortName || stockNo,
        z: fmt(m.regularMarketPrice),
        y: fmt(m.chartPreviousClose ?? m.previousClose),
        o: fmt(q.open?.[i]),
        h: fmt(q.high?.[i]),
        l: fmt(q.low?.[i]),
        v: q.volume?.[i] != null ? Math.round(q.volume[i] / 1000).toLocaleString() : '--',
      };
    } catch { /* try next suffix */ }
  }
  return null;
}

async function getHistory(stockNo, months) {
  // Yahoo Finance 用 .TW（上市）；上櫃需 .TWO 後綴。先試 .TW，沒資料再 fallback .TWO。
  const range = months <= 1 ? '1mo' : months <= 3 ? '3mo' : months <= 6 ? '6mo' : '1y';

  for (const suffix of ['TW', 'TWO']) {
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${stockNo}.${suffix}?interval=1d&range=${range}`;
    try {
      const resp = await fetch(p(url));
      const data = await resp.json();
      const result = data?.chart?.result?.[0];
      if (!result || !result.timestamp?.length) continue;

      const ts = result.timestamp;
      const q  = result.indicators?.quote?.[0] || {};

      const rows = ts.map((sec, i) => {
        const d = new Date(sec * 1000);
        const time = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
        return {
          time,
          open:   q.open?.[i],
          high:   q.high?.[i],
          low:    q.low?.[i],
          close:  q.close?.[i],
          volume: q.volume?.[i] || 0,
        };
      }).filter(r => r.open != null && r.close != null);

      if (rows.length) return rows;
    } catch { /* try next suffix */ }
  }
  return [];
}

// ── Technical Indicators ─────────────────────────────────────

function movingAvg(closes, n) {
  return closes.map((_, i) => {
    const slice = closes.slice(Math.max(0, i - n + 1), i + 1);
    return slice.reduce((a, b) => a + b, 0) / slice.length;
  });
}

function calcEMA(data, n) {
  const k = 2 / (n + 1);
  const result = [];
  for (let i = 0; i < data.length; i++) {
    result.push(i === 0 ? data[0] : data[i] * k + result[i - 1] * (1 - k));
  }
  return result;
}

function calcRSI(closes, period = 14) {
  return closes.map((_, i) => {
    if (i < period) return null;
    const slice = closes.slice(i - period, i + 1);
    let gains = 0, losses = 0;
    for (let j = 1; j < slice.length; j++) {
      const diff = slice[j] - slice[j - 1];
      if (diff > 0) gains += diff;
      else losses -= diff;
    }
    const ag = gains / period, al = losses / period;
    return al === 0 ? 100 : 100 - 100 / (1 + ag / al);
  });
}

function calcMACD(closes, fast = 12, slow = 26, signal = 9) {
  const emaFast = calcEMA(closes, fast);
  const emaSlow = calcEMA(closes, slow);
  const macdLine = emaFast.map((v, i) => v - emaSlow[i]);
  const signalLine = calcEMA(macdLine, signal);
  const histogram = macdLine.map((v, i) => v - signalLine[i]);
  return { macdLine, signalLine, histogram };
}

function calcBollinger(closes, period = 20, stdDev = 2) {
  return closes.map((_, i) => {
    const w = closes.slice(Math.max(0, i - period + 1), i + 1);
    const ma = w.reduce((a, b) => a + b, 0) / w.length;
    const std = Math.sqrt(w.reduce((a, b) => a + (b - ma) ** 2, 0) / w.length);
    return { upper: ma + stdDev * std, lower: ma - stdDev * std };
  });
}

// ── Charts ───────────────────────────────────────────────────

const LC = LightweightCharts;
let activeCharts = [];
let resizeObserver = null;

const BASE_OPTS = {
  layout: {
    background: { color: '#1e1e2e' },
    textColor: '#aaaacc',
  },
  grid: {
    vertLines: { color: '#252535' },
    horzLines: { color: '#252535' },
  },
  rightPriceScale: { borderColor: '#3a3a5e' },
  timeScale: { borderColor: '#3a3a5e', timeVisible: false },
  crosshair: { mode: LC.CrosshairMode.Normal },
};

function makeChart(id, height) {
  const el = document.getElementById(id);
  el.style.height = height + 'px';
  const chart = LC.createChart(el, { ...BASE_OPTS, width: el.clientWidth, height });
  activeCharts.push({ chart, el });
  return chart;
}

function destroyCharts() {
  if (resizeObserver) { resizeObserver.disconnect(); resizeObserver = null; }
  activeCharts.forEach(({ chart }) => chart.remove());
  activeCharts = [];
}

function drawCharts(stockNo, records) {
  destroyCharts();

  const area = document.getElementById('chartArea');
  const totalH = area.clientHeight - 8;
  const h = [
    Math.round(totalH * 0.43),
    Math.round(totalH * 0.22),
    Math.round(totalH * 0.18),
    Math.round(totalH * 0.17),
  ];

  const times  = records.map(r => r.time);
  const closes = records.map(r => r.close);

  // K線 + MA + Bollinger
  const c1 = makeChart('chart-candle', h[0]);
  const candle = c1.addCandlestickSeries({
    upColor: '#ff4444', downColor: '#44cc88',
    borderUpColor: '#ff4444', borderDownColor: '#44cc88',
    wickUpColor: '#ff4444', wickDownColor: '#44cc88',
  });
  candle.setData(records.map(r => ({
    time: r.time, open: r.open, high: r.high, low: r.low, close: r.close,
  })));

  const addLine = (chart, values, color, dash = false) => {
    const s = chart.addLineSeries({
      color, lineWidth: 1,
      lineStyle: dash ? LC.LineStyle.Dashed : LC.LineStyle.Solid,
      priceLineVisible: false, lastValueVisible: false,
    });
    s.setData(times.map((t, i) => ({ time: t, value: values[i] })));
    return s;
  };

  addLine(c1, movingAvg(closes, 5),  '#ffaa33');
  addLine(c1, movingAvg(closes, 20), '#5599ff');
  const bb = calcBollinger(closes);
  addLine(c1, bb.map(b => b.upper), '#cc88ff', true);
  addLine(c1, bb.map(b => b.lower), '#cc88ff', true);
  c1.applyOptions({ watermark: { visible: true, text: `${stockNo}  MA5  MA20  BB`, color: 'rgba(150,150,200,0.35)', fontSize: 13, horzAlign: 'left', vertAlign: 'top' } });

  // MACD
  const c2 = makeChart('chart-macd', h[1]);
  const { macdLine, signalLine, histogram } = calcMACD(closes);
  const hist = c2.addHistogramSeries({ priceLineVisible: false, lastValueVisible: false });
  hist.setData(times.map((t, i) => ({
    time: t, value: histogram[i],
    color: histogram[i] >= 0 ? '#ff4444' : '#44cc88',
  })));
  addLine(c2, macdLine,   '#ffaa33');
  addLine(c2, signalLine, '#5599ff');
  c2.applyOptions({ watermark: { visible: true, text: 'MACD (12, 26, 9)', color: 'rgba(150,150,200,0.35)', fontSize: 13, horzAlign: 'left', vertAlign: 'top' } });

  // RSI
  const c3 = makeChart('chart-rsi', h[2]);
  const rsi = calcRSI(closes);
  const rsiSeries = c3.addLineSeries({ color: '#33ddcc', lineWidth: 1, priceLineVisible: false, lastValueVisible: false });
  rsiSeries.setData(
    times.map((t, i) => rsi[i] !== null ? { time: t, value: rsi[i] } : null).filter(Boolean)
  );
  rsiSeries.createPriceLine({ price: 70, color: '#ff4444', lineWidth: 1, lineStyle: LC.LineStyle.Dashed, axisLabelVisible: true });
  rsiSeries.createPriceLine({ price: 30, color: '#44cc88', lineWidth: 1, lineStyle: LC.LineStyle.Dashed, axisLabelVisible: true });
  c3.applyOptions({ watermark: { visible: true, text: 'RSI (14)', color: 'rgba(150,150,200,0.35)', fontSize: 13, horzAlign: 'left', vertAlign: 'top' } });

  // 成交量
  const c4 = makeChart('chart-volume', h[3]);
  const vol = c4.addHistogramSeries({ priceLineVisible: false, lastValueVisible: false });
  vol.setData(records.map(r => ({
    time: r.time, value: r.volume,
    color: r.close >= r.open ? '#ff4444' : '#44cc88',
  })));
  c4.applyOptions({ watermark: { visible: true, text: '成交量', color: 'rgba(150,150,200,0.35)', fontSize: 13, horzAlign: 'left', vertAlign: 'top' } });

  // 同步時間軸
  const allCharts = [c1, c2, c3, c4];
  allCharts.forEach((c, idx) => {
    c.timeScale().subscribeVisibleLogicalRangeChange(range => {
      if (!range) return;
      allCharts.forEach((other, j) => {
        if (j !== idx) other.timeScale().setVisibleLogicalRange(range);
      });
    });
  });

  // 視窗縮放
  resizeObserver = new ResizeObserver(() => {
    activeCharts.forEach(({ chart, el }) => {
      chart.applyOptions({ width: el.clientWidth });
    });
  });
  resizeObserver.observe(area);
}

// ── Info Panel ───────────────────────────────────────────────

function updateInfoPanel(info, stockNo) {
  if (!info) return;
  const price = info.z || '--';
  const ref   = info.y || '--';
  let change = '--', color = 'white';
  const diff = parseFloat(price) - parseFloat(ref);
  if (!isNaN(diff)) {
    change = (diff >= 0 ? '+' : '') + diff.toFixed(2);
    color  = diff >= 0 ? '#ff6666' : '#66ff99';
  }

  document.getElementById('info-name').textContent  = `${info.n || stockNo}（${stockNo}）`;
  const priceEl = document.getElementById('info-price');
  priceEl.textContent = price;
  priceEl.style.color = color;
  document.getElementById('info-open').textContent   = info.o || '--';
  document.getElementById('info-high').textContent   = info.h || '--';
  document.getElementById('info-low').textContent    = info.l || '--';
  document.getElementById('info-ref').textContent    = ref;
  document.getElementById('info-vol').textContent    = info.v || '--';
  const changeEl = document.getElementById('info-change');
  changeEl.textContent = change;
  changeEl.style.color = color;
}

// ── Query ────────────────────────────────────────────────────

async function onQuery() {
  const stockNo = document.getElementById('stockInput').value.trim();
  if (!stockNo) return;
  const months = Math.max(1, Math.min(12, parseInt(document.getElementById('monthsInput').value) || 3));

  const btn    = document.getElementById('queryBtn');
  const status = document.getElementById('statusLabel');
  btn.disabled = true;
  btn.textContent = '查詢中...';
  status.textContent = '正在下載資料，請稍候...';

  try {
    const [info, records] = await Promise.all([
      getRealtimeInfo(stockNo),
      getHistory(stockNo, months),
    ]);

    updateInfoPanel(info, stockNo);

    const parts = [];
    if (records.length === 0) {
      parts.push('查無資料，請確認股票代號');
    } else {
      drawCharts(stockNo, records);
      parts.push(`共 ${records.length} 筆資料`);
    }
    if (!info) parts.push('（即時資料查詢失敗）');
    status.textContent = parts.join(' ');
  } catch (err) {
    status.textContent = `錯誤：${err.message}`;
  } finally {
    btn.disabled = false;
    btn.textContent = '查詢 + 繪圖';
  }
}

document.getElementById('queryBtn').addEventListener('click', onQuery);
document.getElementById('stockInput').addEventListener('keydown', e => {
  if (e.key === 'Enter') onQuery();
});

// ── Tab switching ─────────────────────────────────────────────

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const tab = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(`tab-${tab}`).classList.add('active');
    if (tab === 'alerts') loadAlerts();
    if (tab === 'holdings') loadHoldings();
    if (tab === 'market') startMarket();
    else stopMarket();          // 離開看盤頁就停掉自動刷新，省請求
  });
});

// 從其他分頁（如存股清單）直接切到圖表分析並繪圖
function showChartFor(stockNo) {
  if (!stockNo) return;
  stopMarket();   // 若從看盤頁跳來，停掉自動刷新計時器
  // 切到「圖表分析」分頁
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
  document.querySelector('.tab-btn[data-tab="chart"]').classList.add('active');
  document.getElementById('tab-chart').classList.add('active');
  // 填入代號並查詢繪圖
  document.getElementById('stockInput').value = stockNo;
  onQuery();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

// ── 圖表頁「我的持股」快選清單 ─────────────────────────
async function loadQuickHoldings() {
  const el = document.getElementById('quickHoldingsList');
  if (!el) return;
  try {
    if (!holdingsState.holdings.length) {
      const data = await gasGet();
      holdingsState.holdings = data.holdings || [];
    }
    renderQuickHoldings();
  } catch {
    el.innerHTML = '<span class="quick-empty" style="color:#ff8888">載入失敗</span>';
  }
}

function renderQuickHoldings() {
  const el = document.getElementById('quickHoldingsList');
  if (!el) return;
  const seen = new Set();
  const uniq = [];
  for (const h of holdingsState.holdings) {
    if (seen.has(h.stockNo)) continue;   // 同代號多帳戶只列一次
    seen.add(h.stockNo);
    uniq.push(h);
  }
  if (!uniq.length) {
    el.innerHTML = '<span class="quick-empty">尚無持股</span>';
    return;
  }
  el.innerHTML = uniq.map(h =>
    `<button class="quick-holding" onclick="showChartFor('${h.stockNo}')" title="${escapeHtml(h.stockName || '')}">` +
      `<span class="qh-code">${h.stockNo}</span>` +
      `<span class="qh-name">${escapeHtml(h.stockName || '')}</span>` +
    `</button>`
  ).join('');
}

// ── Alerts Management ─────────────────────────────────────────

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

const alertsState = { alerts: [], holdings: [], extra: [] };  // extra：手動加入監控但尚未存的股票

async function loadAlerts() {
  const { telegramToken, telegramChatId } = storageGet(['telegramToken', 'telegramChatId']);
  document.getElementById('telegramTokenInput').value  = telegramToken || '';
  document.getElementById('telegramChatIdInput').value = telegramChatId || '';

  const tbody = document.getElementById('alertTbody');
  tbody.innerHTML = '<tr><td colspan="7" class="no-holdings">載入中...</td></tr>';
  try {
    const data = await gasGet();
    alertsState.alerts   = data.alerts || [];
    alertsState.holdings = data.holdings || [];
    renderAlertTable();
  } catch {
    tbody.innerHTML = '<tr><td colspan="7" class="no-holdings" style="color:#ff8888">載入失敗，請確認 GAS 部署</td></tr>';
  }
}

// 把 持股 + 既有警報 + 手動加入的股票，整理成「每檔一列」，買點掛 lte、賣點掛 gte
function buildAlertRows() {
  const map = new Map();
  const ensure = (stockNo, stockName) => {
    if (!map.has(stockNo)) map.set(stockNo, { stockNo, stockName: stockName || '', buy: null, sell: null });
    else if (stockName && !map.get(stockNo).stockName) map.get(stockNo).stockName = stockName;
    return map.get(stockNo);
  };
  for (const h of alertsState.holdings) ensure(h.stockNo, h.stockName);
  for (const e of alertsState.extra)    ensure(e.stockNo, e.stockName);
  for (const a of alertsState.alerts) {
    const row = ensure(a.stockNo, a.stockName);
    if (a.condition === 'lte')      row.buy  = a;
    else if (a.condition === 'gte') row.sell = a;
  }
  return Array.from(map.values());
}

function renderAlertTable() {
  const tbody = document.getElementById('alertTbody');
  const badge = document.getElementById('alertCountBadge');
  const rows  = buildAlertRows();

  const active = alertsState.alerts.filter(a => a.enabled && !a.triggered).length;
  badge.textContent = active ? `${active} 監控中` : '';

  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="7" class="no-holdings">尚無持股 — 可在上方搜尋加入股票</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map(r => {
    const buyVal  = r.buy  ? r.buy.targetPrice  : '';
    const sellVal = r.sell ? r.sell.targetPrice : '';
    return `
      <tr>
        <td class="code"><a class="code-link" onclick="showChartFor('${r.stockNo}')">${r.stockNo}</a></td>
        <td class="name">${escapeHtml(r.stockName || '')}</td>
        <td class="num" id="acur-${r.stockNo}">--</td>
        <td class="num"><input type="number" class="alert-price-input buy" id="buy-${r.stockNo}" value="${buyVal}" step="0.1" min="0" placeholder="跌到"></td>
        <td class="num"><input type="number" class="alert-price-input sell" id="sell-${r.stockNo}" value="${sellVal}" step="0.1" min="0" placeholder="漲到"></td>
        <td>${alertStatusHtml(r)}</td>
        <td>
          <div class="row-actions">
            <button class="btn-save" onclick="saveStockAlerts('${r.stockNo}')">儲存</button>
            <button class="btn-delete" onclick="clearStockAlerts('${r.stockNo}')">清除</button>
          </div>
        </td>
      </tr>`;
  }).join('');

  rows.forEach(r => fetchAlertRowPrice(r.stockNo));
}

function alertStatusHtml(r) {
  const one = (a, label) => {
    if (!a) return '';
    if (a.triggered)  return `<span class="alert-badge hit">${label}觸發 ${a.triggeredPrice}（${a.triggeredAt || ''}）</span>`;
    if (!a.enabled)   return `<span class="alert-badge paused">${label}暫停</span>`;
    return `<span class="alert-badge on">${label}監控中</span>`;
  };
  const html = [one(r.buy, '買點'), one(r.sell, '賣點')].filter(Boolean).join(' ');
  return html || '<span style="color:#555">—</span>';
}

async function fetchAlertRowPrice(stockNo) {
  const el = document.getElementById(`acur-${stockNo}`);
  if (!el) return;
  const info = await getRealtimeInfo(stockNo);
  let priceStr = null;
  if (info) {
    if (info.z && info.z !== '-' && info.z !== '--')      priceStr = info.z;
    else if (info.y && info.y !== '--')                   priceStr = info.y; // 休市用昨收
  }
  if (!priceStr) { el.textContent = '--'; return; }
  const price = parseFloat(priceStr);
  const ref   = info ? parseFloat(info.y) : NaN;
  const diff  = price - ref;
  el.textContent = priceStr;
  el.style.color = isNaN(diff) ? '#ccccee' : (diff >= 0 ? '#ff6666' : '#44cc88');
}

// 儲存某檔的買點/賣點（用既有 add/delete 動作 upsert，不需改後端）
async function saveStockAlerts(stockNo) {
  const row    = buildAlertRows().find(r => r.stockNo === stockNo) || { stockName: '' };
  const buyEl  = document.getElementById(`buy-${stockNo}`);
  const sellEl = document.getElementById(`sell-${stockNo}`);
  const status = document.getElementById('addAlertStatus');

  status.style.color = '#aaaacc';
  status.textContent = '儲存中...';
  try {
    await upsertAlert(stockNo, row.stockName, 'lte', buyEl.value,  row.buy);
    await upsertAlert(stockNo, row.stockName, 'gte', sellEl.value, row.sell);
    alertsState.extra = alertsState.extra.filter(e => e.stockNo !== stockNo);
    status.style.color = '#66ff99';
    status.textContent = `✓ 已更新 ${stockNo} ${row.stockName || ''} 的買賣點`;
    setTimeout(() => status.textContent = '', 2500);
    loadAlerts();
  } catch (err) {
    status.style.color = '#ff8888';
    status.textContent = `✗ 儲存失敗：${err.message}`;
  }
}

async function upsertAlert(stockNo, stockName, condition, priceStr, existing) {
  const val = parseFloat(priceStr);
  const valid = !isNaN(val) && val > 0;
  if (!valid) {                                   // 清空 → 刪除既有
    if (existing) await gasPost({ action: 'delete', id: existing.id });
    return;
  }
  // 價格相同、監控中且未觸發 → 不動；否則刪舊建新（順便重新武裝已觸發者）
  if (existing && Number(existing.targetPrice) === val && existing.enabled && !existing.triggered) return;
  if (existing) await gasPost({ action: 'delete', id: existing.id });
  const res = await gasPost({ action: 'add', alert: { id: genId(), stockNo, stockName, condition, targetPrice: val } });
  if (!res || !res.ok) throw new Error((res && res.error) || '伺服器錯誤');
}

async function clearStockAlerts(stockNo) {
  const row = buildAlertRows().find(r => r.stockNo === stockNo);
  if ((row && (row.buy || row.sell)) && !confirm(`清除 ${stockNo} 的買賣點警報？`)) return;
  const status = document.getElementById('addAlertStatus');
  status.style.color = '#aaaacc';
  status.textContent = '清除中...';
  try {
    if (row && row.buy)  await gasPost({ action: 'delete', id: row.buy.id });
    if (row && row.sell) await gasPost({ action: 'delete', id: row.sell.id });
    alertsState.extra = alertsState.extra.filter(e => e.stockNo !== stockNo);
    status.style.color = '#66ff99';
    status.textContent = `✓ 已清除 ${stockNo}`;
    setTimeout(() => status.textContent = '', 2000);
    loadAlerts();
  } catch (err) {
    status.style.color = '#ff8888';
    status.textContent = `✗ 清除失敗：${err.message}`;
  }
}

// 從搜尋加入一檔（非持股）到監控表
function addMonitorStock(stockNo, stockName) {
  const exists = buildAlertRows().some(r => r.stockNo === stockNo);
  if (!exists) alertsState.extra.push({ stockNo, stockName });
  renderAlertTable();
  setTimeout(() => {
    const el = document.getElementById(`buy-${stockNo}`);
    if (el) { el.focus(); el.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
  }, 50);
}

// ══════════════ 即時看盤（第四分頁）══════════════

const watchState = { list: [], quotes: {}, timer: null };
const MARKET_REFRESH_MS = 30000;   // 盤中每 30 秒
const UP_C = '#ff5555', DOWN_C = '#44cc88', FLAT_C = '#ccccee';

function saveWatchlistLocal() {
  localStorage.setItem('watchlist', JSON.stringify(watchState.list));
}
function loadWatchlistLocal() {
  try { watchState.list = JSON.parse(localStorage.getItem('watchlist') || '[]'); }
  catch { watchState.list = []; }
}

// 從後端載入觀察名單（跨裝置同步）；後端若尚未支援則退回本機
async function loadWatchlistFromServer() {
  loadWatchlistLocal();                       // 先用本機快取即時顯示
  try {
    const d = await gasGet();
    if (Array.isArray(d.watchlist)) {          // 後端已支援 watchlist
      if (d.watchlist.length === 0 && watchState.list.length > 0) {
        // 首次遷移：把本機既有清單上傳到後端
        for (const w of watchState.list) {
          try { await gasPost({ action: 'addWatch', watch: { stockNo: w.stockNo, stockName: w.stockName } }); } catch {}
        }
      } else {
        watchState.list = d.watchlist;         // 以後端為準
        saveWatchlistLocal();
      }
    }
  } catch { /* 離線／後端未更新：沿用本機 */ }
}

// 台北時間（不依賴使用者時區）
function twNow() {
  return new Date(Date.now() + new Date().getTimezoneOffset() * 60000 + 8 * 3600 * 1000);
}
function isTwTradingHours() {
  const t = twNow();
  const day = t.getDay(), hm = t.getHours() * 100 + t.getMinutes();
  return day >= 1 && day <= 5 && hm >= 900 && hm <= 1330;
}
function hms() {
  const t = twNow();
  const p = n => String(n).padStart(2, '0');
  return `${p(t.getHours())}:${p(t.getMinutes())}:${p(t.getSeconds())}`;
}

async function startMarket() {
  await loadWatchlistFromServer();
  renderWatchTable();
  refreshMarket();
  stopMarket();
  watchState.timer = setInterval(() => { if (!document.hidden) refreshMarket(); }, MARKET_REFRESH_MS);
}
function stopMarket() {
  if (watchState.timer) { clearInterval(watchState.timer); watchState.timer = null; }
}

async function refreshMarket() {
  updateMarketStatus();
  refreshIndices();
  await refreshWatchQuotes();
  const el = document.getElementById('watchUpdated');
  if (el) el.textContent = '更新 ' + hms();
}

function updateMarketStatus() {
  const el = document.getElementById('marketStatus');
  if (!el) return;
  const open = isTwTradingHours();
  el.textContent = open ? '🟢 盤中' : '⚪ 休市';
  el.style.color = open ? '#66ff99' : '#888899';
}

// 指數（^TWII 等）即時
async function getIndexInfo(symbol) {
  try {
    const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d&range=1d`;
    const resp = await fetch(p(url));
    const data = await resp.json();
    const r = data && data.chart && data.chart.result && data.chart.result[0];
    const m = r && r.meta;
    if (!m || m.regularMarketPrice == null) return null;
    return {
      price: Number(m.regularMarketPrice),
      prev:  Number(m.chartPreviousClose != null ? m.chartPreviousClose : m.previousClose),
    };
  } catch { return null; }
}

// 台指期近月（由 GAS 後端代打 TAIFEX）
async function getTaifexInfo() {
  try {
    const resp = await fetch(GAS_URL + '?taifex=1');
    const d = await resp.json();
    if (!d || d.error || d.price == null || isNaN(d.price)) return null;
    return d;
  } catch { return null; }
}

async function refreshIndices() {
  // ^ 必須先 percent-encode 成 %5E，否則 GAS 的 UrlFetchApp 會擋（引數無效）
  const twii = await getIndexInfo('%5ETWII');
  updateIndexCard('idxTwii', twii);

  const otc = await getIndexInfo('%5ETWOII');   // 櫃買 OTC 指數
  updateIndexCard('idxOtc', otc);

  const tx = await getTaifexInfo();
  if (tx) {
    updateIndexCard('idxTx', { price: tx.price, prev: tx.prev });
  } else {
    // 抓不到（可能休市或 TAIFEX 擋 Google IP）
    const c = document.getElementById('idxTx');
    if (c) {
      c.querySelector('.idx-price').textContent = '--';
      const chg = c.querySelector('.idx-change');
      chg.textContent = '暫無資料'; chg.style.color = '#888899';
    }
  }
}

function fmtIdx(v) {
  return Number(v).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function updateIndexCard(id, info) {
  const card = document.getElementById(id);
  if (!card) return;
  const priceEl = card.querySelector('.idx-price');
  const chgEl   = card.querySelector('.idx-change');
  if (!info || info.price == null) { priceEl.textContent = '--'; return; }
  const diff = (info.prev != null && !isNaN(info.prev)) ? info.price - info.prev : null;
  const pct  = (diff != null && info.prev) ? diff / info.prev * 100 : null;
  const col  = diff == null ? FLAT_C : (diff >= 0 ? UP_C : DOWN_C);
  priceEl.textContent = fmtIdx(info.price);
  priceEl.style.color = col;
  if (diff != null) {
    const arrow = diff >= 0 ? '▲' : '▼';
    chgEl.textContent = `${arrow} ${diff >= 0 ? '+' : ''}${diff.toFixed(2)}　(${pct >= 0 ? '+' : ''}${pct.toFixed(2)}%)`;
    chgEl.style.color = col;
  }
}

async function refreshWatchQuotes() {
  if (!watchState.list.length) return;
  const results = await Promise.all(watchState.list.map(async w => {
    const info = await getRealtimeInfo(w.stockNo);
    return [w.stockNo, info];
  }));
  watchState.quotes = Object.fromEntries(results);
  renderWatchTable();
}

function renderWatchTable() {
  const tbody = document.getElementById('watchTbody');
  const badge = document.getElementById('watchCountBadge');
  if (!tbody) return;
  badge.textContent = watchState.list.length ? `${watchState.list.length} 檔` : '';

  if (!watchState.list.length) {
    tbody.innerHTML = '<tr><td colspan="10" class="no-holdings">觀察名單是空的 — 上方搜尋加入股票</td></tr>';
    return;
  }

  tbody.innerHTML = watchState.list.map(w => {
    const info = watchState.quotes[w.stockNo];
    const name = escapeHtml(w.stockName || (info && info.n) || '');
    let priceStr = '--', diffStr = '--', pctStr = '--', col = FLAT_C;
    let o = '--', h = '--', l = '--', vol = '--';
    if (info) {
      o = info.o || '--'; h = info.h || '--'; l = info.l || '--'; vol = info.v || '--';
      const price = parseFloat(info.z);
      const ref   = parseFloat(info.y);
      if (!isNaN(price)) {
        priceStr = info.z;
        if (!isNaN(ref)) {
          const diff = price - ref;
          const pct  = ref ? diff / ref * 100 : 0;
          col = diff >= 0 ? UP_C : DOWN_C;
          diffStr = `${diff >= 0 ? '+' : ''}${diff.toFixed(2)}`;
          pctStr  = `${pct >= 0 ? '+' : ''}${pct.toFixed(2)}%`;
        }
      }
    }
    return `
      <tr data-no="${w.stockNo}">
        <td class="code"><a class="code-link" onclick="showChartFor('${w.stockNo}')">${w.stockNo}</a></td>
        <td class="name">${name}</td>
        <td class="num" style="color:${col};font-weight:bold">${priceStr}</td>
        <td class="num" style="color:${col}">${diffStr}</td>
        <td class="num" style="color:${col}">${pctStr}</td>
        <td class="num">${o}</td>
        <td class="num">${h}</td>
        <td class="num">${l}</td>
        <td class="num">${vol}</td>
        <td>
          <div class="row-actions">
            <button class="btn-chart" onclick="showChartFor('${w.stockNo}')">看圖</button>
            <button class="btn-delete" onclick="removeWatch('${w.stockNo}')">移除</button>
          </div>
        </td>
      </tr>`;
  }).join('');
}

async function addWatch(stockNo, stockName) {
  if (watchState.list.some(w => w.stockNo === stockNo)) return;   // 不重複
  watchState.list.push({ stockNo, stockName });                   // 先本機更新（即時）
  saveWatchlistLocal();
  renderWatchTable();
  refreshWatchQuotes();
  try { await gasPost({ action: 'addWatch', watch: { stockNo, stockName } }); } catch {}  // 同步到後端
}

async function removeWatch(stockNo) {
  watchState.list = watchState.list.filter(w => w.stockNo !== stockNo);
  delete watchState.quotes[stockNo];
  saveWatchlistLocal();
  renderWatchTable();
  try { await gasPost({ action: 'deleteWatch', stockNo }); } catch {}
}

document.getElementById('refreshWatchBtn').addEventListener('click', refreshMarket);

// 觀察名單搜尋（加入股票）
(function initWatchSearch() {
  const searchInput = document.getElementById('watchSearchInput');
  const dropdown    = document.getElementById('watchSearchDropdown');
  let debounceTimer = null, activeIndex = -1, currentResults = [];

  function showDropdown(items, msg) {
    dropdown.innerHTML = ''; activeIndex = -1;
    if (msg) { dropdown.innerHTML = `<div class="search-msg">${msg}</div>`; dropdown.classList.add('open'); return; }
    items.forEach(item => {
      const el = document.createElement('div');
      el.className = 'search-item';
      el.innerHTML = `<span class="search-code">${item.code}</span><span class="search-name">${item.name}</span>`;
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(item); });
      dropdown.appendChild(el);
    });
    currentResults = items; dropdown.classList.add('open');
  }
  function hideDropdown() { dropdown.classList.remove('open'); activeIndex = -1; }
  function selectItem(item) { searchInput.value = ''; hideDropdown(); addWatch(item.code, item.name); }
  function highlightItem(idx) { dropdown.querySelectorAll('.search-item').forEach((el, i) => el.classList.toggle('active', i === idx)); }

  searchInput.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    const q = searchInput.value.trim();
    if (!q) { hideDropdown(); return; }
    debounceTimer = setTimeout(async () => {
      showDropdown([], '搜尋中...');
      const results = await searchStock(q);
      results.length === 0 ? showDropdown([], '查無結果') : showDropdown(results);
    }, 350);
  });
  searchInput.addEventListener('keydown', e => {
    const items = dropdown.querySelectorAll('.search-item');
    if (e.key === 'ArrowDown') { e.preventDefault(); activeIndex = Math.min(activeIndex + 1, items.length - 1); highlightItem(activeIndex); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); activeIndex = Math.max(activeIndex - 1, 0); highlightItem(activeIndex); }
    else if (e.key === 'Enter' && activeIndex >= 0 && currentResults[activeIndex]) { e.preventDefault(); selectItem(currentResults[activeIndex]); }
    else if (e.key === 'Escape') hideDropdown();
  });
  searchInput.addEventListener('blur', () => setTimeout(hideDropdown, 150));
  document.addEventListener('click', e => { if (!e.target.closest('#watchSearchContainer')) hideDropdown(); });
})();

// 自動取得 Chat ID
document.getElementById('getChatIdBtn').addEventListener('click', async () => {
  const token  = document.getElementById('telegramTokenInput').value.trim();
  const status = document.getElementById('tokenStatus');
  if (!token) { status.style.color = '#ff8888'; status.textContent = '請先輸入 Bot Token'; return; }
  status.style.color = '#aaaacc';
  status.textContent = '取得中，請確認已傳訊息給 Bot...';
  try {
    const resp = await fetch(`https://api.telegram.org/bot${token}/getUpdates`);
    const data = await resp.json();
    const updates = data.result || [];
    if (!updates.length) {
      status.style.color = '#ffaa33';
      status.textContent = '找不到訊息，請先在 Telegram 對你的 Bot 傳送任意訊息後再試';
      return;
    }
    const chatId = updates[updates.length - 1].message?.chat?.id;
    if (!chatId) {
      status.style.color = '#ff8888';
      status.textContent = '無法解析 Chat ID，請手動輸入';
      return;
    }
    document.getElementById('telegramChatIdInput').value = chatId;
    status.style.color = '#66ff99';
    status.textContent = `✓ Chat ID：${chatId}（請記得按儲存）`;
  } catch {
    status.style.color = '#ff8888';
    status.textContent = '連線失敗，請確認 Token 是否正確';
  }
});

// 儲存 Telegram 設定
document.getElementById('saveTokenBtn').addEventListener('click', () => {
  const token  = document.getElementById('telegramTokenInput').value.trim();
  const chatId = document.getElementById('telegramChatIdInput').value.trim();
  const status = document.getElementById('tokenStatus');
  if (!token || !chatId) { status.style.color = '#ff8888'; status.textContent = '請填入 Token 與 Chat ID'; return; }
  storageSet({ telegramToken: token, telegramChatId: chatId });
  status.style.color = '#66ff99';
  status.textContent = '✓ 已儲存';
  setTimeout(() => status.textContent = '', 2000);
});

// 測試 Telegram 通知
document.getElementById('testTelegramBtn').addEventListener('click', async () => {
  const token  = document.getElementById('telegramTokenInput').value.trim();
  const chatId = document.getElementById('telegramChatIdInput').value.trim();
  const status = document.getElementById('tokenStatus');
  if (!token || !chatId) { status.style.color = '#ff8888'; status.textContent = '請先填入 Token 與 Chat ID'; return; }
  status.style.color = '#aaaacc';
  status.textContent = '傳送中...';
  try {
    const resp = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ chat_id: chatId, text: '【台股分析工具】Telegram 通知測試成功！' }),
    });
    const data = await resp.json();
    if (data.ok) {
      status.style.color = '#66ff99';
      status.textContent = '✓ 通知已送出，請確認 Telegram';
    } else {
      status.style.color = '#ff8888';
      status.textContent = `✗ 失敗：${data.description || '請確認 Token 與 Chat ID'}`;
    }
  } catch {
    status.style.color = '#ff8888';
    status.textContent = '✗ 連線失敗';
  }
});

// ── Alert Search ──────────────────────────────────────────────

(function initAlertSearch() {
  const searchInput  = document.getElementById('alertSearchInput');
  const dropdown     = document.getElementById('alertSearchDropdown');
  let debounceTimer  = null;
  let activeIndex    = -1;
  let currentResults = [];

  function showDropdown(items, msg) {
    dropdown.innerHTML = '';
    activeIndex = -1;
    if (msg) {
      dropdown.innerHTML = `<div class="search-msg">${msg}</div>`;
      dropdown.classList.add('open');
      return;
    }
    items.forEach(item => {
      const el = document.createElement('div');
      el.className = 'search-item';
      el.innerHTML = `<span class="search-code">${item.code}</span><span class="search-name">${item.name}</span>`;
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(item); });
      dropdown.appendChild(el);
    });
    currentResults = items;
    dropdown.classList.add('open');
  }

  function hideDropdown() { dropdown.classList.remove('open'); activeIndex = -1; }

  function selectItem(item) {
    searchInput.value = '';
    hideDropdown();
    addMonitorStock(item.code, item.name);   // 加入監控表並聚焦到買點欄
  }

  function highlightItem(idx) {
    dropdown.querySelectorAll('.search-item').forEach((el, i) => el.classList.toggle('active', i === idx));
  }

  searchInput.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    const q = searchInput.value.trim();
    if (!q) { hideDropdown(); return; }
    debounceTimer = setTimeout(async () => {
      showDropdown([], '搜尋中...');
      const results = await searchStock(q);
      results.length === 0 ? showDropdown([], '查無結果') : showDropdown(results);
    }, 350);
  });

  searchInput.addEventListener('keydown', e => {
    const items = dropdown.querySelectorAll('.search-item');
    if (e.key === 'ArrowDown') { e.preventDefault(); activeIndex = Math.min(activeIndex + 1, items.length - 1); highlightItem(activeIndex); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); activeIndex = Math.max(activeIndex - 1, 0); highlightItem(activeIndex); }
    else if (e.key === 'Enter' && activeIndex >= 0 && currentResults[activeIndex]) { e.preventDefault(); selectItem(currentResults[activeIndex]); }
    else if (e.key === 'Escape') hideDropdown();
  });

  searchInput.addEventListener('blur', () => setTimeout(hideDropdown, 150));
  document.addEventListener('click', e => { if (!e.target.closest('#alertSearchContainer')) hideDropdown(); });
})();

// ── Stock Search ─────────────────────────────────────────────

// 台股代號↔中文簡稱對照表（來源：TWSE + TPEX 開放資料），首次載入後在 localStorage 快取 24 小時
let stockListCache = null;
const STOCK_LIST_TTL = 24 * 60 * 60 * 1000;

async function loadStockList() {
  if (stockListCache) return stockListCache;

  const cached   = localStorage.getItem('stockList');
  const cachedAt = parseInt(localStorage.getItem('stockListAt') || '0', 10);
  if (cached && Date.now() - cachedAt < STOCK_LIST_TTL) {
    try {
      stockListCache = JSON.parse(cached);
      return stockListCache;
    } catch { /* cache 壞了，重抓 */ }
  }

  try {
    const resp = await fetch('lib/taiwan-stocks.json');
    stockListCache = await resp.json();
    localStorage.setItem('stockList', JSON.stringify(stockListCache));
    localStorage.setItem('stockListAt', String(Date.now()));
    return stockListCache;
  } catch {
    return [];
  }
}

async function searchStock(query) {
  const q = query.trim();
  if (!q) return [];
  const list = await loadStockList();

  const exact = [], codeStarts = [], nameContains = [];
  for (const s of list) {
    if (!s.code || !s.name) continue;
    if (s.code === q) exact.push(s);
    else if (s.code.startsWith(q)) codeStarts.push(s);
    else if (s.name.includes(q)) nameContains.push(s);
  }
  return [...exact, ...codeStarts, ...nameContains]
    .slice(0, 20)
    .map(s => ({ code: s.code, name: s.name }));
}

(function initSearch() {
  const searchInput  = document.getElementById('searchInput');
  const dropdown     = document.getElementById('searchDropdown');
  const stockInput   = document.getElementById('stockInput');
  let debounceTimer  = null;
  let activeIndex    = -1;
  let currentResults = [];

  function showDropdown(items, msg) {
    dropdown.innerHTML = '';
    activeIndex = -1;
    if (msg) {
      dropdown.innerHTML = `<div class="search-msg">${msg}</div>`;
      dropdown.classList.add('open');
      return;
    }
    items.forEach(item => {
      const el = document.createElement('div');
      el.className = 'search-item';
      el.innerHTML = `<span class="search-code">${item.code}</span><span class="search-name">${item.name}</span>`;
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(item); });
      dropdown.appendChild(el);
    });
    currentResults = items;
    dropdown.classList.add('open');
  }

  function hideDropdown() { dropdown.classList.remove('open'); activeIndex = -1; }

  function selectItem(item) {
    stockInput.value = item.code;
    searchInput.value = `${item.code} ${item.name}`;
    hideDropdown();
    onQuery();
  }

  function highlightItem(idx) {
    dropdown.querySelectorAll('.search-item').forEach((el, i) => el.classList.toggle('active', i === idx));
  }

  searchInput.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    const q = searchInput.value.trim();
    if (!q) { hideDropdown(); return; }
    debounceTimer = setTimeout(async () => {
      showDropdown([], '搜尋中...');
      const results = await searchStock(q);
      results.length === 0 ? showDropdown([], '查無結果') : showDropdown(results);
    }, 350);
  });

  searchInput.addEventListener('keydown', e => {
    const items = dropdown.querySelectorAll('.search-item');
    if (e.key === 'ArrowDown') { e.preventDefault(); activeIndex = Math.min(activeIndex + 1, items.length - 1); highlightItem(activeIndex); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); activeIndex = Math.max(activeIndex - 1, 0); highlightItem(activeIndex); }
    else if (e.key === 'Enter' && activeIndex >= 0 && currentResults[activeIndex]) { e.preventDefault(); selectItem(currentResults[activeIndex]); }
    else if (e.key === 'Escape') hideDropdown();
  });

  searchInput.addEventListener('blur', () => setTimeout(hideDropdown, 150));
  document.addEventListener('click', e => { if (!e.target.closest('#searchContainer')) hideDropdown(); });
})();

// ── Holdings (存股) ──────────────────────────────────────────

const holdingsState = {
  holdings: [],
  dividends: {},   // stockNo -> [{date: epochSec, amount}]
  quotes: {},      // stockNo -> {price, name}
  sortKey: 'yield',
  sortDir: 'desc',
  activeTags: new Set(),
  period: 'ttm',
};

async function fetchDividends(stockNo) {
  if (holdingsState.dividends[stockNo]) return holdingsState.dividends[stockNo];
  for (const suffix of ['TW', 'TWO']) {
    try {
      const url = `https://query1.finance.yahoo.com/v8/finance/chart/${stockNo}.${suffix}?interval=1d&range=2y&events=div`;
      const resp = await fetch(p(url));
      const data = await resp.json();
      const result = data?.chart?.result?.[0];
      if (!result) continue;
      const divsObj = result.events?.dividends || {};
      const divs = Object.values(divsObj).map(d => ({ date: d.date, amount: d.amount }));
      const meta = result.meta || {};
      holdingsState.dividends[stockNo] = divs;
      holdingsState.quotes[stockNo] = {
        price: meta.regularMarketPrice,
        name: meta.shortName || meta.longName || stockNo,
      };
      return divs;
    } catch { /* try next */ }
  }
  holdingsState.dividends[stockNo] = [];
  return [];
}

function sumDividends(divs, period) {
  const now = new Date();
  const year = now.getFullYear();
  let from, to;
  if (period === 'ttm') {
    to = now.getTime() / 1000;
    from = to - 365 * 86400;
  } else if (period === 'thisYear') {
    from = Date.UTC(year, 0, 1) / 1000;
    to   = Date.UTC(year + 1, 0, 1) / 1000;
  } else { // lastYear
    from = Date.UTC(year - 1, 0, 1) / 1000;
    to   = Date.UTC(year, 0, 1) / 1000;
  }
  return divs.filter(d => d.date >= from && d.date < to).reduce((a, d) => a + d.amount, 0);
}

function computeRow(h, period) {
  const q = holdingsState.quotes[h.stockNo] || {};
  const divs = holdingsState.dividends[h.stockNo] || [];
  const price = q.price;
  const divPerShare = sumDividends(divs, period);
  const annualDiv = divPerShare * h.shares;
  const marketValue = price != null ? price * h.shares : null;
  const yieldPct = (price && divPerShare) ? (divPerShare / price * 100) : null;
  const cost = (h.costPrice != null && h.costPrice > 0) ? h.costPrice : null;
  const costValue = cost != null ? cost * h.shares : null;
  const personalYield = (cost != null && divPerShare) ? (divPerShare / cost * 100) : null;
  const profit = (marketValue != null && costValue != null) ? (marketValue - costValue) : null;
  const profitPct = (cost != null && price != null) ? ((price - cost) / cost * 100) : null;
  return { price, divPerShare, annualDiv, marketValue, yield: yieldPct, costValue, personalYield, profit, profitPct };
}

function fmtMoney(v) {
  if (v == null || isNaN(v)) return '--';
  return Math.round(v).toLocaleString();
}

function fmtNum(v, digits = 2) {
  if (v == null || isNaN(v)) return '--';
  return v.toFixed(digits);
}

function yieldCls(y) {
  if (y == null) return 'loading';
  if (y >= 5) return 'yield-good';
  if (y >= 3) return 'yield-mid';
  return 'yield-low';
}

async function loadHoldings() {
  const tbody = document.getElementById('holdingsTbody');
  tbody.innerHTML = '<tr><td colspan="13" class="no-holdings">載入中...</td></tr>';
  try {
    const data = await gasGet();
    holdingsState.holdings = data.holdings || [];
  } catch {
    tbody.innerHTML = '<tr><td colspan="13" class="no-holdings" style="color:#ff8888">載入失敗，請確認 GAS 部署</td></tr>';
    return;
  }

  if (holdingsState.holdings.length === 0) {
    renderHoldings();
    renderSummary();
    renderTagFilter();
    renderPies();
    return;
  }

  // 先用 loading 狀態渲染一次
  renderHoldings();
  renderTagFilter();
  renderPies();

  // 並行抓所有股利資料
  await Promise.all(holdingsState.holdings.map(h => fetchDividends(h.stockNo)));

  renderHoldings();
  renderSummary();
  renderPies();
}

function renderHoldings() {
  const tbody = document.getElementById('holdingsTbody');
  const badge = document.getElementById('holdingCountBadge');
  const list = holdingsState.holdings;
  badge.textContent = list.length ? `${list.length} 檔` : '';
  renderQuickHoldings();   // 同步更新圖表頁的持股快選

  if (!list.length) {
    tbody.innerHTML = '<tr><td colspan="13" class="no-holdings">尚無持股 — 從上方「新增持股」開始</td></tr>';
    return;
  }

  // 套用標籤篩選
  const active = holdingsState.activeTags;
  const filtered = active.size === 0
    ? list
    : list.filter(h => (h.tags || []).some(t => active.has(t)));

  if (!filtered.length) {
    tbody.innerHTML = '<tr><td colspan="13" class="no-holdings">沒有符合篩選的持股</td></tr>';
    return;
  }

  // 計算每列再排序
  const rows = filtered.map(h => ({ h, ...computeRow(h, holdingsState.period) }));

  const key = holdingsState.sortKey;
  const dir = holdingsState.sortDir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    let av, bv;
    if (key === 'stockNo')        { av = a.h.stockNo; bv = b.h.stockNo; }
    else if (key === 'stockName') { av = a.h.stockName || ''; bv = b.h.stockName || ''; }
    else if (key === 'account')   { av = a.h.account || ''; bv = b.h.account || ''; }
    else if (key === 'tags')      { av = (a.h.tags || []).join(','); bv = (b.h.tags || []).join(','); }
    else if (key === 'shares')    { av = a.h.shares; bv = b.h.shares; }
    else if (key === 'costPrice') { av = a.h.costPrice; bv = b.h.costPrice; }
    else { av = a[key]; bv = b[key]; }
    if (av == null && bv == null) return 0;
    if (av == null) return 1;
    if (bv == null) return -1;
    if (typeof av === 'string') return av.localeCompare(bv) * dir;
    return (av - bv) * dir;
  });

  // 排序指示器
  document.querySelectorAll('#holdingsTable th').forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.dataset.sort === key) {
      th.classList.add(holdingsState.sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
    }
  });

  tbody.innerHTML = rows.map(({ h, price, divPerShare, annualDiv, marketValue, yield: y, personalYield, profit, profitPct }) => {
    const tagsHtml = (h.tags || []).map(t => `<span class="row-tag">${escapeHtml(t)}</span>`).join('');
    const accountHtml = h.account ? `<span class="row-account">${escapeHtml(h.account)}</span>` : '<span style="color:#555">—</span>';
    const loaded = holdingsState.dividends[h.stockNo] !== undefined;
    const costStr = (h.costPrice != null && h.costPrice > 0) ? fmtNum(h.costPrice) : '<span style="color:#555">—</span>';
    const profitCell = profit == null
      ? '<span style="color:#555">—</span>'
      : `${profit >= 0 ? '+' : ''}${fmtMoney(profit)}` +
        (profitPct != null ? `<br><span class="profit-pct">${profitPct >= 0 ? '+' : ''}${profitPct.toFixed(2)}%</span>` : '');
    const profitCls = profit == null ? '' : (profit >= 0 ? 'profit-up' : 'profit-down');
    return `
      <tr>
        <td class="code"><a class="code-link" onclick="showChartFor('${h.stockNo}')" title="查看 ${h.stockNo} 圖表">${h.stockNo}</a></td>
        <td class="name">${escapeHtml(h.stockName || '')}</td>
        <td>${accountHtml}</td>
        <td>${tagsHtml || '<span style="color:#555">—</span>'}</td>
        <td class="num">${h.shares.toLocaleString()}</td>
        <td class="num">${costStr}</td>
        <td class="num ${loaded ? '' : 'loading'}">${fmtNum(price)}</td>
        <td class="num">${fmtMoney(marketValue)}</td>
        <td class="num ${profitCls}">${profitCell}</td>
        <td class="num">${fmtMoney(annualDiv)}</td>
        <td class="num ${yieldCls(y)}">${y != null ? y.toFixed(2) + '%' : '--'}</td>
        <td class="num ${yieldCls(personalYield)}">${personalYield != null ? personalYield.toFixed(2) + '%' : '<span style="color:#555">—</span>'}</td>
        <td>
          <div class="row-actions">
            <button class="btn-chart" onclick="showChartFor('${h.stockNo}')">看圖</button>
            <button onclick="editHoldingShares('${h.id}')">股數變更</button>
            <button onclick="editHoldingCost('${h.id}')">成本</button>
            <button onclick="editHoldingAccount('${h.id}')">帳戶</button>
            <button onclick="editHoldingTags('${h.id}')">標籤</button>
            <button class="btn-delete" onclick="deleteHolding('${h.id}')">刪除</button>
          </div>
        </td>
      </tr>`;
  }).join('');

  // tfoot 總計（依目前篩選後的結果）
  let totalShares = 0, totalMV = 0, totalDiv = 0, totalCost = 0, divForCostYield = 0;
  let totalProfit = 0, costForProfit = 0, hasProfit = false;
  for (const r of rows) {
    totalShares += r.h.shares;
    if (r.marketValue != null) totalMV += r.marketValue;
    if (r.annualDiv != null)   totalDiv += r.annualDiv;
    if (r.costValue != null) {
      totalCost += r.costValue;
      divForCostYield += r.annualDiv || 0;
    }
    if (r.profit != null) {
      totalProfit += r.profit;
      costForProfit += r.costValue;
      hasProfit = true;
    }
  }
  const wYield  = totalMV   > 0 ? (totalDiv / totalMV * 100) : null;
  const wPYield = totalCost > 0 ? (divForCostYield / totalCost * 100) : null;
  const totalProfitPct = costForProfit > 0 ? (totalProfit / costForProfit * 100) : null;
  const totalProfitCls = !hasProfit ? '' : (totalProfit >= 0 ? 'profit-up' : 'profit-down');
  const totalProfitCell = !hasProfit
    ? '—'
    : `${totalProfit >= 0 ? '+' : ''}${fmtMoney(totalProfit)}` +
      (totalProfitPct != null ? `<br><span class="profit-pct">${totalProfitPct >= 0 ? '+' : ''}${totalProfitPct.toFixed(2)}%</span>` : '');
  document.getElementById('holdingsTfoot').innerHTML = `
    <tr class="total-row">
      <td colspan="4" class="label">總計（${rows.length} 檔）</td>
      <td class="num">${totalShares.toLocaleString()}</td>
      <td class="num">${totalCost > 0 ? fmtMoney(totalCost) : '—'}</td>
      <td class="num">—</td>
      <td class="num">${fmtMoney(totalMV)}</td>
      <td class="num ${totalProfitCls}">${totalProfitCell}</td>
      <td class="num">${fmtMoney(totalDiv)}</td>
      <td class="num">${wYield != null ? wYield.toFixed(2) + '%' : '--'}</td>
      <td class="num">${wPYield != null ? wPYield.toFixed(2) + '%' : '—'}</td>
      <td>—</td>
    </tr>`;
}

function renderSummary() {
  const list = holdingsState.holdings;
  let totalMV = 0, totalDiv = 0, validMV = 0;
  let totalProfit = 0, costForProfit = 0, hasProfit = false;
  for (const h of list) {
    const { marketValue, annualDiv, profit, costValue } = computeRow(h, holdingsState.period);
    if (marketValue != null) { totalMV += marketValue; validMV += marketValue; }
    if (annualDiv != null) totalDiv += annualDiv;
    if (profit != null) {
      totalProfit += profit;
      costForProfit += costValue;
      hasProfit = true;
    }
  }
  const avgYield = validMV > 0 ? (totalDiv / validMV * 100) : null;
  const totalProfitPct = costForProfit > 0 ? (totalProfit / costForProfit * 100) : null;

  document.getElementById('sumMarketValue').textContent = list.length ? fmtMoney(totalMV) : '--';
  document.getElementById('sumAnnualDiv').textContent   = list.length ? fmtMoney(totalDiv) : '--';
  document.getElementById('sumAvgYield').textContent    = avgYield != null ? avgYield.toFixed(2) + '%' : '--';
  document.getElementById('sumCount').textContent       = list.length || '--';

  const profitEl = document.getElementById('sumProfit');
  if (!hasProfit) {
    profitEl.textContent = '--';
    profitEl.classList.remove('profit-up', 'profit-down');
  } else {
    const pctStr = totalProfitPct != null ? ` (${totalProfit >= 0 ? '+' : ''}${totalProfitPct.toFixed(2)}%)` : '';
    profitEl.textContent = `${totalProfit >= 0 ? '+' : ''}${fmtMoney(totalProfit)}${pctStr}`;
    profitEl.classList.toggle('profit-up', totalProfit >= 0);
    profitEl.classList.toggle('profit-down', totalProfit < 0);
  }
}

function renderTagFilter() {
  const container = document.getElementById('tagFilterChips');
  const allTags = new Set();
  holdingsState.holdings.forEach(h => (h.tags || []).forEach(t => allTags.add(t)));

  if (allTags.size === 0) {
    container.innerHTML = '<span style="color:#666;font-size:12px">（尚無標籤）</span>';
    return;
  }

  container.innerHTML = [...allTags].sort().map(t => {
    const active = holdingsState.activeTags.has(t) ? 'active' : '';
    return `<span class="tag-chip ${active}" data-tag="${escapeHtml(t)}">${escapeHtml(t)}</span>`;
  }).join('');

  container.querySelectorAll('.tag-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const t = chip.dataset.tag;
      if (holdingsState.activeTags.has(t)) holdingsState.activeTags.delete(t);
      else holdingsState.activeTags.add(t);
      renderTagFilter();
      renderHoldings();
    });
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

// ── Pie Chart ────────────────────────────────────────────

const PIE_COLORS = [
  '#5a7fff', '#66ff99', '#ffcc55', '#ff8888', '#cc88ff',
  '#33ddcc', '#ffaa33', '#ff66bb', '#88aaff', '#aabb77',
];

function renderPie(elId, slices) {
  const el = document.getElementById(elId);
  const total = slices.reduce((a, s) => a + s.value, 0);
  if (total <= 0) {
    el.innerHTML = '<div class="pie-empty">尚無資料</div>';
    return;
  }

  const cx = 100, cy = 100, r = 90;
  let angle = -Math.PI / 2; // 12 點方向

  const paths = slices.map((s, i) => {
    const sweep = (s.value / total) * 2 * Math.PI;
    const x1 = cx + r * Math.cos(angle);
    const y1 = cy + r * Math.sin(angle);
    angle += sweep;
    const x2 = cx + r * Math.cos(angle);
    const y2 = cy + r * Math.sin(angle);
    const large = sweep > Math.PI ? 1 : 0;
    // 整圓情況（只有一塊）特別處理避免 path 縮成 0
    const d = slices.length === 1
      ? `M${cx - r},${cy} A${r},${r} 0 1 1 ${cx + r},${cy} A${r},${r} 0 1 1 ${cx - r},${cy} Z`
      : `M${cx},${cy} L${x1},${y1} A${r},${r} 0 ${large} 1 ${x2},${y2} Z`;
    const pct = (s.value / total * 100);
    return `<path d="${d}" fill="${s.color}" stroke="#1e1e2e" stroke-width="1.5"><title>${escapeHtml(s.label)}：${Math.round(s.value).toLocaleString()}（${pct.toFixed(1)}%）</title></path>`;
  }).join('');

  const legend = slices.map(s => {
    const pct = (s.value / total * 100);
    return `<div class="pie-legend-item">
      <span class="pie-legend-color" style="background:${s.color}"></span>
      <span class="pie-legend-label" title="${escapeHtml(s.label)}">${escapeHtml(s.label)}</span>
      <span class="pie-legend-value">${pct.toFixed(1)}%</span>
    </div>`;
  }).join('');

  el.innerHTML = `<svg viewBox="0 0 200 200" class="pie-svg">${paths}</svg><div class="pie-legend">${legend}</div>`;
}

function buildPieSlices(valueKey) {
  // 同一檔股票（同 stockNo）聚合成一塊，名稱用第一筆的
  const byStock = new Map();
  for (const h of holdingsState.holdings) {
    const r = computeRow(h, holdingsState.period);
    const v = r[valueKey];
    if (v == null) continue;
    const key = h.stockNo;
    if (!byStock.has(key)) {
      byStock.set(key, { stockNo: h.stockNo, stockName: h.stockName || '', value: 0 });
    }
    byStock.get(key).value += v;
  }

  const sorted = [...byStock.values()].filter(s => s.value > 0).sort((a, b) => b.value - a.value);

  // 超過 9 檔合併「其他」
  const CAP = 9;
  let display = sorted;
  if (sorted.length > CAP) {
    const top = sorted.slice(0, CAP - 1);
    const rest = sorted.slice(CAP - 1);
    const restSum = rest.reduce((a, s) => a + s.value, 0);
    display = [...top, { stockNo: '', stockName: `其他 ${rest.length} 檔`, value: restSum }];
  }

  return display.map((s, i) => ({
    label: s.stockNo ? `${s.stockNo} ${s.stockName}` : s.stockName,
    value: s.value,
    color: PIE_COLORS[i % PIE_COLORS.length],
  }));
}

function renderPies() {
  renderPie('pieMarketValue', buildPieSlices('marketValue'));
  renderPie('pieAnnualDiv',   buildPieSlices('annualDiv'));
}

async function deleteHolding(id) {
  const h = holdingsState.holdings.find(x => x.id === id);
  if (!h) return;
  if (!confirm(`確定刪除 ${h.stockNo} ${h.stockName || ''}？`)) return;
  await gasPost({ action: 'deleteHolding', id });
  loadHoldings();
}

async function editHoldingShares(id) {
  const h = holdingsState.holdings.find(x => x.id === id);
  if (!h) return;
  const input = prompt(`${h.stockNo} ${h.stockName || ''} 的新股數：`, h.shares);
  if (input == null) return;
  const shares = parseInt(input, 10);
  if (isNaN(shares) || shares <= 0) { alert('股數需為正整數'); return; }
  await gasPost({ action: 'updateHolding', id, shares });
  loadHoldings();
}

async function editHoldingTags(id) {
  const h = holdingsState.holdings.find(x => x.id === id);
  if (!h) return;
  const current = (h.tags || []).join(',');
  const input = prompt(`${h.stockNo} ${h.stockName || ''} 的新標籤（逗號分隔，留空清除）：`, current);
  if (input == null) return;
  const tags = input.split(/[,，]/).map(t => t.trim()).filter(Boolean);
  await gasPost({ action: 'updateHolding', id, tags });
  loadHoldings();
}

async function editHoldingCost(id) {
  const h = holdingsState.holdings.find(x => x.id === id);
  if (!h) return;
  const current = h.costPrice != null ? String(h.costPrice) : '';
  const input = prompt(`${h.stockNo} ${h.stockName || ''} 的成本價 / 股（留空清除）：`, current);
  if (input == null) return;
  const trimmed = input.trim();
  let costPrice = null;
  if (trimmed !== '') {
    costPrice = parseFloat(trimmed);
    if (isNaN(costPrice) || costPrice <= 0) { alert('成本價需為正數'); return; }
  }
  await gasPost({ action: 'updateHolding', id, costPrice });
  loadHoldings();
}

async function editHoldingAccount(id) {
  const h = holdingsState.holdings.find(x => x.id === id);
  if (!h) return;
  const current = h.account || '';
  const input = prompt(`${h.stockNo} ${h.stockName || ''} 的帳戶（留空清除）：`, current);
  if (input == null) return;
  await gasPost({ action: 'updateHolding', id, account: input.trim() });
  loadHoldings();
}

// 排序 header
document.querySelectorAll('#holdingsTable th[data-sort]').forEach(th => {
  th.addEventListener('click', () => {
    const key = th.dataset.sort;
    if (holdingsState.sortKey === key) {
      holdingsState.sortDir = holdingsState.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      holdingsState.sortKey = key;
      holdingsState.sortDir = ['stockNo', 'stockName', 'tags'].includes(key) ? 'asc' : 'desc';
    }
    renderHoldings();
  });
});

// 期間切換
document.getElementById('dividendPeriod').addEventListener('change', e => {
  holdingsState.period = e.target.value;
  renderHoldings();
  renderSummary();
  renderPies();
});

// 重新整理
document.getElementById('refreshHoldingsBtn').addEventListener('click', () => {
  holdingsState.dividends = {};
  holdingsState.quotes = {};
  loadHoldings();
});

// 新增持股 — 搜尋下拉
(function initHoldingSearch() {
  const searchInput = document.getElementById('holdingSearchInput');
  const dropdown    = document.getElementById('holdingSearchDropdown');
  const stockNoEl   = document.getElementById('holdingStockNo');
  const stockNameEl = document.getElementById('holdingStockName');
  let debounceTimer = null;
  let activeIndex = -1;
  let currentResults = [];

  function showDropdown(items, msg) {
    dropdown.innerHTML = '';
    activeIndex = -1;
    if (msg) {
      dropdown.innerHTML = `<div class="search-msg">${msg}</div>`;
      dropdown.classList.add('open');
      return;
    }
    items.forEach(item => {
      const el = document.createElement('div');
      el.className = 'search-item';
      el.innerHTML = `<span class="search-code">${item.code}</span><span class="search-name">${item.name}</span>`;
      el.addEventListener('mousedown', e => { e.preventDefault(); selectItem(item); });
      dropdown.appendChild(el);
    });
    currentResults = items;
    dropdown.classList.add('open');
  }

  function hideDropdown() { dropdown.classList.remove('open'); activeIndex = -1; }

  function selectItem(item) {
    stockNoEl.value = item.code;
    stockNameEl.value = item.name;
    searchInput.value = `${item.code} ${item.name}`;
    hideDropdown();
    document.getElementById('holdingShares').focus();
  }

  function highlightItem(idx) {
    dropdown.querySelectorAll('.search-item').forEach((el, i) => el.classList.toggle('active', i === idx));
  }

  searchInput.addEventListener('input', () => {
    clearTimeout(debounceTimer);
    const q = searchInput.value.trim();
    stockNoEl.value = ''; stockNameEl.value = '';
    if (!q) { hideDropdown(); return; }
    debounceTimer = setTimeout(async () => {
      showDropdown([], '搜尋中...');
      const results = await searchStock(q);
      results.length === 0 ? showDropdown([], '查無結果') : showDropdown(results);
    }, 350);
  });

  searchInput.addEventListener('keydown', e => {
    const items = dropdown.querySelectorAll('.search-item');
    if (e.key === 'ArrowDown') { e.preventDefault(); activeIndex = Math.min(activeIndex + 1, items.length - 1); highlightItem(activeIndex); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); activeIndex = Math.max(activeIndex - 1, 0); highlightItem(activeIndex); }
    else if (e.key === 'Enter' && activeIndex >= 0 && currentResults[activeIndex]) { e.preventDefault(); selectItem(currentResults[activeIndex]); }
    else if (e.key === 'Escape') hideDropdown();
  });

  searchInput.addEventListener('blur', () => setTimeout(hideDropdown, 150));
  document.addEventListener('click', e => { if (!e.target.closest('#holdingSearchContainer')) hideDropdown(); });
})();

// 新增持股按鈕
document.getElementById('addHoldingBtn').addEventListener('click', async () => {
  const stockNo   = document.getElementById('holdingStockNo').value.trim();
  const stockName = document.getElementById('holdingStockName').value.trim();
  const shares    = parseInt(document.getElementById('holdingShares').value, 10);
  const costRaw   = document.getElementById('holdingCostPrice').value.trim();
  const account   = document.getElementById('holdingAccount').value.trim();
  const tagsRaw   = document.getElementById('holdingTags').value.trim();
  const status    = document.getElementById('addHoldingStatus');

  if (!stockNo || isNaN(shares) || shares <= 0) {
    status.style.color = '#ff8888';
    status.textContent = '請選股票並填入正整數股數';
    return;
  }

  let costPrice = null;
  if (costRaw !== '') {
    costPrice = parseFloat(costRaw);
    if (isNaN(costPrice) || costPrice <= 0) {
      status.style.color = '#ff8888';
      status.textContent = '成本價需為正數（或留空）';
      return;
    }
  }

  const tags = tagsRaw.split(/[,，]/).map(t => t.trim()).filter(Boolean);

  status.style.color = '#aaaacc';
  status.textContent = '新增中...';
  try {
    const result = await gasPost({
      action: 'addHolding',
      holding: { stockNo, stockName, shares, costPrice, account, tags },
    });
    if (!result.ok) throw new Error(result.error || '伺服器錯誤');

    document.getElementById('holdingSearchInput').value = '';
    document.getElementById('holdingStockNo').value     = '';
    document.getElementById('holdingStockName').value   = '';
    document.getElementById('holdingShares').value      = '';
    document.getElementById('holdingCostPrice').value   = '';
    document.getElementById('holdingAccount').value     = '';
    document.getElementById('holdingTags').value        = '';
    status.style.color = '#66ff99';
    status.textContent = `✓ 已新增 ${stockNo} ${stockName || ''} ${shares} 股`;
    setTimeout(() => status.textContent = '', 3000);
    loadHoldings();
  } catch (err) {
    status.style.color = '#ff8888';
    status.textContent = `✗ 新增失敗：${err.message}`;
  }
});

// ── 啟動：載入圖表頁的「我的持股」快選 ─────────────────
loadQuickHoldings();

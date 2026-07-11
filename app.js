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
  });
});

// ── Alerts Management ─────────────────────────────────────────

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

async function loadAlerts() {
  const { telegramToken, telegramChatId } = storageGet(['telegramToken', 'telegramChatId']);
  document.getElementById('telegramTokenInput').value  = telegramToken || '';
  document.getElementById('telegramChatIdInput').value = telegramChatId || '';

  const list = document.getElementById('alertList');
  list.innerHTML = '<div class="no-alerts">載入中...</div>';
  try {
    const data = await gasGet();
    renderAlertList(data.alerts || []);
  } catch {
    list.innerHTML = '<div class="no-alerts" style="color:#ff8888">載入失敗，請確認 GAS 部署是否正確</div>';
  }
}

function renderAlertList(alerts) {
  const list  = document.getElementById('alertList');
  const badge = document.getElementById('alertCountBadge');
  const active = alerts.filter(a => a.enabled && !a.triggered).length;
  badge.textContent = active ? `${active} 監控中` : '';

  if (!alerts.length) {
    list.innerHTML = '<div class="no-alerts">尚無警報</div>';
    return;
  }

  list.innerHTML = alerts.map((a, i) => {
    const arrow = a.condition === 'lte' ? '↓' : '↑';
    const cls   = a.triggered ? 'triggered' : (!a.enabled ? 'paused' : '');
    const tag   = a.triggered
      ? `<span class="alert-tag">已觸發 ${a.triggeredPrice} 元（${a.triggeredAt || ''}）</span>`
      : '';
    return `
      <div class="alert-item ${cls}">
        <div class="alert-info">
          <span class="alert-stock">${a.stockNo} ${a.stockName || ''}</span>
          <span class="alert-condition">${arrow}</span>
          <span class="alert-target">${a.targetPrice} 元</span>
          <span class="alert-current" id="cur-${a.id}">現價 --</span>
          ${tag}
        </div>
        <div class="alert-actions">
          ${!a.triggered ? `<button onclick="toggleAlert(${i})">${a.enabled ? '暫停' : '啟用'}</button>` : ''}
          <button class="btn-delete" onclick="deleteAlert(${i})">刪除</button>
        </div>
      </div>`;
  }).join('');

  // 非同步更新每筆現價
  alerts.forEach(a => fetchAlertPrice(a));
}

async function fetchAlertPrice(alert) {
  const el = document.getElementById(`cur-${alert.id}`);
  if (!el) return;
  const info = await getRealtimeInfo(alert.stockNo);
  if (!info || !info.z || info.z === '-' || info.z === '--') { el.textContent = '現價 --'; return; }
  const price = parseFloat(info.z);
  if (isNaN(price)) { el.textContent = '現價 --'; return; }
  const near  = alert.condition === 'lte'
    ? price <= alert.targetPrice * 1.05
    : price >= alert.targetPrice * 0.95;
  el.textContent = `現價 ${info.z}`;
  el.style.color = near ? '#ffaa33' : '#7788aa';
}

async function toggleAlert(idx) {
  const list = document.getElementById('alertList');
  const data = await gasGet();
  const alerts = data.alerts || [];
  const alert  = alerts[idx];
  if (!alert) return;
  await gasPost({ action: 'toggle', id: alert.id });
  loadAlerts();
}

async function deleteAlert(idx) {
  const data = await gasGet();
  const alerts = data.alerts || [];
  const alert  = alerts[idx];
  if (!alert) return;
  if (!confirm(`確定刪除 ${alert.stockNo} ${alert.stockName || ''} 的警報？`)) return;
  await gasPost({ action: 'delete', id: alert.id });
  loadAlerts();
}

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
  const stockNoEl    = document.getElementById('alertStockNo');
  const stockNameEl  = document.getElementById('alertStockName');
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
    stockNoEl.value   = item.code;
    stockNameEl.value = item.name;
    searchInput.value = `${item.code} ${item.name}`;
    hideDropdown();
    fetchFormPrice(item.code);
  }

  async function fetchFormPrice(stockNo) {
    const el = document.getElementById('alertCurrentPrice');
    el.textContent = '查詢中...';
    el.style.color = '#7788aa';
    const info = await getRealtimeInfo(stockNo);
    if (!info) { el.textContent = ''; return; }

    if (!info.z || info.z === '-' || info.z === '--') {
      // 休市：用 y（昨收）顯示
      if (info.y && info.y !== '--') {
        el.textContent = `昨收 ${info.y} 元`;
        el.style.color = '#aaaacc';
      } else {
        el.textContent = '';
      }
      return;
    }
    const price = parseFloat(info.z);
    const ref   = parseFloat(info.y);
    const diff  = price - ref;
    el.textContent = `現價 ${info.z} 元`;
    el.style.color = isNaN(diff) ? '#aaaacc' : (diff >= 0 ? '#ff6666' : '#44cc88');
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
  document.addEventListener('click', e => { if (!e.target.closest('#alertSearchContainer')) hideDropdown(); });
})();

// 新增警報
document.getElementById('addAlertBtn').addEventListener('click', async () => {
  const stockNo   = document.getElementById('alertStockNo').value.trim();
  const stockName = document.getElementById('alertStockName').value.trim();
  const condition = document.getElementById('alertCondition').value;
  const price     = parseFloat(document.getElementById('alertPrice').value);
  const status    = document.getElementById('addAlertStatus');

  if (!stockNo || isNaN(price) || price <= 0) {
    status.style.color = '#ff8888';
    status.textContent = '請填入股票代號與目標價格';
    return;
  }

  status.style.color = '#aaaacc';
  status.textContent = '新增中...';
  try {
    const result = await gasPost({
      action: 'add',
      alert: { id: genId(), stockNo, stockName, condition, targetPrice: price },
    });
    if (!result.ok) throw new Error(result.error || '伺服器錯誤');

    document.getElementById('alertSearchInput').value = '';
    document.getElementById('alertStockNo').value     = '';
    document.getElementById('alertStockName').value   = '';
    document.getElementById('alertCurrentPrice').textContent = '';
    document.getElementById('alertPrice').value       = '';
    status.style.color = '#66ff99';
    status.textContent = `✓ 已新增 ${stockNo} ${stockName || ''} 目標 ${price}`;
    setTimeout(() => status.textContent = '', 3000);
    loadAlerts();
  } catch (err) {
    status.style.color = '#ff8888';
    status.textContent = `✗ 新增失敗：${err.message}`;
  }
});

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
        <td class="code">${h.stockNo}</td>
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

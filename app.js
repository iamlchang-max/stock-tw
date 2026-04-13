// ── CORS Proxy ────────────────────────────────────────────────
// 用免費 proxy 繞過 TWSE 的跨網域限制
const PROXY = 'https://corsproxy.io/?url=';
function p(url) {
  return PROXY + encodeURIComponent(url);
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
  const url = `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${stockNo}.tw&json=1&delay=0`;
  try {
    const resp = await fetch(p(url));
    const data = await resp.json();
    return (data.msgArray || [])[0] || null;
  } catch {
    return null;
  }
}

async function getHistory(stockNo, months) {
  const allRows = [];
  const today = new Date();

  for (let i = months; i >= 0; i--) {
    const d = new Date(today.getFullYear(), today.getMonth() - i, 1);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const dateStr = `${year}${month}01`;
    const url = `https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=${dateStr}&stockNo=${stockNo}`;
    try {
      const resp = await fetch(p(url));
      const data = await resp.json();
      if (data.stat === 'OK') allRows.push(...(data.data || []));
    } catch { /* ignore */ }
    await sleep(300);
  }

  const dateMap = new Map();
  for (const row of allRows) {
    try {
      const parts = row[0].split('/');
      const year = parseInt(parts[0]) + 1911;
      const month = parts[1].padStart(2, '0');
      const day = parts[2].padStart(2, '0');
      const time = `${year}-${month}-${day}`;
      if (!dateMap.has(time)) {
        dateMap.set(time, {
          time,
          open:   parseNum(row[3]),
          high:   parseNum(row[4]),
          low:    parseNum(row[5]),
          close:  parseNum(row[6]),
          volume: parseInt(String(row[1]).replace(/,/g, '')),
        });
      }
    } catch { /* ignore */ }
  }

  return Array.from(dateMap.values()).sort((a, b) => a.time.localeCompare(b.time));
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
  if (!info) {
    document.getElementById('statusLabel').textContent += '（即時資料查詢失敗）';
    return;
  }
  const price = info.z || '--';
  const ref   = info.y || '--';
  let change = '--', color = 'white';
  try {
    const diff = parseFloat(price) - parseFloat(ref);
    change = (diff >= 0 ? '+' : '') + diff.toFixed(2);
    color  = diff >= 0 ? '#ff6666' : '#66ff99';
  } catch { /* market closed */ }

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

    if (records.length === 0) {
      status.textContent = '查無資料，請確認股票代號';
    } else {
      drawCharts(stockNo, records);
      status.textContent = `共 ${records.length} 筆資料`;
    }
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
  });
});

// ── Alerts Management ─────────────────────────────────────────

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function loadAlerts() {
  const { alerts, telegramToken, telegramChatId } = storageGet(['alerts', 'telegramToken', 'telegramChatId']);
  document.getElementById('telegramTokenInput').value  = telegramToken || '';
  document.getElementById('telegramChatIdInput').value = telegramChatId || '';
  renderAlertList(alerts || []);
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
  try {
    const url = `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${alert.stockNo}.tw&json=1&delay=0`;
    const resp = await fetch(p(url));
    const data = await resp.json();
    const info = (data.msgArray || [])[0];
    if (!info || !info.z || info.z === '-') { el.textContent = '現價 --'; return; }
    const price = parseFloat(info.z);
    const near  = alert.condition === 'lte'
      ? price <= alert.targetPrice * 1.05
      : price >= alert.targetPrice * 0.95;
    el.textContent = `現價 ${info.z}`;
    el.style.color = near ? '#ffaa33' : '#7788aa';
  } catch {
    el.textContent = '現價 --';
  }
}

function toggleAlert(idx) {
  const { alerts = [] } = storageGet(['alerts']);
  alerts[idx].enabled = !alerts[idx].enabled;
  storageSet({ alerts });
  renderAlertList(alerts);
}

function deleteAlert(idx) {
  const { alerts = [] } = storageGet(['alerts']);
  alerts.splice(idx, 1);
  storageSet({ alerts });
  renderAlertList(alerts);
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
    try {
      const url = `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${stockNo}.tw&json=1&delay=0`;
      const resp = await fetch(p(url));
      const data = await resp.json();
      const info = (data.msgArray || [])[0];
      if (!info || !info.z || info.z === '-') {
        // 休市時改抓最近收盤價
        const today = new Date();
        const dateStr = `${today.getFullYear()}${String(today.getMonth()+1).padStart(2,'0')}01`;
        const hUrl = `https://www.twse.com.tw/exchangeReport/STOCK_DAY?response=json&date=${dateStr}&stockNo=${stockNo}`;
        const hResp = await fetch(p(hUrl));
        const hData = await hResp.json();
        const rows = hData.data || [];
        if (rows.length) {
          const last = rows[rows.length - 1];
          const close = parseNum(last[6]);
          el.textContent = `收盤 ${close} 元`;
          el.style.color = '#aaaacc';
        } else {
          el.textContent = '';
        }
        return;
      }
      const price = parseFloat(info.z);
      const ref   = parseFloat(info.y);
      const diff  = price - ref;
      const color = diff >= 0 ? '#ff6666' : '#44cc88';
      el.textContent = `現價 ${info.z} 元`;
      el.style.color = color;
    } catch {
      el.textContent = '';
    }
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
document.getElementById('addAlertBtn').addEventListener('click', () => {
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

  try {
    const { alerts = [] } = storageGet(['alerts']);
    alerts.push({ id: genId(), stockNo, stockName, condition, targetPrice: price, enabled: true, triggered: false });
    storageSet({ alerts });

    document.getElementById('alertSearchInput').value = '';
    document.getElementById('alertStockNo').value     = '';
    document.getElementById('alertStockName').value   = '';
    document.getElementById('alertPrice').value       = '';
    status.style.color = '#66ff99';
    status.textContent = `✓ 已新增 ${stockNo} ${stockName || ''} 目標 ${price}`;
    setTimeout(() => status.textContent = '', 3000);
    renderAlertList(alerts);
  } catch (err) {
    status.style.color = '#ff8888';
    status.textContent = `✗ 新增失敗：${err.message}`;
  }
});

// ── Stock Search ─────────────────────────────────────────────

async function searchStock(query) {
  const url = `https://www.twse.com.tw/zh/api/codeQuery?query=${encodeURIComponent(query)}`;
  try {
    const resp = await fetch(p(url));
    const data = await resp.json();
    return (data.suggestions || [])
      .filter(s => s && s !== 'bar')
      .map(s => {
        const parts = s.split('\t');
        return { code: parts[0]?.trim(), name: parts[1]?.trim() || '' };
      })
      .filter(s => s.code && /^\d+$/.test(s.code));
  } catch {
    return [];
  }
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

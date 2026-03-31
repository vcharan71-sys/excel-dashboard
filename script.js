const PALETTE = [
  "#c9410a","#1a6b3e","#1a3a6b","#8b3a6e","#6b5a1a",
  "#3a6b8b","#6b1a1a","#1a6b6b","#6b3a1a","#3a1a6b",
  "#4a8b3a","#8b6b1a"
];

let sheetsData = [];
let activeSheet = 0;
const chartInstances = {};

// DOM refs
const dropZone    = document.getElementById('drop-zone');
const fileInput   = document.getElementById('file-input');
const fileBar     = document.getElementById('file-bar');
const fileNameEl  = document.getElementById('file-name-display');
const fileMetaEl  = document.getElementById('file-meta-display');
const errorBox    = document.getElementById('error-box');
const sheetTabs   = document.getElementById('sheet-tabs');
const results     = document.getElementById('results');
const chartsGrid  = document.getElementById('charts-grid');
const emptyMsg    = document.getElementById('empty-msg');
const resultsTitle = document.getElementById('results-title');
const resultsCount = document.getElementById('results-count');

// Upload interactions
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragging'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragging'));
dropZone.addEventListener('drop', e => { e.preventDefault(); dropZone.classList.remove('dragging'); handleFile(e.dataTransfer.files[0]); });
fileInput.addEventListener('change', e => handleFile(e.target.files[0]));

function showError(msg) {
  errorBox.textContent = '⚠️ ' + msg;
  errorBox.classList.add('visible');
}
function clearError() { errorBox.classList.remove('visible'); }

function handleFile(file) {
  clearError();
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) {
    showError('Please upload an .xlsx, .xls, or .csv file.');
    return;
  }

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      sheetsData = wb.SheetNames.map(name => {
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
        const columns = extractCategorical(rows);
        return { name, rows: rows.length, columns };
      });
      fileNameEl.textContent = file.name;
      fileMetaEl.textContent = `${sheetsData.length} sheet${sheetsData.length > 1 ? 's' : ''} · ${(file.size / 1024).toFixed(1)} KB`;
      fileBar.classList.add('visible');
      renderSheetTabs();
      activeSheet = 0;
      renderCharts();
    } catch (err) {
      showError('Could not parse file. Check it is a valid Excel or CSV file.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function extractCategorical(rows) {
  if (!rows.length) return [];
  return Object.keys(rows[0]).map(col => {
    const vals = rows.map(r => r[col]).filter(v => v !== '' && v != null);
    if (!vals.length) return null;
    const numericCount = vals.filter(v => typeof v === 'number' || (!isNaN(Number(v)) && String(v).trim() !== '')).length;
    const isNumeric = numericCount / vals.length > 0.85;
    const unique = [...new Set(vals.map(v => String(v)))];
    if (isNumeric && unique.length > 15) return null;
    const freq = {};
    vals.forEach(v => { const k = String(v); freq[k] = (freq[k] || 0) + 1; });
    const data = Object.entries(freq)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 12)
      .map(([label, value]) => ({ label, value }));
    if (data.length < 2) return null;
    return { col, data };
  }).filter(Boolean);
}

function renderSheetTabs() {
  sheetTabs.innerHTML = '';
  if (sheetsData.length <= 1) { sheetTabs.classList.remove('visible'); return; }
  sheetTabs.classList.add('visible');
  sheetsData.forEach((s, i) => {
    const btn = document.createElement('button');
    btn.className = 'tab-btn' + (i === activeSheet ? ' active' : '');
    btn.textContent = s.name;
    btn.addEventListener('click', () => { activeSheet = i; renderSheetTabs(); renderCharts(); });
    sheetTabs.appendChild(btn);
  });
}

function renderCharts() {
  Object.values(chartInstances).forEach(c => c.destroy());
  for (const k in chartInstances) delete chartInstances[k];

  chartsGrid.innerHTML = '';
  results.classList.add('visible');

  const sheet = sheetsData[activeSheet];
  if (!sheet || !sheet.columns.length) {
    emptyMsg.classList.add('visible');
    chartsGrid.style.display = 'none';
    resultsTitle.textContent = sheet?.name || 'Results';
    resultsCount.textContent = '';
    return;
  }

  emptyMsg.classList.remove('visible');
  chartsGrid.style.display = 'grid';
  resultsTitle.textContent = sheet.name;
  resultsCount.textContent = `${sheet.columns.length} chart${sheet.columns.length > 1 ? 's' : ''} · ${sheet.rows.toLocaleString()} rows`;

  sheet.columns.forEach((col, idx) => {
    const card = buildCard(col, idx);
    chartsGrid.appendChild(card);
    card.style.animationDelay = (idx * 0.07) + 's';
  });
}

function buildCard({ col, data }, idx) {
  const total = data.reduce((s, d) => s + d.value, 0);
  const pct = v => total > 0 ? ((v / total) * 100).toFixed(1) : '0.0';
  const colors = data.map((_, i) => PALETTE[i % PALETTE.length]);

  const card = document.createElement('div');
  card.className = 'chart-card';
  const canvasId = `chart-${idx}`;

  card.innerHTML = `
    <div class="card-eyebrow">Column Analysis</div>
    <div class="card-title">${escHtml(col)}</div>
    <div class="card-meta">${data.length} categories · ${total.toLocaleString()} total values</div>
    <div class="chart-row">
      <div class="chart-wrap">
        <canvas id="${canvasId}"></canvas>
        <div class="chart-center-label">
          <div class="cl-num">${data.length}</div>
          <div class="cl-txt">categories</div>
        </div>
      </div>
      <div class="legend">
        ${data.map((d, i) => `
          <div class="legend-item">
            <div class="legend-dot" style="background:${colors[i]}"></div>
            <div class="legend-label" title="${escHtml(d.label)}">${escHtml(d.label)}</div>
            <div class="legend-pct">${pct(d.value)}%</div>
          </div>
        `).join('')}
      </div>
    </div>
    <table class="pct-table">
      <thead><tr><th>Category</th><th>Share</th><th>Count</th></tr></thead>
      <tbody>
        ${data.map((d, i) => `
          <tr>
            <td>
              <div class="bar-cell">
                <div class="mini-bar" style="width:${Math.max(4, (d.value / total) * 100)}px;background:${colors[i]}"></div>
                ${escHtml(d.label)}
              </div>
            </td>
            <td><strong style="color:${colors[i]}">${pct(d.value)}%</strong></td>
            <td>${d.value.toLocaleString()}</td>
          </tr>
        `).join('')}
        <tr>
          <td><strong>Total</strong></td>
          <td><strong>100%</strong></td>
          <td><strong>${total.toLocaleString()}</strong></td>
        </tr>
      </tbody>
    </table>
  `;

  requestAnimationFrame(() => {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    chartInstances[canvasId] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: data.map(d => d.label),
        datasets: [{
          data: data.map(d => d.value),
          backgroundColor: colors,
          borderColor: '#fff',
          borderWidth: 2,
          hoverOffset: 6
        }]
      },
      options: {
        cutout: '62%',
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => ` ${ctx.label}: ${ctx.raw.toLocaleString()} (${pct(ctx.raw)}%)`
            }
          }
        },
        animation: { animateRotate: true, duration: 700 }
      }
    });
  });

  return card;
}

function escHtml(str) {
  return String(str).replace(/[&<>"']/g, m => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[m]));
}

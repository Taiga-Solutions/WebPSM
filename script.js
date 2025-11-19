// ---------- State ----------
let currentData = null;
let chart = null;
const cdfCache = new Map(); // ⭐ CDF-cache

// ---------- Demo data ----------
const demoData = {
    tooCheap: [2.15, 4.15, 2.64, 4.53, 4.76, 1.18, 3.11, 4.57, 3.21, 2.83, 4.83, 2.81, 3.71, 3.29, 1.41, 4.60, 1.98, 1.17, 2.31, 4.82],
    bargain: [5.56, 4.77, 4.56, 5.98, 4.62, 4.83, 4.18, 4.38, 3.16, 2.59, 5.85, 5.61, 4.76, 5.18, 2.10, 3.91, 5.03, 2.87, 3.27, 2.93],
    expensive: [3.57, 4.66, 4.65, 4.48, 3.61, 3.56, 3.93, 4.86, 4.06, 6.43, 3.18, 4.77, 6.20, 3.49, 5.24, 3.83, 3.51, 6.01, 6.58, 4.50],
    tooExpensive: [6.66, 4.38, 5.54, 5.10, 7.26, 5.79, 7.24, 7.25, 7.18, 5.76, 7.02, 6.52, 6.84, 4.00, 5.90, 4.88, 5.52, 6.45, 5.41, 4.44]
};

// ---------- Helpers ----------
function $(id) { return document.getElementById(id); }

function showTab(tabId) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    const btn = document.querySelector('.tab[data-tab="' + tabId + '"]');
    if (btn) btn.classList.add('active');
    const content = document.getElementById(tabId);
    if (content) content.classList.add('active');
}

document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => showTab(btn.dataset.tab));
});

// ---------- CDF with cache ----------
function calculateCDF(data, x) {
    if (!Array.isArray(data) || data.length === 0) return 0;
    const key = `${data.join(',')}|${x}`;
    if (cdfCache.has(key)) return cdfCache.get(key);
    const sorted = [...data].sort((a, b) => a - b);
    let count = 0;
    for (const val of sorted) {
        if (val <= x) count++;
        else break;
    }
    const result = count / data.length;
    cdfCache.set(key, result);
    return result;
}

// ---------- File input ----------
const fileInput = $('fileInput');
$('openFileBtn').addEventListener('click', () => fileInput.click());

let fileRawData = null;

fileInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = ev => {
        try {
            const data = new Uint8Array(ev.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
            if (json.length < 2) return alert('Tiedosto on tyhjä tai virheellinen.');
            fileRawData = json;
            populateColumnSelects(json[0].map(h => String(h)));
            $('columnMapping').classList.add('active');
            showTab('upload');
        } catch {
            alert('Tiedoston lukeminen epäonnistui.');
        }
    };
    reader.readAsArrayBuffer(file);
});

function populateColumnSelects(headers) {
    ['tooCheapCol', 'bargainCol', 'expensiveCol', 'tooExpensiveCol'].forEach(id => {
        const sel = $(id);
        sel.innerHTML = '<option value="">-- Valitse --</option>';
        headers.forEach((h, idx) => {
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = h || `Sarake ${idx + 1}`;
            sel.appendChild(opt);
        });
    });
}

// ---------- Demo ----------
$('demoBtn').addEventListener('click', () => {
    cdfCache.clear(); // ⭐
    currentData = JSON.parse(JSON.stringify(demoData));
    $('columnMapping').classList.remove('active');
    populateDataTable();
    $('dataTab').style.display = 'inline-block';
    $('resultsTabBtn').style.display = 'inline-block';
    performAnalysis();
    showTab('results');
});

// ---------- Analyze ----------
$('analyzeBtn').addEventListener('click', analyzeData);

function analyzeData() {
    if (!fileRawData) return alert('Lataa ensin Excel-tiedosto.');
    const idxs = {
        tooCheap: $('tooCheapCol').value,
        bargain: $('bargainCol').value,
        expensive: $('expensiveCol').value,
        tooExpensive: $('tooExpensiveCol').value
    };
    if (Object.values(idxs).some(v => v === '')) return alert('Valitse kaikki sarakkeet!');
    const extracted = { tooCheap: [], bargain: [], expensive: [], tooExpensive: [] };
    for (let i = 1; i < fileRawData.length; i++) {
        const row = fileRawData[i];
        const tc = parseNumberSafe(row[idxs.tooCheap]);
        const b = parseNumberSafe(row[idxs.bargain]);
        const e = parseNumberSafe(row[idxs.expensive]);
        const te = parseNumberSafe(row[idxs.tooExpensive]);
        if (!isNaN(tc)) extracted.tooCheap.push(tc);
        if (!isNaN(b)) extracted.bargain.push(b);
        if (!isNaN(e)) extracted.expensive.push(e);
        if (!isNaN(te)) extracted.tooExpensive.push(te);
    }
    cdfCache.clear(); // ⭐
    currentData = extracted;
    $('columnMapping').classList.remove('active');
    populateDataTable();
    $('dataTab').style.display = 'inline-block';
    $('resultsTabBtn').style.display = 'inline-block';
    performAnalysis();
    showTab('results');
}

function parseNumberSafe(v) {
    if (v === null || v === undefined || v === '') return NaN;
    const n = parseFloat(String(v).replace(',', '.'));
    return isNaN(n) ? NaN : n;
}

// ---------- Data table ----------
function populateDataTable() {
    const tbody = $('dataTableBody');
    tbody.innerHTML = '';
    if (!currentData) return;
    const len = Math.max(currentData.tooCheap.length, currentData.bargain.length, currentData.expensive.length, currentData.tooExpensive.length);
    for (let i = 0; i < len; i++) {
        addRowToTable(i,
            currentData.tooCheap[i] ?? '',
            currentData.bargain[i] ?? '',
            currentData.expensive[i] ?? '',
            currentData.tooExpensive[i] ?? ''
        );
    }
}

function addRowToTable(index, tooCheap = '', bargain = '', expensive = '', tooExpensive = '') {
    const tbody = $('dataTableBody');
    const row = tbody.insertRow();
    row.innerHTML = `
        <td>${index + 1}</td>
        <td><input type="number" step="0.01" value="${tooCheap}" data-col="tooCheap" data-row="${index}"></td>
        <td><input type="number" step="0.01" value="${bargain}" data-col="bargain" data-row="${index}"></td>
        <td><input type="number" step="0.01" value="${expensive}" data-col="expensive" data-row="${index}"></td>
        <td><input type="number" step="0.01" value="${tooExpensive}" data-col="tooExpensive" data-row="${index}"></td>
    `;
}

$('addRowBtn').addEventListener('click', () => addRowToTable($('dataTableBody').rows.length));
$('removeRowBtn').addEventListener('click', () => {
    const tbody = $('dataTableBody');
    if (tbody.rows.length > 0) tbody.deleteRow(tbody.rows.length - 1);
});

$('updateBtn').addEventListener('click', updateFromTable);

function updateFromTable() {
    const tbody = $('dataTableBody');
    const extracted = { tooCheap: [], bargain: [], expensive: [], tooExpensive: [] };
    for (let i = 0; i < tbody.rows.length; i++) {
        const inputs = tbody.rows[i].querySelectorAll('input');
        const tc = parseNumberSafe(inputs[0].value);
        const b = parseNumberSafe(inputs[1].value);
        const e = parseNumberSafe(inputs[2].value);
        const te = parseNumberSafe(inputs[3].value);
        if (!isNaN(tc)) extracted.tooCheap.push(tc);
        if (!isNaN(b)) extracted.bargain.push(b);
        if (!isNaN(e)) extracted.expensive.push(e);
        if (!isNaN(te)) extracted.tooExpensive.push(te);
    }
    cdfCache.clear(); // ⭐
    currentData = extracted;
    performAnalysis();
    showTab('results');
}

// ---------- Analysis ----------
function performAnalysis() {
    if (!currentData) return;
    const all = [
        ...currentData.tooCheap || [],
        ...currentData.bargain || [],
        ...currentData.expensive || [],
        ...currentData.tooExpensive || []
    ].filter(v => typeof v === 'number' && !isNaN(v));
    if (all.length === 0) return alert('Ei kelvollisia hintatietoja.');
    const minPrice = Math.min(...all);
    const maxPrice = Math.max(...all);
    const PMC = findIntersection(currentData.tooCheap, currentData.tooExpensive);
    const OPP = findIntersection(currentData.bargain, currentData.expensive);
    const APRL = findIntersection(currentData.tooCheap, currentData.expensive);
    const APRU = findIntersection(currentData.bargain, currentData.tooExpensive);
    const priceRange = linspace(minPrice, maxPrice, 201);
    const tooCheapCurve = priceRange.map(x => calculateCDF(currentData.tooCheap, x) * 100);
    const bargainCurve = priceRange.map(x => calculateCDF(currentData.bargain, x) * 100);
    const expensiveCurve = priceRange.map(x => (1 - calculateCDF(currentData.expensive, x)) * 100);
    const tooExpensiveCurve = priceRange.map(x => (1 - calculateCDF(currentData.tooExpensive, x)) * 100);
    drawChart(priceRange, tooCheapCurve, bargainCurve, expensiveCurve, tooExpensiveCurve, PMC, OPP, APRL, APRU);
    displayPricePoints(PMC, OPP, APRL, APRU);
}

function linspace(min, max, n) {
    if (n <= 1) return [min];
    const out = [];
    const step = (max - min) / (n - 1);
    for (let i = 0; i < n; i++) out.push(min + step * i);
    return out;
}

/*
function median(values) {
    if (values.length === 0) return null;

    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);

    if (sorted.length % 2 === 0) {
        // Parillinen määrä: kahden keskimmäisen keskiarvo
        return (sorted[mid - 1] + sorted[mid]) / 2;
    } else {
        // Pariton määrä: keskimmäinen arvo
        return sorted[mid];
    }
}*/

/**
 * Etsii funktion f(x) = |1 - CDF1(x) - CDF2(x)| minimin
 * ja palauttaa minimin arvon sekä x-arvon (tai peräkkäisten minien keskipisteen)
 *
 * @param {number[]} data1 - Ensimmäinen järjestetty datasetti
 * @param {number[]} data2 - Toinen järjestetty datasetti
 * @param {object} options - Valinnaiset asetukset
 * @returns {{x: number, value: number}}
 */
function findKSStatistic(data1, data2, options = {}) {
    // Oletusasetukset
    const {
        steps = 10000,       // Mitä enemmän askelia, sitä tarkempi
        margin = 0.01         // Marginaali datan ääriarvojen ulkopuolelle
    } = options;

    if (data1.length === 0 || data2.length === 0) {
        throw new Error("Datasetit eivät saa olla tyhjiä");
    }

    const allValues = [...data1, ...data2];
    const minX = Math.min(...allValues) - margin;
    const maxX = Math.max(...allValues) + margin;

    let bestX = null;
    let bestValue = Infinity;
    let plateauStart = null;  // Jos useita peräkkäisiä minimejä

    // Käydään läpi tiheä määrä pisteitä
    for (let i = 0; i <= steps; i++) {
        const x = minX + (maxX - minX) * i / steps;
        const value = Math.abs(1 - calculateCDF(data1, x) - calculateCDF(data2, x));

        if (value < bestValue - 1e-12) {  // Selvästi pienempi
            bestValue = value;
            bestX = x;
            plateauStart = null;         // Aloitetaan uusi plateau
        } else if (Math.abs(value - bestValue) < 1e-12) {  // Sama minimiarvo (kelluvapilkku huomioiden)
            if (plateauStart === null) {
                plateauStart = x;        // Plateau alkaa
            }
            bestX = x;                   // Päivitetään ääripää
        } else if (plateauStart !== null) {
            // Plateau loppui edellisellä kierroksella → laske keskiarvo ja lopeta plateau
            const plateauMid = (plateauStart + bestX) / 2;
            bestX = plateauMid;
            plateauStart = null;
        }
    }

    // Jos plateau jäi kesken loppuun asti
    if (plateauStart !== null) {
        bestX = (plateauStart + bestX) / 2;
    }

    return {
        x: bestX,
        value: bestValue
    };
}

function findIntersection(data1, data2) {
    const max_of_mins = Math.max(Math.min(...data1), Math.min(...data2));
    const min_of_maxs = Math.min(Math.max(...data1), Math.max(...data2));

    if (max_of_mins > min_of_maxs) return ((min_of_maxs + max_of_mins) / 2);

    let { x: bestX } = findKSStatistic(data1, data2)

    return bestX;
}

// ---------- Chart ----------
function drawChart(prices, tooCheap, bargain, expensive, tooExpensive, pmc, opp, aprl, apru) {
    const ctx = $('psmChart').getContext('2d');
    if (chart) chart.destroy();
    chart = new Chart(ctx, {
        type: 'line',
        data: {
            datasets: [
                {
                    label: 'Liian halpa',
                    data: prices.map((p, i) => ({ x: p, y: tooCheap[i] })),
                    fill: false,
                    borderColor: '#3182ce',
                    backgroundColor: 'rgba(49,130,206,0.1)',
                    tension: 0.3
                },
                {
                    label: 'Edullinen',
                    data: prices.map((p, i) => ({ x: p, y: bargain[i] })),
                    fill: false,
                    borderColor: '#48bb78',
                    backgroundColor: 'rgba(72,187,120,0.1)',
                    tension: 0.3
                },
                {
                    label: 'Kallis',
                    data: prices.map((p, i) => ({ x: p, y: expensive[i] })),
                    fill: false,
                    borderColor: '#ed8936',
                    backgroundColor: 'rgba(237,137,54,0.1)',
                    tension: 0.3
                },
                {
                    label: 'Liian kallis',
                    data: prices.map((p, i) => ({ x: p, y: tooExpensive[i] })),
                    fill: false,
                    borderColor: '#e53e3e',
                    backgroundColor: 'rgba(229,62,62,0.1)',
                    tension: 0.3
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: 'index', intersect: false },
            plugins: {
                legend: { position: 'top' },
                title: { display: true, text: 'Van Westendorp Price Sensitivity Meter', font: { size: 16 } },
                annotation: {
                    annotations: {
                        pmcLine: {
                            type: 'line',
                            xMin: pmc.toFixed(2),
                            xMax: pmc.toFixed(2),
                            borderColor: '#9f7aea',
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: { content: ['PMC', Number(pmc.toFixed(2)) + ' €'], enabled: true, position: 'start' }
                        },
                        oppLine: {
                            type: 'line',
                            xMin: opp.toFixed(2),
                            xMax: opp.toFixed(2),
                            borderColor: '#38b2ac',
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: { content: ['OPP', Number(opp.toFixed(2)) + ' €'], enabled: true, position: 'end' }
                        },
                        aprlLine: {
                            type: 'line',
                            xMin: aprl.toFixed(2),
                            xMax: aprl.toFixed(2),
                            borderColor: '#ecc94b',
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: { content: ['APRL', Number(aprl.toFixed(2)) + ' €'], enabled: true, position: 'start', yAdjust: 20 }
                        },
                        apruLine: {
                            type: 'line',
                            xMin: apru.toFixed(2),
                            xMax: apru.toFixed(2),
                            borderColor: '#f56565',
                            borderWidth: 2,
                            borderDash: [6, 4],
                            label: { content: ['APRU', Number(apru.toFixed(2)) + ' €'], enabled: true, position: 'end', yAdjust: 20 }
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'linear', // ⭐ Lisää tämä
                    title: { display: true, text: 'Hinta (€)' },
                    ticks: { maxTicksLimit: 10 }
                },
                y: { title: { display: true, text: 'Prosentti (%)' }, min: 0, max: 100 }
            }
        },
        plugins: [Chart.registry.getPlugin('annotation')]
    });
}

function displayPricePoints(pmc, opp, aprl, apru) {
    const container = $('pricePoints');
    const pmcNum = Number(pmc);
    const oppNum = Number(opp);
    const aprlNum = Number(aprl);
    const apruNum = Number(apru);

    const points = [
        { name: 'Acceptable Price Range Lower (APRL)', price: aprlNum.toFixed(2) + ' €', desc: 'Liian halpa ∩ Kallis' },
        { name: 'Optimal Price Point (OPP)', price: oppNum.toFixed(2) + ' €', desc: 'Edullinen ∩ Kallis' },
        { name: 'Point of Marginal Cheapness (PMC)', price: pmcNum.toFixed(2) + ' €', desc: 'Liian halpa ∩ Liian kallis' },
        { name: 'Acceptable Price Range Upper (APRU)', price: apruNum.toFixed(2) + ' €', desc: 'Edullinen ∩ Liian kallis' }
    ];
    container.innerHTML = points.map(p => `
        <div class="price-card">
            <h3>${p.name}</h3>
            <div class="price">${p.price}</div>
            <p style="margin-top:10px;font-size:0.9em;opacity:0.9;">${p.desc}</p>
        </div>
    `).join('');
}

// ---------- Export ----------
$('exportExcelBtn').addEventListener('click', exportToExcel);
$('exportPptBtn').addEventListener('click', exportToPPT);

function exportToExcel() {
    if (!currentData) return alert('Ei dataa vietäväksi.');

    const wb = XLSX.utils.book_new();

    // Raw Data Sheet
    const len = Math.max(currentData.tooCheap.length, currentData.bargain.length, currentData.expensive.length, currentData.tooExpensive.length);
    const rawData = [['Liian halpa', 'Edullinen', 'Kallis', 'Liian kallis']];
    for (let i = 0; i < len; i++) {
        rawData.push([
            currentData.tooCheap[i] ?? '',
            currentData.bargain[i] ?? '',
            currentData.expensive[i] ?? '',
            currentData.tooExpensive[i] ?? ''
        ]);
    }
    const wsRaw = XLSX.utils.aoa_to_sheet(rawData);
    XLSX.utils.book_append_sheet(wb, wsRaw, "Data");

    // Results Sheet
    const pmc = findIntersection(currentData.tooCheap, currentData.tooExpensive);
    const opp = findIntersection(currentData.bargain, currentData.expensive);
    const aprl = findIntersection(currentData.tooCheap, currentData.expensive);
    const apru = findIntersection(currentData.bargain, currentData.tooExpensive);

    const resultsData = [
        ['Mittari', 'Hinta (€)', 'Kuvaus'],
        ['APRL', aprl.toFixed(2), 'Acceptable Price Range Lower'],
        ['OPP', opp.toFixed(2), 'Optimal Price Point'],
        ['PMC', pmc.toFixed(2), 'Point of Marginal Cheapness'],
        ['APRU', apru.toFixed(2), 'Acceptable Price Range Upper']
    ];
    const wsRes = XLSX.utils.aoa_to_sheet(resultsData);
    XLSX.utils.book_append_sheet(wb, wsRes, "Tulokset");

    XLSX.writeFile(wb, "PSM_Analyysi.xlsx");
}

function exportToPPT() {
    if (!currentData) return alert('Ei dataa vietäväksi.');

    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_16x9';

    // Slide 1: Title
    let slide1 = pptx.addSlide();
    slide1.addText("Van Westendorp Price Sensitivity Meter", { x: 1, y: 1, w: '80%', fontSize: 24, bold: true, color: '363636' });
    slide1.addText("Analyysin tulokset", { x: 1, y: 2, fontSize: 18, color: '718096' });
    slide1.addText(`Luotu: ${new Date().toLocaleDateString()}`, { x: 1, y: 3, fontSize: 14, color: 'A0AEC0' });

    // Slide 2: Results
    let slide2 = pptx.addSlide();
    slide2.addText("Tulokset", { x: 0.5, y: 0.5, fontSize: 18, bold: true, color: '363636' });

    const pmc = findIntersection(currentData.tooCheap, currentData.tooExpensive);
    const opp = findIntersection(currentData.bargain, currentData.expensive);
    const aprl = findIntersection(currentData.tooCheap, currentData.expensive);
    const apru = findIntersection(currentData.bargain, currentData.tooExpensive);

    const rows = [
        ['Mittari', 'Hinta (€)', 'Kuvaus'],
        ['APRL', aprl.toFixed(2) + ' €', 'Acceptable Price Range Lower'],
        ['OPP', opp.toFixed(2) + ' €', 'Optimal Price Point'],
        ['PMC', pmc.toFixed(2) + ' €', 'Point of Marginal Cheapness'],
        ['APRU', apru.toFixed(2) + ' €', 'Acceptable Price Range Upper']
    ];

    slide2.addTable(rows, { x: 1, y: 1.5, w: 8, colW: [2, 2, 4], border: { pt: 1, color: 'E2E8F0' }, fill: 'F7FAFC' });

    // Slide 3: Chart
    let slide3 = pptx.addSlide();
    slide3.addText("Kuvaaja", { x: 0.5, y: 0.5, fontSize: 18, bold: true, color: '363636' });

    // Capture chart image
    const chartCanvas = document.getElementById('psmChart');
    const chartImg = chartCanvas.toDataURL('image/png');
    slide3.addImage({ data: chartImg, x: 1, y: 1, w: 8, h: 4.5 });

    pptx.writeFile({ fileName: "PSM_Analyysi.pptx" });
}

// ---------- Init ----------
(function init() {
    showTab('upload');
    $('dataTab').style.display = 'none';
    $('resultsTabBtn').style.display = 'none';
})();

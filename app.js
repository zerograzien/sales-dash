// Global State Storage
let globalCustomerData = {};
let globalBrandData = {};
let globalProductData = {};
let globalMonthlyData = {};

let totalSalesVal = 0;
let totalQtyVal = 0;
let totalCasesVal = 0;
let totalTransactionsVal = 0;

let activeChartMetric = 'sales'; // 'sales' | 'qty' | 'cases'
let activeLeaderboardTab = 'customers'; // 'customers' | 'brands'
let trendChartInstance = null;

// Utility Formatters
const formatCurrency = (val) => new Intl.NumberFormat('en-PH', { style: 'currency', currency: 'PHP' }).format(val);
const formatNumber = (val) => new Intl.NumberFormat('en-US').format(Math.round(val));

/**
 * Header Index Resolver
 * Uses STRICT EQUALITY on sanitized strings to ensure exact column matching
 * (e.g. 'customer' will NOT match 'customer code').
 */
function getColumnIndex(headers, targets) {
    return headers.findIndex(h => {
        if (!h) return false;
        const cleanHeader = h.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
        return targets.some(t => {
            const cleanTarget = t.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, '');
            return cleanHeader === cleanTarget;
        });
    });
}

/**
 * Excel & JS Safe Date Parser
 * Handles Excel numeric serials (e.g. 45306), native JS Date objects, and strings.
 */
function parseExcelDate(val) {
    if (!val) return null;
    
    // JS Date Object
    if (val instanceof Date && !isNaN(val.getTime())) {
        return val;
    }
    
    // Excel Numeric Serial Number
    if (typeof val === 'number') {
        const dateObj = new Date((val - (25567 + 2)) * 86400 * 1000);
        return isNaN(dateObj.getTime()) ? null : dateObj;
    }
    
    // Standard Date String
    const dateObj = new Date(val);
    return isNaN(dateObj.getTime()) ? null : dateObj;
}

// File Upload Handler
document.getElementById('excel-file')?.addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.csv')) {
        Papa.parse(file, {
            complete: (results) => processParsedRows(results.data),
            error: (err) => alert("Error parsing CSV file: " + err.message)
        });
    } else {
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            // cellDates: true enables SheetJS to parse dates directly
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            processParsedRows(rows);
        };
        reader.readAsArrayBuffer(file);
    }
}

function processParsedRows(rows) {
    if (!rows || rows.length < 2) {
        alert("Spreadsheet appears to be empty or missing data rows.");
        return;
    }

    const headers = rows[0];

    // Column Mapping
    const salesIdx = getColumnIndex(headers, ['nettotal', 'net total']);
    const qtyIdx = getColumnIndex(headers, ['quantitysold', 'quantity sold']);
    const casesIdx = getColumnIndex(headers, ['qtyfactor', 'qty factor']);
    const custIdx = getColumnIndex(headers, ['customer']);
    const brandIdx = getColumnIndex(headers, ['subbrand', 'sub brand']);
    const prodIdx = getColumnIndex(headers, ['itemdescription', 'item description']);
    const dateIdx = getColumnIndex(headers, ['transdate', 'trans date']);
    const addrIdx = getColumnIndex(headers, ['customeraddress', 'customer address']);

    // Reset Counters
    globalCustomerData = {};
    globalBrandData = {};
    globalProductData = {};
    globalMonthlyData = {};

    totalSalesVal = 0;
    totalQtyVal = 0;
    totalCasesVal = 0;
    totalTransactionsVal = rows.length - 1;

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;

        const rawSales = salesIdx !== -1 ? row[salesIdx] : 0;
        const rawQty = qtyIdx !== -1 ? row[qtyIdx] : 0;
        const rawCases = casesIdx !== -1 ? row[casesIdx] : 0;

        const sales = parseFloat(String(rawSales).replace(/[^0-9.-]+/g, '')) || 0;
        const qty = parseFloat(String(rawQty).replace(/[^0-9.-]+/g, '')) || 0;
        const cases = parseFloat(String(rawCases).replace(/[^0-9.-]+/g, '')) || 0;

        const custName = (custIdx !== -1 && row[custIdx] !== undefined && row[custIdx] !== null) 
            ? String(row[custIdx]).trim() 
            : "Unassigned Customer";

        const brandName = (brandIdx !== -1 && row[brandIdx] !== undefined && row[brandIdx] !== null) 
            ? String(row[brandIdx]).trim() 
            : "Unassigned Brand";

        const prodName = (prodIdx !== -1 && row[prodIdx] !== undefined && row[prodIdx] !== null) 
            ? String(row[prodIdx]).trim() 
            : "Unassigned Product";

        const rawDate = (dateIdx !== -1 && row[dateIdx] !== undefined) ? row[dateIdx] : null;
        const address = (addrIdx !== -1 && row[addrIdx] !== undefined && row[addrIdx] !== null) 
            ? String(row[addrIdx]).trim() 
            : "No Address Recorded";

        totalSalesVal += sales;
        totalQtyVal += qty;
        totalCasesVal += cases;

        // Month Trend Aggregation
        let monthKey = "Other";
        if (rawDate) {
            const parsedDate = parseExcelDate(rawDate);
            if (parsedDate) {
                monthKey = parsedDate.toLocaleString('default', { month: 'short', year: '2-digit' });
            }
        }

        if (!globalMonthlyData[monthKey]) {
            globalMonthlyData[monthKey] = { sales: 0, qty: 0, cases: 0 };
        }
        globalMonthlyData[monthKey].sales += sales;
        globalMonthlyData[monthKey].qty += qty;
        globalMonthlyData[monthKey].cases += cases;

        // Customer Aggregation
        if (!globalCustomerData[custName]) {
            globalCustomerData[custName] = {
                sales: 0, qty: 0, cases: 0, orders: 0,
                address: address,
                firstDate: rawDate,
                lastDate: rawDate,
                products: {}
            };
        }
        globalCustomerData[custName].sales += sales;
        globalCustomerData[custName].qty += qty;
        globalCustomerData[custName].cases += cases;
        globalCustomerData[custName].orders += 1;

        if (address !== "No Address Recorded") {
            globalCustomerData[custName].address = address;
        }

        if (!globalCustomerData[custName].products[prodName]) {
            globalCustomerData[custName].products[prodName] = { qty: 0, cases: 0, sales: 0 };
        }
        globalCustomerData[custName].products[prodName].qty += qty;
        globalCustomerData[custName].products[prodName].cases += cases;
        globalCustomerData[custName].products[prodName].sales += sales;

        // Sub Brand Aggregation
        if (!globalBrandData[brandName]) {
            globalBrandData[brandName] = { sales: 0, qty: 0, cases: 0 };
        }
        globalBrandData[brandName].sales += sales;
        globalBrandData[brandName].qty += qty;
        globalBrandData[brandName].cases += cases;

        // Product Aggregation
        if (!globalProductData[prodName]) {
            globalProductData[prodName] = { sales: 0, qty: 0, cases: 0 };
        }
        globalProductData[prodName].sales += sales;
        globalProductData[prodName].qty += qty;
        globalProductData[prodName].cases += cases;
    }

    // Toggle UI State
    document.getElementById('empty-state')?.classList.add('hidden');
    document.getElementById('dashboard')?.classList.remove('hidden');

    updateKPIs();
    renderProductBreakdown();
    renderLeaderboard();
    updateTrendChart();
}

// KPI Dashboard Updates
function updateKPIs() {
    document.getElementById('kpi-sales').textContent = formatCurrency(totalSalesVal);
    document.getElementById('kpi-transactions').textContent = formatNumber(totalTransactionsVal);
    document.getElementById('kpi-items').textContent = formatNumber(totalQtyVal);
    document.getElementById('kpi-cases').textContent = formatNumber(totalCasesVal);

    updateTargetStatus();
}

// Dynamic Target Calculation & Visual Feedback
function updateTargetStatus() {
    const inputEl = document.getElementById('target-input');
    const rawTarget = inputEl?.value.replace(/[^0-9.-]+/g, '') || "0";
    const targetVal = parseFloat(rawTarget) || 40000000;

    const pct = targetVal > 0 ? (totalSalesVal / targetVal) * 100 : 0;
    const remaining = Math.max(targetVal - totalSalesVal, 0);

    const pctEl = document.getElementById('kpi-target-pct');
    const barEl = document.getElementById('kpi-target-bar');
    const remEl = document.getElementById('kpi-target-rem');

    if (pctEl) pctEl.textContent = `${pct.toFixed(2)}%`;
    if (barEl) barEl.style.width = `${Math.min(pct, 100)}%`;

    if (pct >= 100) {
        // Target Achieved State
        if (barEl) {
            barEl.classList.remove('bg-brand');
            barEl.classList.add('bg-emerald-500');
        }
        if (pctEl) {
            pctEl.classList.remove('text-gray-900');
            pctEl.classList.add('text-emerald-600');
        }
        const surplus = totalSalesVal - targetVal;
        if (remEl) {
            remEl.innerHTML = `<span class="text-emerald-600 font-bold">🎉 Target Achieved!</span> (+${formatCurrency(surplus)})`;
        }
    } else {
        // In Progress State
        if (barEl) {
            barEl.classList.remove('bg-emerald-500');
            barEl.classList.add('bg-brand');
        }
        if (pctEl) {
            pctEl.classList.remove('text-emerald-600');
            pctEl.classList.add('text-gray-900');
        }
        if (remEl) {
            remEl.textContent = `${formatCurrency(remaining)} remaining`;
        }
    }
}

document.getElementById('target-input')?.addEventListener('change', updateTargetStatus);
document.getElementById('target-input')?.addEventListener('input', updateTargetStatus);

// Metric Switch Controls
document.getElementById('toggle-sales')?.addEventListener('click', () => setMetric('sales'));
document.getElementById('toggle-qty')?.addEventListener('click', () => setMetric('qty'));
document.getElementById('toggle-cases')?.addEventListener('click', () => setMetric('cases'));

function setMetric(metric) {
    activeChartMetric = metric;

    const btnSales = document.getElementById('toggle-sales');
    const btnQty = document.getElementById('toggle-qty');
    const btnCases = document.getElementById('toggle-cases');

    [btnSales, btnQty, btnCases].forEach(btn => {
        if (btn) btn.className = "px-3 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    });

    if (metric === 'sales' && btnSales) btnSales.className = "px-3 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
    if (metric === 'qty' && btnQty) btnQty.className = "px-3 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
    if (metric === 'cases' && btnCases) btnCases.className = "px-3 py-1.5 text-xs font-semibold rounded-md transition-all bg-purple-600 text-white";

    renderProductBreakdown();
    renderLeaderboard();
    updateTrendChart();
}

// Leaderboard Tab Selection
document.getElementById('tab-customers')?.addEventListener('click', () => {
    activeLeaderboardTab = 'customers';
    document.getElementById('tab-customers').className = "px-3 py-1 text-xs font-bold rounded-md transition-all bg-brand text-white";
    document.getElementById('tab-brands').className = "px-3 py-1 text-xs font-bold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderLeaderboard();
});

document.getElementById('tab-brands')?.addEventListener('click', () => {
    activeLeaderboardTab = 'brands';
    document.getElementById('tab-brands').className = "px-3 py-1 text-xs font-bold rounded-md transition-all bg-brand text-white";
    document.getElementById('tab-customers').className = "px-3 py-1 text-xs font-bold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderLeaderboard();
});

// Render Product Breakdown List
function renderProductBreakdown() {
    const listEl = document.getElementById('items-scroll-list');
    if (!listEl) return;

    const prods = Object.entries(globalProductData).map(([name, data]) => ({ name, ...data }));

    prods.sort((a, b) => {
        if (activeChartMetric === 'sales') return b.sales - a.sales;
        if (activeChartMetric === 'cases') return b.cases - a.cases;
        return b.qty - a.qty;
    });

    const maxValue = prods.length > 0 ? (
        activeChartMetric === 'sales' ? prods[0].sales : 
        activeChartMetric === 'cases' ? prods[0].cases : prods[0].qty
    ) : 1;

    let barColor = 'bg-brand';
    if (activeChartMetric === 'qty') barColor = 'bg-emerald-600';
    if (activeChartMetric === 'cases') barColor = 'bg-purple-600';

    let html = '';
    prods.forEach(prod => {
        const val = activeChartMetric === 'sales' ? prod.sales : (activeChartMetric === 'cases' ? prod.cases : prod.qty);
        const formattedVal = activeChartMetric === 'sales' ? formatCurrency(val) : `${formatNumber(val)} ${activeChartMetric === 'cases' ? 'cases' : 'units'}`;
        const pctWidth = Math.max((val / maxValue) * 100, 1.5);

        html += `
            <div class="bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5">
                <div class="flex items-center justify-between w-full">
                    <p class="text-sm font-semibold text-gray-800 truncate">${prod.name}</p>
                    <span class="text-xs font-extrabold text-gray-900 ml-2 whitespace-nowrap">${formattedVal}</span>
                </div>
                <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                    <div class="h-full ${barColor} rounded-full" style="width: ${pctWidth}%"></div>
                </div>
            </div>`;
    });

    listEl.innerHTML = html || `<p class="text-xs text-gray-400 py-4 text-center">No products mapped.</p>`;
}

// Render Leaderboards
function renderLeaderboard() {
    const listEl = document.getElementById('leaderboard-list');
    const titleEl = document.getElementById('leaderboard-title');
    const subtextEl = document.getElementById('leaderboard-subtext');
    if (!listEl) return;

    const isSales = activeChartMetric === 'sales';
    const isCases = activeChartMetric === 'cases';

    let barColor = 'bg-brand';
    if (activeChartMetric === 'qty') barColor = 'bg-emerald-600';
    if (activeChartMetric === 'cases') barColor = 'bg-purple-600';

    const getVal = (item) => isSales ? item.sales : (isCases ? item.cases : item.qty);
    const formatVal = (val) => isSales ? formatCurrency(val) : `${formatNumber(val)} ${isCases ? 'cases' : 'units'}`;

    if (activeLeaderboardTab === 'customers') {
        const totalCusts = Object.keys(globalCustomerData).length;
        titleEl.textContent = `All Customers (${formatNumber(totalCusts)})`;
        subtextEl.textContent = `Ranked by ${isSales ? 'Revenue' : (isCases ? 'Cases' : 'Units')} (Click to view profile)`;

        const custs = Object.entries(globalCustomerData).map(([name, data]) => ({ name, ...data }));
        custs.sort((a, b) => getVal(b) - getVal(a));
        const maxVal = custs.length > 0 ? getVal(custs[0]) : 1;

        let html = '';
        custs.forEach((cust, idx) => {
            const rank = idx + 1;
            const val = getVal(cust);
            const pctWidth = Math.max((val / maxVal) * 100, 1.5);

            let rankBadge = 'bg-gray-100 text-gray-600 font-medium';
            if (rank === 1) rankBadge = 'bg-amber-400 text-amber-950 font-bold';
            else if (rank === 2) rankBadge = 'bg-slate-300 text-slate-900 font-bold';
            else if (rank === 3) rankBadge = 'bg-orange-300 text-orange-950 font-bold';

            html += `
                <button onclick="openCustomerProfile('${encodeURIComponent(cust.name)}')" 
                        class="w-full text-left bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5 hover:shadow-md hover:border-gray-300 transition focus:outline-none cursor-pointer">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadge}">${rank}</span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${cust.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2 flex items-center gap-1.5">
                            ${formatVal(val)}
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-3 w-3 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 5l7 7-7 7" /></svg>
                        </span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${barColor} rounded-full" style="width: ${pctWidth}%"></div>
                    </div>
                </button>`;
        });
        listEl.innerHTML = html || `<p class="text-xs text-gray-400 py-4 text-center">No customers found.</p>`;

    } else {
        const totalBrands = Object.keys(globalBrandData).length;
        titleEl.textContent = `Sub Brand Performance (${formatNumber(totalBrands)})`;
        subtextEl.textContent = `Ranked by ${isSales ? 'Revenue' : (isCases ? 'Cases' : 'Units')}`;

        const brands = Object.entries(globalBrandData).map(([name, data]) => ({ name, ...data }));
        brands.sort((a, b) => getVal(b) - getVal(a));
        const maxVal = brands.length > 0 ? getVal(brands[0]) : 1;

        let html = '';
        brands.forEach((brand, idx) => {
            const rank = idx + 1;
            const val = getVal(brand);
            const pctWidth = Math.max((val / maxVal) * 100, 1.5);

            let rankBadge = 'bg-gray-100 text-gray-600 font-medium';
            if (rank === 1) rankBadge = 'bg-amber-400 text-amber-950 font-bold';
            else if (rank === 2) rankBadge = 'bg-slate-300 text-slate-900 font-bold';
            else if (rank === 3) rankBadge = 'bg-orange-300 text-orange-950 font-bold';

            html += `
                <div class="w-full bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadge}">${rank}</span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${brand.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2">${formatVal(val)}</span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${barColor} rounded-full" style="width: ${pctWidth}%"></div>
                    </div>
                </div>`;
        });
        listEl.innerHTML = html || `<p class="text-xs text-gray-400 py-4 text-center">No sub-brands found.</p>`;
    }
}

// Render Monthly Trend Chart
function updateTrendChart() {
    const canvas = document.getElementById('trendChart');
    if (!canvas) return;

    const labels = Object.keys(globalMonthlyData);
    const dataPoints = labels.map(m => {
        if (activeChartMetric === 'sales') return globalMonthlyData[m].sales;
        if (activeChartMetric === 'cases') return globalMonthlyData[m].cases;
        return globalMonthlyData[m].qty;
    });

    let lineColor = '#6C1527';
    let bgColor = 'rgba(108, 21, 39, 0.1)';
    if (activeChartMetric === 'qty') { lineColor = '#059669'; bgColor = 'rgba(5, 150, 105, 0.1)'; }
    if (activeChartMetric === 'cases') { lineColor = '#9333EA'; bgColor = 'rgba(147, 51, 234, 0.1)'; }

    if (trendChartInstance) {
        trendChartInstance.destroy();
    }

    trendChartInstance = new Chart(canvas, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: activeChartMetric === 'sales' ? 'Net Total (PHP)' : (activeChartMetric === 'cases' ? 'Cases Sold' : 'Units Sold'),
                data: dataPoints,
                borderColor: lineColor,
                backgroundColor: bgColor,
                fill: true,
                tension: 0.3,
                borderWidth: 2,
                pointRadius: 4,
                pointHoverRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { display: false }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        callback: (value) => activeChartMetric === 'sales' ? '₱' + (value / 1000).toFixed(0) + 'k' : formatNumber(value)
                    }
                }
            }
        }
    });
}

// Customer Drawer Functionality
window.openCustomerProfile = function(encodedName) {
    const custName = decodeURIComponent(encodedName);
    const custData = globalCustomerData[custName];
    if (!custData) return;

    document.getElementById('drawer-cust-name').textContent = custName;
    document.getElementById('drawer-cust-address').textContent = custData.address || "No Address Recorded";
    document.getElementById('drawer-cust-sales').textContent = formatCurrency(custData.sales);
    document.getElementById('drawer-cust-qty').textContent = formatNumber(custData.qty);
    document.getElementById('drawer-cust-cases').textContent = formatNumber(custData.cases);
    document.getElementById('drawer-cust-orders').textContent = `${formatNumber(custData.orders)} transactions`;

    const parsedFirst = parseExcelDate(custData.firstDate);
    const parsedLast = parseExcelDate(custData.lastDate);

    document.getElementById('drawer-cust-first').textContent = parsedFirst ? parsedFirst.toLocaleDateString() : 'N/A';
    document.getElementById('drawer-cust-last').textContent = parsedLast ? parsedLast.toLocaleDateString() : 'N/A';

    const tbody = document.getElementById('drawer-cust-products');
    if (tbody) {
        let rowsHtml = '';
        Object.entries(custData.products).forEach(([pName, pStats]) => {
            rowsHtml += `
                <tr class="hover:bg-gray-50/80">
                    <td class="px-4 py-2.5 font-medium text-gray-800">${pName}</td>
                    <td class="px-4 py-2.5 text-right font-semibold">${formatNumber(pStats.qty)}</td>
                    <td class="px-4 py-2.5 text-right font-semibold text-purple-700">${formatNumber(pStats.cases)}</td>
                    <td class="px-4 py-2.5 text-right font-bold text-gray-900">${formatCurrency(pStats.sales)}</td>
                </tr>`;
        });
        tbody.innerHTML = rowsHtml || `<tr><td colspan="4" class="p-4 text-center text-xs text-gray-400">No products mapped for account.</td></tr>`;
    }

    document.getElementById('customer-drawer')?.classList.remove('hidden');
};

// Close Drawer Listeners
document.getElementById('close-drawer-btn')?.addEventListener('click', closeDrawer);
document.getElementById('close-drawer-backdrop')?.addEventListener('click', closeDrawer);

function closeDrawer() {
    document.getElementById('customer-drawer')?.classList.add('hidden');
}

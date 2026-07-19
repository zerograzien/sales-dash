// DOM References
const fileInput = document.getElementById('excel-file');
const dashboard = document.getElementById('dashboard');
const emptyState = document.getElementById('empty-state');

// KPI Card References
const kpiSales = document.getElementById('kpi-sales');
const kpiSalesSubtext = document.getElementById('kpi-sales-subtext');
const kpiTransactions = document.getElementById('kpi-transactions');
const kpiItems = document.getElementById('kpi-items');
const kpiItemsSubtext = document.getElementById('kpi-items-subtext');

// Target Tracker Elements
const targetInput = document.getElementById('target-input');
const kpiTargetPct = document.getElementById('kpi-target-pct');
const kpiTargetBar = document.getElementById('kpi-target-bar');
const kpiTargetRem = document.getElementById('kpi-target-rem');

// Metric/Tab Controls
const toggleSalesBtn = document.getElementById('toggle-sales');
const toggleQtyBtn = document.getElementById('toggle-qty');
const itemsScrollList = document.getElementById('items-scroll-list');
const leaderboardList = document.getElementById('leaderboard-list');
const leaderboardTitle = document.getElementById('leaderboard-title');
const leaderboardSubtext = document.getElementById('leaderboard-subtext');
const tabCustomers = document.getElementById('tab-customers');
const tabBrands = document.getElementById('tab-brands');

// Slide-over Drawer References
const customerDrawer = document.getElementById('customer-drawer');
const drawerCustName = document.getElementById('drawer-cust-name');
const drawerCustAddress = document.getElementById('drawer-cust-address');
const drawerCustSales = document.getElementById('drawer-cust-sales');
const drawerCustQty = document.getElementById('drawer-cust-qty');
const drawerCustOrders = document.getElementById('drawer-cust-orders');
const drawerCustFirst = document.getElementById('drawer-cust-first');
const drawerCustLast = document.getElementById('drawer-cust-last');
const drawerCustProducts = document.getElementById('drawer-cust-products');
const closeDrawerBtn = document.getElementById('close-drawer-btn');
const closeDrawerBackdrop = document.getElementById('close-drawer-backdrop');

// Global Datasets & Instances
let globalHeaders = [];
let globalData = [];
let globalItemData = {}; 
let globalCustomerData = {}; 
let globalBrandData = {}; 
let cachedTotalSales = 0; 
let trendChartInstance = null; 

// Global Trend Stores
let globalChartLabels = [];
let globalSalesTrendData = [];
let globalQtyTrendData = [];

let activeChartMetric = 'sales'; // 'sales' | 'qty'
let activeLeaderboardTab = 'customers'; // 'customers' | 'brands'

// Cached Core Column Keys (Prevents lookup overhead inside loops)
let salesCol = null;
let qtyCol = null;
let itemCol = null;
let customerCol = null;
let brandCol = null;
let dateCol = null;

// 1. File Upload Logic (Optimized Parsing Speed)
fileInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        
        // SPEED OPTIMIZATION: Ignore styles, formulas, and string formatting structures
        const workbook = XLSX.read(data, { 
            type: 'array',
            cellDates: true,   // Parse dates natively straight out of sheet engine
            cellStyles: true, // Huge memory/speed improvement
            cellFormulas: false,
            cellNF: false 
        });
        
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: true });

        if (jsonData.length === 0) {
            alert("The uploaded spreadsheet is empty.");
            return;
        }

        const rowMetadata = worksheet['!rows'] || [];
        globalData = jsonData.filter((row, index) => {
            const excelRowMeta = rowMetadata[index + 1];
            if (excelRowMeta && (excelRowMeta.hidden || excelRowMeta.h)) return false; 
            return true;
        });

        globalHeaders = Object.keys(jsonData[0]);
        
        // Run column detection exactly once before calculating
        cacheColumnPositions();

        emptyState.classList.add('hidden');
        dashboard.classList.remove('hidden');

        setupChartToggleUI();
        calculateKPIs();
    };
});

// Cache target column indices cleanly upfront
function cacheColumnPositions() {
    salesCol = findColumnByNames(['net total', 'net_total', 'total sales', 'sales', 'total']);
    qtyCol = findColumnByNames(['quantity sold', 'qty sold', 'quantity', 'qty', 'items sold', 'units sold']);
    itemCol = findColumnByNames(['item description', 'item_description', 'description', 'item', 'product name', 'product']);
    customerCol = findColumnByNames(['customer', 'customer name', 'client', 'buyer', 'account']);
    brandCol = findColumnByNames(['sub brand', 'sub_brand', 'subbrand', 'brand']);
    dateCol = findColumnByNames(['date', 'transaction date', 'order date']);
}

// Helper to inject chart filter buttons dynamically
function setupChartToggleUI() {
    const chartHeadingBlock = document.querySelector('#trendChart').parentElement.previousElementSibling;
    if (chartHeadingBlock && !document.getElementById('chart-toggle-container')) {
        const toggleDiv = document.createElement('div');
        toggleDiv.id = 'chart-toggle-container';
        toggleDiv.className = 'inline-flex rounded-lg border border-gray-200 bg-white p-1 shadow-sm mt-3 sm:mt-0';
        toggleDiv.innerHTML = `
            <button id="chart-toggle-sales" class="px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white">Net Total (PHP)</button>
            <button id="chart-toggle-qty" class="px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900">Quantity Sold</button>
        `;
        
        chartHeadingBlock.classList.add('flex', 'flex-col', 'sm:flex-row', 'justify-between', 'items-start', 'sm:items-center');
        chartHeadingBlock.appendChild(toggleDiv);

        document.getElementById('chart-toggle-sales').addEventListener('click', function() {
            activeChartMetric = 'sales';
            this.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
            document.getElementById('chart-toggle-qty').className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
            
            toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
            toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
            
            updateChartMetricView();
            renderScrollItemsList();
            renderLeaderboard();
        });

        document.getElementById('chart-toggle-qty').addEventListener('click', function() {
            activeChartMetric = 'qty';
            this.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
            document.getElementById('chart-toggle-sales').className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
            
            toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
            toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
            
            updateChartMetricView();
            renderScrollItemsList();
            renderLeaderboard();
        });
    }
}

// 2. Aggregate Calculations Engine
function calculateKPIs() {
    cachedTotalSales = 0;
    let totalItems = 0;
    let totalTransactions = 0;
    
    globalItemData = {}; 
    globalCustomerData = {}; 
    globalBrandData = {}; 
    let dateTrendMap = {}; 

    if (salesCol) kpiSalesSubtext.innerHTML = `Sum of <strong class="text-brand font-bold">"${salesCol}"</strong>`;
    if (qtyCol) kpiItemsSubtext.innerHTML = `Sum of <strong class="text-emerald-700 font-bold">"${qtyCol}"</strong>`;

    const dataLength = globalData.length;
    for (let i = 0; i < dataLength; i++) {
        const row = globalData[i];
        if (!row) continue;

        // Fast path summary checker (Look at specific columns instead of looping all properties)
        if (customerCol && row[customerCol]) {
            const checkVal = String(row[customerCol]).toLowerCase();
            if (checkVal.includes('total') || checkVal.includes('summary')) continue;
        }

        totalTransactions++;
        const saleValue = salesCol ? parseNumericValue(row[salesCol]) : 0;
        const qtyValue = qtyCol ? parseNumericValue(row[qtyCol]) : 0;

        cachedTotalSales += saleValue;
        totalItems += qtyValue;

        // Dynamic Monthly Grouping
        if (dateCol && row[dateCol]) {
            let parsedDate = parseExcelDate(row[dateCol]);
            if (parsedDate) {
                let year = parsedDate.getFullYear();
                let month = String(parsedDate.getMonth() + 1).padStart(2, '0');
                let monthKey = `${year}-${month}`;
                
                if (!dateTrendMap[monthKey]) {
                    dateTrendMap[monthKey] = { sales: 0, qty: 0 };
                }
                dateTrendMap[monthKey].sales += saleValue;
                dateTrendMap[monthKey].qty += qtyValue;
            }
        }

        if (itemCol && row[itemCol]) {
            const itemName = String(row[itemCol]).trim();
            if (!globalItemData[itemName]) globalItemData[itemName] = { sales: 0, qty: 0 };
            globalItemData[itemName].sales += saleValue;
            globalItemData[itemName].qty += qtyValue;
        }

        if (customerCol && row[customerCol]) {
            const customerName = String(row[customerCol]).trim();
            if (!globalCustomerData[customerName]) globalCustomerData[customerName] = { sales: 0, qty: 0 };
            globalCustomerData[customerName].sales += saleValue;
            globalCustomerData[customerName].qty += qtyValue;
        }

        if (brandCol && row[brandCol]) {
            const brandName = String(row[brandCol]).trim();
            if (!globalBrandData[brandName]) globalBrandData[brandName] = { sales: 0, qty: 0 };
            globalBrandData[brandName].sales += saleValue;
            globalBrandData[brandName].qty += qtyValue;
        }
    }

    kpiSales.textContent = formatCurrency(cachedTotalSales);
    kpiTransactions.textContent = totalTransactions.toLocaleString();
    kpiItems.textContent = totalItems.toLocaleString();

    // Cache historical timeline arrays
    const sortedMonthKeys = Object.keys(dateTrendMap).sort();
    globalChartLabels = sortedMonthKeys.map(key => {
        const [year, month] = key.split('-');
        const dateObj = new Date(year, parseInt(month) - 1, 1);
        return dateObj.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
    });
    globalSalesTrendData = sortedMonthKeys.map(key => dateTrendMap[key].sales);
    globalQtyTrendData = sortedMonthKeys.map(key => dateTrendMap[key].qty);

    updateTargetMetrics();
    updateChartMetricView(); 

    // SPEED OPTIMIZATION: Yield threads using requestAnimationFrame to clear DOM bottleneck
    requestAnimationFrame(() => {
        renderScrollItemsList();
        requestAnimationFrame(() => {
            renderLeaderboard();
        });
    });
}

// 3. Update Chart Dataset Dynamically
function updateChartMetricView() {
    const ctx = document.getElementById('trendChart').getContext('2d');
    
    if (trendChartInstance) {
        trendChartInstance.destroy();
    }

    const isSales = activeChartMetric === 'sales';
    
    const activeDataset = isSales ? {
        label: 'Net Total (PHP)',
        data: globalSalesTrendData,
        borderColor: '#6C1527',
        backgroundColor: 'rgba(108, 21, 39, 0.04)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.2
    } : {
        label: 'Quantity Sold (Units)',
        data: globalQtyTrendData,
        borderColor: '#059669',
        backgroundColor: 'rgba(5, 150, 105, 0.04)',
        fill: true,
        borderWidth: 2.5,
        pointRadius: 4,
        pointHoverRadius: 6,
        tension: 0.2
    };

    trendChartInstance = new Chart(ctx, {
        type: 'line',
        data: {
            labels: globalChartLabels,
            datasets: [activeDataset]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                x: {
                    grid: { display: false },
                    ticks: { font: { size: 11 } }
                },
                y: {
                    type: 'linear',
                    display: true,
                    title: { 
                        display: true, 
                        text: isSales ? 'Net Total Revenue' : 'Units Volume', 
                        font: { weight: 'bold', size: 11 } 
                    },
                    ticks: {
                        callback: val => isSales ? '₱' + val.toLocaleString(undefined, {maximumFractionDigits: 0}) : val.toLocaleString(),
                        font: { size: 10 }
                    }
                }
            },
            plugins: {
                legend: { display: false }, 
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    padding: 10,
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) label += ': ';
                            label += isSales ? '₱' + context.parsed.y.toLocaleString(undefined, { minimumFractionDigits: 2 }) : context.parsed.y.toLocaleString() + ' units';
                            return label;
                        }
                    }
                }
            }
        }
    });
}

// 4. Dynamic Manual Target Calculations
function updateTargetMetrics() {
    const rawTargetValue = targetInput.value.replace(/,/g, '');
    const numericTarget = parseFloat(rawTargetValue) || 0;

    if (numericTarget <= 0) {
        kpiTargetPct.textContent = "0.00%";
        kpiTargetBar.style.width = "0%";
        kpiTargetRem.textContent = "Set valid target threshold";
        return;
    }

    const percentage = (cachedTotalSales / numericTarget) * 100;
    kpiTargetPct.textContent = percentage.toFixed(2) + "%";
    kpiTargetBar.style.width = Math.min(percentage, 100) + "%";

    if (percentage >= 100) {
        kpiTargetBar.className = "h-full bg-emerald-600 rounded-full transition-all duration-500";
        kpiTargetRem.textContent = "Target Goal Achieved! 🎉";
    } else {
        kpiTargetBar.className = "h-full bg-brand rounded-full transition-all duration-500";
        const remaining = numericTarget - cachedTotalSales;
        kpiTargetRem.textContent = `${formatCurrency(remaining)} remaining`;
    }
}

targetInput.addEventListener('keyup', function() {
    let cleanValue = this.value.replace(/[^0-9.]/g, '');
    if (cleanValue) {
        let parts = cleanValue.split('.');
        parts[0] = parseInt(parts[0], 10).toLocaleString('en-US');
        this.value = parts.join('.');
    }
    updateTargetMetrics();
});

// 5. Scrollable Item Performance Breakdown (Optimized via Top-30 Slice)
function renderScrollItemsList() {
    if (!globalItemData || Object.keys(globalItemData).length === 0) {
        itemsScrollList.innerHTML = `<div class="text-center py-8 text-gray-400 text-sm">No uploaded data items mapped.</div>`;
        return;
    }

    const isSales = activeChartMetric === 'sales';
    const listColor = isSales ? 'bg-brand' : 'bg-emerald-600';

    const itemsArray = Object.entries(globalItemData).map(([name, data]) => ({
        name, sales: data.sales, qty: data.qty
    }));

    itemsArray.sort((a, b) => isSales ? b.sales - a.sales : b.qty - a.qty);
    
    // SPEED OPTIMIZATION: Render only top 30 elements to eliminate large scale DOM payload drops
    const visibleItems = itemsArray.slice(0, 30);
    const maxItemValue = itemsArray.length > 0 ? (isSales ? itemsArray[0].sales : itemsArray[0].qty) : 1;

    let html = '';
    const itemsCount = visibleItems.length;
    for (let i = 0; i < itemsCount; i++) {
        const item = visibleItems[i];
        const rank = i + 1;
        const currentVal = isSales ? item.sales : item.qty;
        const percentWidth = Math.max((currentVal / maxItemValue) * 100, 1.5);
        const displayMetricValue = isSales ? formatCurrency(currentVal) : currentVal.toLocaleString() + ' units';

        html += `
            <div class="w-full bg-white p-3 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1">
                <div class="flex items-center justify-between w-full text-xs">
                    <div class="flex items-center gap-2 min-w-0">
                        <span class="text-gray-400 font-bold w-5">#${rank}</span>
                        <p class="font-semibold text-gray-800 truncate" title="${item.name}">${item.name}</p>
                    </div>
                    <span class="font-bold text-gray-900 whitespace-nowrap pl-2">${displayMetricValue}</span>
                </div>
                <div class="w-full bg-gray-100 h-1 rounded-full overflow-hidden">
                    <div class="h-full ${listColor} rounded-full" style="width: ${percentWidth}%"></div>
                </div>
            </div>`;
    }
    itemsScrollList.innerHTML = html;
}

// 6. HTML Leaderboards Render Logic
function renderLeaderboard() {
    const isSales = activeChartMetric === 'sales';
    const progressBarColor = isSales ? 'bg-brand' : 'bg-emerald-600';

    if (activeLeaderboardTab === 'customers') {
        leaderboardTitle.textContent = "Top 15 Customers";
        leaderboardSubtext.textContent = `Ranked by ${isSales ? 'Net Total' : 'Units Bought'} (Click to open profile)`;

        if (!globalCustomerData || Object.keys(globalCustomerData).length === 0) {
            leaderboardList.innerHTML = `<div class="text-center py-8 text-gray-400 text-sm">No customer data rows mapped.</div>`;
            return;
        }

        const customersArray = Object.entries(globalCustomerData).map(([name, data]) => ({
            name, sales: data.sales, qty: data.qty
        }));
        customersArray.sort((a, b) => isSales ? b.sales - a.sales : b.qty - a.qty);
        const topCustomers = customersArray.slice(0, 15);
        const maxValue = topCustomers.length > 0 ? (isSales ? topCustomers[0].sales : topCustomers[0].qty) : 1;

        let html = '';
        topCustomers.forEach((cust, index) => {
            const rank = index + 1;
            const currentVal = isSales ? cust.sales : cust.qty;
            const percentWidth = Math.max((currentVal / maxValue) * 100, 2);

            let rankBadgeClass = 'bg-gray-200 text-gray-700 font-medium';
            if (rank === 1) rankBadgeClass = 'bg-amber-400 text-amber-950 font-bold shadow-sm';
            else if (rank === 2) rankBadgeClass = 'bg-slate-300 text-slate-900 font-bold shadow-sm';
            else if (rank === 3) rankBadgeClass = 'bg-orange-300 text-orange-950 font-bold shadow-sm';

            html += `
                <button onclick="openCustomerProfile('${encodeURIComponent(cust.name)}')" 
                        class="w-full text-left bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5 transition duration-150 hover:shadow-md hover:border-gray-300 hover:scale-[1.01] focus:outline-none cursor-pointer">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadgeClass}">${rank}</span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${cust.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2 flex items-center gap-1.5">
                            ${isSales ? formatCurrency(currentVal) : currentVal.toLocaleString() + ' units'}
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-3 w-3 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 5l7 7-7 7" /></svg>
                        </span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${progressBarColor} rounded-full" style="width: ${percentWidth}%"></div>
                    </div>
                </button>`;
        });
        leaderboardList.innerHTML = html;

    } else if (activeLeaderboardTab === 'brands') {
        leaderboardTitle.textContent = "Sub Brand Performance";
        leaderboardSubtext.textContent = `Ranked by ${isSales ? 'Revenue metrics' : 'Quantities sold'}`;

        if (!globalBrandData || Object.keys(globalBrandData).length === 0) {
            leaderboardList.innerHTML = `<div class="text-center py-8 text-gray-400 text-sm">No "Sub Brand" column mapped.</div>`;
            return;
        }

        const brandsArray = Object.entries(globalBrandData).map(([name, data]) => ({
            name, sales: data.sales, qty: data.qty
        }));
        brandsArray.sort((a, b) => isSales ? b.sales - a.sales : b.qty - a.qty);
        const topBrands = brandsArray.slice(0, 15);
        const maxValue = topBrands.length > 0 ? (isSales ? topBrands[0].sales : topBrands[0].qty) : 1;

        let html = '';
        topBrands.forEach((brand, index) => {
            const rank = index + 1;
            const currentVal = isSales ? brand.sales : brand.qty;
            const percentWidth = Math.max((currentVal / maxValue) * 100, 2);

            let rankBadgeClass = 'bg-gray-200 text-gray-700 font-medium';
            if (rank === 1) rankBadgeClass = 'bg-amber-400 text-amber-950 font-bold shadow-sm';
            else if (rank === 2) rankBadgeClass = 'bg-slate-300 text-slate-900 font-bold shadow-sm';
            else if (rank === 3) rankBadgeClass = 'bg-orange-300 text-orange-950 font-bold shadow-sm';

            html += `
                <div class="w-full bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadgeClass}">${rank}</span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${brand.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2">
                            ${isSales ? formatCurrency(currentVal) : brand.qty.toLocaleString() + ' units'}
                        </span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${progressBarColor} rounded-full" style="width: ${percentWidth}%"></div>
                    </div>
                </div>`;
        });
        leaderboardList.innerHTML = html;
    }
}

// 7. Deep Account Drawer Logic Handler
function openCustomerProfile(encodedName) {
    const customerName = decodeURIComponent(encodedName);
    const addressCol = findColumnByNames(['customer address', 'customer_address', 'address', 'client address', 'location']);

    const customerTransactions = globalData.filter(row => row && customerCol && String(row[customerCol]).trim() === customerName);

    let totalSales = 0, totalQty = 0, resolvedAddress = '';
    let dates = [], productsMap = {};

    customerTransactions.forEach(row => {
        const sale = salesCol ? parseNumericValue(row[salesCol]) : 0;
        const qty = qtyCol ? parseNumericValue(row[qtyCol]) : 0;
        totalSales += sale; totalQty += qty;

        if (addressCol && row[addressCol] && !resolvedAddress) resolvedAddress = String(row[addressCol]).trim();
        if (dateCol && row[dateCol]) {
            let parsedDate = parseExcelDate(row[dateCol]);
            if (parsedDate) dates.push(parsedDate);
        }
        if (itemCol && row[itemCol]) {
            const product = String(row[itemCol]).trim();
            if (!productsMap[product]) productsMap[product] = { sales: 0, qty: 0 };
            productsMap[product].sales += sale;
            productsMap[product].qty += qty;
        }
    });

    drawerCustName.textContent = customerName;
    drawerCustAddress.textContent = resolvedAddress || "No address on record";
    drawerCustSales.textContent = formatCurrency(totalSales);
    drawerCustQty.textContent = totalQty.toLocaleString() + ' units';
    drawerCustOrders.textContent = customerTransactions.length.toLocaleString() + ' recorded transactions';

    if (dates.length > 0) {
        dates.sort((a, b) => a - b);
        const formatOpt = { year: 'numeric', month: 'short', day: 'numeric' };
        drawerCustFirst.textContent = dates[0].toLocaleDateString('en-US', formatOpt);
        drawerCustLast.textContent = dates[dates.length - 1].toLocaleDateString('en-US', formatOpt);
    } else {
        drawerCustFirst.textContent = drawerCustLast.textContent = 'N/A';
    }

    const sortedProducts = Object.entries(productsMap).sort((a, b) => b[1].sales - a[1].sales);
    let productsHTML = '';
    if (sortedProducts.length === 0) {
        productsHTML = `<tr><td colspan="3" class="px-4 py-4 text-center text-xs text-gray-400">No items detected.</td></tr>`;
    } else {
        sortedProducts.forEach(([prodName, metrics]) => {
            productsHTML += `
                <tr class="hover:bg-gray-50/50">
                    <td class="px-4 py-2.5 text-xs font-semibold text-gray-900 truncate max-w-[240px]">${prodName}</td>
                    <td class="px-4 py-2.5 text-xs text-right font-medium text-gray-500">${metrics.qty.toLocaleString()}</td>
                    <td class="px-4 py-2.5 text-xs text-right font-bold text-gray-900">${formatCurrency(metrics.sales)}</td>
                </tr>`;
        });
    }
    drawerCustProducts.innerHTML = productsHTML;
    customerDrawer.classList.remove('hidden');
    document.body.classList.add('overflow-hidden');
}

function closeDrawer() {
    customerDrawer.classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}
closeDrawerBtn.addEventListener('click', closeDrawer);
closeDrawerBackdrop.addEventListener('click', closeDrawer);

// Side Panels Metric Toggle Event Hooks
toggleSalesBtn.addEventListener('click', () => {
    activeChartMetric = 'sales';
    toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
    toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    
    const cSales = document.getElementById('chart-toggle-sales');
    const cQty = document.getElementById('chart-toggle-qty');
    if(cSales && cQty) {
        cSales.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
        cQty.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    }
    
    updateChartMetricView();
    renderScrollItemsList();
    renderLeaderboard();
});

toggleQtyBtn.addEventListener('click', () => {
    activeChartMetric = 'qty';
    toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
    toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    
    const cSales = document.getElementById('chart-toggle-sales');
    const cQty = document.getElementById('chart-toggle-qty');
    if(cSales && cQty) {
        cQty.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
        cSales.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    }
    
    updateChartMetricView();
    renderScrollItemsList();
    renderLeaderboard();
});

tabCustomers.addEventListener('click', () => {
    activeLeaderboardTab = 'customers';
    tabCustomers.className = "px-3 py-1 text-xs font-bold rounded-md transition-all bg-brand text-white";
    tabBrands.className = "px-3 py-1 text-xs font-bold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderLeaderboard();
});

tabBrands.addEventListener('click', () => {
    activeLeaderboardTab = 'brands';
    tabBrands.className = "px-3 py-1 text-xs font-bold rounded-md transition-all bg-brand text-white";
    tabCustomers.className = "px-3 py-1 text-xs font-bold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderLeaderboard();
});

// Utilities Parsing Functions
function formatCurrency(val) {
    return new Intl.NumberFormat('en-US', { 
        style: 'currency', currency: 'PHP', minimumFractionDigits: 2, maximumFractionDigits: 2 
    }).format(val);
}

function findColumnByNames(possibleNames) {
    const exactMatch = globalHeaders.find(header => {
        const normalized = header.toLowerCase().replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();
        return possibleNames.some(name => normalized === name);
    });
    if (exactMatch) return exactMatch;
    return globalHeaders.find(header => {
        const normalized = header.toLowerCase().replace(/[\r\n\t]/g, ' ').replace(/\s+/g, ' ').trim();
        return possibleNames.some(name => normalized.includes(name));
    }) || null;
}

function parseNumericValue(val) {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    let strVal = String(val).trim();
    let isNegative = strVal.startsWith('-') || (strVal.startsWith('(') && strVal.endsWith(')'));
    const cleaned = strVal.replace(/,/g, '').replace(/\s/g, '').replace(/[^0-9.]/g, '');
    let parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0 : (isNegative ? -parsed : parsed);
}

function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') return new Date((val - 25569) * 86400 * 1000);
    let parsed = Date.parse(String(val).trim());
    return isNaN(parsed) ? null : new Date(parsed);
}

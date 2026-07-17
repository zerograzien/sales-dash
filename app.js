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

// Chart Toggle References
const toggleSalesBtn = document.getElementById('toggle-sales');
const toggleQtyBtn = document.getElementById('toggle-qty');

// Leaderboard References
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

// Global variables to store active dataset
let globalHeaders = [];
let globalData = [];
let globalItemData = {}; 
let globalCustomerData = {}; 
let globalBrandData = {}; // Storage bucket for Sub Brand metrics
let itemsChartInstance = null;

let activeChartMetric = 'sales'; // 'sales' | 'qty'
let activeLeaderboardTab = 'customers'; // 'customers' | 'brands'

// 1. File Upload Handler
fileInput.addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.readAsArrayBuffer(file);
    
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        
        const workbook = XLSX.read(data, { 
            type: 'array',
            cellHTML: false, 
            cellFormula: false, 
            cellStyles: true 
        });
        
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

        if (jsonData.length === 0) {
            alert("The uploaded spreadsheet is empty.");
            return;
        }

        const rowMetadata = worksheet['!rows'] || [];

        globalData = jsonData.filter((row, index) => {
            const excelRowMeta = rowMetadata[index + 1];
            if (excelRowMeta && (excelRowMeta.hidden || excelRowMeta.h)) {
                return false; 
            }
            return true;
        });

        globalHeaders = Object.keys(jsonData[0]);

        emptyState.classList.add('hidden');
        dashboard.classList.remove('hidden');

        calculateKPIs();
    };
});

// 2. Dynamic KPI & Data Aggregator Calculator
function calculateKPIs() {
    let totalSales = 0;
    let totalItems = 0;
    let totalTransactions = 0;
    
    globalItemData = {}; 
    globalCustomerData = {}; 
    globalBrandData = {}; // Reset Brand storage on recalculation

    const salesCol = findColumnByNames(['net total', 'net_total', 'total sales', 'sales', 'total']);
    const qtyCol = findColumnByNames(['quantity sold', 'qty sold', 'quantity', 'qty', 'items sold', 'units sold']);
    const itemCol = findColumnByNames(['item description', 'item_description', 'description', 'item', 'product name', 'product']);
    const customerCol = findColumnByNames(['customer', 'customer name', 'client', 'buyer', 'account']);
    const brandCol = findColumnByNames(['sub brand', 'sub_brand', 'subbrand', 'brand']); // Sub brand identifier

    if (salesCol) {
        kpiSalesSubtext.innerHTML = `Summing column: <strong class="text-brand font-bold">"${salesCol}"</strong>`;
    } else {
        kpiSalesSubtext.textContent = `No matching sales column found ("Net Total")`;
    }

    if (qtyCol) {
        kpiItemsSubtext.innerHTML = `Summing column: <strong class="text-emerald-700 font-bold">"${qtyCol}"</strong>`;
    } else {
        kpiItemsSubtext.textContent = `No matching items column found ("Quantity Sold")`;
    }

    globalData.forEach(row => {
        if (!row || Object.keys(row).length === 0) return;

        // SAFE ROW FILTERING: Only skip true Excel summary/subtotal rows.
        // We ignore cells belonging to Brand, Customer, or Item Description when searching for "total" 
        // to prevent discarding actual data rows like "Total Care Brand".
        let isSummaryRow = false;
        for (const key in row) {
            if (row.hasOwnProperty(key)) {
                // If the current cell is part of our main identifier columns, skip checking it for "total"
                if (key === brandCol || key === customerCol || key === itemCol) {
                    continue;
                }

                const cellValue = String(row[key] || '').toLowerCase().trim();
                if (
                    cellValue === 'total' || 
                    cellValue === 'subtotal' || 
                    cellValue === 'grand total' || 
                    cellValue === 'grand_total' ||
                    cellValue === 'average' ||
                    cellValue === 'summary'
                ) {
                    isSummaryRow = true;
                    break;
                }
            }
        }

        if (isSummaryRow) return;

        totalTransactions++;

        const saleValue = salesCol ? parseNumericValue(row[salesCol]) : 0;
        const qtyValue = qtyCol ? parseNumericValue(row[qtyCol]) : 0;

        totalSales += saleValue;
        totalItems += qtyValue;

        // Group by Item Description
        if (itemCol) {
            const itemName = String(row[itemCol] || 'Unknown Item').trim();
            if (itemName) {
                if (!globalItemData[itemName]) {
                    globalItemData[itemName] = { sales: 0, qty: 0 };
                }
                globalItemData[itemName].sales += saleValue;
                globalItemData[itemName].qty += qtyValue;
            }
        }

        // Group by Customer
        if (customerCol) {
            const customerName = String(row[customerCol] || 'Unknown Customer').trim();
            if (customerName) {
                if (!globalCustomerData[customerName]) {
                    globalCustomerData[customerName] = { sales: 0, qty: 0 };
                }
                globalCustomerData[customerName].sales += saleValue;
                globalCustomerData[customerName].qty += qtyValue;
            }
        }

        // Group by Sub Brand
        if (brandCol) {
            const brandName = String(row[brandCol] || 'Unbranded / Other').trim();
            if (brandName) {
                if (!globalBrandData[brandName]) {
                    globalBrandData[brandName] = { sales: 0, qty: 0 };
                }
                globalBrandData[brandName].sales += saleValue;
                globalBrandData[brandName].qty += qtyValue;
            }
        }
    });

    kpiSales.textContent = new Intl.NumberFormat('en-US', { 
        style: 'currency', 
        currency: 'PHP',
        minimumFractionDigits: 3,
        maximumFractionDigits: 3
    }).format(totalSales);
    
    kpiTransactions.textContent = totalTransactions.toLocaleString();
    kpiItems.textContent = totalItems.toLocaleString();

    renderChart();
    renderLeaderboard();
}

// 3. Horizontal Bar Chart (Top 10 Items)
function renderChart() {
    if (!globalItemData || Object.keys(globalItemData).length === 0) return;

    const dataArray = Object.entries(globalItemData).map(([name, data]) => ({
        name,
        sales: data.sales,
        qty: data.qty
    }));

    if (activeChartMetric === 'sales') {
        dataArray.sort((a, b) => b.sales - a.sales);
    } else {
        dataArray.sort((a, b) => b.qty - a.qty);
    }

    const topData = dataArray.slice(0, 10);
    const labels = topData.map(d => d.name);
    const values = topData.map(d => activeChartMetric === 'sales' ? d.sales : d.qty);

    const ctx = document.getElementById('itemsChart').getContext('2d');
    const isSales = activeChartMetric === 'sales';
    const datasetLabel = isSales ? 'Net Total (PHP)' : 'Quantity Sold';
    const barColor = isSales ? '#6C1527' : '#059669'; 

    if (itemsChartInstance) {
        itemsChartInstance.data.labels = labels;
        itemsChartInstance.data.datasets[0].label = datasetLabel;
        itemsChartInstance.data.datasets[0].data = values;
        itemsChartInstance.data.datasets[0].backgroundColor = barColor;
        itemsChartInstance.data.datasets[0].borderColor = barColor;
        itemsChartInstance.update();
    } else {
        itemsChartInstance = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: labels,
                datasets: [{
                    label: datasetLabel,
                    data: values,
                    backgroundColor: barColor,
                    borderColor: barColor,
                    borderWidth: 1,
                    borderRadius: 6,
                    barThickness: 24
                }]
            },
            options: {
                indexAxis: 'y',
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let val = context.raw || 0;
                                if (activeChartMetric === 'sales') {
                                    return ' Net Total: ₱' + val.toLocaleString(undefined, {minimumFractionDigits: 3, maximumFractionDigits: 3});
                                }
                                return ' Quantity Sold: ' + val.toLocaleString();
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        grid: { color: '#F3F4F6' },
                        ticks: {
                            callback: function(value) {
                                if (activeChartMetric === 'sales') {
                                    return '₱' + value.toLocaleString(undefined, {maximumFractionDigits: 0});
                                }
                                return value.toLocaleString();
                            },
                            color: '#6B7280',
                            font: { size: 11 }
                        }
                    },
                    y: {
                        grid: { display: false },
                        ticks: {
                            color: '#374151',
                            font: { size: 11, weight: '500' }
                        }
                    }
                }
            }
        });
    }
}

// 4. HTML Leaderboards Render Logic (Supports Customers & Sub Brands)
function renderLeaderboard() {
    const isSales = activeChartMetric === 'sales';
    const progressBarColor = isSales ? 'bg-brand' : 'bg-emerald-600';

    // RENDER BRANCH 1: Customers View
    if (activeLeaderboardTab === 'customers') {
        leaderboardTitle.textContent = "Top 15 Customers";
        leaderboardSubtext.textContent = isSales 
            ? 'Ranked by Net Total (Click to open account profile)' 
            : 'Ranked by Quantity Sold (Click to open account profile)';

        if (!globalCustomerData || Object.keys(globalCustomerData).length === 0) {
            leaderboardList.innerHTML = `
                <div class="text-center py-8 text-gray-400 text-sm">
                    No customer column ("Customer") identified in dataset.
                </div>`;
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

            // Consistent Metallic Medals
            let rankBadgeClass = 'bg-gray-200 text-gray-700 font-medium';
            if (rank === 1) rankBadgeClass = 'bg-amber-400 text-amber-950 font-bold shadow-sm'; // Gold
            else if (rank === 2) rankBadgeClass = 'bg-slate-300 text-slate-900 font-bold shadow-sm'; // Silver
            else if (rank === 3) rankBadgeClass = 'bg-orange-300 text-orange-950 font-bold shadow-sm'; // Bronze

            const formattedVal = isSales 
                ? '₱' + currentVal.toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })
                : currentVal.toLocaleString() + ' units';

            html += `
                <button onclick="openCustomerProfile('${encodeURIComponent(cust.name)}')" 
                        class="w-full text-left bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5 transition duration-150 hover:shadow-md hover:border-gray-300 hover:scale-[1.01] active:scale-[0.99] focus:outline-none focus:ring-2 focus:ring-brand/20">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadgeClass}">
                                ${rank}
                            </span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${cust.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2 flex items-center gap-1.5">
                            ${formattedVal}
                            <svg xmlns="http://www.w3.org/2000/svg" class="h-3 w-3 text-gray-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M9 5l7 7-7 7" />
                            </svg>
                        </span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${progressBarColor} rounded-full transition-all duration-500" style="width: ${percentWidth}%"></div>
                    </div>
                </button>`;
        });
        leaderboardList.innerHTML = html;

    // RENDER BRANCH 2: Sub Brands View
    } else if (activeLeaderboardTab === 'brands') {
        leaderboardTitle.textContent = "Sub Brand Performance";
        leaderboardSubtext.textContent = isSales ? 'Ranked by Net Total Revenue' : 'Ranked by Quantity of Units Sold';

        if (!globalBrandData || Object.keys(globalBrandData).length === 0) {
            leaderboardList.innerHTML = `
                <div class="text-center py-8 text-gray-400 text-sm">
                    No "Sub Brand" or "Brand" column identified in dataset.
                </div>`;
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

            // Consistent Metallic Medals
            let rankBadgeClass = 'bg-gray-200 text-gray-700 font-medium';
            if (rank === 1) rankBadgeClass = 'bg-amber-400 text-amber-950 font-bold shadow-sm'; // Gold
            else if (rank === 2) rankBadgeClass = 'bg-slate-300 text-slate-900 font-bold shadow-sm'; // Silver
            else if (rank === 3) rankBadgeClass = 'bg-orange-300 text-orange-950 font-bold shadow-sm'; // Bronze

            const formattedVal = isSales 
                ? '₱' + currentVal.toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })
                : currentVal.toLocaleString() + ' units';

            html += `
                <div class="w-full bg-white p-3.5 rounded-lg border border-gray-150 shadow-sm flex flex-col space-y-1.5">
                    <div class="flex items-center justify-between w-full">
                        <div class="flex items-center gap-3 min-w-0">
                            <span class="w-6 h-6 flex items-center justify-center text-xs rounded-full shrink-0 ${rankBadgeClass}">
                                ${rank}
                            </span>
                            <p class="text-sm font-semibold text-gray-800 truncate">${brand.name}</p>
                        </div>
                        <span class="text-xs font-extrabold text-gray-900 whitespace-nowrap pl-2">
                            ${formattedVal}
                        </span>
                    </div>
                    <div class="w-full bg-gray-100 h-1.5 rounded-full overflow-hidden">
                        <div class="h-full ${progressBarColor} rounded-full transition-all duration-500" style="width: ${percentWidth}%"></div>
                    </div>
                </div>`;
        });
        leaderboardList.innerHTML = html;
    }
}

// 5. Customer Profile Account Drawer Logic
function openCustomerProfile(encodedName) {
    const customerName = decodeURIComponent(encodedName);
    
    // Find reference columns
    const customerCol = findColumnByNames(['customer', 'customer name', 'client', 'buyer', 'account']);
    const addressCol = findColumnByNames(['customer address', 'customer_address', 'address', 'client address', 'location', 'billing address']);
    const salesCol = findColumnByNames(['net total', 'net_total', 'total sales', 'sales', 'total']);
    const qtyCol = findColumnByNames(['quantity sold', 'qty sold', 'quantity', 'qty', 'items sold', 'units sold']);
    const itemCol = findColumnByNames(['item description', 'item_description', 'description', 'item', 'product name', 'product']);
    const dateCol = findColumnByNames(['date', 'transaction date', 'order date', 'posting date']);

    // Filter raw data strictly related to this selected customer profile
    const customerTransactions = globalData.filter(row => {
        if (!row || !customerCol) return false;
        return String(row[customerCol] || '').trim() === customerName;
    });

    let totalSales = 0;
    let totalQty = 0;
    let orderCount = customerTransactions.length;
    let dates = [];
    let productsMap = {};
    let resolvedAddress = '';

    customerTransactions.forEach(row => {
        const sale = salesCol ? parseNumericValue(row[salesCol]) : 0;
        const qty = qtyCol ? parseNumericValue(row[qtyCol]) : 0;

        totalSales += sale;
        totalQty += qty;

        // Extract first valid address found in the rows for this customer
        if (addressCol && row[addressCol] && !resolvedAddress) {
            resolvedAddress = String(row[addressCol]).trim();
        }

        // Extract dates if formatting is compliant
        if (dateCol && row[dateCol]) {
            let parsedDate = parseExcelDate(row[dateCol]);
            if (parsedDate) dates.push(parsedDate);
        }

        // Segment purchases by product descriptions
        if (itemCol) {
            const product = String(row[itemCol] || 'Unknown Item').trim();
            if (product) {
                if (!productsMap[product]) {
                    productsMap[product] = { sales: 0, qty: 0 };
                }
                productsMap[product].sales += sale;
                productsMap[product].qty += qty;
            }
        }
    });

    // Populate profile cards with data
    drawerCustName.textContent = customerName;
    
    // Render dynamic Address or show warning if column is missing
    if (resolvedAddress) {
        drawerCustAddress.textContent = resolvedAddress;
        drawerCustAddress.className = "truncate font-medium text-gray-600";
    } else {
        drawerCustAddress.textContent = "No address on record";
        drawerCustAddress.className = "truncate font-medium text-gray-400 italic";
    }
    
    drawerCustSales.textContent = new Intl.NumberFormat('en-US', { 
        style: 'currency', 
        currency: 'PHP',
        minimumFractionDigits: 3,
        maximumFractionDigits: 3
    }).format(totalSales);

    drawerCustQty.textContent = totalQty.toLocaleString() + ' units';
    drawerCustOrders.textContent = orderCount.toLocaleString() + ' recorded transactions';

    // Timeline calculation
    if (dates.length > 0) {
        dates.sort((a, b) => a - b);
        drawerCustFirst.textContent = dates[0].toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
        drawerCustLast.textContent = dates[dates.length - 1].toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
    } else {
        drawerCustFirst.textContent = 'N/A';
        drawerCustLast.textContent = 'N/A';
    }

    // Build itemized purchase breakdown
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
                    <td class="px-4 py-2.5 text-xs text-right font-bold text-gray-900">
                        ₱${metrics.sales.toLocaleString(undefined, { minimumFractionDigits: 3, maximumFractionDigits: 3 })}
                    </td>
                </tr>`;
        });
    }
    drawerCustProducts.innerHTML = productsHTML;

    // Show dynamic side sheet
    customerDrawer.classList.remove('hidden');
    document.body.classList.add('overflow-hidden'); 
}

// Close drawer handlers
function closeDrawer() {
    customerDrawer.classList.add('hidden');
    document.body.classList.remove('overflow-hidden');
}

closeDrawerBtn.addEventListener('click', closeDrawer);
closeDrawerBackdrop.addEventListener('click', closeDrawer);

// Helper to convert Excel serialized dates or string values to real Javascript Dates
function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    
    if (typeof val === 'number') {
        const dateUtc = new Date((val - 25569) * 86400 * 1000);
        return isNaN(dateUtc.getTime()) ? null : dateUtc;
    }

    let parsed = Date.parse(String(val).trim());
    return isNaN(parsed) ? null : new Date(parsed);
}

// 6. Metric Toggle Click Handlers
toggleSalesBtn.addEventListener('click', () => {
    activeChartMetric = 'sales';
    toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-brand text-white";
    toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderChart();
    renderLeaderboard();
});

toggleQtyBtn.addEventListener('click', () => {
    activeChartMetric = 'qty';
    toggleQtyBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all bg-emerald-600 text-white";
    toggleSalesBtn.className = "px-4 py-1.5 text-xs font-semibold rounded-md transition-all text-gray-600 hover:text-gray-900";
    renderChart();
    renderLeaderboard();
});

// Leaderboard Section Tabs
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

// Search utilities strictly prioritizing exact headers
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

// Converts standard numeric metrics or currency string values into floats (strictly sanitizes commas/spaces)
function parseNumericValue(val) {
    if (typeof val === 'number') return val;
    if (val === undefined || val === null) return 0;
    
    let strVal = String(val).trim();
    if (!strVal || strVal === '-' || strVal === '–' || strVal === '—') return 0;
    
    let isNegative = false;
    
    if (strVal.startsWith('-')) {
        isNegative = true;
        strVal = strVal.slice(1);
    }
    
    if (strVal.startsWith('(') && strVal.endsWith(')')) {
        isNegative = true;
        strVal = strVal.slice(1, -1);
    }
    
    // Remove commas, spaces, and non-numeric characters EXCEPT the decimal point
    const cleaned = strVal.replace(/,/g, '').replace(/\s/g, '').replace(/[^0-9.]/g, '');
    let parsed = parseFloat(cleaned);
    
    if (isNaN(parsed)) return 0;
    
    return isNegative ? -parsed : parsed;
}

<!DOCTYPE html>
<html lang="en" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Textile Accounts Payable & Yarn Receiving Dashboard</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    colors: {
                        brand: {
                            50: '#f0f7ff',
                            100: '#e0effe',
                            500: '#0284c7',
                            600: '#0265d2',
                            700: '#0369a1',
                            900: '#0c4a6e',
                        },
                        textile: {
                            gold: '#d97706',
                            emerald: '#10b981',
                            rose: '#f43f5e',
                            darkbg: '#0f172a',
                            carddark: '#1e293b'
                        }
                    },
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                        mono: ['JetBrains Mono', 'Fira Code', 'monospace']
                    }
                }
            }
        }
    </script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">
    <!-- SheetJS (xlsx) for Excel processing -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Chart.js for visualization -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    
    <style>
        body {
            font-family: 'Inter', sans-serif;
            transition: background-color 0.3s, color 0.3s;
        }
        .financial-num {
            font-family: 'JetBrains Mono', monospace;
        }
        /* Custom scrollbar */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: rgba(0, 0, 0, 0.05);
        }
        .dark ::-webkit-scrollbar-track {
            background: rgba(255, 255, 255, 0.05);
        }
        ::-webkit-scrollbar-thumb {
            background: rgba(0, 0, 0, 0.2);
            border-radius: 4px;
        }
        .dark ::-webkit-scrollbar-thumb {
            background: rgba(255, 255, 255, 0.2);
            border-radius: 4px;
        }
        /* Print styles */
        @media print {
            .no-print {
                display: none !important;
            }
            body {
                background: white !important;
                color: black !important;
            }
            .print-full-width {
                width: 100% !important;
                max-width: none !important;
                box-shadow: none !important;
                border: none !important;
            }
        }
        .tab-btn.active {
            border-bottom: 3px solid #0284c7;
            color: #0284c7;
            font-weight: 600;
        }
        .dark .tab-btn.active {
            border-bottom: 3px solid #38bdf8;
            color: #38bdf8;
        }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 dark:bg-slate-950 dark:text-slate-100 min-h-screen flex flex-col">

    <!-- TOP HEADER / NAVBAR -->
    <header class="bg-white dark:bg-slate-900 border-b border-slate-200 dark:border-slate-800 sticky top-0 z-30 shadow-sm transition-colors duration-200 no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex items-center justify-between h-16">
                
                <!-- Logo & Brand -->
                <div class="flex items-center space-x-3">
                    <div class="bg-gradient-to-tr from-brand-600 to-indigo-600 text-white p-2.5 rounded-xl shadow-md">
                        <i class="fa-solid font-bold fa-industry text-xl"></i>
                    </div>
                    <div>
                        <h1 class="font-extrabold text-lg sm:text-xl tracking-tight text-slate-900 dark:text-white flex items-center gap-2">
                            TexPay <span class="text-xs px-2 py-0.5 rounded-full bg-brand-100 dark:bg-brand-900/60 text-brand-700 dark:text-brand-300 font-semibold uppercase">Yarn ERP</span>
                        </h1>
                        <p class="text-xs text-slate-500 dark:text-slate-400 hidden sm:block">Accounts Payable & Yarn Receiving Dashboard</p>
                    </div>
                </div>

                <!-- Global Controls & Actions -->
                <div class="flex items-center space-x-2 sm:space-x-3">
                    <!-- Sample Data Quick Load -->
                    <button id="btn-load-sample" onclick="loadSampleData()" class="px-3 py-1.5 text-xs sm:text-sm font-medium rounded-lg bg-indigo-50 dark:bg-indigo-950/60 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-800 hover:bg-indigo-100 dark:hover:bg-indigo-900/80 transition flex items-center gap-1.5 shadow-sm">
                        <i class="fa-solid fa-bolt text-amber-500"></i>
                        <span>Load Sample Data</span>
                    </button>

                    <!-- Dynamic Header Row Selector Dropdown -->
                    <div class="flex items-center bg-slate-100 dark:bg-slate-800 rounded-lg px-2 py-1 border border-slate-200 dark:border-slate-700 text-xs">
                        <label for="header-row-idx" class="text-slate-500 dark:text-slate-400 mr-1.5 whitespace-nowrap">Header Row:</label>
                        <select id="header-row-idx" onchange="onHeaderRowChange()" class="bg-transparent font-medium focus:outline-none text-slate-800 dark:text-slate-200 cursor-pointer">
                            <option value="0" class="dark:bg-slate-800">Row 1 (0)</option>
                            <option value="1" class="dark:bg-slate-800">Row 2 (1)</option>
                            <option value="2" class="dark:bg-slate-800">Row 3 (2)</option>
                            <option value="3" class="dark:bg-slate-800">Row 4 (3)</option>
                        </select>
                    </div>

                    <!-- Upload Excel File Trigger -->
                    <label for="excel-file-input" class="cursor-pointer px-3 py-1.5 text-xs sm:text-sm font-medium rounded-lg bg-brand-600 hover:bg-brand-700 text-white transition flex items-center gap-1.5 shadow-sm">
                        <i class="fa-solid fa-file-excel"></i>
                        <span class="hidden sm:inline">Upload Excel/CSV</span>
                    </label>
                    <input type="file" id="excel-file-input" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleFileUpload(event)">

                    <!-- Theme Toggle -->
                    <button onclick="toggleTheme()" class="p-2 text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition">
                        <i id="theme-toggle-dark-icon" class="fa-solid fa-moon text-base hidden"></i>
                        <i id="theme-toggle-light-icon" class="fa-solid fa-sun text-base"></i>
                    </button>
                </div>
            </div>
        </div>
    </header>

    <!-- SHEET SELECTOR MODAL / BANNER (Shown if multiple sheets in uploaded file) -->
    <div id="sheet-selector-container" class="hidden bg-brand-50 dark:bg-slate-900 border-b border-brand-200 dark:border-slate-800 py-2.5 px-4 text-xs">
        <div class="max-w-7xl mx-auto flex items-center justify-between gap-4">
            <div class="flex items-center gap-2 text-brand-900 dark:text-brand-200">
                <i class="fa-solid fa-layer-group text-brand-600"></i>
                <span class="font-semibold">Multiple Worksheets Detected:</span>
            </div>
            <div id="sheet-buttons" class="flex flex-wrap gap-2"></div>
        </div>
    </div>

    <!-- MAIN CONTENT BODY -->
    <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">

        <!-- GLOBAL TOP EXECUTIVE KPI CARDS -->
        <section class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
            
            <!-- KPI 1: Total Invoice Amount -->
            <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm relative overflow-hidden group">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Total Invoiced</p>
                        <h3 id="kpi-total-invoice" class="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white mt-1 financial-num">PKR 0</h3>
                    </div>
                    <div class="p-3 bg-blue-50 dark:bg-blue-950/60 text-blue-600 dark:text-blue-400 rounded-xl group-hover:scale-110 transition-transform">
                        <i class="fa-solid fa-file-invoice-dollar text-xl"></i>
                    </div>
                </div>
                <div class="mt-2 flex items-center text-xs text-slate-500 dark:text-slate-400">
                    <span id="kpi-invoice-count" class="font-semibold text-slate-700 dark:text-slate-300">0 Invoices</span>
                </div>
                <div class="absolute bottom-0 left-0 right-0 h-1 bg-blue-500"></div>
            </div>

            <!-- KPI 2: Total Paid -->
            <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm relative overflow-hidden group">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Total Paid</p>
                        <h3 id="kpi-total-paid" class="text-xl sm:text-2xl font-bold text-emerald-600 dark:text-emerald-400 mt-1 financial-num">PKR 0</h3>
                    </div>
                    <div class="p-3 bg-emerald-50 dark:bg-emerald-950/60 text-emerald-600 dark:text-emerald-400 rounded-xl group-hover:scale-110 transition-transform">
                        <i class="fa-solid fa-circle-check text-xl"></i>
                    </div>
                </div>
                <div class="mt-2 flex items-center text-xs text-slate-500 dark:text-slate-400">
                    <span id="kpi-paid-percentage" class="font-semibold text-emerald-600 dark:text-emerald-400">0% Paid</span>
                </div>
                <div class="absolute bottom-0 left-0 right-0 h-1 bg-emerald-500"></div>
            </div>

            <!-- KPI 3: Balance Payable -->
            <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm relative overflow-hidden group">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Balance Payable</p>
                        <h3 id="kpi-total-balance" class="text-xl sm:text-2xl font-bold text-rose-600 dark:text-rose-400 mt-1 financial-num">PKR 0</h3>
                    </div>
                    <div class="p-3 bg-rose-50 dark:bg-rose-950/60 text-rose-600 dark:text-rose-400 rounded-xl group-hover:scale-110 transition-transform">
                        <i class="fa-solid fa-clock-rotate-left text-xl"></i>
                    </div>
                </div>
                <div class="mt-2 flex items-center text-xs text-slate-500 dark:text-slate-400">
                    <span id="kpi-unpaid-count" class="font-semibold text-rose-500">0 Pending Clearance</span>
                </div>
                <div class="absolute bottom-0 left-0 right-0 h-1 bg-rose-500"></div>
            </div>

            <!-- KPI 4: Yarn Received (Qty) -->
            <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm relative overflow-hidden group">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Total Quantity</p>
                        <h3 id="kpi-total-qty" class="text-xl sm:text-2xl font-bold text-amber-600 dark:text-amber-400 mt-1 financial-num">0 Lbs/Bags</h3>
                    </div>
                    <div class="p-3 bg-amber-50 dark:bg-amber-950/60 text-amber-600 dark:text-amber-400 rounded-xl group-hover:scale-110 transition-transform">
                        <i class="fa-solid fa-boxes-stacked text-xl"></i>
                    </div>
                </div>
                <div class="mt-2 flex items-center text-xs text-slate-500 dark:text-slate-400">
                    <span>Yarn Receiving Records</span>
                </div>
                <div class="absolute bottom-0 left-0 right-0 h-1 bg-amber-500"></div>
            </div>

            <!-- KPI 5: Active Yarn Suppliers -->
            <div class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm relative overflow-hidden group">
                <div class="flex items-center justify-between">
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400 uppercase tracking-wider">Active Parties</p>
                        <h3 id="kpi-party-count" class="text-xl sm:text-2xl font-bold text-indigo-600 dark:text-indigo-400 mt-1">0</h3>
                    </div>
                    <div class="p-3 bg-indigo-50 dark:bg-indigo-950/60 text-indigo-600 dark:text-indigo-400 rounded-xl group-hover:scale-110 transition-transform">
                        <i class="fa-solid fa-users-viewfinder text-xl"></i>
                    </div>
                </div>
                <div class="mt-2 flex items-center text-xs text-slate-500 dark:text-slate-400">
                    <span>Suppliers / Mills</span>
                </div>
                <div class="absolute bottom-0 left-0 right-0 h-1 bg-indigo-500"></div>
            </div>

        </section>

        <!-- FILTERS AND GLOBAL SEARCH BAR -->
        <section class="bg-white dark:bg-slate-900 rounded-xl p-4 border border-slate-200/80 dark:border-slate-800 shadow-sm space-y-3 no-print">
            <div class="flex flex-col md:flex-row gap-3 items-stretch md:items-center justify-between">
                
                <!-- Search Input -->
                <div class="relative flex-1">
                    <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-sm"></i>
                    <input type="text" id="filter-search" oninput="applyFilters()" placeholder="Search supplier party, invoice #, yarn quality, or notes..." class="w-full pl-9 pr-4 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:outline-none focus:ring-2 focus:ring-brand-500 text-slate-800 dark:text-slate-200">
                </div>

                <div class="flex flex-wrap items-center gap-2.5">
                    <!-- Party Filter Dropdown -->
                    <select id="filter-party" onchange="applyFilters()" class="px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:outline-none focus:ring-2 focus:ring-brand-500 text-slate-800 dark:text-slate-200">
                        <option value="ALL">All Parties / Suppliers</option>
                    </select>

                    <!-- Month Filter Dropdown -->
                    <select id="filter-month" onchange="applyFilters()" class="px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:outline-none focus:ring-2 focus:ring-brand-500 text-slate-800 dark:text-slate-200">
                        <option value="ALL">All Months</option>
                    </select>

                    <!-- Payment Status Filter -->
                    <select id="filter-status" onchange="applyFilters()" class="px-3 py-2 text-sm bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg focus:outline-none focus:ring-2 focus:ring-brand-500 text-slate-800 dark:text-slate-200">
                        <option value="ALL">All Payment Status</option>
                        <option value="Paid">Paid Only</option>
                        <option value="Partial">Partial Paid</option>
                        <option value="Unpaid">Unpaid / Pending</option>
                    </select>

                    <!-- Currency / Unit Display Prefix Toggle -->
                    <button onclick="toggleCurrencyPrefix()" id="btn-currency-prefix" class="px-3 py-2 text-sm font-semibold rounded-lg bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 hover:bg-slate-200 dark:hover:bg-slate-700 text-slate-700 dark:text-slate-300 transition flex items-center gap-1.5" title="Toggle PKR Prefix display">
                        <i class="fa-solid fa-coins text-amber-500"></i>
                        <span id="lbl-currency-prefix">Prefix: PKR</span>
                    </button>

                    <!-- Reset Filters Button -->
                    <button onclick="resetFilters()" class="p-2 text-slate-400 hover:text-slate-600 dark:hover:text-slate-200 rounded-lg hover:bg-slate-100 dark:hover:bg-slate-800 transition" title="Reset All Filters">
                        <i class="fa-solid fa-rotate-right"></i>
                    </button>
                </div>
            </div>
        </section>

        <!-- MULTI-TAB NAVIGATION -->
        <div class="border-b border-slate-200 dark:border-slate-800 no-print">
            <nav class="flex space-x-6 overflow-x-auto pb-px" aria-label="Tabs">
                <button onclick="switchTab('tab-pivot')" id="tab-btn-pivot" class="tab-btn active py-3 px-1 text-sm font-medium border-b-2 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-table-cells text-base"></i>
                    <span>📊 Executive Pivot Summary</span>
                </button>
                <button onclick="switchTab('tab-details')" id="tab-btn-details" class="tab-btn py-3 px-1 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-receipt text-base"></i>
                    <span>🧾 Invoice-wise Breakdown</span>
                </button>
                <button onclick="switchTab('tab-analytics')" id="tab-btn-analytics" class="tab-btn py-3 px-1 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-chart-column text-base"></i>
                    <span>📈 Party Analytics & Charts</span>
                </button>
                <button onclick="switchTab('tab-receivables')" id="tab-btn-receivables" class="tab-btn py-3 px-1 text-sm font-medium border-b-2 border-transparent text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-slate-200 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-hand-holding-dollar text-base"></i>
                    <span>💰 Receivables & Ledger</span>
                </button>
            </nav>
        </div>

        <!-- TAB CONTENT SECTION -->
        <section id="tab-content-container">

            <!-- TAB 1: EXECUTIVE PIVOT SUMMARY (PARTY X MONTH MATRIX) -->
            <div id="tab-pivot" class="tab-pane space-y-4">
                
                <!-- Pivot Controls & Value Toggle -->
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 bg-white dark:bg-slate-900 p-3.5 rounded-xl border border-slate-200/80 dark:border-slate-800 shadow-sm no-print">
                    <div class="flex items-center gap-2">
                        <span class="text-xs font-semibold uppercase text-slate-500 dark:text-slate-400 tracking-wider">Matrix Values:</span>
                        <div class="inline-flex rounded-lg p-1 bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700">
                            <button onclick="setPivotValueMode('amount')" id="pivot-mode-amount" class="px-3 py-1 text-xs font-medium rounded-md bg-white dark:bg-slate-700 text-brand-600 dark:text-brand-300 shadow-sm transition">
                                Invoice Amount (PKR)
                            </button>
                            <button onclick="setPivotValueMode('qty')" id="pivot-mode-qty" class="px-3 py-1 text-xs font-medium rounded-md text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 transition">
                                Yarn Qty (Bags/Kg)
                            </button>
                            <button onclick="setPivotValueMode('paid')" id="pivot-mode-paid" class="px-3 py-1 text-xs font-medium rounded-md text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 transition">
                                Paid Amount (PKR)
                            </button>
                        </div>
                    </div>

                    <!-- Export Matrix Buttons -->
                    <div class="flex items-center gap-2 w-full sm:w-auto justify-end">
                        <button onclick="exportPivotToCSV()" class="px-3 py-1.5 text-xs font-medium rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-file-csv"></i>
                            <span>Export Matrix CSV</span>
                        </button>
                        <button onclick="triggerPrint()" class="px-3 py-1.5 text-xs font-medium rounded-lg bg-slate-700 hover:bg-slate-800 dark:bg-slate-800 dark:hover:bg-slate-700 text-white transition flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-print"></i>
                            <span>Print PDF</span>
                        </button>
                    </div>
                </div>

                <!-- Pivot Table Container -->
                <div class="bg-white dark:bg-slate-900 rounded-xl border border-slate-200/80 dark:border-slate-800 shadow-sm overflow-hidden print-full-width">
                    <div class="overflow-x-auto max-h-[600px] relative">
                        <table class="w-full text-xs text-left text-slate-700 dark:text-slate-300 border-collapse" id="pivot-matrix-table">
                            <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-800/90 text-slate-700 dark:text-slate-300 sticky top-0 z-20 backdrop-blur-sm">
                                <tr id="pivot-header-row">
                                    <!-- Dynamic header injected here -->
                                </tr>
                            </thead>
                            <tbody id="pivot-table-body" class="divide-y divide-slate-100 dark:divide-slate-800/60 font-mono">
                                <!-- Dynamic rows injected here -->
                            </tbody>
                            <tfoot class="bg-slate-100/90 dark:bg-slate-800/90 font-bold sticky bottom-0 z-10 font-mono" id="pivot-footer">
                                <tr id="pivot-footer-row">
                                    <!-- Grand total row injected here -->
                                </tr>
                            </tfoot>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 2: INVOICE-WISE DETAILED BREAKDOWN TABLE -->
            <div id="tab-details" class="tab-pane hidden space-y-4">
                <div class="flex justify-between items-center no-print">
                    <p class="text-xs text-slate-500 dark:text-slate-400">
                        Showing line-item yarn ledger records. Click column headers to sort.
                    </p>
                    <button onclick="exportDetailsToCSV()" class="px-3 py-1.5 text-xs font-medium rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 shadow-sm">
                        <i class="fa-solid fa-file-excel"></i>
                        <span>Export Detail Records</span>
                    </button>
                </div>

                <div class="bg-white dark:bg-slate-900 rounded-xl border border-slate-200/80 dark:border-slate-800 shadow-sm overflow-hidden">
                    <div class="overflow-x-auto max-h-[620px]">
                        <table class="w-full text-xs text-left text-slate-700 dark:text-slate-300">
                            <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 sticky top-0 z-10">
                                <tr>
                                    <th scope="col" class="px-4 py-3 cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('partyName')">Party / Mill <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('month')">Month <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('invoiceNo')">Inv / Gate Pass # <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('yarnQuality')">Yarn Quality / Spec <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 text-right cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('quantity')">Qty (Lbs/Bags) <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 text-right cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('invoiceAmount')">Inv Amount <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 text-right cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('paidAmount')">Paid <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 text-right cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('balance')">Balance <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                    <th scope="col" class="px-4 py-3 text-center cursor-pointer hover:bg-slate-200 dark:hover:bg-slate-700" onclick="sortTable('status')">Status <i class="fa-solid fa-sort text-slate-400 ml-1"></i></th>
                                </tr>
                            </thead>
                            <tbody id="detailed-table-body" class="divide-y divide-slate-100 dark:divide-slate-800 font-mono">
                                <!-- Rows injected dynamically -->
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>

            <!-- TAB 3: PARTY ANALYTICS & CHARTS -->
            <div id="tab-analytics" class="tab-pane hidden space-y-6">
                <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    
                    <!-- Chart 1: Top Suppliers by Balance Payable -->
                    <div class="bg-white dark:bg-slate-900 rounded-xl p-5 border border-slate-200/80 dark:border-slate-800 shadow-sm">
                        <h4 class="font-semibold text-sm text-slate-800 dark:text-slate-200 mb-4 flex items-center gap-2">
                            <i class="fa-solid fa-chart-bar text-brand-500"></i>
                            Top Suppliers by Balance Payable (PKR)
                        </h4>
                        <div class="h-72">
                            <canvas id="chart-top-payables"></canvas>
                        </div>
                    </div>

                    <!-- Chart 2: Month-over-Month Receiving & Payments -->
                    <div class="bg-white dark:bg-slate-900 rounded-xl p-5 border border-slate-200/80 dark:border-slate-800 shadow-sm">
                        <h4 class="font-semibold text-sm text-slate-800 dark:text-slate-200 mb-4 flex items-center gap-2">
                            <i class="fa-solid fa-chart-line text-emerald-500"></i>
                            Monthly Invoiced vs Paid Trend (Chronological)
                        </h4>
                        <div class="h-72">
                            <canvas id="chart-monthly-trend"></canvas>
                        </div>
                    </div>

                    <!-- Chart 3: Yarn Receiving Quantities by Party -->
                    <div class="bg-white dark:bg-slate-900 rounded-xl p-5 border border-slate-200/80 dark:border-slate-800 shadow-sm lg:col-span-2">
                        <h4 class="font-semibold text-sm text-slate-800 dark:text-slate-200 mb-4 flex items-center gap-2">
                            <i class="fa-solid fa-boxes-stacked text-amber-500"></i>
                            Yarn Receiving Volume Distribution by Supplier
                        </h4>
                        <div class="h-80">
                            <canvas id="chart-yarn-quantities"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- TAB 4: ACCOUNTS RECEIVABLE & LEDGER RECONCILIATION -->
            <div id="tab-receivables" class="tab-pane hidden space-y-4">
                <div class="bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-900 rounded-xl p-6 text-white shadow-md relative overflow-hidden">
                    <div class="relative z-10 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                        <div>
                            <span class="bg-indigo-500/30 text-indigo-200 text-xs px-2.5 py-1 rounded-full border border-indigo-400/30 font-medium">Textile Module</span>
                            <h3 class="text-xl font-bold mt-2">Accounts Receivable & Fabric Sales Ledger Tracker</h3>
                            <p class="text-slate-300 text-xs mt-1 max-w-xl">
                                Track woven grey fabric sales, customer receivables, mill deductions, and customer balance reconciliations alongside payables.
                            </p>
                        </div>
                        <button onclick="alert('Receivables Entry form loaded. Connect to sales ledger backend or import sales Excel.')" class="px-4 py-2 bg-brand-500 hover:bg-brand-600 text-white text-xs font-semibold rounded-lg shadow transition">
                            + Add Customer Receivable
                        </button>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-900 rounded-xl p-6 border border-slate-200/80 dark:border-slate-800 shadow-sm text-center py-12">
                    <i class="fa-solid fa-scale-balanced text-4xl text-slate-300 dark:text-slate-700 mb-3"></i>
                    <h4 class="font-semibold text-slate-800 dark:text-slate-200 text-base">Yarn Payable vs Customer Receivable Net Position</h4>
                    <p class="text-xs text-slate-500 dark:text-slate-400 max-w-md mx-auto mt-1">
                        Use this section to reconcile yarn supplier liabilities against grey fabric customer receivables to monitor your net textile working capital.
                    </p>
                    <div class="mt-6 inline-flex gap-4">
                        <div class="text-left bg-slate-50 dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 min-w-[160px]">
                            <p class="text-xs text-slate-500">Yarn Payables</p>
                            <p id="rec-total-payable" class="text-lg font-bold text-rose-500 financial-num">PKR 0</p>
                        </div>
                        <div class="text-left bg-slate-50 dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 min-w-[160px]">
                            <p class="text-xs text-slate-500">Sales Receivables</p>
                            <p id="rec-total-receivable" class="text-lg font-bold text-emerald-500 financial-num">PKR 0</p>
                        </div>
                    </div>
                </div>
            </div>

        </section>
    </main>

    <!-- FOOTER -->
    <footer class="bg-white dark:bg-slate-900 border-t border-slate-200 dark:border-slate-800 py-4 mt-auto text-xs text-slate-500 dark:text-slate-400 no-print">
        <div class="max-w-7xl mx-auto px-4 flex flex-col sm:flex-row justify-between items-center gap-2">
            <div>
                <span class="font-semibold text-slate-700 dark:text-slate-300">TexPay ERP</span> &copy; 2026 Textile Financial Dashboard
            </div>
            <div class="flex items-center gap-4">
                <span>Chronological Sorting Active</span>
                <span>•</span>
                <span>Supports .XLSX / .CSV Upload</span>
            </div>
        </div>
    </footer>

    <!-- APPLICATION LOGIC & DATA PROCESSING -->
    <script>
        // GLOBAL STATE
        let rawWorksheetData = [];
        let parsedRecords = [];
        let filteredRecords = [];
        let workbookRef = null;
        let selectedHeaderRowIndex = 0;
        let showCurrencyPrefix = true;
        let pivotValueMode = 'amount'; // 'amount', 'qty', 'paid'
        let sortState = { key: 'partyName', direction: 'asc' };

        // Standardized Month Order mapping for Strict Chronological Order
        const CHRONOLOGICAL_MONTHS = [
            'JAN-2026', 'FEB-2026', 'MAR-2026', 'APR-2026', 'MAY-2026', 'JUN-2026',
            'JUL-2026', 'AUG-2026', 'SEP-2026', 'OCT-2026', 'NOV-2026', 'DEC-2026'
        ];

        // MOCK SAMPLE TEXTILE DATA
        const MOCK_TEXTILE_DATA = [
            { partyName: 'Kohinoor Textile Mills', month: 'JAN-2026', invoiceNo: 'INV-1092', yarnQuality: '20/1 Carded Cotton', quantity: 12500, invoiceAmount: 4850000, paidAmount: 4850000, status: 'Paid' },
            { partyName: 'Nishat Spun Mills', month: 'JAN-2026', invoiceNo: 'INV-3041', yarnQuality: '30/1 Combed Cotton', quantity: 18000, invoiceAmount: 7200000, paidAmount: 5000000, status: 'Partial' },
            { partyName: 'Sapphire Fibres Ltd', month: 'FEB-2026', invoiceNo: 'INV-8812', yarnQuality: '40/1 PC 52/48 Blend', quantity: 22000, invoiceAmount: 8900000, paidAmount: 0, status: 'Unpaid' },
            { partyName: 'Fazal Cloth Mills', month: 'FEB-2026', invoiceNo: 'INV-5520', yarnQuality: '10/1 Ring Spun Open End', quantity: 15000, invoiceAmount: 5100000, paidAmount: 5100000, status: 'Paid' },
            { partyName: 'Kohinoor Textile Mills', month: 'MAR-2026', invoiceNo: 'INV-1204', yarnQuality: '20/1 Carded Cotton', quantity: 14000, invoiceAmount: 5460000, paidAmount: 2000000, status: 'Partial' },
            { partyName: 'Ghani Textile Spinning', month: 'MAR-2026', invoiceNo: 'INV-9901', yarnQuality: '30/1 PV Fancy Slub', quantity: 9500, invoiceAmount: 4200000, paidAmount: 4200000, status: 'Paid' },
            { partyName: 'Nishat Spun Mills', month: 'APR-2026', invoiceNo: 'INV-3199', yarnQuality: '30/1 Combed Cotton', quantity: 19500, invoiceAmount: 7800000, paidAmount: 0, status: 'Unpaid' },
            { partyName: 'Sapphire Fibres Ltd', month: 'APR-2026', invoiceNo: 'INV-9020', yarnQuality: '60/1 Compact Egyptian', quantity: 8000, invoiceAmount: 6400000, paidAmount: 6400000, status: 'Paid' },
            { partyName: 'Fazal Cloth Mills', month: 'MAY-2026', invoiceNo: 'INV-5701', yarnQuality: '16/1 Carded Weft', quantity: 21000, invoiceAmount: 7350000, paidAmount: 3500000, status: 'Partial' },
            { partyName: 'Kohinoor Textile Mills', month: 'MAY-2026', invoiceNo: 'INV-1380', yarnQuality: '30/1 Carded Weft', quantity: 16000, invoiceAmount: 6240000, paidAmount: 0, status: 'Unpaid' },
            { partyName: 'Ghani Textile Spinning', month: 'JUN-2026', invoiceNo: 'INV-1011', yarnQuality: '40/1 Polyester Poly', quantity: 11000, invoiceAmount: 4800000, paidAmount: 4800000, status: 'Paid' },
            { partyName: 'Ibrahim Fibres Spinning', month: 'JUL-2026', invoiceNo: 'INV-7004', yarnQuality: '24/1 Combed Warp', quantity: 25000, invoiceAmount: 9750000, paidAmount: 4000000, status: 'Partial' },
            { partyName: 'Ibrahim Fibres Spinning', month: 'AUG-2026', invoiceNo: 'INV-7210', yarnQuality: '30/1 PC Jet Spun', quantity: 20000, invoiceAmount: 8100000, paidAmount: 0, status: 'Unpaid' },
            { partyName: 'Sapphire Fibres Ltd', month: 'SEP-2026', invoiceNo: 'INV-9410', yarnQuality: '40/1 Combed Cotton', quantity: 17500, invoiceAmount: 7175000, paidAmount: 7175000, status: 'Paid' },
            { partyName: 'Nishat Spun Mills', month: 'OCT-2026', invoiceNo: 'INV-3550', yarnQuality: '30/1 Melange Grey', quantity: 13000, invoiceAmount: 5850000, paidAmount: 0, status: 'Unpaid' }
        ];

        // INIT ON PAGE LOAD
        window.onload = function () {
            // Dark mode initialization
            if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
                document.documentElement.classList.add('dark');
            }
            updateThemeIcons();
            
            // Load sample data by default
            loadSampleData();
        };

        // THEME TOGGLE LOGIC
        function toggleTheme() {
            document.documentElement.classList.toggle('dark');
            updateThemeIcons();
            renderCharts(); // Re-render charts with correct text colors
        }

        function updateThemeIcons() {
            const isDark = document.documentElement.classList.contains('dark');
            document.getElementById('theme-toggle-dark-icon').classList.toggle('hidden', isDark);
            document.getElementById('theme-toggle-light-icon').classList.toggle('hidden', !isDark);
        }

        // LOAD SAMPLE DATA
        function loadSampleData() {
            parsedRecords = [...MOCK_TEXTILE_DATA];
            document.getElementById('sheet-selector-container').classList.add('hidden');
            populateFilterDropdowns();
            applyFilters();
        }

        // FILE UPLOAD AND HEADER ROW SELECTION
        function handleFileUpload(e) {
            const file = e.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function (evt) {
                try {
                    const data = new Uint8Array(evt.target.result);
                    workbookRef = XLSX.read(data, { type: 'array', cellDates: true });
                    
                    // Show Sheet Selector if multiple sheets exist
                    if (workbookRef.SheetNames.length > 1) {
                        const btnContainer = document.getElementById('sheet-buttons');
                        btnContainer.innerHTML = '';
                        workbookRef.SheetNames.forEach((sheetName, index) => {
                            const btn = document.createElement('button');
                            btn.className = `px-2.5 py-1 text-xs font-medium rounded-md transition ${index === 0 ? 'bg-brand-600 text-white' : 'bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-300 border border-slate-200 dark:border-slate-700'}`;
                            btn.innerText = sheetName;
                            btn.onclick = () => selectSheet(sheetName);
                            btnContainer.appendChild(btn);
                        });
                        document.getElementById('sheet-selector-container').classList.remove('hidden');
                    } else {
                        document.getElementById('sheet-selector-container').classList.add('hidden');
                    }

                    // Process First Sheet
                    selectSheet(workbookRef.SheetNames[0]);
                } catch (err) {
                    alert('Error parsing file: ' + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function selectSheet(sheetName) {
            if (!workbookRef) return;
            const sheet = workbookRef.Sheets[sheetName];
            rawWorksheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
            processRawDataWithHeaderIndex();
        }

        function onHeaderRowChange() {
            selectedHeaderRowIndex = parseInt(document.getElementById('header-row-idx').value) || 0;
            if (rawWorksheetData.length > 0) {
                processRawDataWithHeaderIndex();
            }
        }

        // FLEXIBLE HEADER PARSING & NORMALIZATION
        function processRawDataWithHeaderIndex() {
            if (!rawWorksheetData || rawWorksheetData.length <= selectedHeaderRowIndex) {
                alert('Invalid Header Row Selected.');
                return;
            }

            const headerRow = rawWorksheetData[selectedHeaderRowIndex].map(h => String(h || '').trim().toLowerCase());
            
            // Find Column Index map
            const partyIdx = headerRow.findIndex(h => h.includes('party') || h.includes('supplier') || h.includes('mill') || h.includes('vendor'));
            const monthIdx = headerRow.findIndex(h => h.includes('month') || h.includes('date') || h.includes('period'));
            const invNoIdx = headerRow.findIndex(h => h.includes('inv') || h.includes('bill') || h.includes('gate') || h.includes('number'));
            const qualityIdx = headerRow.findIndex(h => h.includes('quality') || h.includes('spec') || h.includes('yarn') || h.includes('count'));
            const qtyIdx = headerRow.findIndex(h => h.includes('qty') || h.includes('quantity') || h.includes('bag') || h.includes('lbs') || h.includes('kg'));
            const amountIdx = headerRow.findIndex(h => h.includes('amount') || h.includes('invoice') || h.includes('total') || h.includes('val'));
            const paidIdx = headerRow.findIndex(h => h.includes('paid') || h.includes('cleared') || h.includes('payment'));

            parsedRecords = [];

            for (let i = selectedHeaderRowIndex + 1; i < rawWorksheetData.length; i++) {
                const row = rawWorksheetData[i];
                if (!row || row.length === 0) continue;

                const partyName = partyIdx !== -1 && row[partyIdx] ? String(row[partyIdx]).trim() : 'Unspecified Party';
                let rawMonth = monthIdx !== -1 && row[monthIdx] ? String(row[monthIdx]).trim().toUpperCase() : 'JAN-2026';
                
                // Month formatting normalizer
                rawMonth = normalizeMonthString(rawMonth);

                const invoiceNo = invNoIdx !== -1 && row[invNoIdx] ? String(row[invNoIdx]).trim() : 'INV-' + (i + 100);
                const yarnQuality = qualityIdx !== -1 && row[qualityIdx] ? String(row[qualityIdx]).trim() : 'Yarn Consignment';
                const quantity = qtyIdx !== -1 ? parseFloat(row[qtyIdx]) || 0 : 0;
                const invoiceAmount = amountIdx !== -1 ? parseFloat(row[amountIdx]) || 0 : 0;
                const paidAmount = paidIdx !== -1 ? parseFloat(row[paidIdx]) || 0 : 0;
                
                let status = 'Unpaid';
                if (paidAmount >= invoiceAmount && invoiceAmount > 0) status = 'Paid';
                else if (paidAmount > 0) status = 'Partial';

                parsedRecords.push({
                    partyName,
                    month: rawMonth,
                    invoiceNo,
                    yarnQuality,
                    quantity,
                    invoiceAmount,
                    paidAmount,
                    status
                });
            }

            populateFilterDropdowns();
            applyFilters();
        }

        // Helper to normalize months to format JAN-2026
        function normalizeMonthString(monthStr) {
            monthStr = monthStr.replace(/[^A-Z0-9-]/gi, '').toUpperCase();
            
            // Check if monthStr matches standard CHRONOLOGICAL_MONTHS
            const matched = CHRONOLOGICAL_MONTHS.find(m => monthStr.includes(m.substring(0, 3)));
            if (matched) return matched;

            return 'JAN-2026';
        }

        // POPULATE DROPDOWNS
        function populateFilterDropdowns() {
            const partySelect = document.getElementById('filter-party');
            const monthSelect = document.getElementById('filter-month');

            const currentParty = partySelect.value;
            const currentMonth = monthSelect.value;

            // Unique Parties
            const parties = Array.from(new Set(parsedRecords.map(r => r.partyName))).sort();
            partySelect.innerHTML = '<option value="ALL">All Parties / Suppliers</option>';
            parties.forEach(p => {
                partySelect.innerHTML += `<option value="${p}">${p}</option>`;
            });
            partySelect.value = parties.includes(currentParty) ? currentParty : 'ALL';

            // Chronologically ordered months present in dataset
            const monthsInDataset = Array.from(new Set(parsedRecords.map(r => r.month)));
            monthsInDataset.sort((a, b) => CHRONOLOGICAL_MONTHS.indexOf(a) - CHRONOLOGICAL_MONTHS.indexOf(b));

            monthSelect.innerHTML = '<option value="ALL">All Months</option>';
            monthsInDataset.forEach(m => {
                monthSelect.innerHTML += `<option value="${m}">${m}</option>`;
            });
            monthSelect.value = monthsInDataset.includes(currentMonth) ? currentMonth : 'ALL';
        }

        // APPLY FILTERS & RE-RENDER
        function applyFilters() {
            const search = document.getElementById('filter-search').value.toLowerCase();
            const party = document.getElementById('filter-party').value;
            const month = document.getElementById('filter-month').value;
            const status = document.getElementById('filter-status').value;

            filteredRecords = parsedRecords.filter(r => {
                const matchesSearch = !search || 
                    r.partyName.toLowerCase().includes(search) || 
                    r.invoiceNo.toLowerCase().includes(search) || 
                    r.yarnQuality.toLowerCase().includes(search);

                const matchesParty = party === 'ALL' || r.partyName === party;
                const matchesMonth = month === 'ALL' || r.month === month;
                const matchesStatus = status === 'ALL' || r.status === status;

                return matchesSearch && matchesParty && matchesMonth && matchesStatus;
            });

            renderKPIs();
            renderPivotMatrix();
            renderDetailedTable();
            renderCharts();
        }

        function resetFilters() {
            document.getElementById('filter-search').value = '';
            document.getElementById('filter-party').value = 'ALL';
            document.getElementById('filter-month').value = 'ALL';
            document.getElementById('filter-status').value = 'ALL';
            applyFilters();
        }

        // RENDER EXECUTIVE KPIS
        function renderKPIs() {
            const totalInvoice = filteredRecords.reduce((sum, r) => sum + r.invoiceAmount, 0);
            const totalPaid = filteredRecords.reduce((sum, r) => sum + r.paidAmount, 0);
            const totalBalance = Math.max(0, totalInvoice - totalPaid);
            const totalQty = filteredRecords.reduce((sum, r) => sum + r.quantity, 0);
            const activeParties = new Set(filteredRecords.map(r => r.partyName)).size;
            
            const unpaidCount = filteredRecords.filter(r => r.status !== 'Paid').length;
            const paidPct = totalInvoice > 0 ? Math.round((totalPaid / totalInvoice) * 100) : 0;

            document.getElementById('kpi-total-invoice').innerText = formatCurrency(totalInvoice);
            document.getElementById('kpi-total-paid').innerText = formatCurrency(totalPaid);
            document.getElementById('kpi-total-balance').innerText = formatCurrency(totalBalance);
            document.getElementById('kpi-total-qty').innerText = formatNumber(totalQty) + ' Bags/Lbs';
            document.getElementById('kpi-party-count').innerText = activeParties;

            document.getElementById('kpi-invoice-count').innerText = `${filteredRecords.length} Invoice Vouchers`;
            document.getElementById('kpi-paid-percentage').innerText = `${paidPct}% Settled`;
            document.getElementById('kpi-unpaid-count').innerText = `${unpaidCount} Pending Vouchers`;

            // Tab 4 ledger net numbers
            document.getElementById('rec-total-payable').innerText = formatCurrency(totalBalance);
            document.getElementById('rec-total-receivable').innerText = formatCurrency(totalInvoice * 0.85); // Mock relative receivable
        }

        // TOGGLE CURRENCY PREFIX (PKR)
        function toggleCurrencyPrefix() {
            showCurrencyPrefix = !showCurrencyPrefix;
            document.getElementById('lbl-currency-prefix').innerText = `Prefix: ${showCurrencyPrefix ? 'PKR' : 'Off'}`;
            renderKPIs();
            renderPivotMatrix();
            renderDetailedTable();
        }

        // PIVOT MATRIX VALUE TOGGLE
        function setPivotValueMode(mode) {
            pivotValueMode = mode;
            ['amount', 'qty', 'paid'].forEach(m => {
                const btn = document.getElementById(`pivot-mode-${m}`);
                if (m === mode) {
                    btn.className = "px-3 py-1 text-xs font-medium rounded-md bg-white dark:bg-slate-700 text-brand-600 dark:text-brand-300 shadow-sm transition";
                } else {
                    btn.className = "px-3 py-1 text-xs font-medium rounded-md text-slate-600 dark:text-slate-400 hover:text-slate-900 dark:hover:text-slate-200 transition";
                }
            });
            renderPivotMatrix();
        }

        // RENDER EXECUTIVE PIVOT MATRIX (CHRONOLOGICAL MONTH COLS x PARTY ROWS)
        function renderPivotMatrix() {
            const headerRow = document.getElementById('pivot-header-row');
            const tbody = document.getElementById('pivot-table-body');
            const footerRow = document.getElementById('pivot-footer-row');

            // Unique Parties
            const parties = Array.from(new Set(filteredRecords.map(r => r.partyName))).sort();

            // Strict Chronological Month Columns active in current filtered dataset
            const activeMonths = CHRONOLOGICAL_MONTHS.filter(m => 
                filteredRecords.some(r => r.month === m)
            );

            // 1. Build Header
            let headerHTML = `<th scope="col" class="px-4 py-3 min-w-[200px] border-r border-slate-200 dark:border-slate-700 font-bold bg-slate-100 dark:bg-slate-800">Supplier / Mill Party</th>`;
            activeMonths.forEach(month => {
                headerHTML += `<th scope="col" class="px-3 py-3 text-right min-w-[110px] border-r border-slate-200 dark:border-slate-700/60 font-semibold">${month}</th>`;
            });
            headerHTML += `<th scope="col" class="px-4 py-3 text-right min-w-[130px] font-bold bg-slate-200/60 dark:bg-slate-800">Total ${getValueModeLabel()}</th>`;
            headerRow.innerHTML = headerHTML;

            // 2. Build Matrix Rows
            let tbodyHTML = '';
            const monthTotals = {};
            activeMonths.forEach(m => monthTotals[m] = 0);
            let grandMatrixTotal = 0;

            if (parties.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${activeMonths.length + 2}" class="px-4 py-8 text-center text-slate-400">No yarn records matching active filters.</td></tr>`;
                footerRow.innerHTML = '';
                return;
            }

            parties.forEach(party => {
                tbodyHTML += `<tr class="hover:bg-slate-50 dark:hover:bg-slate-800/50 transition">`;
                tbodyHTML += `<td class="px-4 py-3 font-semibold text-slate-900 dark:text-white border-r border-slate-200 dark:border-slate-800">${party}</td>`;

                let partyRowTotal = 0;

                activeMonths.forEach(month => {
                    // Filter records for this party & month
                    const cellRecords = filteredRecords.filter(r => r.partyName === party && r.month === month);
                    
                    let cellVal = 0;
                    if (pivotValueMode === 'amount') cellVal = cellRecords.reduce((s, r) => s + r.invoiceAmount, 0);
                    else if (pivotValueMode === 'qty') cellVal = cellRecords.reduce((s, r) => s + r.quantity, 0);
                    else if (pivotValueMode === 'paid') cellVal = cellRecords.reduce((s, r) => s + r.paidAmount, 0);

                    partyRowTotal += cellVal;
                    monthTotals[month] += cellVal;

                    const formattedVal = cellVal === 0 ? '<span class="text-slate-300 dark:text-slate-600">-</span>' : formatValueByMode(cellVal);
                    tbodyHTML += `<td class="px-3 py-3 text-right border-r border-slate-100 dark:border-slate-800/60">${formattedVal}</td>`;
                });

                grandMatrixTotal += partyRowTotal;
                tbodyHTML += `<td class="px-4 py-3 text-right font-bold bg-slate-50/80 dark:bg-slate-800/40 text-brand-700 dark:text-brand-300">${formatValueByMode(partyRowTotal)}</td>`;
                tbodyHTML += `</tr>`;
            });

            tbody.innerHTML = tbodyHTML;

            // 3. Build Footer (Grand Totals Row)
            let footerHTML = `<td class="px-4 py-3 border-r border-slate-200 dark:border-slate-700 uppercase font-bold text-slate-900 dark:text-white">Grand Total</td>`;
            activeMonths.forEach(month => {
                footerHTML += `<td class="px-3 py-3 text-right border-r border-slate-200 dark:border-slate-700 text-slate-900 dark:text-white font-bold">${formatValueByMode(monthTotals[month])}</td>`;
            });
            footerHTML += `<td class="px-4 py-3 text-right text-brand-700 dark:text-brand-300 font-extrabold bg-slate-200 dark:bg-slate-700/80">${formatValueByMode(grandMatrixTotal)}</td>`;
            footerRow.innerHTML = footerHTML;
        }

        function getValueModeLabel() {
            if (pivotValueMode === 'qty') return 'Qty (Bags)';
            if (pivotValueMode === 'paid') return 'Paid';
            return 'Invoiced';
        }

        function formatValueByMode(val) {
            if (pivotValueMode === 'qty') return formatNumber(val);
            return formatCurrency(val);
        }

        // RENDER INVOICE-WISE DETAILED BREAKDOWN TABLE
        function renderDetailedTable() {
            const tbody = document.getElementById('detailed-table-body');
            
            // Sort filtered records
            const sorted = [...filteredRecords].sort((a, b) => {
                let valA = a[sortState.key];
                let valB = b[sortState.key];
                if (typeof valA === 'string') valA = valA.toLowerCase();
                if (typeof valB === 'string') valB = valB.toLowerCase();

                if (valA < valB) return sortState.direction === 'asc' ? -1 : 1;
                if (valA > valB) return sortState.direction === 'asc' ? 1 : -1;
                return 0;
            });

            if (sorted.length === 0) {
                tbody.innerHTML = `<tr><td colspan="9" class="px-4 py-8 text-center text-slate-400">No records found matching current search/filters.</td></tr>`;
                return;
            }

            let html = '';
            sorted.forEach(r => {
                const balance = Math.max(0, r.invoiceAmount - r.paidAmount);
                
                let badgeClass = 'bg-slate-100 text-slate-700 dark:bg-slate-800 dark:text-slate-300';
                if (r.status === 'Paid') badgeClass = 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300';
                else if (r.status === 'Partial') badgeClass = 'bg-amber-100 text-amber-800 dark:bg-amber-950 dark:text-amber-300';
                else if (r.status === 'Unpaid') badgeClass = 'bg-rose-100 text-rose-800 dark:bg-rose-950 dark:text-rose-300';

                html += `
                    <tr class="hover:bg-slate-50 dark:hover:bg-slate-800/40 transition">
                        <td class="px-4 py-3 font-semibold text-slate-900 dark:text-slate-100">${r.partyName}</td>
                        <td class="px-4 py-3"><span class="px-2 py-0.5 rounded bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 text-[11px] font-sans">${r.month}</span></td>
                        <td class="px-4 py-3 text-slate-600 dark:text-slate-400">${r.invoiceNo}</td>
                        <td class="px-4 py-3 text-slate-500 dark:text-slate-400 font-sans text-xs">${r.yarnQuality}</td>
                        <td class="px-4 py-3 text-right">${formatNumber(r.quantity)}</td>
                        <td class="px-4 py-3 text-right font-medium">${formatCurrency(r.invoiceAmount)}</td>
                        <td class="px-4 py-3 text-right text-emerald-600 dark:text-emerald-400">${formatCurrency(r.paidAmount)}</td>
                        <td class="px-4 py-3 text-right font-semibold text-rose-600 dark:text-rose-400">${formatCurrency(balance)}</td>
                        <td class="px-4 py-3 text-center">
                            <span class="px-2.5 py-1 rounded-full text-[11px] font-bold font-sans ${badgeClass}">${r.status}</span>
                        </td>
                    </tr>
                `;
            });

            tbody.innerHTML = html;
        }

        function sortTable(key) {
            if (sortState.key === key) {
                sortState.direction = sortState.direction === 'asc' ? 'desc' : 'asc';
            } else {
                sortState.key = key;
                sortState.direction = 'asc';
            }
            renderDetailedTable();
        }

        // CHARTS RENDERING (CHART.JS)
        let chart1Instance = null;
        let chart2Instance = null;
        let chart3Instance = null;

        function renderCharts() {
            const isDark = document.documentElement.classList.contains('dark');
            const textColor = isDark ? '#94a3b8' : '#64748b';
            const gridColor = isDark ? 'rgba(255, 255, 255, 0.05)' : 'rgba(0, 0, 0, 0.05)';

            // Destroy existing charts
            if (chart1Instance) chart1Instance.destroy();
            if (chart2Instance) chart2Instance.destroy();
            if (chart3Instance) chart3Instance.destroy();

            // CHART 1: Top Suppliers by Balance Payable
            const partyBalances = {};
            filteredRecords.forEach(r => {
                const bal = Math.max(0, r.invoiceAmount - r.paidAmount);
                partyBalances[r.partyName] = (partyBalances[r.partyName] || 0) + bal;
            });

            const topPayableParties = Object.keys(partyBalances)
                .sort((a, b) => partyBalances[b] - partyBalances[a])
                .slice(0, 6);

            const ctx1 = document.getElementById('chart-top-payables').getContext('2d');
            chart1Instance = new Chart(ctx1, {
                type: 'bar',
                data: {
                    labels: topPayableParties,
                    datasets: [{
                        label: 'Balance Payable (PKR)',
                        data: topPayableParties.map(p => partyBalances[p]),
                        backgroundColor: '#f43f5e',
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { ticks: { color: textColor }, grid: { display: false } },
                        y: { ticks: { color: textColor }, grid: { color: gridColor } }
                    }
                }
            });

            // CHART 2: Chronological Monthly Invoiced vs Paid
            const activeMonths = CHRONOLOGICAL_MONTHS.filter(m => filteredRecords.some(r => r.month === m));
            const monthlyInvoiced = activeMonths.map(m => filteredRecords.filter(r => r.month === m).reduce((s, r) => s + r.invoiceAmount, 0));
            const monthlyPaid = activeMonths.map(m => filteredRecords.filter(r => r.month === m).reduce((s, r) => s + r.paidAmount, 0));

            const ctx2 = document.getElementById('chart-monthly-trend').getContext('2d');
            chart2Instance = new Chart(ctx2, {
                type: 'line',
                data: {
                    labels: activeMonths,
                    datasets: [
                        {
                            label: 'Total Invoiced',
                            data: monthlyInvoiced,
                            borderColor: '#0284c7',
                            backgroundColor: 'rgba(2, 132, 199, 0.1)',
                            fill: true,
                            tension: 0.3
                        },
                        {
                            label: 'Total Paid',
                            data: monthlyPaid,
                            borderColor: '#10b981',
                            backgroundColor: 'transparent',
                            borderDash: [5, 5],
                            tension: 0.3
                        }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { labels: { color: textColor } } },
                    scales: {
                        x: { ticks: { color: textColor }, grid: { color: gridColor } },
                        y: { ticks: { color: textColor }, grid: { color: gridColor } }
                    }
                }
            });

            // CHART 3: Yarn Quantities Distribution
            const partyQtys = {};
            filteredRecords.forEach(r => {
                partyQtys[r.partyName] = (partyQtys[r.partyName] || 0) + r.quantity;
            });
            const partiesList = Object.keys(partyQtys);

            const ctx3 = document.getElementById('chart-yarn-quantities').getContext('2d');
            chart3Instance = new Chart(ctx3, {
                type: 'bar',
                data: {
                    labels: partiesList,
                    datasets: [{
                        label: 'Yarn Received Qty (Bags/Lbs)',
                        data: partiesList.map(p => partyQtys[p]),
                        backgroundColor: '#d97706',
                        borderRadius: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: {
                        x: { ticks: { color: textColor }, grid: { display: false } },
                        y: { ticks: { color: textColor }, grid: { color: gridColor } }
                    }
                }
            });
        }

        // TAB SWITCHING LOGIC
        function switchTab(tabId) {
            document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

            document.getElementById(tabId).classList.remove('hidden');
            
            const btnMap = {
                'tab-pivot': 'tab-btn-pivot',
                'tab-details': 'tab-btn-details',
                'tab-analytics': 'tab-btn-analytics',
                'tab-receivables': 'tab-btn-receivables'
            };
            document.getElementById(btnMap[tabId]).classList.add('active');

            if (tabId === 'tab-analytics') {
                renderCharts(); // ensure chart layout recalculation on tab display
            }
        }

        // EXPORT TO CSV / EXCEL
        function exportPivotToCSV() {
            const parties = Array.from(new Set(filteredRecords.map(r => r.partyName))).sort();
            const activeMonths = CHRONOLOGICAL_MONTHS.filter(m => filteredRecords.some(r => r.month === m));

            let csv = ['Supplier Party,' + activeMonths.join(',') + ',Total ' + getValueModeLabel()];

            parties.forEach(party => {
                let row = [`"${party}"`];
                let rowSum = 0;
                activeMonths.forEach(m => {
                    const cellRecords = filteredRecords.filter(r => r.partyName === party && r.month === m);
                    let val = 0;
                    if (pivotValueMode === 'amount') val = cellRecords.reduce((s, r) => s + r.invoiceAmount, 0);
                    else if (pivotValueMode === 'qty') val = cellRecords.reduce((s, r) => s + r.quantity, 0);
                    else if (pivotValueMode === 'paid') val = cellRecords.reduce((s, r) => s + r.paidAmount, 0);
                    
                    rowSum += val;
                    row.push(val);
                });
                row.push(rowSum);
                csv.push(row.join(','));
            });

            downloadCSVFile(csv.join('\n'), `Accounts_Payable_Pivot_Matrix_${pivotValueMode}.csv`);
        }

        function exportDetailsToCSV() {
            let csv = ['Party / Mill,Month,Invoice No,Yarn Quality,Quantity,Invoice Amount,Paid Amount,Balance,Status'];
            
            filteredRecords.forEach(r => {
                const bal = Math.max(0, r.invoiceAmount - r.paidAmount);
                csv.push(`"${r.partyName}","${r.month}","${r.invoiceNo}","${r.yarnQuality}",${r.quantity},${r.invoiceAmount},${r.paidAmount},${bal},"${r.status}"`);
            });

            downloadCSVFile(csv.join('\n'), 'Accounts_Payable_Detailed_Ledger.csv');
        }

        function downloadCSVFile(csvContent, fileName) {
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', fileName);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        function triggerPrint() {
            window.print();
        }

        // UTILITY FORMATTERS
        function formatCurrency(num) {
            const formatted = Math.round(num).toLocaleString('en-US');
            return showCurrencyPrefix ? `PKR ${formatted}` : formatted;
        }

        function formatNumber(num) {
            return Math.round(num).toLocaleString('en-US');
        }
    </script>
</body>
</html>

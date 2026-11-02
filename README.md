[光學製造生產數據管理系統-V22.html](https://github.com/user-attachments/files/34257008/-V22.html)
# TEST<!DOCTYPE html>
<html lang="zh-TW" class="h-full bg-slate-900 text-slate-100">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>光學/製造生產數據輸入與分析系統</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- Font Awesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <!-- Chart.js -->
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <!-- Chart.js Plugin Datalabels -->
  <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
  <!-- SheetJS (XLSX) library -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script>
    tailwind.config = {
      theme: {
        extend: {
          colors: {
            brand: {
              50: '#f0f7ff',
              100: '#e0effe',
              500: '#0284c7',
              600: '#0369a1',
              700: '#075985',
              800: '#0c4a6e',
              900: '#0f172a',
            }
          }
        }
      }
    }
  </script>
  <style>
    /* Custom Scrollbar */
    ::-webkit-scrollbar {
      width: 8px;
      height: 8px;
    }
    ::-webkit-scrollbar-track {
      background: #0f172a;
    }
    ::-webkit-scrollbar-thumb {
      background: #334155;
      border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #475569;
    }
    /* Sticky Column Shadow */
    .sticky-col {
      position: sticky;
      left: 0;
      z-index: 10;
    }
    /* Date picker dark mode color scheme adjustment */
    input[type="date"]::-webkit-calendar-picker-indicator,
    input[type="time"]::-webkit-calendar-picker-indicator {
      filter: invert(0.8);
      cursor: pointer;
    }

    /* 虚式「茈」能量光束動畫特效 */
    @keyframes hollowPurpleGlow {
      0%, 100% {
        box-shadow: 0 0 15px rgba(168, 85, 247, 0.6), inset 0 0 15px rgba(59, 130, 246, 0.6);
      }
      50% {
        box-shadow: 0 0 35px rgba(217, 70, 239, 0.9), inset 0 0 25px rgba(147, 51, 234, 0.9);
      }
    }
    .hollow-purple-bar {
      background: linear-gradient(90deg, #3b82f6, #9333ea, #d946ef, #ffffff);
      background-size: 200% 100%;
      animation: purpleFlow 2s linear infinite, hollowPurpleGlow 1.5s ease-in-out infinite;
    }
    @keyframes purpleFlow {
      0% { background-position: 100% 0; }
      100% { background-position: -100% 0; }
    }
  </style>
</head>
<body class="h-full font-sans antialiased bg-slate-950 text-slate-100 flex flex-col">

  <!-- Top Navigation Header -->
  <header class="bg-slate-900 border-b border-slate-800 px-6 py-3 flex flex-wrap justify-between items-center shrink-0">
    <div class="flex items-center space-x-3">
      <div class="bg-blue-600 text-white p-2 rounded-lg shadow-lg shadow-blue-500/20">
        <i class="fa-solid fa-microchip text-xl"></i>
      </div>
      <div>
        <h1 class="text-lg font-bold text-white tracking-wide">光學製造生產數據管理系統</h1>
        <p class="text-xs text-slate-400">Optics Manufacturing Data Entry & Analytics Platform</p>
      </div>
    </div>

    <!-- Global Actions -->
    <div class="flex items-center space-x-3 mt-2 sm:mt-0">
      <label class="cursor-pointer bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-2 rounded-lg border border-slate-700 text-sm font-medium flex items-center space-x-2 transition">
        <i class="fa-solid fa-file-import text-blue-400"></i>
        <span>匯入 Excel (.xlsx)</span>
        <input type="file" id="excelFileInput" accept=".xlsx, .xls" class="hidden" onchange="importExcel(event)">
      </label>

      <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-emerald-900/30">
        <i class="fa-solid fa-file-excel"></i>
        <span>匯出數據 Excel</span>
      </button>

      <button onclick="addNewRowModal()" class="bg-blue-600 hover:bg-blue-500 text-white px-4 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-blue-900/30">
        <i class="fa-solid fa-plus"></i>
        <span>新增生產紀錄</span>
      </button>
    </div>
  </header>

  <!-- Main Content Layout -->
  <main class="flex-1 overflow-y-auto p-4 md:p-6 space-y-6">

    <!-- KPI Summary Cards -->
    <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-sm relative overflow-hidden">
        <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">總生產投入數</div>
        <div class="text-2xl font-black text-white" id="kpi-total-input">0</div>
        <div class="text-xs text-slate-500 mt-1"><i class="fa-solid fa-cubes text-blue-400 mr-1"></i>批次累積數量</div>
        <div class="absolute -right-2 -bottom-2 opacity-10 text-5xl text-blue-400"><i class="fa-solid fa-boxes-stacked"></i></div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-sm relative overflow-hidden">
        <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">平均良率</div>
        <div class="text-2xl font-black text-emerald-400" id="kpi-avg-yield">0%</div>
        <div class="text-xs text-slate-500 mt-1"><i class="fa-solid fa-chart-line text-emerald-400 mr-1"></i>綜合良率</div>
        <div class="absolute -right-2 -bottom-2 opacity-10 text-5xl text-emerald-400"><i class="fa-solid fa-chart-pie"></i></div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-sm relative overflow-hidden">
        <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">平均直通率</div>
        <div class="text-2xl font-black text-cyan-400" id="kpi-avg-fpyn">0%</div>
        <div class="text-xs text-slate-500 mt-1"><i class="fa-solid fa-bolt text-cyan-400 mr-1"></i>一次完成率</div>
        <div class="absolute -right-2 -bottom-2 opacity-10 text-5xl text-cyan-400"><i class="fa-solid fa-industry"></i></div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-sm relative overflow-hidden">
        <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">AOI 總不良率</div>
        <div class="text-2xl font-black text-rose-400" id="kpi-aoi-ng-rate">0%</div>
        <div class="text-xs text-slate-500 mt-1"><i class="fa-solid fa-triangle-exclamation text-rose-400 mr-1"></i>綜合不良率</div>
        <div class="absolute -right-2 -bottom-2 opacity-10 text-5xl text-rose-400"><i class="fa-solid fa-bug"></i></div>
      </div>

      <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 shadow-sm relative overflow-hidden">
        <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-1">主要瑕疵 TOP 1</div>
        <div class="text-lg font-bold text-amber-400 truncate" id="kpi-top-defect">無數據</div>
        <div class="text-xs text-slate-500 mt-1" id="kpi-top-defect-count">累計 0 件</div>
        <div class="absolute -right-2 -bottom-2 opacity-10 text-5xl text-amber-400"><i class="fa-solid fa-magnifying-glass-chart"></i></div>
      </div>
    </div>

    <!-- Analytics Filter Bar (Hole No & Date Range) -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-wrap items-center justify-between gap-4 shadow-sm">
      <div class="flex items-center space-x-2 text-sm font-bold text-slate-200">
        <i class="fa-solid fa-filter text-blue-500"></i>
        <span>圖表統計篩選 (穴號與日期)：</span>
      </div>
      <div class="flex flex-wrap items-center gap-3 text-xs">
        <div class="flex items-center space-x-2 bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5">
          <span class="text-slate-400">穴號:</span>
          <select id="chartHoleFilter" onchange="updateDashboard()" class="bg-transparent text-amber-300 font-semibold focus:outline-none cursor-pointer">
            <option value="ALL">全部穴號</option>
          </select>
        </div>
        <div class="flex items-center space-x-2 bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5">
          <span class="text-slate-400">開始日期:</span>
          <input type="date" id="chartStartDate" onchange="updateDashboard()" oninput="updateDashboard()" onwheel="handleDateWheel(event)" class="bg-transparent text-slate-200 focus:outline-none cursor-pointer">
        </div>
        <span class="text-slate-500">至</span>
        <div class="flex items-center space-x-2 bg-slate-950 border border-slate-800 rounded-lg px-3 py-1.5">
          <span class="text-slate-400">結束日期:</span>
          <input type="date" id="chartEndDate" onchange="updateDashboard()" oninput="updateDashboard()" onwheel="handleDateWheel(event)" class="bg-transparent text-slate-200 focus:outline-none cursor-pointer">
        </div>
        <button onclick="resetChartFilters()" class="bg-slate-800 hover:bg-slate-700 text-slate-300 px-3 py-1.5 rounded-lg border border-slate-700 transition flex items-center space-x-1">
          <i class="fa-solid fa-filter-circle-xmark"></i>
          <span>重置篩選</span>
        </button>
      </div>
    </div>

    <!-- Analytics Charts Section -->
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <!-- Line Yield/FPY/Defect Chart -->
      <div class="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex justify-between items-center mb-4">
          <h2 class="text-sm font-bold text-slate-200 flex items-center space-x-2">
            <i class="fa-solid fa-diagram-project text-blue-500"></i>
            <span>各線別良率 / 直通率 / 不良率分析</span>
          </h2>
        </div>
        <div class="h-72">
          <canvas id="lineYieldChart"></canvas>
        </div>
      </div>

      <!-- Machine Yield/FPY/Defect Chart -->
      <div class="bg-slate-900 border border-slate-800 rounded-xl p-5 shadow-sm">
        <div class="flex justify-between items-center mb-4">
          <h2 class="text-sm font-bold text-slate-200 flex items-center space-x-2">
            <i class="fa-solid fa-industry text-emerald-500"></i>
            <span>各機台良率 / 直通率 / 不良率分析</span>
          </h2>
        </div>
        <div class="h-72">
          <canvas id="machineYieldChart"></canvas>
        </div>
      </div>
    </div>

    <!-- Table Section -->
    <div class="bg-slate-900 border border-slate-800 rounded-xl shadow-lg flex flex-col overflow-hidden">
      <!-- Table Filter Bar -->
      <div class="p-4 border-b border-slate-800 flex flex-wrap gap-4 items-center justify-between bg-slate-900/50">
        <div class="flex items-center space-x-3 flex-1 min-w-[280px]">
          <form onsubmit="handleSearchSubmit(event)" class="relative w-full">
            <i class="fa-solid fa-magnifying-glass absolute left-3 top-1/2 -translate-y-1/2 text-slate-500 text-sm"></i>
            <input type="text" id="searchInput" oninput="resetPageAndRender()" onkeydown="handleSearchKeyDown(event)" placeholder="搜尋 RUN CARD、機台、線別、穴號、操作員、備註..." class="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-4 py-2 text-sm text-slate-200 focus:outline-none focus:border-blue-500 transition">
          </form>
        </div>

        <div class="flex items-center space-x-3 text-sm">
          <label class="text-slate-400">班別篩選:</label>
          <select id="shiftFilter" onchange="resetPageAndRender()" class="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-blue-500">
            <option value="ALL">全部班別</option>
            <option value="D">D 班</option>
            <option value="N">N 班</option>
          </select>

          <label class="text-slate-400 ml-2">每頁顯示:</label>
          <select id="pageSizeSelect" onchange="changePageSize(this.value)" class="bg-slate-950 border border-slate-800 rounded-lg px-2 py-1.5 text-slate-200 focus:outline-none focus:border-blue-500 text-xs">
            <option value="20">20 筆</option>
            <option value="50" selected>50 筆</option>
            <option value="100">100 筆</option>
            <option value="200">200 筆</option>
            <option value="500">500 筆</option>
          </select>

          <button onclick="resetData()" class="text-slate-400 hover:text-slate-200 px-3 py-2 rounded-lg border border-slate-800 hover:bg-slate-800 transition flex items-center space-x-1">
            <i class="fa-solid fa-rotate-left"></i>
            <span>重置預設數據</span>
          </button>
        </div>
      </div>

      <!-- Data Table Scroll Container -->
      <div class="overflow-x-auto overflow-y-auto max-h-[500px]">
        <table class="w-full text-xs text-left border-collapse whitespace-nowrap">
          <thead class="bg-slate-950 text-slate-300 font-semibold sticky top-0 z-20 border-b border-slate-800">
            <tr>
              <th class="p-2 border-r border-slate-800 text-center sticky-col bg-slate-950">操作</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">班別</th>
              
              <!-- Sorted Header: 前段生產日期 -->
              <th onclick="toggleSort('preDate')" class="p-2 border-r border-slate-800 bg-amber-900/40 text-amber-200 cursor-pointer hover:bg-amber-800/50 transition select-none">
                <div class="flex items-center justify-between space-x-1">
                  <span>前段生產日期</span>
                  <i id="sortIcon-preDate" class="fa-solid fa-sort text-amber-400/60"></i>
                </div>
              </th>

              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">穴號</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">線別</th>
              <th class="p-2 border-r border-slate-800 bg-amber-400 text-slate-900 font-bold">RUN CARD</th>

              <!-- Sorted Header: 日期 -->
              <th onclick="toggleSort('date')" class="p-2 border-r border-slate-800 bg-amber-900/40 text-amber-200 cursor-pointer hover:bg-amber-800/50 transition select-none">
                <div class="flex items-center justify-between space-x-1">
                  <span>日期</span>
                  <i id="sortIcon-date" class="fa-solid fa-sort text-amber-400/60"></i>
                </div>
              </th>

              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">機台</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">開始時間</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">結束時間</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200">操作員</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">投入數</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">圖紋種類 Pattern type</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">總數</th>
              <th class="p-2 border-r border-slate-800 bg-emerald-900/40 text-emerald-300">良率</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">入片數</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">AOI-OK</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">AOI-OK入片</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">AOI-NG</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">AOI總不良率</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">直通率</th>
              <th class="p-2 border-r border-slate-800 bg-cyan-900/40 text-cyan-300">CPK</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">備註</th>
              <th class="p-2 border-r border-slate-800 bg-emerald-900/40 text-emerald-200">度數</th>
              <!-- Defect Header Group -->
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">1.留公</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">2.油殘 YOSAI</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">3.圖紋密</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">4.暗視野</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">5.積墨 CHIMO</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">6.邊破</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">7.氣泡 CHIPA</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">8.附加物</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">混料:穴號:顆粒</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">混料:顆粒</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">混入穴號</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">混料誤判</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">空洞 EMPTY</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">反面 TAOBAO</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">穴外亮點 OUTSIDE</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">疊片 DOUBLE</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">碎片 POPIE</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">鏡片偏移</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300">留RING</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200">備註</th>
            </tr>
          </thead>
          <tbody id="dataTableBody" class="divide-y divide-slate-800 text-slate-300">
            <!-- Dynamic rows rendered here -->
          </tbody>
        </table>
      </div>

      <!-- Pagination Toolbar Footer -->
      <div class="p-3 bg-slate-950 border-t border-slate-800 text-xs text-slate-400 flex flex-wrap gap-3 justify-between items-center">
        <span id="rowCountInfo">顯示 0 筆資料</span>

        <!-- Page Controls -->
        <div class="flex items-center space-x-1" id="paginationControls">
          <!-- Rendered dynamically -->
        </div>
      </div>
    </div>
  </main>

  <!-- Add/Edit Record Modal -->
  <div id="dataModal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-4 z-50 overflow-y-auto">
    <div class="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[90vh] flex flex-col shadow-2xl">
      <!-- Modal Header -->
      <div class="px-6 py-4 border-b border-slate-800 flex justify-between items-center">
        <h3 class="text-base font-bold text-slate-100 flex items-center space-x-2">
          <i class="fa-solid fa-pen-to-square text-blue-500"></i>
          <span id="modalTitle">生產紀錄</span>
        </h3>
        <button onclick="closeModal()" class="text-slate-400 hover:text-white text-lg px-2">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <!-- Modal Body Form -->
      <form id="recordForm" onsubmit="handleFormSubmit(event)" class="p-6 overflow-y-auto space-y-6 flex-1 text-xs">
        <!-- Section 1: Basic Production Details -->
        <div>
          <h4 class="text-xs font-bold text-blue-400 uppercase tracking-wider mb-3 flex items-center space-x-2">
            <i class="fa-solid fa-info-circle"></i>
            <span>生產基本資訊</span>
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label class="block text-slate-400 mb-1">班別</label>
              <select name="shift" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
                <option value="">--請選擇--</option>
                <option value="D">D 班</option>
                <option value="N">N 班</option>
              </select>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">前段生產日期 (滾輪)</label>
              <input type="date" name="preDate" onwheel="handleDateWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">穴號</label>
              <input type="number" name="holeNo" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">線別</label>
              <select name="lineNo" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
                <option value="">--請選擇--</option>
                <option value="107">107</option>
                <option value="108">108</option>
                <option value="109">109</option>
                <option value="110">110</option>
                <option value="111">111</option>
                <option value="112">112</option>
                <option value="113">113</option>
                <option value="114">114</option>
                <option value="115">115</option>
                <option value="116">116</option>
                <option value="117">117</option>
                <option value="118">118</option>
                <option value="119">119</option>
                <option value="120">120</option>
                <option value="121">121</option>
              </select>
            </div>
            <div class="col-span-2">
              <label class="block text-amber-400 font-semibold mb-1">RUN CARD</label>
              <input type="text" name="runCard" required placeholder="" class="w-full bg-slate-950 border border-amber-500/50 rounded p-2 text-amber-200 font-mono">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">日期 (滾輪)</label>
              <input type="date" name="date" onwheel="handleDateWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">機台</label>
              <select name="machine" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
                <option value="">--請選擇--</option>
                <option value="116">116</option>
                <option value="117">117</option>
                <option value="118">118</option>
                <option value="119">119</option>
                <option value="120">120</option>
                <option value="121">121</option>
                <option value="122">122</option>
                <option value="123">123</option>
                <option value="124">124</option>
                <option value="125">125</option>
                <option value="126">126</option>
                <option value="127">127</option>
                <option value="128">128</option>
                <option value="129">129</option>
                <option value="130">130</option>
                <option value="131">131</option>
                <option value="132">132</option>
                <option value="133">133</option>
                <option value="134">134</option>
                <option value="135">135</option>
                <option value="136">136</option>
                <option value="137">137</option>
                <option value="138">138</option>
                <option value="139">139</option>
                <option value="140">140</option>
              </select>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">開始時間 (滾輪)</label>
              <input type="time" name="startTime" onwheel="handleTimeWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">結束時間 (滾輪)</label>
              <input type="time" name="endTime" onwheel="handleTimeWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">操作員</label>
              <input type="text" name="operator" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">圖紋種類 Pattern</label>
              <input type="text" name="patternType" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
          </div>
        </div>

        <!-- Section 2: Quantities and Measurements -->
        <div>
          <h4 class="text-xs font-bold text-emerald-400 uppercase tracking-wider mb-3 flex items-center space-x-2">
            <i class="fa-solid fa-calculator"></i>
            <span>數量與檢測指標</span>
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label class="block text-slate-400 mb-1">投入數</label>
              <input type="number" id="inputNum" name="inputNum" placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">總數</label>
              <input type="number" id="totalNum" name="totalNum" placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-emerald-400 mb-1">良率</label>
              <input type="text" id="yieldVal" name="yieldVal" placeholder="" readonly class="w-full bg-slate-950/60 border border-slate-800 rounded p-2 text-emerald-400 font-bold cursor-not-allowed">
              <span id="yieldWarning" class="text-rose-400 text-[10px] hidden font-semibold mt-0.5 block">⚠️ 良率不能大於 100%</span>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">入片數</label>
              <input type="number" id="inPiece" name="inPiece" placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-OK</label>
              <input type="number" id="aoiOk" name="aoiOk" placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-OK入片</label>
              <input type="number" name="aoiOkIn" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-NG</label>
              <input type="number" id="aoiNg" name="aoiNg" placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-rose-400 mb-1">AOI總不良率</label>
              <input type="text" id="aoiTotalNgYield" name="aoiTotalNgYield" placeholder="" readonly class="w-full bg-slate-950/60 border border-slate-800 rounded p-2 text-rose-400 font-bold cursor-not-allowed">
            </div>
            <div>
              <label class="block text-cyan-400 mb-1">直通率</label>
              <input type="text" id="fpynVal" name="fpynVal" placeholder="" readonly class="w-full bg-slate-950/60 border border-slate-800 rounded p-2 text-cyan-400 font-bold cursor-not-allowed">
              <span id="fpynWarning" class="text-rose-400 text-[10px] hidden font-semibold mt-0.5 block">⚠️ 直通率不能大於 100%</span>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">CPK</label>
              <input type="text" id="cpkVal" name="cpk" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">備註</label>
              <input type="text" name="ipqcNote" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">度數</label>
              <input type="text" name="degree" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
          </div>
        </div>

        <!-- Section 3: Defects Count -->
        <div>
          <h4 class="text-xs font-bold text-rose-400 uppercase tracking-wider mb-3 flex items-center space-x-2">
            <i class="fa-solid fa-bug"></i>
            <span>瑕疵項目統計 (數量)</span>
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-950/40 p-4 border border-slate-800/80 rounded-xl">
            <div><label class="block text-slate-400 mb-1">1.留公</label><input type="number" name="d_liugong" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">2.油殘 YOSAI</label><input type="number" name="d_yosai" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">3.圖紋密</label><input type="number" name="d_pattern_dense" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">4.暗視野</label><input type="number" name="d_darkfield" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">5.積墨 CHIMO</label><input type="number" name="d_chimo" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">6.邊破</label><input type="number" name="d_edge_break" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">7.氣泡 CHIPA</label><input type="number" name="d_chipa" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">8.附加物</label><input type="number" name="d_attach" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">空洞 EMPTY</label><input type="number" name="d_empty" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">反面 TAOBAO</label><input type="number" name="d_taobao" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">穴外亮點 OUTSIDE</label><input type="number" name="d_outside" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">疊片 DOUBLE</label><input type="number" name="d_double" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">碎片 POPIE</label><input type="number" name="d_popie" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">鏡片偏移</label><input type="number" name="d_shift" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">留RING</label><input type="number" name="d_ring" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">通用備註</label><input type="text" name="note" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
          </div>
        </div>

        <input type="hidden" name="editingId" id="editingId" value="">

        <!-- Modal Footer Actions -->
        <div class="flex justify-end space-x-3 pt-4 border-t border-slate-800">
          <button type="button" onclick="closeModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition">取消</button>
          <button type="submit" id="saveSubmitBtn" class="px-5 py-2 bg-blue-600 hover:bg-blue-500 text-white font-medium rounded-lg shadow-lg shadow-blue-900/30 transition flex items-center space-x-2">
            <i class="fa-solid fa-floppy-disk"></i>
            <span>儲存修改</span>
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- Loading Overlay with Gojo Hollow Purple (虚式「茈」) Progress Bar -->
  <div id="loadingOverlay" class="fixed inset-0 bg-slate-950/85 backdrop-blur-md hidden flex flex-col items-center justify-center z-50 p-4">
    <div class="bg-slate-900/90 border border-purple-500/40 p-6 rounded-2xl shadow-[0_0_50px_rgba(168,85,247,0.3)] max-w-md w-full flex flex-col items-center space-y-5 text-center relative overflow-hidden">
      
      <!-- Gojo Theme Header & Glowing Orb Representation -->
      <div class="relative flex items-center justify-center w-20 h-20 rounded-full bg-purple-950/50 border border-purple-400/60 shadow-[0_0_25px_rgba(217,70,239,0.5)]">
        <div class="absolute inset-0 rounded-full bg-gradient-to-tr from-blue-600 via-purple-600 to-fuchsia-500 opacity-75 animate-pulse"></div>
        <i class="fa-solid fa-bolt text-2xl text-white relative z-10 drop-shadow-[0_0_10px_#ffffff]"></i>
      </div>

      <div class="space-y-1">
        <h4 class="text-base font-black text-transparent bg-clip-text bg-gradient-to-r from-blue-400 via-purple-300 to-fuchsia-400 tracking-wider">
          虚式「 茈 」發射中...
        </h4>
        <p class="text-xs text-slate-400" id="loadingDesc">正在解析巨量光學生產數據...</p>
      </div>

      <!-- Custom Progress Bar -->
      <div class="w-full space-y-2">
        <div class="w-full bg-slate-950 border border-purple-900/60 rounded-full h-3.5 p-0.5 overflow-hidden shadow-inner">
          <div id="hollowPurpleProgressBar" class="h-full rounded-full hollow-purple-bar transition-all duration-300 ease-out" style="width: 0%;"></div>
        </div>
        <div class="flex justify-between items-center text-[11px] font-mono text-purple-300 px-1">
          <span id="loadingStepText">領域展開：資料讀取</span>
          <span id="loadingPercentText">0%</span>
        </div>
      </div>

    </div>
  </div>

  <!-- Notification Message Box -->
  <div id="toast" class="fixed bottom-5 right-5 bg-slate-800 text-slate-100 border border-slate-700 px-4 py-3 rounded-xl shadow-2xl hidden flex items-center space-x-3 z-50">
    <i class="fa-solid fa-circle-check text-emerald-400 text-lg" id="toastIcon"></i>
    <span id="toastMsg" class="text-sm">操作成功！</span>
  </div>

  <script>
    // Initial Dataset based exactly on image structure and content
    const defaultData = [
      {
        id: 1, shift: 'D', preDate: '2026-10-20', holeNo: '10', lineNo: '109', runCard: 'v5121-261004078-015.1',
        date: '2026-10-23', machine: '122', startTime: '03:37', endTime: '08:16', operator: 'VF1338',
        inputNum: 4095, patternType: 'CF55T V2-2208', totalNum: 4095, yieldVal: '100%', inPiece: 4092,
        aoiOk: 3785, aoiOkIn: 3785, aoiNg: 307, aoiTotalNgYield: '7.5%', fpynVal: '92.4%',
        cpk: '', ipqcNote: '烤2小時,靜置12小時, IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: 16,
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 2, shift: 'D', preDate: '2026-10-20', holeNo: '10', lineNo: '109', runCard: 'v5121-261003904-011.1',
        date: '2026-10-23', machine: '117', startTime: '03:39', endTime: '08:22', operator: 'PC2388',
        inputNum: 4088, patternType: 'CF55T V2-2238', totalNum: 4088, yieldVal: '100%', inPiece: 4079,
        aoiOk: 3959, aoiOkIn: 3959, aoiNg: 120, aoiTotalNgYield: '2.9%', fpynVal: '96.8%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: 2,
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 3, shift: 'D', preDate: '2026-10-20', holeNo: '10', lineNo: '109', runCard: 'v5121-261003903-010.1',
        date: '2026-10-23', machine: '125', startTime: '04:45', endTime: '08:22', operator: 'FT0005',
        inputNum: 4096, patternType: 'CF55T V2-2238', totalNum: 4096, yieldVal: '98.1%', inPiece: 4019,
        aoiOk: 3855, aoiOkIn: 3855, aoiNg: 164, aoiTotalNgYield: '4.1%', fpynVal: '94.1%',
        cpk: '', ipqcNote: '烤2小時,靜置12小時, IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: '',
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 4, shift: 'D', preDate: '2026-10-20', holeNo: '15', lineNo: '108', runCard: 'v5121-261004877-010.1',
        date: '2026-10-20', machine: '120', startTime: '05:08', endTime: '09:02', operator: 'VF1338',
        inputNum: 4093, patternType: 'CF55T V2-2209', totalNum: 4093, yieldVal: '99.2%', inPiece: 4059,
        aoiOk: 3891, aoiOkIn: 3891, aoiNg: 168, aoiTotalNgYield: '4.1%', fpynVal: '95.1%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: 18,
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 5, shift: 'D', preDate: '2026-10-20', holeNo: '11', lineNo: '112', runCard: 'v5121-261004877-009.1',
        date: '2026-10-20', machine: '116', startTime: '05:01', endTime: '09:17', operator: 'FW0298',
        inputNum: 4096, patternType: 'CF55T V2-2209', totalNum: 4096, yieldVal: '99.8%', inPiece: 4088,
        aoiOk: 3944, aoiOkIn: 3921, aoiNg: 144, aoiTotalNgYield: '3.5%', fpynVal: '96.3%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: 1,
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 6, shift: 'D', preDate: '2026-10-20', holeNo: '10', lineNo: '112', runCard: 'v5121-261004963-006.1',
        date: '2026-10-20', machine: '123', startTime: '07:18', endTime: '09:31', operator: 'FS0162',
        inputNum: 2059, patternType: 'MLK-2419', totalNum: 2059, yieldVal: '99.1%', inPiece: 2040,
        aoiOk: 1938, aoiOkIn: 1938, aoiNg: 102, aoiTotalNgYield: '5.0%', fpynVal: '94.1%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: '',
        d_chipa: '', d_attach: 5, d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 7, shift: 'D', preDate: '2026-10-21', holeNo: '10', lineNo: '109', runCard: 'v5121-261003117-012.1',
        date: '2026-10-21', machine: '126', startTime: '06:30', endTime: '10:15', operator: 'PC2515',
        inputNum: 4096, patternType: 'CF55T V2-2125', totalNum: 4096, yieldVal: '99.8%', inPiece: 4087,
        aoiOk: 3847, aoiOkIn: 3847, aoiNg: 240, aoiTotalNgYield: '5.9%', fpynVal: '93.9%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: '',
        d_chipa: '', d_attach: 5, d_mix_hole_particle: 2, d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      },
      {
        id: 8, shift: 'D', preDate: '2026-10-21', holeNo: '9', lineNo: '109', runCard: 'v5121-261003116-011.1',
        date: '2026-10-21', machine: '118', startTime: '06:53', endTime: '10:31', operator: 'VF1338',
        inputNum: 4098, patternType: 'CF55T V2-2125', totalNum: 4098, yieldVal: '99.9%', inPiece: 4092,
        aoiOk: 3936, aoiOkIn: 3936, aoiNg: 156, aoiTotalNgYield: '3.8%', fpynVal: '96.0%',
        cpk: '', ipqcNote: 'IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: '',
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      }
    ];

    // Load dataset from localStorage if available, otherwise defaultData
    let dataset = [];
    try {
      const saved = localStorage.getItem('optics_production_dataset');
      if (saved) {
        dataset = JSON.parse(saved);
      } else {
        dataset = [...defaultData];
      }
    } catch (e) {
      dataset = [...defaultData];
    }

    // Save dataset to localStorage
    function saveDatasetToStorage() {
      try {
        localStorage.setItem('optics_production_dataset', JSON.stringify(dataset));
      } catch (e) {
        console.error('LocalStorage save failed:', e);
      }
    }

    let lineChartObj = null;
    let machineChartObj = null;

    // Pagination State
    let currentPage = 1;
    let pageSize = 50;

    // Sorting State
    let sortColumn = null; // 'preDate' or 'date'
    let sortDirection = null; // 'asc' or 'desc'

    // Defect Mapping Keys & Labels
    const defectKeys = [
      { key: 'd_liugong', label: '1.留公' },
      { key: 'd_yosai', label: '2.油殘 YOSAI' },
      { key: 'd_pattern_dense', label: '3.圖紋密' },
      { key: 'd_darkfield', label: '4.暗視野' },
      { key: 'd_chimo', label: '5.積墨 CHIMO' },
      { key: 'd_edge_break', label: '6.邊破' },
      { key: 'd_chipa', label: '7.氣泡 CHIPA' },
      { key: 'd_attach', label: '8.附加物' },
      { key: 'd_mix_hole_particle', label: '混料:穴號:顆粒' },
      { key: 'd_mix_particle', label: '混料:顆粒' },
      { key: 'd_mix_hole', label: '混入穴號' },
      { key: 'd_mix_false', label: '混料誤判' },
      { key: 'd_empty', label: '空洞 EMPTY' },
      { key: 'd_taobao', label: '反面 TAOBAO' },
      { key: 'd_outside', label: '穴外亮點 OUTSIDE' },
      { key: 'd_double', label: '疊片 DOUBLE' },
      { key: 'd_popie', label: '碎片 POPIE' },
      { key: 'd_shift', label: '鏡片偏移' },
      { key: 'd_ring', label: '留RING' }
    ];

    // Initialize Page
    window.onload = function() {
      if (typeof ChartDataLabels !== 'undefined') {
        Chart.register(ChartDataLabels);
      }
      updateHoleFilterOptions();
      renderTable();
      initCharts();
      updateDashboard();
    };

    // Populate Hole Filter Dropdown dynamically from dataset
    function updateHoleFilterOptions() {
      const holeSelect = document.getElementById('chartHoleFilter');
      if (!holeSelect) return;
      const currentSelected = holeSelect.value;

      // Extract unique holes
      const holesSet = new Set();
      dataset.forEach(d => {
        if (d.holeNo !== undefined && d.holeNo !== null && String(d.holeNo).trim() !== '') {
          holesSet.add(String(d.holeNo).trim());
        }
      });

      // Sort holes numerically if possible
      const sortedHoles = Array.from(holesSet).sort((a, b) => {
        const na = parseFloat(a);
        const nb = parseFloat(b);
        if (!isNaN(na) && !isNaN(nb)) return na - nb;
        return a.localeCompare(b);
      });

      let html = '<option value="ALL">全部穴號</option>';
      sortedHoles.forEach(h => {
        html += `<option value="${h}">穴號 ${h}</option>`;
      });

      holeSelect.innerHTML = html;
      if (sortedHoles.includes(currentSelected) || currentSelected === 'ALL') {
        holeSelect.value = currentSelected;
      } else {
        holeSelect.value = 'ALL';
      }
    }

    // Helper to format date display
    function formatDateDisplay(dateStr) {
      if (!dateStr) return '';
      if (dateStr.includes('-')) {
        const parts = dateStr.split('-');
        if (parts.length === 3) {
          return `${parseInt(parts[1], 10)}月${parseInt(parts[2], 10)}日`;
        }
      }
      return dateStr;
    }

    // Helper to convert date string to normalized timestamp for comparisons
    function parseNormalizedDateTimestamp(dateStr) {
      if (!dateStr) return 0;
      const str = String(dateStr).trim();
      
      if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        return new Date(str + 'T00:00:00').getTime();
      }
      
      if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(str)) {
        const parts = str.split('/');
        const y = parts[0];
        const m = parts[1].padStart(2, '0');
        const d = parts[2].padStart(2, '0');
        return new Date(`${y}-${m}-${d}T00:00:00`).getTime();
      }

      const matchCn = str.match(/(?:(\d{4})年)?\s*(\d{1,2})月\s*(\d{1,2})日/);
      if (matchCn) {
        const currentYear = new Date().getFullYear();
        const y = matchCn[1] || currentYear;
        const m = matchCn[2].padStart(2, '0');
        const d = matchCn[3].padStart(2, '0');
        return new Date(`${y}-${m}-${d}T00:00:00`).getTime();
      }

      const parsed = Date.parse(str);
      return isNaN(parsed) ? 0 : parsed;
    }

    // Toggle Sorting on Table Headers
    function toggleSort(columnKey) {
      if (sortColumn === columnKey) {
        if (sortDirection === 'asc') {
          sortDirection = 'desc';
        } else if (sortDirection === 'desc') {
          sortColumn = null;
          sortDirection = null;
        }
      } else {
        sortColumn = columnKey;
        sortDirection = 'asc';
      }

      updateSortIcons();
      currentPage = 1;
      renderTable();
    }

    function updateSortIcons() {
      const preDateIcon = document.getElementById('sortIcon-preDate');
      const dateIcon = document.getElementById('sortIcon-date');

      if (preDateIcon) preDateIcon.className = 'fa-solid fa-sort text-amber-400/60';
      if (dateIcon) dateIcon.className = 'fa-solid fa-sort text-amber-400/60';

      if (sortColumn === 'preDate' && preDateIcon) {
        preDateIcon.className = sortDirection === 'asc' ? 'fa-solid fa-sort-up text-amber-300 font-bold' : 'fa-solid fa-sort-down text-amber-300 font-bold';
      } else if (sortColumn === 'date' && dateIcon) {
        dateIcon.className = sortDirection === 'asc' ? 'fa-solid fa-sort-up text-amber-300 font-bold' : 'fa-solid fa-sort-down text-amber-300 font-bold';
      }
    }

    function resetPageAndRender() {
      currentPage = 1;
      renderTable();
    }

    function changePageSize(newSize) {
      pageSize = parseInt(newSize, 10) || 50;
      currentPage = 1;
      renderTable();
    }

    function goToPage(p) {
      currentPage = p;
      renderTable();
    }

    function handleSearchSubmit(e) {
      e.preventDefault();
      resetPageAndRender();
      document.getElementById('searchInput').blur();
    }

    function handleSearchKeyDown(e) {
      if (e.key === 'Enter') {
        e.preventDefault();
        resetPageAndRender();
        e.target.blur();
      }
    }

    function handleDateWheel(e) {
      e.preventDefault();
      const input = e.target;
      if (!input.value) {
        const today = new Date().toISOString().slice(0, 10);
        input.value = today;
        input.dispatchEvent(new Event('change'));
        return;
      }
      
      const current = new Date(input.value);
      if (isNaN(current.getTime())) return;

      if (e.deltaY < 0) {
        current.setDate(current.getDate() + 1);
      } else {
        current.setDate(current.getDate() - 1);
      }

      input.value = current.toISOString().slice(0, 10);
      input.dispatchEvent(new Event('change'));
    }

    function handleTimeWheel(e) {
      e.preventDefault();
      const input = e.target;
      let hh = 8, mm = 0;
      if (input.value) {
        const parts = input.value.split(':');
        hh = parseInt(parts[0], 10) || 0;
        mm = parseInt(parts[1], 10) || 0;
      }

      let totalMinutes = hh * 60 + mm;
      if (e.deltaY < 0) {
        totalMinutes = (totalMinutes + 15) % 1440;
      } else {
        totalMinutes = (totalMinutes - 15 + 1440) % 1440;
      }

      const newH = String(Math.floor(totalMinutes / 60)).padStart(2, '0');
      const newM = String(totalMinutes % 60).padStart(2, '0');
      input.value = `${newH}:${newM}`;
    }

    function getFilteredAndSortedData() {
      const rawSearch = document.getElementById('searchInput').value.trim();
      const search = rawSearch.toLowerCase();
      const shiftFilter = document.getElementById('shiftFilter').value;

      let result = dataset.filter(item => {
        const matchesShift = shiftFilter === 'ALL' || item.shift === shiftFilter;
        if (!search) return matchesShift;

        const searchableText = [
          item.runCard,
          item.machine,
          item.lineNo,
          item.holeNo,
          item.shift,
          item.operator,
          item.patternType,
          item.ipqcNote,
          item.note,
          item.preDate,
          item.date
        ].map(v => (v !== undefined && v !== null) ? String(v).toLowerCase() : '').join(' ');

        return matchesShift && searchableText.includes(search);
      });

      if (sortColumn && sortDirection) {
        result.sort((a, b) => {
          const tsA = parseNormalizedDateTimestamp(a[sortColumn]);
          const tsB = parseNormalizedDateTimestamp(b[sortColumn]);
          return sortDirection === 'asc' ? tsA - tsB : tsB - tsA;
        });
      }

      return result;
    }

    function renderTable() {
      const tbody = document.getElementById('dataTableBody');
      const filteredSorted = getFilteredAndSortedData();
      const totalCount = filteredSorted.length;
      const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));

      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;

      const startIndex = (currentPage - 1) * pageSize;
      const endIndex = Math.min(startIndex + pageSize, totalCount);
      const displayList = filteredSorted.slice(startIndex, endIndex);

      document.getElementById('rowCountInfo').innerText = `顯示第 ${totalCount > 0 ? startIndex + 1 : 0} 至 ${endIndex} 筆 (符合條件 ${totalCount} 筆，總計 ${dataset.length} 筆)`;

      tbody.innerHTML = displayList.map(row => `
        <tr class="hover:bg-slate-800/60 transition group">
          <td class="p-2 border-r border-slate-800 text-center sticky-col bg-slate-900 group-hover:bg-slate-800">
            <div class="flex items-center justify-center space-x-1">
              <button onclick="editRow(${row.id})" class="p-1 text-blue-400 hover:text-blue-300 transition" title="編輯修改">
                <i class="fa-solid fa-pen-to-square"></i>
              </button>
              <button onclick="deleteRow(${row.id})" class="p-1 text-rose-400 hover:text-rose-300 transition" title="刪除">
                <i class="fa-solid fa-trash-can"></i>
              </button>
            </div>
          </td>
          <td class="p-2 border-r border-slate-800 text-center font-bold text-amber-300">${row.shift || ''}</td>
          <td class="p-2 border-r border-slate-800">${formatDateDisplay(row.preDate)}</td>
          <td class="p-2 border-r border-slate-800 text-center font-bold text-amber-200">${row.holeNo || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-semibold text-blue-300">${row.lineNo || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono font-bold text-amber-300 bg-amber-950/10">${row.runCard || ''}</td>
          <td class="p-2 border-r border-slate-800">${formatDateDisplay(row.date)}</td>
          <td class="p-2 border-r border-slate-800 text-center font-semibold text-slate-200">${row.machine || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-400">${row.startTime || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-400">${row.endTime || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.operator || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-slate-200">${row.inputNum || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-slate-300">${row.patternType || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-slate-200">${row.totalNum || ''}</td>
          <td class="p-2 border-r border-slate-800 font-bold text-emerald-400 bg-emerald-950/20 text-center">${row.yieldVal || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.inPiece || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.aoiOk || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.aoiOkIn || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-rose-400 font-semibold">${row.aoiNg || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.aoiTotalNgYield || ''}</td>
          <td class="p-2 border-r border-slate-800 font-bold text-cyan-400 bg-cyan-950/20 text-center">${row.fpynVal || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.cpk || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-300">${row.ipqcNote || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center">${row.degree || ''}</td>
          
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_liugong ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_liugong || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_yosai ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_yosai || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_pattern_dense ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_pattern_dense || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_darkfield ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_darkfield || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_chimo ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_chimo || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_edge_break ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_edge_break || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_chipa ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_chipa || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_attach ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_attach || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_mix_hole_particle ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_mix_hole_particle || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_mix_particle ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_mix_particle || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_mix_hole ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_mix_hole || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_mix_false ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_mix_false || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_empty ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_empty || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_taobao ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_taobao || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_outside ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_outside || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_double ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_double || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_popie ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_popie || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_shift ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_shift || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono ${row.d_ring ? 'text-rose-400 font-bold bg-rose-950/30' : 'text-slate-600'}">${row.d_ring || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-400">${row.note || ''}</td>
        </tr>
      `).join('');

      renderPaginationControls(totalPages);
    }

    function renderPaginationControls(totalPages) {
      const container = document.getElementById('paginationControls');
      if (!container) return;

      let html = '';
      html += `
        <button onclick="goToPage(1)" ${currentPage === 1 ? 'disabled' : ''} class="px-2 py-1 rounded border border-slate-800 bg-slate-900 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition">
          <i class="fa-solid fa-angles-left"></i>
        </button>
        <button onclick="goToPage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''} class="px-2.5 py-1 rounded border border-slate-800 bg-slate-900 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition">
          <i class="fa-solid fa-angle-left"></i>
        </button>
      `;

      let startP = Math.max(1, currentPage - 2);
      let endP = Math.min(totalPages, startP + 4);
      if (endP - startP < 4) {
        startP = Math.max(1, endP - 4);
      }

      for (let p = startP; p <= endP; p++) {
        const isActive = p === currentPage;
        html += `
          <button onclick="goToPage(${p})" class="px-3 py-1 rounded text-xs font-semibold border transition ${isActive ? 'bg-blue-600 border-blue-500 text-white shadow-md' : 'bg-slate-900 border-slate-800 text-slate-300 hover:bg-slate-800'}">
            ${p}
          </button>
        `;
      }

      html += `
        <button onclick="goToPage(${currentPage + 1})" ${currentPage === totalPages ? 'disabled' : ''} class="px-2.5 py-1 rounded border border-slate-800 bg-slate-900 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition">
          <i class="fa-solid fa-angle-right"></i>
        </button>
        <button onclick="goToPage(${totalPages})" ${currentPage === totalPages ? 'disabled' : ''} class="px-2 py-1 rounded border border-slate-800 bg-slate-900 hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition">
          <i class="fa-solid fa-angles-right"></i>
        </button>
        <span class="ml-2 text-slate-500 text-[11px]">第 ${currentPage} / ${totalPages} 頁</span>
      `;

      container.innerHTML = html;
    }

    function resetChartFilters() {
      document.getElementById('chartHoleFilter').value = 'ALL';
      document.getElementById('chartStartDate').value = '';
      document.getElementById('chartEndDate').value = '';
      updateDashboard();
    }

    function updateDashboard() {
      const holeFilterVal = document.getElementById('chartHoleFilter').value;
      const startDateVal = document.getElementById('chartStartDate').value;
      const endDateVal = document.getElementById('chartEndDate').value;

      const startTimestamp = startDateVal ? new Date(startDateVal + 'T00:00:00').getTime() : null;
      const endTimestamp = endDateVal ? new Date(endDateVal + 'T23:59:59').getTime() : null;

      const analyticsDataset = dataset.filter(d => {
        // Hole filter check
        if (holeFilterVal !== 'ALL') {
          if (String(d.holeNo).trim() !== String(holeFilterVal).trim()) return false;
        }

        // Date filter check
        const itemDateStr = d.date || d.preDate;
        if (!itemDateStr) return true;

        const itemTs = parseNormalizedDateTimestamp(itemDateStr);
        if (itemTs === 0) return true;

        if (startTimestamp && itemTs < startTimestamp) return false;
        if (endTimestamp && itemTs > endTimestamp) return false;
        return true;
      });

      const totalInput = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.inputNum) || 0), 0);
      const totalAoiNg = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.aoiNg) || 0), 0);

      const sumInPiece = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.inPiece) || 0), 0);
      const sumTotalNum = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.totalNum) || 0), 0);
      const avgYield = sumTotalNum > 0 ? ((sumInPiece / sumTotalNum) * 100).toFixed(1) : 0;

      const sumAoiOk = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.aoiOk) || 0), 0);
      const avgFpyn = sumTotalNum > 0 ? ((sumAoiOk / sumTotalNum) * 100).toFixed(1) : 0;

      const avgAoiNgRate = sumInPiece > 0 ? ((totalAoiNg / sumInPiece) * 100).toFixed(1) : 0;

      const defectTotals = {};
      defectKeys.forEach(dk => {
        defectTotals[dk.label] = analyticsDataset.reduce((sum, item) => sum + (Number(item[dk.key]) || 0), 0);
      });
      const sortedDefects = Object.entries(defectTotals).sort((a, b) => b[1] - a[1]);
      const topDefect = sortedDefects[0] || ['無數據', 0];

      document.getElementById('kpi-total-input').innerText = totalInput.toLocaleString();
      document.getElementById('kpi-avg-yield').innerText = `${avgYield}%`;
      document.getElementById('kpi-avg-fpyn').innerText = `${avgFpyn}%`;
      document.getElementById('kpi-aoi-ng-rate').innerText = `${avgAoiNgRate}%`;
      document.getElementById('kpi-top-defect').innerText = topDefect[1] > 0 ? topDefect[0] : '無瑕疵';
      document.getElementById('kpi-top-defect-count').innerText = `累計 ${topDefect[1]} 件`;

      const lineMap = {};
      analyticsDataset.forEach(d => {
        const line = d.lineNo ? `線別 ${d.lineNo}` : '未分類';
        if (!lineMap[line]) {
          lineMap[line] = { totalNum: 0, inPiece: 0, aoiOk: 0, aoiNg: 0 };
        }
        lineMap[line].totalNum += Number(d.totalNum) || 0;
        lineMap[line].inPiece += Number(d.inPiece) || 0;
        lineMap[line].aoiOk += Number(d.aoiOk) || 0;
        lineMap[line].aoiNg += Number(d.aoiNg) || 0;
      });

      const lineLabels = Object.keys(lineMap).sort();
      const lineYields = [], lineFpys = [], lineDefects = [];
      lineLabels.forEach(l => {
        const m = lineMap[l];
        const y = m.totalNum > 0 ? (m.inPiece / m.totalNum * 100) : 0;
        const f = m.totalNum > 0 ? (m.aoiOk / m.totalNum * 100) : 0;
        const d = m.inPiece > 0 ? (m.aoiNg / m.inPiece * 100) : 0;
        lineYields.push(parseFloat(y.toFixed(1)));
        lineFpys.push(parseFloat(f.toFixed(1)));
        lineDefects.push(parseFloat(d.toFixed(1)));
      });

      const machineMap = {};
      analyticsDataset.forEach(d => {
        const mac = d.machine ? `機台 ${d.machine}` : '未分類';
        if (!machineMap[mac]) {
          machineMap[mac] = { totalNum: 0, inPiece: 0, aoiOk: 0, aoiNg: 0 };
        }
        machineMap[mac].totalNum += Number(d.totalNum) || 0;
        machineMap[mac].inPiece += Number(d.inPiece) || 0;
        machineMap[mac].aoiOk += Number(d.aoiOk) || 0;
        machineMap[mac].aoiNg += Number(d.aoiNg) || 0;
      });

      const machineLabels = Object.keys(machineMap).sort();
      const machineYields = [], machineFpys = [], machineDefects = [];
      machineLabels.forEach(mKey => {
        const m = machineMap[mKey];
        const y = m.totalNum > 0 ? (m.inPiece / m.totalNum * 100) : 0;
        const f = m.totalNum > 0 ? (m.aoiOk / m.totalNum * 100) : 0;
        const d = m.inPiece > 0 ? (m.aoiNg / m.inPiece * 100) : 0;
        machineYields.push(parseFloat(y.toFixed(1)));
        machineFpys.push(parseFloat(f.toFixed(1)));
        machineDefects.push(parseFloat(d.toFixed(1)));
      });

      if (lineChartObj) {
        lineChartObj.data.labels = lineLabels;
        lineChartObj.data.datasets[0].data = lineYields;
        lineChartObj.data.datasets[1].data = lineFpys;
        lineChartObj.data.datasets[2].data = lineDefects;
        lineChartObj.update();
      }

      if (machineChartObj) {
        machineChartObj.data.labels = machineLabels;
        machineChartObj.data.datasets[0].data = machineYields;
        machineChartObj.data.datasets[1].data = machineFpys;
        machineChartObj.data.datasets[2].data = machineDefects;
        machineChartObj.update();
      }
    }

    function initCharts() {
      const commonDatalabelOptions = {
        anchor: 'end',
        align: 'top',
        color: '#f1f5f9',
        font: {
          size: 9,
          weight: 'bold'
        },
        formatter: function(value) {
          return value !== null && value !== undefined && !isNaN(value) ? value + '%' : '';
        },
        offset: 1
      };

      const ctx1 = document.getElementById('lineYieldChart').getContext('2d');
      lineChartObj = new Chart(ctx1, {
        type: 'bar',
        data: {
          labels: [],
          datasets: [
            {
              label: '良率 (%)',
              data: [],
              backgroundColor: 'rgba(16, 185, 129, 0.75)',
              borderColor: '#10b981',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            },
            {
              label: '直通率 (%)',
              data: [],
              backgroundColor: 'rgba(6, 182, 212, 0.75)',
              borderColor: '#06b6d4',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            },
            {
              label: '不良率 (%)',
              data: [],
              backgroundColor: 'rgba(244, 63, 94, 0.75)',
              borderColor: '#f43f5e',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: { padding: { top: 22 } },
          plugins: {
            legend: { labels: { color: '#94a3b8', font: { size: 11 } } },
            datalabels: { display: true }
          },
          scales: {
            x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } },
            y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' }, max: 105, beginAtZero: true }
          }
        }
      });

      const ctx2 = document.getElementById('machineYieldChart').getContext('2d');
      machineChartObj = new Chart(ctx2, {
        type: 'bar',
        data: {
          labels: [],
          datasets: [
            {
              label: '良率 (%)',
              data: [],
              backgroundColor: 'rgba(16, 185, 129, 0.75)',
              borderColor: '#10b981',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            },
            {
              label: '直通率 (%)',
              data: [],
              backgroundColor: 'rgba(6, 182, 212, 0.75)',
              borderColor: '#06b6d4',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            },
            {
              label: '不良率 (%)',
              data: [],
              backgroundColor: 'rgba(244, 63, 94, 0.75)',
              borderColor: '#f43f5e',
              borderWidth: 1,
              borderRadius: 4,
              datalabels: commonDatalabelOptions
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          layout: { padding: { top: 22 } },
          plugins: {
            legend: { labels: { color: '#94a3b8', font: { size: 11 } } },
            datalabels: { display: true }
          },
          scales: {
            x: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' } },
            y: { ticks: { color: '#94a3b8', font: { size: 10 } }, grid: { color: '#1e293b' }, max: 105, beginAtZero: true }
          }
        }
      });
    }

    function addNewRowModal() {
      document.getElementById('recordForm').reset();
      document.getElementById('editingId').value = '';
      document.getElementById('modalTitle').innerText = '新增生產紀錄';
      calcMetrics();
      document.getElementById('dataModal').classList.remove('hidden');
    }

    function editRow(id) {
      const item = dataset.find(d => d.id === id);
      if (!item) return;

      const form = document.getElementById('recordForm');
      document.getElementById('editingId').value = item.id;
      document.getElementById('modalTitle').innerText = `修改生產紀錄 (${item.runCard || '項次 ' + id})`;

      for (let key in item) {
        if (form.elements[key]) {
          form.elements[key].value = item[key] || '';
        }
      }

      calcMetrics();
      document.getElementById('dataModal').classList.remove('hidden');
    }

    function closeModal() {
      document.getElementById('dataModal').classList.add('hidden');
    }

    function calcMetrics() {
      const totalNum = parseFloat(document.getElementById('totalNum').value) || 0;
      const inPiece = parseFloat(document.getElementById('inPiece').value) || 0;
      const aoiOk = parseFloat(document.getElementById('aoiOk').value) || 0;
      const aoiNg = parseFloat(document.getElementById('aoiNg').value) || 0;

      let yieldNum = 0;
      let fpynNum = 0;
      let hasError = false;

      const yieldField = document.getElementById('yieldVal');
      const yieldWarning = document.getElementById('yieldWarning');
      if (totalNum > 0 && inPiece >= 0) {
        yieldNum = (inPiece / totalNum) * 100;
        yieldField.value = yieldNum.toFixed(1) + '%';
        if (yieldNum > 100) {
          hasError = true;
          yieldField.classList.add('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
          yieldWarning.classList.remove('hidden');
        } else {
          yieldField.classList.remove('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
          yieldWarning.classList.add('hidden');
        }
      } else {
        yieldField.value = '';
        yieldField.classList.remove('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
        yieldWarning.classList.add('hidden');
      }

      const aoiTotalNgYieldField = document.getElementById('aoiTotalNgYield');
      if (inPiece > 0 && aoiNg >= 0) {
        const aoiTotalNgYield = ((aoiNg / inPiece) * 100).toFixed(1) + '%';
        aoiTotalNgYieldField.value = aoiTotalNgYield;
      } else {
        aoiTotalNgYieldField.value = '';
      }

      const fpynField = document.getElementById('fpynVal');
      const fpynWarning = document.getElementById('fpynWarning');
      if (totalNum > 0 && aoiOk >= 0) {
        fpynNum = (aoiOk / totalNum) * 100;
        fpynField.value = fpynNum.toFixed(1) + '%';
        if (fpynNum > 100) {
          hasError = true;
          fpynField.classList.add('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
          fpynWarning.classList.remove('hidden');
        } else {
          fpynField.classList.remove('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
          fpynWarning.classList.add('hidden');
        }
      } else {
        fpynField.value = '';
        fpynField.classList.remove('border-rose-500', 'text-rose-400', 'bg-rose-950/40');
        fpynWarning.classList.add('hidden');
      }

      const saveBtn = document.getElementById('saveSubmitBtn');
      if (hasError) {
        saveBtn.disabled = true;
        saveBtn.classList.add('opacity-50', 'cursor-not-allowed', 'bg-slate-700');
        saveBtn.classList.remove('bg-blue-600', 'hover:bg-blue-500', 'shadow-blue-900/30');
      } else {
        saveBtn.disabled = false;
        saveBtn.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-slate-700');
        saveBtn.classList.add('bg-blue-600', 'hover:bg-blue-500', 'shadow-blue-900/30');
      }
    }

    function handleFormSubmit(e) {
      e.preventDefault();
      
      const totalNum = parseFloat(document.getElementById('totalNum').value) || 0;
      const inPiece = parseFloat(document.getElementById('inPiece').value) || 0;
      const aoiOk = parseFloat(document.getElementById('aoiOk').value) || 0;

      if (totalNum > 0 && (inPiece / totalNum > 1 || aoiOk / totalNum > 1)) {
        alert('錯誤：良率或直通率大於 100%，無法儲存！');
        return;
      }

      const formData = new FormData(e.target);
      const editingId = document.getElementById('editingId').value;

      const record = {};
      formData.forEach((value, key) => {
        record[key] = value;
      });

      if (editingId) {
        const index = dataset.findIndex(d => d.id === Number(editingId));
        if (index !== -1) {
          dataset[index] = { ...dataset[index], ...record, id: Number(editingId) };
          showToast('修改生產資料成功！');
        }
      } else {
        const newId = dataset.length ? Math.max(...dataset.map(d => d.id)) + 1 : 1;
        dataset.unshift({ ...record, id: newId });
        showToast('新增生產資料成功！');
      }

      saveDatasetToStorage();
      updateHoleFilterOptions();
      closeModal();
      resetPageAndRender();
      updateDashboard();
    }

    function deleteRow(id) {
      if (confirm('確定要刪除這筆生產資料嗎？')) {
        dataset = dataset.filter(d => d.id !== id);
        saveDatasetToStorage();
        updateHoleFilterOptions();
        renderTable();
        updateDashboard();
        showToast('資料已成功刪除');
      }
    }

    function resetData() {
      if (confirm('確定重置為預設數據？這將清除所有自訂修改與匯入紀錄。')) {
        dataset = [...defaultData];
        saveDatasetToStorage();
        sortColumn = null;
        sortDirection = null;
        updateSortIcons();
        updateHoleFilterOptions();
        resetPageAndRender();
        updateDashboard();
        showToast('已重置為預設數據');
      }
    }

    function exportToExcel() {
      const headers = [
        ["班別", "前段生產日期", "穴號", "線別", "RUN CARD", "日期", "機台", "開始時間", "結束時間", "操作員", 
         "投入數", "圖紋種類 Pattern type", "總數", "良率", "入片數", "AOI-OK", "AOI-OK入片", "AOI-NG", "AOI總不良率", 
         "直通率", "CPK", "備註", "度數", "1.留公", "2.油殘 YOSAI", "3.圖紋密", "4.暗視野", "5.積墨 CHIMO", "6.邊破", 
         "7.氣泡 CHIPA", "8.附加物", "混料:穴號:顆粒", "混料:顆粒", "混入穴號", "混料誤判", "空洞 EMPTY", "反面 TAOBAO", "穴外亮點 OUTSIDE", "疊片 DOUBLE", "碎片 POPIE", "鏡片偏移", "留RING", "備註"]
      ];

      const rows = dataset.map(d => [
        d.shift, d.preDate, d.holeNo, d.lineNo, d.runCard, d.date, d.machine, d.startTime, d.endTime, d.operator,
        d.inputNum, d.patternType, d.totalNum, d.yieldVal, d.inPiece, d.aoiOk, d.aoiOkIn, d.aoiNg, d.aoiTotalNgYield,
        d.fpynVal, d.cpk, d.ipqcNote, d.degree, d.d_liugong, d.d_yosai, d.d_pattern_dense, d.d_darkfield, d.d_chimo, d.d_edge_break,
        d.d_chipa, d.d_attach, d.d_mix_hole_particle, d.d_mix_particle, d.d_mix_hole, d.d_mix_false, d.d_empty, d.d_taobao, d.d_outside, d.d_double, d.d_popie, d.d_shift, d.d_ring, d.note
      ]);

      const worksheet = XLSX.utils.aoa_to_sheet([...headers, ...rows]);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "生產紀錄表");

      const fileName = `光學生產數據表_${new Date().toISOString().slice(0,10)}.xlsx`;
      XLSX.writeFile(workbook, fileName);
      showToast('Excel 檔案匯出成功！');
    }

    function updateProgress(percent, stepText, descText) {
      const progressBar = document.getElementById('hollowPurpleProgressBar');
      const percentText = document.getElementById('loadingPercentText');
      const stepEl = document.getElementById('loadingStepText');
      const descEl = document.getElementById('loadingDesc');

      if (progressBar) progressBar.style.width = percent + '%';
      if (percentText) percentText.innerText = Math.round(percent) + '%';
      if (stepText && stepEl) stepEl.innerText = stepText;
      if (descText && descEl) descEl.innerText = descText;
    }

    function importExcel(e) {
      const file = e.target.files[0];
      if (!file) return;

      const loadingOverlay = document.getElementById('loadingOverlay');
      loadingOverlay.classList.remove('hidden');
      updateProgress(10, '領域展開：讀取檔案', `檔案大小: ${(file.size / (1024 * 1024)).toFixed(2)} MB`);

      const reader = new FileReader();
      
      reader.onload = function(evt) {
        updateProgress(40, '術式順轉：解析 Excel 結構', '正在解析試算表內容...');

        setTimeout(() => {
          try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            updateProgress(70, '術式反轉：轉換數據陣列', '萃取生產與瑕疵紀錄...');

            setTimeout(() => {
              try {
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

                if (jsonRows.length < 2) {
                  loadingOverlay.classList.add('hidden');
                  alert('Excel 內容為空或無數據行！');
                  e.target.value = '';
                  return;
                }

                updateProgress(90, '虚式「 茈 」：重構數據模型', `總計 ${jsonRows.length - 1} 筆紀錄，渲染中...`);

                setTimeout(() => {
                  try {
                    function parseExcelDate(val) {
                      if (val === undefined || val === null || val === '') return '';
                      if (typeof val === 'number' && val > 30000 && val < 60000) {
                        const dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
                        return dateObj.toISOString().slice(0, 10);
                      }
                      return String(val);
                    }

                    function parseExcelTime(val) {
                      if (val === undefined || val === null || val === '') return '';
                      if (typeof val === 'number' && val >= 0 && val < 1) {
                        const totalSeconds = Math.round(val * 86400);
                        const hours = Math.floor(totalSeconds / 3600);
                        const minutes = Math.floor((totalSeconds % 3600) / 60);
                        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
                      }
                      return String(val);
                    }

                    function parseExcelPercent(val, decimals = 1) {
                      if (val === undefined || val === null || val === '') return '';
                      const strVal = String(val).trim();
                      if (strVal.endsWith('%')) return strVal;
                      const num = parseFloat(strVal);
                      if (!isNaN(num)) {
                        if (num <= 1 && num >= 0) {
                          return (num * 100).toFixed(decimals) + '%';
                        } else if (num > 1) {
                          return num.toFixed(decimals) + '%';
                        }
                      }
                      return strVal;
                    }

                    const newDataset = [];
                    for (let i = 1; i < jsonRows.length; i++) {
                      const r = jsonRows[i];
                      if (!r || r.length === 0) continue;

                      newDataset.push({
                        id: Date.now() + i,
                        shift: r[0] || '',
                        preDate: parseExcelDate(r[1]),
                        holeNo: r[2] || '',
                        lineNo: r[3] || '',
                        runCard: r[4] || '',
                        date: parseExcelDate(r[5]),
                        machine: r[6] || '',
                        startTime: parseExcelTime(r[7]),
                        endTime: parseExcelTime(r[8]),
                        operator: r[9] || '',
                        inputNum: r[10] || '',
                        patternType: r[11] || '',
                        totalNum: r[12] || '',
                        yieldVal: parseExcelPercent(r[13], 1),
                        inPiece: r[14] || '',
                        aoiOk: r[15] || '',
                        aoiOkIn: r[16] || '',
                        aoiNg: r[17] || '',
                        aoiTotalNgYield: parseExcelPercent(r[18], 1),
                        fpynVal: parseExcelPercent(r[19], 1),
                        cpk: r[20] || '',
                        ipqcNote: r[21] || '',
                        degree: r[22] || '',
                        d_liugong: r[23] || '',
                        d_yosai: r[24] || '',
                        d_pattern_dense: r[25] || '',
                        d_darkfield: r[26] || '',
                        d_chimo: r[27] || '',
                        d_edge_break: r[28] || '',
                        d_chipa: r[29] || '',
                        d_attach: r[30] || '',
                        d_mix_hole_particle: r[31] || '',
                        d_mix_particle: r[32] || '',
                        d_mix_hole: r[33] || '',
                        d_mix_false: r[34] || '',
                        d_empty: r[35] || '',
                        d_taobao: r[36] || '',
                        d_outside: r[37] || '',
                        d_double: r[38] || '',
                        d_popie: r[39] || '',
                        d_shift: r[40] || '',
                        d_ring: r[41] || '',
                        note: r[42] || ''
                      });
                    }

                    if (newDataset.length > 0) {
                      dataset = newDataset;
                      saveDatasetToStorage();
                      updateHoleFilterOptions();
                      resetPageAndRender();
                      updateDashboard();
                      updateProgress(100, '發射完成！', `成功導入 ${newDataset.length} 筆生產數據`);
                      setTimeout(() => {
                        loadingOverlay.classList.add('hidden');
                        showToast(`成功導入 ${newDataset.length} 筆 Excel 生產數據！`);
                      }, 400);
                    }
                  } catch (innerErr) {
                    console.error(innerErr);
                    alert('解析資料時發生錯誤。');
                    loadingOverlay.classList.add('hidden');
                  } finally {
                    e.target.value = '';
                  }
                }, 50);

              } catch (parseErr) {
                console.error(parseErr);
                alert('解析 Excel 表格失敗。');
                loadingOverlay.classList.add('hidden');
                e.target.value = '';
              }
            }, 50);

          } catch (err) {
            console.error(err);
            alert('讀取 Excel 檔案時發生錯誤，請確保格式正確。');
            loadingOverlay.classList.add('hidden');
            e.target.value = '';
          }
        }, 50);
      };
      reader.readAsArrayBuffer(file);
    }

    function showToast(msg) {
      const toast = document.getElementById('toast');
      document.getElementById('toastMsg').innerText = msg;
      toast.classList.remove('hidden');
      setTimeout(() => {
        toast.classList.add('hidden');
      }, 3000);
    }
  </script>
</body>
</html>

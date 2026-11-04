<!DOCTYPE html>
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
    <div class="flex items-center space-x-3 mt-2 sm:mt-0 flex-wrap gap-y-2">
      <label class="cursor-pointer bg-slate-800 hover:bg-slate-700 text-slate-200 px-3 py-2 rounded-lg border border-slate-700 text-sm font-medium flex items-center space-x-2 transition">
        <i class="fa-solid fa-file-import text-blue-400"></i>
        <span>匯入 Excel (.xlsx)</span>
        <input type="file" id="excelFileInput" accept=".xlsx, .xls" class="hidden" onchange="importExcel(event)">
      </label>

      <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white px-3 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-emerald-900/30">
        <i class="fa-solid fa-file-excel"></i>
        <span>匯出數據 Excel</span>
      </button>

      <button onclick="manualBackupDownload()" class="bg-indigo-600 hover:bg-indigo-500 text-white px-3 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-indigo-900/30" title="備份資料到本機電腦 JSON">
        <i class="fa-solid fa-hard-drive"></i>
        <span>備份 JSON 檔</span>
      </button>

      <button onclick="openCloudConfigModal()" class="bg-teal-600 hover:bg-teal-500 text-white px-3 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-teal-900/30" title="設定網路 Server 或 Google Sheet 上傳">
        <i class="fa-solid fa-cloud-arrow-up"></i>
        <span>雲端同步設定</span>
      </button>

      <button onclick="openGithubModal()" class="bg-slate-800 hover:bg-slate-700 text-white px-3 py-2 rounded-lg border border-slate-700 text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-slate-900/30" title="同步與備份資料至 GitHub 倉庫">
        <i class="fa-brands fa-github text-purple-400 text-base"></i>
        <span>GitHub 雲端同步</span>
      </button>

      <button onclick="openAuditLogModal()" class="bg-purple-600 hover:bg-purple-500 text-white px-3 py-2 rounded-lg text-sm font-medium flex items-center space-x-2 transition shadow-lg shadow-purple-900/30">
        <i class="fa-solid fa-clock-rotate-left"></i>
        <span>生產記錄增加與修改</span>
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
            <input type="text" id="searchInput" oninput="resetPageAndRender()" onkeydown="handleSearchKeyDown(event)" placeholder="搜尋 工單、機台、線別、穴號、作業人員、備註..." class="w-full bg-slate-950 border border-slate-800 rounded-lg pl-9 pr-4 py-2 text-sm text-slate-200 focus:outline-none focus:border-blue-500 transition">
          </form>
        </div>

        <div class="flex items-center space-x-3 text-sm flex-wrap">
          <label class="text-slate-400">班別(SHIFT)篩選:</label>
          <select id="shiftFilter" onchange="resetPageAndRender()" class="bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-slate-200 focus:outline-none focus:border-blue-500">
            <option value="ALL">全部班別</option>
            <option value="D">D 班</option>
            <option value="N">N 班</option>
          </select>

          <label class="text-slate-400 ml-2">直通率低於(%):</label>
          <input type="number" id="fpyThresholdInput" oninput="resetPageAndRender()" placeholder="例: 95" class="w-20 bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-slate-200 text-xs focus:outline-none focus:border-cyan-500">

          <!-- 新增機台快速搜尋選項 -->
          <label class="text-slate-400 ml-2">機台(Equipment NO):</label>
          <input type="text" id="machineFilterInput" oninput="resetPageAndRender()" placeholder="例: 122" class="w-20 bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-slate-200 text-xs focus:outline-none focus:border-blue-500">

          <!-- 新增線別快速搜尋選項 -->
          <label class="text-slate-400 ml-2">線別(Front-end equipment NO):</label>
          <input type="text" id="lineFilterInput" oninput="resetPageAndRender()" placeholder="例: 109" class="w-20 bg-slate-950 border border-slate-800 rounded-lg px-2.5 py-1.5 text-slate-200 text-xs focus:outline-none focus:border-blue-500">

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

      <!-- Data Table Scroll Container with Mouse Drag Scroll Support -->
      <div id="tableScrollContainer" class="overflow-x-auto overflow-y-auto max-h-[500px] select-none cursor-grab active:cursor-grabbing">
        <table class="w-full text-xs text-left border-collapse whitespace-nowrap">
          <thead class="bg-slate-950 text-slate-300 font-semibold sticky top-0 z-20 border-b border-slate-800">
            <tr>
              <th class="p-2 border-r border-slate-800 text-center sticky-col bg-slate-950">操作</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">班別<br>(SHIFT)</th>
              
              <!-- Sorted Header: 前段生產日期 -->
              <th onclick="toggleSort('preDate')" class="p-2 border-r border-slate-800 bg-amber-900/40 text-amber-200 cursor-pointer hover:bg-amber-800/50 transition select-none">
                <div class="flex items-center justify-between space-x-1">
                  <span>前段生產日期<br>(Front-end equipment production time)</span>
                  <i id="sortIcon-preDate" class="fa-solid fa-sort text-amber-400/60"></i>
                </div>
              </th>

              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">穴號<br>(Model Number)</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">線別<br>(Front-end equipment NO)</th>
              
              <!-- 日期、機台、進站、出站、作業人員 移至前面 -->
              <!-- Sorted Header: 日期 -->
              <th onclick="toggleSort('date')" class="p-2 border-r border-slate-800 bg-amber-900/40 text-amber-200 cursor-pointer hover:bg-amber-800/50 transition select-none">
                <div class="flex items-center justify-between space-x-1">
                  <span>日期<br>(Date)</span>
                  <i id="sortIcon-date" class="fa-solid fa-sort text-amber-400/60"></i>
                </div>
              </th>

              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">機台<br>(Equipment NO)</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">進站時間<br>(Check in time)</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">出站時間<br>(Check out time)</th>
              <th class="p-2 border-r border-slate-800 bg-amber-900/30 text-amber-200 text-center">作業人員<br>(Employee No)</th>

              <!-- 度數、圖紋種類、工單 移至作業人員下方 -->
              <th class="p-2 border-r border-slate-800 bg-emerald-900/40 text-emerald-200 text-center">度數</th>
              <th class="p-2 border-r border-slate-800 bg-purple-900/40 text-purple-200 text-center">圖紋種類<br>(Pattern No)</th>
              <th class="p-2 border-r border-slate-800 bg-amber-400 text-slate-900 font-bold text-center">工單<br>(RUN CARD)</th>

              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">生產投入數量<br>(Current Unit No)</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">良率<br>(DM)</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">入片數</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">AOI-OK</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">AOI-OK入片</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">AOI-NG</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">AOI總不良率</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">直通率</th>
              <th class="p-2 border-r border-slate-800 bg-emerald-900/40 text-emerald-300 text-center">完工數量良品<br>(good Qty)</th>
              <th class="p-2 border-r border-slate-800 bg-cyan-900/40 text-cyan-300 text-center">CPK</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">備註<br>(Remark)</th>
              
              <!-- Defect Header Group -->
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">1.留公</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">2.油殘<br>(YOSAN)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">3.圖紋密度<br>(Pattern Density)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">4.暗視野亮點<br>(Dark field bright spot)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">5.積墨<br>(CHIMO)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">6.邊破<br>(PIENCHE/POPIEN)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">7.氣泡<br>(CHIPAO)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">8.附加物<br>(FU-CHA-WU)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">混料-穴號 顆數<br>(Model Number mixed count)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">混料顆粒<br>(Mixed count)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">(Mixed Model Number)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">混料誤判</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">9.空洞<br>(EMPTY)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">10.反面<br>(Flip)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">11.穴外<br>(Outside)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">12.疊片<br>(Double)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">13.碎片<br>(POPIE)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">14.鏡片偏移<br>(Shifting)</th>
              <th class="p-2 border-r border-slate-800 bg-rose-950/60 text-rose-300 text-center">15.留 ring</th>
              <th class="p-2 border-r border-slate-800 bg-slate-800 text-slate-200 text-center">16.通用備註<br>(Remark)</th>
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
              <label class="block text-slate-400 mb-1">班別 (SHIFT) <span class="text-rose-400">*</span></label>
              <select name="shift" required class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
                <option value="">--請選擇--</option>
                <option value="D">D 班</option>
                <option value="N">N 班</option>
              </select>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">前段生產日期 (Front-end production time) (滾輪) <span class="text-rose-400">*</span></label>
              <input type="date" name="preDate" required onwheel="handleDateWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">穴號 (Model Number) <span class="text-rose-400">*</span></label>
              <input type="number" name="holeNo" required placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">線別 (Front-end equipment NO) <span class="text-rose-400">*</span></label>
              <select name="lineNo" required class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
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
            <div>
              <label class="block text-slate-400 mb-1">日期 (Date) (滾輪) <span class="text-rose-400">*</span></label>
              <input type="date" name="date" required onwheel="handleDateWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">機台 (Equipment NO) <span class="text-rose-400">*</span></label>
              <select name="machine" required class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-blue-500">
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
              <label class="block text-slate-400 mb-1">進站時間 (Check in time) (滾輪) <span class="text-rose-400">*</span></label>
              <input type="time" name="startTime" required onwheel="handleTimeWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">出站時間 (Check out time) (滾輪) <span class="text-rose-400">*</span></label>
              <input type="time" name="endTime" required onwheel="handleTimeWheel(event)" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 cursor-pointer">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">作業人員 (Employee No) <span class="text-rose-400">*</span></label>
              <input type="text" name="operator" required placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-emerald-400 font-semibold mb-1">度數 <span class="text-rose-400">*</span></label>
              <input type="text" name="degree" required placeholder="" class="w-full bg-slate-950 border border-emerald-500/50 rounded p-2 text-emerald-200 font-mono">
            </div>
            <div>
              <label class="block text-purple-400 font-semibold mb-1">圖紋種類 (Pattern No)</label>
              <input type="text" name="patternNo" placeholder="" class="w-full bg-slate-950 border border-purple-500/50 rounded p-2 text-purple-200 font-mono">
            </div>
            <div class="col-span-1">
              <label class="block text-amber-400 font-semibold mb-1">工單 (RUN CARD) <span class="text-rose-400">*</span></label>
              <input type="text" name="runCard" required placeholder="" class="w-full bg-slate-950 border border-amber-500/50 rounded p-2 text-amber-200 font-mono">
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
              <label class="block text-slate-400 mb-1">生產投入數量 (Current Unit No) <span class="text-rose-400">*</span></label>
              <input type="number" id="inputNum" name="inputNum" required placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-emerald-400 mb-1">良率 (DM)</label>
              <input type="text" id="yieldVal" name="yieldVal" placeholder="" readonly class="w-full bg-slate-950/60 border border-slate-800 rounded p-2 text-emerald-400 font-bold cursor-not-allowed">
              <span id="yieldWarning" class="text-rose-400 text-[10px] hidden font-semibold mt-0.5 block">⚠️ 良率不能大於 100%</span>
            </div>
            <div>
              <label class="block text-slate-400 mb-1">入片數 <span class="text-rose-400">*</span></label>
              <input type="number" id="inPiece" name="inPiece" required placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-OK <span class="text-rose-400">*</span></label>
              <input type="number" id="aoiOk" name="aoiOk" required placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-OK入片 <span class="text-rose-400">*</span></label>
              <input type="number" id="aoiOkIn" name="aoiOkIn" required placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">AOI-NG <span class="text-rose-400">*</span></label>
              <input type="number" id="aoiNg" name="aoiNg" required placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
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
              <label class="block text-emerald-400 mb-1">完工數量良品 (good Qty) <span class="text-rose-400">*</span></label>
              <input type="number" id="totalNum" name="totalNum" required placeholder="" oninput="calcMetrics()" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">CPK <span class="text-rose-400">*</span></label>
              <input type="text" id="cpkVal" name="cpk" required placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
            <div>
              <label class="block text-slate-400 mb-1">備註 (Remark)</label>
              <input type="text" name="ipqcNote" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200">
            </div>
          </div>
        </div>

        <!-- Section 3: Defects Count -->
        <div>
          <h4 class="text-xs font-bold text-rose-400 uppercase tracking-wider mb-3 flex items-center space-x-2">
            <i class="fa-solid fa-bug"></i>
            <span>瑕疵項目統計 (數量與勾選)</span>
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 bg-slate-950/40 p-4 border border-slate-800/80 rounded-xl items-center">
            
            <!-- 1. 留公 (Checkbox) -->
            <div class="flex items-center space-x-2 bg-slate-900/60 p-2 border border-slate-800 rounded">
              <input type="checkbox" id="d_liugong_cb" name="d_liugong" value="1" class="w-4 h-4 accent-rose-500 cursor-pointer">
              <label for="d_liugong_cb" class="text-slate-200 cursor-pointer">1.留公</label>
            </div>

            <!-- 2. 油殘 (YOSAN) -->
            <div><label class="block text-slate-400 mb-1">2.油殘 (YOSAN)</label><input type="number" name="d_yosai" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 3. 圖紋密度 (Pattern Density) -->
            <div><label class="block text-slate-400 mb-1">3.圖紋密度 (Pattern Density)</label><input type="number" name="d_pattern_dense" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 4. 暗視野亮點 (Dark field bright spot) -->
            <div><label class="block text-slate-400 mb-1">4.暗視野亮點 (Dark field bright spot)</label><input type="number" name="d_darkfield" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 5. 積墨 (CHIMO) -->
            <div><label class="block text-slate-400 mb-1">5.積墨 (CHIMO)</label><input type="number" name="d_chimo" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 6. 邊破 (PIENCHE/POPIEN) -->
            <div><label class="block text-slate-400 mb-1">6.邊破 (PIENCHE/POPIEN)</label><input type="number" name="d_edge_break" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 7. 氣泡 (CHIPAO) -->
            <div><label class="block text-slate-400 mb-1">7.氣泡 (CHIPAO)</label><input type="number" name="d_chipa" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 8. 附加物 (FU-CHA-WU) -->
            <div><label class="block text-slate-400 mb-1">8.附加物 (FU-CHA-WU)</label><input type="number" name="d_attach" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <div><label class="block text-slate-400 mb-1">混料-穴號 顆數(Model Number mixed count)</label><input type="number" name="d_mix_hole_particle" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">混料顆粒(Mixed count)</label><input type="number" name="d_mix_particle" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">(Mixed Model Number)</label><input type="number" name="d_mix_hole" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            <div><label class="block text-slate-400 mb-1">混料誤判</label><input type="number" name="d_mix_false" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 9. 空洞 (EMPTY) -->
            <div><label class="block text-slate-400 mb-1">9.空洞 (EMPTY)</label><input type="number" name="d_empty" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 10. 反面 (Flip) -->
            <div><label class="block text-slate-400 mb-1">10.反面 (Flip)</label><input type="number" name="d_taobao" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 11. 穴外 (Outside) -->
            <div><label class="block text-slate-400 mb-1">11.穴外 (Outside)</label><input type="number" name="d_outside" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 12. 疊片 (Double) -->
            <div><label class="block text-slate-400 mb-1">12.疊片 (Double)</label><input type="number" name="d_double" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 13. 碎片 (POPIE) -->
            <div><label class="block text-slate-400 mb-1">13.碎片 (POPIE)</label><input type="number" name="d_popie" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 14. 鏡片偏移 (Shifting) -->
            <div><label class="block text-slate-400 mb-1">14.鏡片偏移 (Shifting)</label><input type="number" name="d_shift" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
            
            <!-- 15. 留 ring (Checkbox) -->
            <div class="flex items-center space-x-2 bg-slate-900/60 p-2 border border-slate-800 rounded">
              <input type="checkbox" id="d_ring_cb" name="d_ring" value="1" class="w-4 h-4 accent-rose-500 cursor-pointer">
              <label for="d_ring_cb" class="text-slate-200 cursor-pointer">15.留 ring</label>
            </div>
            
            <!-- 16. 通用備註 (Remark) -->
            <div><label class="block text-slate-400 mb-1">16.通用備註 (Remark)</label><input type="text" name="note" placeholder="" class="w-full bg-slate-950 border border-slate-800 rounded p-1.5 text-slate-200"></div>
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

  <!-- Audit Log Modal -->
  <div id="auditLogModal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-4 z-50 overflow-y-auto">
    <div class="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-4xl max-h-[85vh] flex flex-col shadow-2xl">
      <div class="px-6 py-4 border-b border-slate-800 flex justify-between items-center">
        <h3 class="text-base font-bold text-slate-100 flex items-center space-x-2">
          <i class="fa-solid fa-clock-rotate-left text-purple-400"></i>
          <span>生產記錄增加與修改歷程</span>
        </h3>
        <button onclick="closeAuditLogModal()" class="text-slate-400 hover:text-white text-lg px-2">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
      <div class="p-6 overflow-y-auto flex-1 text-xs">
        <table class="w-full text-left border-collapse">
          <thead class="bg-slate-950 text-slate-300 font-semibold border-b border-slate-800">
            <tr>
              <th class="p-3 border-r border-slate-800">異動時間</th>
              <th class="p-3 border-r border-slate-800">動作類型</th>
              <th class="p-3 border-r border-slate-800 text-amber-300">工單 (RUN CARD)</th>
              <th class="p-3 border-r border-slate-800">作業人員</th>
              <th class="p-3 text-slate-200">修改內容 / 異動說明</th>
            </tr>
          </thead>
          <tbody id="auditLogTableBody" class="divide-y divide-slate-800 text-slate-300">
            <!-- Rendered dynamically -->
          </tbody>
        </table>
      </div>
      <div class="px-6 py-3 bg-slate-950 border-t border-slate-800 flex justify-end">
        <button onclick="closeAuditLogModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition text-xs">關閉</button>
      </div>
    </div>
  </div>

  <!-- GitHub Settings & Sync Modal -->
  <div id="githubModal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-4 z-50 overflow-y-auto">
    <div class="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg flex flex-col shadow-2xl">
      <div class="px-6 py-4 border-b border-slate-800 flex justify-between items-center">
        <h3 class="text-base font-bold text-slate-100 flex items-center space-x-2">
          <i class="fa-brands fa-github text-purple-400 text-lg"></i>
          <span>GitHub API 數據同步與備份設定</span>
        </h3>
        <button onclick="closeGithubModal()" class="text-slate-400 hover:text-white text-lg px-2">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
      <div class="p-6 space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-semibold mb-1">GitHub Personal Access Token (PAT) <span class="text-rose-400">*</span></label>
          <input type="password" id="ghTokenInput" placeholder="ghp_xxxxxxxxxxxxxxxxxxxx" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-purple-500 font-mono text-[11px]">
          <p class="text-[11px] text-slate-500 mt-1">請建立具備 Repo Contents Read/Write 權限之 Fine-grained Token。</p>
        </div>
        
        <div class="grid grid-cols-2 gap-3">
          <div>
            <label class="block text-slate-300 font-semibold mb-1">Repo 擁有者 (Owner/User) <span class="text-rose-400">*</span></label>
            <input type="text" id="ghOwnerInput" placeholder="例: your-github-username" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-purple-500 font-mono text-[11px]">
          </div>
          <div>
            <label class="block text-slate-300 font-semibold mb-1">Repo 名稱 (Repository) <span class="text-rose-400">*</span></label>
            <input type="text" id="ghRepoInput" placeholder="例: optics-production-data" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-purple-500 font-mono text-[11px]">
          </div>
        </div>

        <div>
          <label class="block text-slate-300 font-semibold mb-1">目標儲存檔案路徑 (Target File Path)</label>
          <input type="text" id="ghPathInput" placeholder="data/production_data.json" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-purple-500 font-mono text-[11px]">
          <p class="text-[11px] text-slate-500 mt-1">預設備份為 JSON 檔案格式，路徑若不存在將會自動新建。</p>
        </div>

        <div class="flex items-center space-x-2 pt-1">
          <input type="checkbox" id="ghAutoSyncCheckbox" class="w-4 h-4 accent-purple-500 cursor-pointer">
          <label for="ghAutoSyncCheckbox" class="text-slate-200 cursor-pointer font-medium">資料變更時自動 Commit/Push 至 GitHub</label>
        </div>
      </div>

      <div class="px-6 py-3 bg-slate-950 border-t border-slate-800 flex justify-between items-center">
        <button onclick="syncToGitHub(true)" class="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded-lg transition text-xs font-semibold flex items-center space-x-2 shadow-lg shadow-purple-900/30">
          <i class="fa-solid fa-cloud-arrow-up"></i>
          <span>立即同步至 GitHub</span>
        </button>
        <div class="space-x-2">
          <button onclick="closeGithubModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition text-xs">取消</button>
          <button onclick="saveGithubConfig()" class="px-4 py-2 bg-blue-600 hover:bg-blue-500 text-white rounded-lg transition text-xs font-semibold">儲存設定</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Cloud / Webhook Settings Modal -->
  <div id="cloudConfigModal" class="fixed inset-0 bg-slate-950/80 backdrop-blur-sm hidden flex items-center justify-center p-4 z-50 overflow-y-auto">
    <div class="bg-slate-900 border border-slate-800 rounded-2xl w-full max-w-lg flex flex-col shadow-2xl">
      <div class="px-6 py-4 border-b border-slate-800 flex justify-between items-center">
        <h3 class="text-base font-bold text-slate-100 flex items-center space-x-2">
          <i class="fa-solid fa-cloud-arrow-up text-teal-400"></i>
          <span>網路 / 雲端伺服器自動上傳設定</span>
        </h3>
        <button onclick="closeCloudConfigModal()" class="text-slate-400 hover:text-white text-lg px-2">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>
      <div class="p-6 space-y-4 text-xs">
        <div>
          <label class="block text-slate-300 font-semibold mb-1">Server API / Webhook URL Endpoint:</label>
          <input type="text" id="webhookUrlInput" placeholder="https://your-company-server.com/api/upload or Google Apps Script URL" class="w-full bg-slate-950 border border-slate-800 rounded p-2 text-slate-200 focus:border-teal-500 font-mono text-[11px]">
          <p class="text-[11px] text-slate-500 mt-1">設定後，每次匯入或修改生產資料時，將自動以 POST JSON 格式同步上傳至此 URL。</p>
        </div>
        <div class="flex items-center space-x-2">
          <input type="checkbox" id="autoSyncCheckbox" class="w-4 h-4 accent-teal-500 cursor-pointer">
          <label for="autoSyncCheckbox" class="text-slate-200 cursor-pointer">變更時自動同步至雲端 Server</label>
        </div>
      </div>
      <div class="px-6 py-3 bg-slate-950 border-t border-slate-800 flex justify-end space-x-2">
        <button onclick="closeCloudConfigModal()" class="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg transition text-xs">取消</button>
        <button onclick="saveCloudConfig()" class="px-4 py-2 bg-teal-600 hover:bg-teal-500 text-white rounded-lg transition text-xs font-semibold">儲存設定</button>
      </div>
    </div>
  </div>

  <!-- Loading Overlay -->
  <div id="loadingOverlay" class="fixed inset-0 bg-slate-950/85 backdrop-blur-md hidden flex flex-col items-center justify-center z-50 p-4">
    <div class="bg-slate-900/90 border border-purple-500/40 p-6 rounded-2xl shadow-[0_0_50px_rgba(168,85,247,0.3)] max-w-md w-full flex flex-col items-center space-y-5 text-center relative overflow-hidden">
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
    const defaultData = [
      {
        id: 1, shift: 'D', preDate: '2026-10-20', holeNo: '10', lineNo: '109', degree: '', patternNo: '', runCard: 'v5121-261004078-015.1',
        date: '2026-10-23', machine: '122', startTime: '03:37', endTime: '08:16', operator: 'VF1338',
        inputNum: 4095, totalNum: 4095, yieldVal: '99.9%', inPiece: 4092,
        aoiOk: 3785, aoiOkIn: 3785, aoiNg: 307, aoiTotalNgYield: '7.5%', fpynVal: '92.4%',
        cpk: '1.33', ipqcNote: '烤2小時,靜置12小時, IPQC 80 PCS',
        d_liugong: '', d_yosai: '', d_pattern_dense: '', d_darkfield: '', d_chimo: '', d_edge_break: 16,
        d_chipa: '', d_attach: '', d_mix_hole_particle: '', d_mix_particle: '', d_mix_hole: '', d_mix_false: '',
        d_empty: '', d_taobao: '', d_outside: '', d_double: '', d_popie: '', d_shift: '', d_ring: '', note: ''
      }
    ];

    const fieldLabelMap = {
      shift: '班別(SHIFT)',
      preDate: '前段生產日期(Front-end equipment production time)',
      holeNo: '穴號(Model Number)',
      lineNo: '線別(Front-end equipment NO)',
      date: '日期(Date)',
      machine: '機台(Equipment NO)',
      startTime: '進站時間(Check in time)',
      endTime: '出站時間(Check out time)',
      operator: '作業人員(Employee No)',
      degree: '度數',
      patternNo: '圖紋種類(Pattern No)',
      runCard: '工單(RUN CARD)',
      inputNum: '生產投入數量(Current Unit No)',
      yieldVal: '良率(DM)',
      inPiece: '入片數',
      aoiOk: 'AOI-OK',
      aoiOkIn: 'AOI-OK入片',
      aoiNg: 'AOI-NG',
      aoiTotalNgYield: 'AOI總不良率',
      fpynVal: '直通率',
      totalNum: '完工數量良品 (good Qty)',
      cpk: 'CPK',
      ipqcNote: '備註(Remark)',
      d_liugong: '1.留公',
      d_yosai: '2.油殘(YOSAN)',
      d_pattern_dense: '3.圖紋密度(Pattern Density)',
      d_darkfield: '4.暗視野亮點(Dark field bright spot)',
      d_chimo: '5.積墨(CHIMO)',
      d_edge_break: '6.邊破(PIENCHE/POPIEN)',
      d_chipa: '7.氣泡(CHIPAO)',
      d_attach: '8.附加物(FU-CHA-WU)',
      d_mix_hole_particle: '混料-穴號 顆數(Model Number mixed count)',
      d_mix_particle: '混料顆粒(Mixed count)',
      d_mix_hole: '(Mixed Model Number)',
      d_mix_false: '混料誤判',
      d_empty: '9.空洞(EMPTY)',
      d_taobao: '10.反面(Flip)',
      d_outside: '11.穴外(Outside)',
      d_double: '12.疊片(Double)',
      d_popie: '13.碎片(POPIE)',
      d_shift: '14.鏡片偏移(Shifting)',
      d_ring: '15.留 ring',
      note: '16.通用備註(Remark)'
    };

    // 系統欄位對應字典（支援關鍵字模糊與軟性比對）
    const fieldMappingRules = {
      shift: ["班別", "shift"],
      preDate: ["前段生產日期", "前段", "front-end equipment production time"],
      holeNo: ["穴號", "model number"],
      lineNo: ["線別", "front-end equipment no"],
      date: ["日期", "date"],
      machine: ["機台", "equipment no"],
      startTime: ["進站時間", "check in time", "開始時間", "起算時間"],
      endTime: ["出站時間", "check out time", "結束時間", "完工時間"],
      operator: ["作業人員", "employee no", "人員", "操作員"],
      degree: ["度數"],
      patternNo: ["圖紋種類", "pattern no"],
      runCard: ["工單", "run card"],
      inputNum: ["生產投入數量", "current unit no", "投入數", "投入數量"],
      yieldVal: ["良率", "dm"],
      inPiece: ["入片數"],
      aoiOk: ["aoi-ok", "aoi ok"],
      aoiOkIn: ["aoi-ok入片", "aoi ok入片"],
      aoiNg: ["aoi-ng", "aoi ng"],
      aoiTotalNgYield: ["aoi總不良率", "aoi 不良率", "總不良率"],
      fpynVal: ["直通率", "fpy"],
      totalNum: ["完工數量良品", "good qty", "良品數"],
      cpk: ["cpk"],
      ipqcNote: ["備註", "remark", "ipqc"],
      d_liugong: ["1.留公", "留公"],
      d_yosai: ["2.油殘", "油殘", "yosan"],
      d_pattern_dense: ["3.圖紋密度", "圖紋密度", "pattern density"],
      d_darkfield: ["4.暗視野亮點", "暗視野亮點", "dark field"],
      d_chimo: ["5.積墨", "積墨", "chimo"],
      d_edge_break: ["6.邊破", "邊破", "pienche", "popien"],
      d_chipa: ["7.氣泡", "氣泡", "chipao"],
      d_attach: ["8.附加物", "附加物", "fu-cha-wu"],
      d_mix_hole_particle: ["混料-穴號 顆數", "model number mixed count"],
      d_mix_particle: ["混料顆粒", "mixed count"],
      d_mix_hole: ["mixed model number", "混料穴號"],
      d_mix_false: ["混料誤判"],
      d_empty: ["9.空洞", "空洞", "empty"],
      d_taobao: ["10.反面", "反面", "flip"],
      d_outside: ["11.穴外", "穴外", "outside"],
      d_double: ["12.疊片", "疊片", "double"],
      d_popie: ["13.碎片", "碎片", "popie"],
      d_shift: ["14.鏡片偏移", "鏡片偏移", "shifting"],
      d_ring: ["15.留 ring", "留 ring", "留ring"],
      note: ["16.通用備註", "通用備註"]
    };

    // 預設欄位索引對應順序（備用退路）
    const defaultIndexMap = {
      shift: 0, preDate: 1, holeNo: 2, lineNo: 3, date: 4, machine: 5,
      startTime: 6, endTime: 7, operator: 8, degree: 9, patternNo: 10,
      runCard: 11, inputNum: 12, yieldVal: 13, inPiece: 14, aoiOk: 15,
      aoiOkIn: 16, aoiNg: 17, aoiTotalNgYield: 18, fpynVal: 19, totalNum: 20,
      cpk: 21, ipqcNote: 22, d_liugong: 23, d_yosai: 24, d_pattern_dense: 25,
      d_darkfield: 26, d_chimo: 27, d_edge_break: 28, d_chipa: 29, d_attach: 30,
      d_mix_hole_particle: 31, d_mix_particle: 32, d_mix_hole: 33, d_mix_false: 34,
      d_empty: 35, d_taobao: 36, d_outside: 37, d_double: 38, d_popie: 39,
      d_shift: 40, d_ring: 41, note: 42
    };

    let dataset = [];
    let auditLogs = [];
    let db = null;

    // --- 1. IndexedDB 瀏覽器高效永久資料庫引擎 ---
    function initIndexedDB() {
      return new Promise((resolve, reject) => {
        const request = indexedDB.open('OpticsProductionDB', 1);
        request.onupgradeneeded = function(e) {
          const dbInstance = e.target.result;
          if (!dbInstance.objectStoreNames.contains('production_store')) {
            dbInstance.createObjectStore('production_store', { keyPath: 'key' });
          }
        };
        request.onsuccess = function(e) {
          db = e.target.result;
          resolve(db);
        };
        request.onerror = function(e) {
          console.error('IndexedDB 開啟失敗:', e);
          resolve(null);
        };
      });
    }

    function saveToIndexedDB(key, val) {
      if (!db) return;
      try {
        const tx = db.transaction('production_store', 'readwrite');
        const store = tx.objectStore('production_store');
        store.put({ key: key, value: val });
      } catch (err) {
        console.error('IndexedDB 寫入錯誤:', err);
      }
    }

    function getFromIndexedDB(key) {
      return new Promise((resolve) => {
        if (!db) return resolve(null);
        try {
          const tx = db.transaction('production_store', 'readonly');
          const store = tx.objectStore('production_store');
          const req = store.get(key);
          req.onsuccess = function() {
            resolve(req.result ? req.result.value : null);
          };
          req.onerror = function() {
            resolve(null);
          };
        } catch (e) {
          resolve(null);
        }
      });
    }

    // --- 2. 儲存與多重備份核心機制 ---
    async function saveDatasetToStorage() {
      // (A) 本機 IndexedDB 永久保存 (支援大容量)
      saveToIndexedDB('dataset', dataset);
      saveToIndexedDB('audit_logs', auditLogs);

      // (B) localStorage 雙重備份
      try {
        localStorage.setItem('optics_production_dataset', JSON.stringify(dataset));
        localStorage.setItem('optics_production_audit_logs', JSON.stringify(auditLogs));
      } catch (e) {
        console.warn('LocalStorage 可能容量已滿，改由 IndexedDB 完全保存。');
      }

      // (C) 雲端 Server / Webhook 自動 POST 上傳
      const autoSync = localStorage.getItem('optics_auto_sync') === 'true';
      const webhookUrl = localStorage.getItem('optics_webhook_url');
      if (autoSync && webhookUrl) {
        uploadToCloudServer(webhookUrl);
      }

      // (D) GitHub 自動同步 Commit & Push
      const ghAutoSync = localStorage.getItem('optics_gh_auto_sync') === 'true';
      if (ghAutoSync) {
        syncToGitHub(false);
      }
    }

    // 雲端同步上傳功能
    function uploadToCloudServer(url) {
      fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          timestamp: new Date().toISOString(),
          dataset: dataset,
          auditLogs: auditLogs
        })
      }).then(res => {
        if (res.ok) console.log('網路 Server 雲端備份同步成功');
      }).catch(err => {
        console.error('網路 Server 自動同步失敗:', err);
      });
    }

    // 手動匯出電腦 JSON 備份檔
    function manualBackupDownload() {
      const backupObj = {
        system: "光學製造生產數據管理系統",
        exportDate: new Date().toISOString(),
        dataset: dataset,
        auditLogs: auditLogs
      };
      const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(backupObj, null, 2));
      const downloadAnchor = document.createElement('a');
      downloadAnchor.setAttribute("href", dataStr);
      downloadAnchor.setAttribute("download", `光學生產數據備份_${new Date().toISOString().slice(0,10)}.json`);
      document.body.appendChild(downloadAnchor);
      downloadAnchor.click();
      downloadAnchor.remove();
      showToast('數據 JSON 備份檔已下載至本機電腦！');
    }

    let lineChartObj = null;
    let machineChartObj = null;
    let currentPage = 1;
    let pageSize = 50;
    let sortColumn = null;
    let sortDirection = null;

    const defectKeys = [
      { key: 'd_liugong', label: '1.留公' },
      { key: 'd_yosai', label: '2.油殘(YOSAN)' },
      { key: 'd_pattern_dense', label: '3.圖紋密度(Pattern Density)' },
      { key: 'd_darkfield', label: '4.暗視野亮點(Dark field bright spot)' },
      { key: 'd_chimo', label: '5.積墨(CHIMO)' },
      { key: 'd_edge_break', label: '6.邊破(PIENCHE/POPIEN)' },
      { key: 'd_chipa', label: '7.氣泡(CHIPAO)' },
      { key: 'd_attach', label: '8.附加物(FU-CHA-WU)' },
      { key: 'd_mix_hole_particle', label: '混料-穴號 顆數(Model Number mixed count)' },
      { key: 'd_mix_particle', label: '混料顆粒(Mixed count)' },
      { key: 'd_mix_hole', label: '(Mixed Model Number)' },
      { key: 'd_mix_false', label: '混料誤判' },
      { key: 'd_empty', label: '9.空洞(EMPTY)' },
      { key: 'd_taobao', label: '10.反面(Flip)' },
      { key: 'd_outside', label: '11.穴外(Outside)' },
      { key: 'd_double', label: '12.疊片(Double)' },
      { key: 'd_popie', label: '13.碎片(POPIE)' },
      { key: 'd_shift', label: '14.鏡片偏移(Shifting)' },
      { key: 'd_ring', label: '15.留 ring' }
    ];

    window.onload = async function() {
      if (typeof ChartDataLabels !== 'undefined') {
        Chart.register(ChartDataLabels);
      }

      // 初始化 IndexedDB 並自動恢復數據
      await initIndexedDB();
      const idbData = await getFromIndexedDB('dataset');
      const idbLogs = await getFromIndexedDB('audit_logs');

      if (idbData && Array.isArray(idbData) && idbData.length > 0) {
        dataset = idbData;
      } else {
        const saved = localStorage.getItem('optics_production_dataset');
        if (saved) {
          try { dataset = JSON.parse(saved); } catch (e) { dataset = [...defaultData]; }
        } else {
          dataset = [...defaultData];
        }
      }

      if (idbLogs && Array.isArray(idbLogs) && idbLogs.length > 0) {
        auditLogs = idbLogs;
      } else {
        const savedLogs = localStorage.getItem('optics_production_audit_logs');
        if (savedLogs) {
          try { auditLogs = JSON.parse(savedLogs); } catch (e) {}
        }
      }

      updateHoleFilterOptions();
      renderTable();
      initCharts();
      updateDashboard();
      initMouseDragScroll();
    };

    function initMouseDragScroll() {
      const slider = document.getElementById('tableScrollContainer');
      if (!slider) return;
      let isDown = false;
      let startX;
      let scrollLeft;

      slider.addEventListener('mousedown', (e) => {
        if (['BUTTON', 'INPUT', 'SELECT', 'I'].includes(e.target.tagName)) return;
        isDown = true;
        slider.classList.add('cursor-grabbing');
        slider.classList.remove('cursor-grab');
        startX = e.pageX - slider.offsetLeft;
        scrollLeft = slider.scrollLeft;
      });

      slider.addEventListener('mouseleave', () => {
        isDown = false;
        slider.classList.remove('cursor-grabbing');
        slider.classList.add('cursor-grab');
      });

      slider.addEventListener('mouseup', () => {
        isDown = false;
        slider.classList.remove('cursor-grabbing');
        slider.classList.add('cursor-grab');
      });

      slider.addEventListener('mousemove', (e) => {
        if (!isDown) return;
        e.preventDefault();
        const x = e.pageX - slider.offsetLeft;
        const walk = (x - startX) * 1.5;
        slider.scrollLeft = scrollLeft - walk;
      });
    }

    function updateHoleFilterOptions() {
      const holeSelect = document.getElementById('chartHoleFilter');
      if (!holeSelect) return;
      const currentSelected = holeSelect.value;
      const holesSet = new Set();
      dataset.forEach(d => {
        if (d.holeNo !== undefined && d.holeNo !== null && String(d.holeNo).trim() !== '') {
          holesSet.add(String(d.holeNo).trim());
        }
      });

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
      const fpyThresholdVal = parseFloat(document.getElementById('fpyThresholdInput').value);
      const machineFilterVal = document.getElementById('machineFilterInput').value.trim().toLowerCase();
      const lineFilterVal = document.getElementById('lineFilterInput').value.trim().toLowerCase();

      let result = dataset.filter(item => {
        const matchesShift = shiftFilter === 'ALL' || item.shift === shiftFilter;
        if (!matchesShift) return false;

        if (!isNaN(fpyThresholdVal)) {
          const itemFpy = parseFloat(String(item.fpynVal).replace('%', ''));
          if (!isNaN(itemFpy) && itemFpy >= fpyThresholdVal) return false;
        }

        if (machineFilterVal) {
          const itemMachine = String(item.machine || '').toLowerCase();
          if (!itemMachine.includes(machineFilterVal)) return false;
        }

        if (lineFilterVal) {
          const itemLine = String(item.lineNo || '').toLowerCase();
          if (!itemLine.includes(lineFilterVal)) return false;
        }

        if (search) {
          const machineStr = String(item.machine || '').toLowerCase();
          const isMachineQuery = !isNaN(search) && machineStr === search;
          if (isMachineQuery) return true;

          const searchableText = [
            item.runCard, item.machine, item.lineNo, item.holeNo, item.shift,
            item.operator, item.ipqcNote, item.note, item.preDate, item.date, item.degree, item.patternNo
          ].map(v => (v !== undefined && v !== null) ? String(v).toLowerCase() : '').join(' ');

          return searchableText.includes(search);
        }

        return true;
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
              <button onclick="editRow('${row.id}')" class="p-1 text-blue-400 hover:text-blue-300 transition" title="編輯修改">
                <i class="fa-solid fa-pen-to-square"></i>
              </button>
              <button onclick="deleteRow('${row.id}')" class="p-1 text-rose-400 hover:text-rose-300 transition" title="刪除">
                <i class="fa-solid fa-trash-can"></i>
              </button>
            </div>
          </td>
          <td class="p-2 border-r border-slate-800 text-center font-bold text-amber-300">${row.shift || ''}</td>
          <td class="p-2 border-r border-slate-800">${formatDateDisplay(row.preDate)}</td>
          <td class="p-2 border-r border-slate-800 text-center font-bold text-amber-200">${row.holeNo || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-semibold text-blue-300">${row.lineNo || ''}</td>
          <td class="p-2 border-r border-slate-800">${formatDateDisplay(row.date)}</td>
          <td class="p-2 border-r border-slate-800 text-center font-semibold text-slate-200">${row.machine || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-400">${row.startTime || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-400">${row.endTime || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.operator || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono text-emerald-300 bg-emerald-950/20">${row.degree || ''}</td>
          <td class="p-2 border-r border-slate-800 text-center font-mono text-purple-300 bg-purple-950/20">${row.patternNo || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono font-bold text-amber-300 bg-amber-950/10">${row.runCard || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-slate-200">${row.inputNum || ''}</td>
          <td class="p-2 border-r border-slate-800 font-bold text-emerald-400 bg-emerald-950/20 text-center">${row.yieldVal || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.inPiece || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.aoiOk || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono">${row.aoiOkIn || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-rose-400 font-semibold">${row.aoiNg || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.aoiTotalNgYield || ''}</td>
          <td class="p-2 border-r border-slate-800 font-bold text-cyan-400 bg-cyan-950/20 text-center">${row.fpynVal || ''}</td>
          <td class="p-2 border-r border-slate-800 font-mono text-emerald-300 bg-emerald-950/20 text-center font-bold">${row.totalNum || ''}</td>
          <td class="p-2 border-r border-slate-800">${row.cpk || ''}</td>
          <td class="p-2 border-r border-slate-800 text-slate-300">${row.ipqcNote || ''}</td>
          
          <!-- Defect columns (Checkboxes as ✓ or empty) -->
          <td class="p-2 border-r border-slate-800 text-center font-bold ${row.d_liugong ? 'text-rose-400 bg-rose-950/30' : 'text-slate-600'}">${row.d_liugong ? '✓' : ''}</td>
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
          <td class="p-2 border-r border-slate-800 text-center font-bold ${row.d_ring ? 'text-rose-400 bg-rose-950/30' : 'text-slate-600'}">${row.d_ring ? '✓' : ''}</td>
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
        if (holeFilterVal !== 'ALL') {
          if (String(d.holeNo).trim() !== String(holeFilterVal).trim()) return false;
        }
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
      const sumInputNum = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.inputNum) || 0), 0);
      
      const avgYield = sumInputNum > 0 ? ((sumInPiece / sumInputNum) * 100).toFixed(1) : 0;
      const sumAoiOk = analyticsDataset.reduce((acc, curr) => acc + (Number(curr.aoiOk) || 0), 0);
      const avgFpyn = sumInputNum > 0 ? ((sumAoiOk / sumInputNum) * 100).toFixed(1) : 0;
      const avgAoiNgRate = sumInPiece > 0 ? ((totalAoiNg / sumInPiece) * 100).toFixed(1) : 0;

      const defectTotals = {};
      defectKeys.forEach(dk => {
        if (dk.key === 'd_liugong' || dk.key === 'd_ring') {
          defectTotals[dk.label] = analyticsDataset.reduce((sum, item) => sum + (item[dk.key] ? 1 : 0), 0);
        } else {
          defectTotals[dk.label] = analyticsDataset.reduce((sum, item) => sum + (Number(item[dk.key]) || 0), 0);
        }
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
          lineMap[line] = { inputNum: 0, inPiece: 0, aoiOk: 0, aoiNg: 0 };
        }
        lineMap[line].inputNum += Number(d.inputNum) || 0;
        lineMap[line].inPiece += Number(d.inPiece) || 0;
        lineMap[line].aoiOk += Number(d.aoiOk) || 0;
        lineMap[line].aoiNg += Number(d.aoiNg) || 0;
      });

      const lineLabels = Object.keys(lineMap).sort();
      const lineYields = [], lineFpys = [], lineDefects = [];
      lineLabels.forEach(l => {
        const m = lineMap[l];
        const y = m.inputNum > 0 ? (m.inPiece / m.inputNum * 100) : 0;
        const f = m.inputNum > 0 ? (m.aoiOk / m.inputNum * 100) : 0;
        const d = m.inPiece > 0 ? (m.aoiNg / m.inPiece * 100) : 0;
        lineYields.push(parseFloat(y.toFixed(1)));
        lineFpys.push(parseFloat(f.toFixed(1)));
        lineDefects.push(parseFloat(d.toFixed(1)));
      });

      const machineMap = {};
      analyticsDataset.forEach(d => {
        const mac = d.machine ? `機台 ${d.machine}` : '未分類';
        if (!machineMap[mac]) {
          machineMap[mac] = { inputNum: 0, inPiece: 0, aoiOk: 0, aoiNg: 0 };
        }
        machineMap[mac].inputNum += Number(d.inputNum) || 0;
        machineMap[mac].inPiece += Number(d.inPiece) || 0;
        machineMap[mac].aoiOk += Number(d.aoiOk) || 0;
        machineMap[mac].aoiNg += Number(d.aoiNg) || 0;
      });

      const machineLabels = Object.keys(machineMap).sort();
      const machineYields = [], machineFpys = [], machineDefects = [];
      machineLabels.forEach(mKey => {
        const m = machineMap[mKey];
        const y = m.inputNum > 0 ? (m.inPiece / m.inputNum * 100) : 0;
        const f = m.inputNum > 0 ? (m.aoiOk / m.inputNum * 100) : 0;
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
        font: { size: 9, weight: 'bold' },
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
            { label: '良率 (%)', data: [], backgroundColor: 'rgba(16, 185, 129, 0.75)', borderColor: '#10b981', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions },
            { label: '直通率 (%)', data: [], backgroundColor: 'rgba(6, 182, 212, 0.75)', borderColor: '#06b6d4', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions },
            { label: '不良率 (%)', data: [], backgroundColor: 'rgba(244, 63, 94, 0.75)', borderColor: '#f43f5e', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions }
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
            { label: '良率 (%)', data: [], backgroundColor: 'rgba(16, 185, 129, 0.75)', borderColor: '#10b981', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions },
            { label: '直通率 (%)', data: [], backgroundColor: 'rgba(6, 182, 212, 0.75)', borderColor: '#06b6d4', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions },
            { label: '不良率 (%)', data: [], backgroundColor: 'rgba(244, 63, 94, 0.75)', borderColor: '#f43f5e', borderWidth: 1, borderRadius: 4, datalabels: commonDatalabelOptions }
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
      const item = dataset.find(d => String(d.id) === String(id));
      if (!item) return;

      const form = document.getElementById('recordForm');
      document.getElementById('editingId').value = String(item.id);
      document.getElementById('modalTitle').innerText = `修改生產紀錄 (${item.runCard || '項次 ' + id})`;

      for (let key in item) {
        if (form.elements[key]) {
          if (key === 'd_liugong' || key === 'd_ring') {
            form.elements[key].checked = Boolean(item[key]);
          } else {
            form.elements[key].value = item[key] || '';
          }
        }
      }

      calcMetrics();
      document.getElementById('dataModal').classList.remove('hidden');
    }

    function closeModal() {
      document.getElementById('dataModal').classList.add('hidden');
    }

    function openAuditLogModal() {
      const tbody = document.getElementById('auditLogTableBody');
      if (auditLogs.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" class="p-4 text-center text-slate-500">目前尚無任何增加或修改紀錄</td></tr>`;
      } else {
        tbody.innerHTML = auditLogs.slice().reverse().map(log => `
          <tr class="hover:bg-slate-800/40 transition">
            <td class="p-3 border-r border-slate-800 text-slate-400 font-mono">${log.time}</td>
            <td class="p-3 border-r border-slate-800 font-bold ${log.action === '新增' ? 'text-emerald-400' : 'text-blue-400'}">${log.action}</td>
            <td class="p-3 border-r border-slate-800 font-mono font-bold text-amber-300">${log.runCard || '-'}</td>
            <td class="p-3 border-r border-slate-800 font-semibold">${log.operator || '-'}</td>
            <td class="p-3 text-slate-300">${log.detail}</td>
          </tr>
        `).join('');
      }
      document.getElementById('auditLogModal').classList.remove('hidden');
    }

    function closeAuditLogModal() {
      document.getElementById('auditLogModal').classList.add('hidden');
    }

    // --- GitHub API 雲端同步核心模組 ---
    function openGithubModal() {
      document.getElementById('ghTokenInput').value = localStorage.getItem('optics_gh_token') || '';
      document.getElementById('ghOwnerInput').value = localStorage.getItem('optics_gh_owner') || '';
      document.getElementById('ghRepoInput').value = localStorage.getItem('optics_gh_repo') || '';
      document.getElementById('ghPathInput').value = localStorage.getItem('optics_gh_path') || 'data/production_data.json';
      document.getElementById('ghAutoSyncCheckbox').checked = localStorage.getItem('optics_gh_auto_sync') === 'true';
      document.getElementById('githubModal').classList.remove('hidden');
    }

    function closeGithubModal() {
      document.getElementById('githubModal').classList.add('hidden');
    }

    function saveGithubConfig() {
      const token = document.getElementById('ghTokenInput').value.trim();
      const owner = document.getElementById('ghOwnerInput').value.trim();
      const repo = document.getElementById('ghRepoInput').value.trim();
      let path = document.getElementById('ghPathInput').value.trim();
      if (!path) path = 'data/production_data.json';
      const autoSync = document.getElementById('ghAutoSyncCheckbox').checked;

      localStorage.setItem('optics_gh_token', token);
      localStorage.setItem('optics_gh_owner', owner);
      localStorage.setItem('optics_gh_repo', repo);
      localStorage.setItem('optics_gh_path', path);
      localStorage.setItem('optics_gh_auto_sync', autoSync ? 'true' : 'false');

      closeGithubModal();
      showToast('GitHub 同步設定已成功儲存！');
    }

    async function syncToGitHub(manualTrigger = false) {
      const token = localStorage.getItem('optics_gh_token') || (document.getElementById('ghTokenInput') ? document.getElementById('ghTokenInput').value.trim() : '');
      const owner = localStorage.getItem('optics_gh_owner') || (document.getElementById('ghOwnerInput') ? document.getElementById('ghOwnerInput').value.trim() : '');
      const repo = localStorage.getItem('optics_gh_repo') || (document.getElementById('ghRepoInput') ? document.getElementById('ghRepoInput').value.trim() : '');
      let path = localStorage.getItem('optics_gh_path') || (document.getElementById('ghPathInput') ? document.getElementById('ghPathInput').value.trim() : '');
      if (!path) path = 'data/production_data.json';

      if (!token || !owner || !repo) {
        if (manualTrigger) {
          alert('請先填寫 GitHub Token、Repo 擁有者及 Repo 名稱！');
          openGithubModal();
        }
        return;
      }

      if (manualTrigger) showToast('開始上傳資料至 GitHub...');

      try {
        const exportData = {
          system: "光學製造生產數據管理系統",
          updatedAt: new Date().toISOString(),
          dataset: dataset,
          auditLogs: auditLogs
        };

        // 安全地處理 UTF-8 Unicode 字元轉成 Base64
        const jsonStr = JSON.stringify(exportData, null, 2);
        const contentBase64 = btoa(unescape(encodeURIComponent(jsonStr)));

        // 1. 查詢 GitHub API 以獲取當前檔案的 SHA (避開衝突並進行更新)
        let sha = "";
        const getUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
        try {
          const getRes = await fetch(getUrl, {
            headers: {
              "Authorization": `Bearer ${token}`,
              "Accept": "application/vnd.github.v3+json"
            }
          });
          if (getRes.ok) {
            const fileInfo = await getRes.json();
            sha = fileInfo.sha;
          }
        } catch (e) {
          console.log("檔案可能為全新，將建立新檔:", e);
        }

        // 2. Commit / Push 上傳資料至 GitHub 倉庫
        const putBody = {
          message: `自動備份生產數據: ${new Date().toLocaleString()}`,
          content: contentBase64
        };
        if (sha) putBody.sha = sha;

        const putRes = await fetch(getUrl, {
          method: "PUT",
          headers: {
            "Authorization": `Bearer ${token}`,
            "Content-Type": "application/json",
            "Accept": "application/vnd.github.v3+json"
          },
          body: JSON.stringify(putBody)
        });

        if (putRes.ok) {
          console.log('GitHub 同步成功！');
          if (manualTrigger) showToast('成功將資料備份 Commit 並 Push 至 GitHub！');
        } else {
          const errData = await putRes.json();
          console.error('GitHub API 上傳失敗:', errData);
          if (manualTrigger) alert(`GitHub 上傳失敗: ${errData.message || '請確認 Token 與 Repo 權限'}`);
        }
      } catch (err) {
        console.error('GitHub 同步發生異常:', err);
        if (manualTrigger) alert('GitHub 同步過程發生網路異常。');
      }
    }

    function openCloudConfigModal() {
      document.getElementById('webhookUrlInput').value = localStorage.getItem('optics_webhook_url') || '';
      document.getElementById('autoSyncCheckbox').checked = localStorage.getItem('optics_auto_sync') === 'true';
      document.getElementById('cloudConfigModal').classList.remove('hidden');
    }

    function closeCloudConfigModal() {
      document.getElementById('cloudConfigModal').classList.add('hidden');
    }

    function saveCloudConfig() {
      const url = document.getElementById('webhookUrlInput').value.trim();
      const autoSync = document.getElementById('autoSyncCheckbox').checked;
      localStorage.setItem('optics_webhook_url', url);
      localStorage.setItem('optics_auto_sync', autoSync ? 'true' : 'false');
      closeCloudConfigModal();
      showToast('雲端同步設定已儲存！');
    }

    function getCurrentDateTimeStr() {
      const now = new Date();
      const y = now.getFullYear();
      const m = String(now.getMonth() + 1).padStart(2, '0');
      const d = String(now.getDate()).padStart(2, '0');
      const hh = String(now.getHours()).padStart(2, '0');
      const mm = String(now.getMinutes()).padStart(2, '0');
      const ss = String(now.getSeconds()).padStart(2, '0');
      return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
    }

    function calcMetrics() {
      const inputNum = parseFloat(document.getElementById('inputNum').value) || 0;
      const inPiece = parseFloat(document.getElementById('inPiece').value) || 0;
      const aoiOk = parseFloat(document.getElementById('aoiOk').value) || 0;
      const aoiNg = parseFloat(document.getElementById('aoiNg').value) || 0;

      let yieldNum = 0;
      let fpynNum = 0;
      let hasError = false;

      const yieldField = document.getElementById('yieldVal');
      const yieldWarning = document.getElementById('yieldWarning');
      if (inputNum > 0 && inPiece >= 0) {
        yieldNum = (inPiece / inputNum) * 100;
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
      if (inputNum > 0 && aoiOk >= 0) {
        fpynNum = (aoiOk / inputNum) * 100;
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
      const form = e.target;
      const requiredFields = [
        { name: 'shift', label: '班別(SHIFT)' },
        { name: 'preDate', label: '前段生產日期' },
        { name: 'holeNo', label: '穴號(Model Number)' },
        { name: 'lineNo', label: '線別(Front-end equipment NO)' },
        { name: 'date', label: '日期(Date)' },
        { name: 'machine', label: '機台(Equipment NO)' },
        { name: 'startTime', label: '進站時間' },
        { name: 'endTime', label: '出站時間' },
        { name: 'operator', label: '作業人員' },
        { name: 'degree', label: '度數' },
        { name: 'runCard', label: '工單(RUN CARD)' },
        { name: 'inputNum', label: '生產投入數量' },
        { name: 'totalNum', label: '完工數量良品' },
        { name: 'inPiece', label: '入片數' },
        { name: 'aoiOk', label: 'AOI-OK' },
        { name: 'aoiOkIn', label: 'AOI-OK入片' },
        { name: 'aoiNg', label: 'AOI-NG' },
        { name: 'cpk', label: 'CPK' }
      ];

      for (let f of requiredFields) {
        const val = form.elements[f.name] ? form.elements[f.name].value.trim() : '';
        if (val === '') {
          alert(`錯誤：[${f.label}] 為必填項目，不能為空！`);
          return;
        }
      }

      const inputNum = parseFloat(document.getElementById('inputNum').value) || 0;
      const inPiece = parseFloat(document.getElementById('inPiece').value) || 0;
      const aoiOk = parseFloat(document.getElementById('aoiOk').value) || 0;

      if (inputNum > 0 && (inPiece / inputNum > 1 || aoiOk / inputNum > 1)) {
        alert('錯誤：良率或直通率大於 100%，無法儲存！');
        return;
      }

      const formData = new FormData(form);
      const editingIdStr = document.getElementById('editingId').value.trim();

      const record = {};
      formData.forEach((value, key) => {
        if (key !== 'editingId') {
          record[key] = value;
        }
      });
      // Handle checkboxes specifically
      record.d_liugong = form.elements['d_liugong'].checked ? 1 : '';
      record.d_ring = form.elements['d_ring'].checked ? 1 : '';

      const nowTime = getCurrentDateTimeStr();

      if (editingIdStr !== "") {
        const index = dataset.findIndex(d => String(d.id) === String(editingIdStr));
        if (index !== -1) {
          const oldItem = dataset[index];
          dataset[index] = { ...oldItem, ...record, id: oldItem.id };
          
          const changes = [];
          for (let key in record) {
            const oldVal = (oldItem[key] !== undefined && oldItem[key] !== null) ? String(oldItem[key]).trim() : '';
            const newVal = (record[key] !== undefined && record[key] !== null) ? String(record[key]).trim() : '';
            if (oldVal !== newVal) {
              const labelName = fieldLabelMap[key] || key;
              changes.push(`${labelName} ${oldVal === '' ? '(無)' : oldVal} 改成 ${newVal === '' ? '(無)' : newVal}`);
            }
          }

          auditLogs.push({
            time: nowTime,
            action: '修改',
            runCard: record.runCard || oldItem.runCard,
            operator: record.operator || oldItem.operator,
            detail: changes.length > 0 ? changes.join('；') : '更新資料內容'
          });
          showToast('修改生產資料成功！');
        } else {
          const newId = dataset.length ? Math.max(...dataset.map(d => Number(d.id) || 0)) + 1 : 1;
          dataset.unshift({ ...record, id: newId });
          auditLogs.push({
            time: nowTime,
            action: '新增',
            runCard: record.runCard || '-',
            operator: record.operator || '-',
            detail: `新增生產工單 (機台: ${record.machine || '-'})`
          });
          showToast('新增生產資料成功！');
        }
      } else {
        const newId = dataset.length ? Math.max(...dataset.map(d => Number(d.id) || 0)) + 1 : 1;
        dataset.unshift({ ...record, id: newId });
        auditLogs.push({
          time: nowTime,
          action: '新增',
          runCard: record.runCard || '-',
          operator: record.operator || '-',
          detail: `新增生產工單 (機台: ${record.machine || '-'})`
        });
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
        const item = dataset.find(d => String(d.id) === String(id));
        if (item) {
          auditLogs.push({
            time: getCurrentDateTimeStr(),
            action: '刪除',
            runCard: item.runCard || '-',
            operator: item.operator || '-',
            detail: `刪除工單紀錄 (機台: ${item.machine || '-'})`
          });
        }
        dataset = dataset.filter(d => String(d.id) !== String(id));
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
        auditLogs = [
          { time: getCurrentDateTimeStr(), action: '重置', runCard: 'ALL', operator: 'System', detail: '重置為系統預設數據集' }
        ];
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
        ["班別(SHIFT)", "前段生產日期(Front-end equipment production time)", "穴號(Model Number)", "線別(Front-end equipment NO)", "日期(Date)", "機台(Equipment NO)", "進站時間(Check in time)", "出站時間(Check out time)", "作業人員(Employee No)", "度數", "圖紋種類(Pattern No)", "工單(RUN CARD)",
         "生產投入數量(Current Unit No)", "良率(DM)", "入片數", "AOI-OK", "AOI-OK入片", "AOI-NG", "AOI總不良率", 
         "直通率", "完工數量良品 (good Qty)", "CPK", "備註(Remark)", "1.留公", "2.油殘(YOSAN)", "3.圖紋密度(Pattern Density)", "4.暗視野亮點(Dark field bright spot)", "5.積墨(CHIMO)", "6.邊破(PIENCHE/POPIEN)", 
         "7.氣泡(CHIPAO)", "8.附加物(FU-CHA-WU)", "混料-穴號 顆數(Model Number mixed count)", "混料顆粒(Mixed count)", "(Mixed Model Number)", "混料誤判", "9.空洞(EMPTY)", "10.反面(Flip)", "11.穴外(Outside)", "12.疊片(Double)", "13.碎片(POPIE)", "14.鏡片偏移(Shifting)", "15.留 ring", "16.通用備註(Remark)"]
      ];

      const rows = dataset.map(d => [
        d.shift, d.preDate, d.holeNo, d.lineNo, d.date, d.machine, d.startTime, d.endTime, d.operator, d.degree, d.patternNo, d.runCard,
        d.inputNum, d.yieldVal, d.inPiece, d.aoiOk, d.aoiOkIn, d.aoiNg, d.aoiTotalNgYield,
        d.fpynVal, d.totalNum, d.cpk, d.ipqcNote, d.d_liugong ? '✓' : '', d.d_yosai, d.d_pattern_dense, d.d_darkfield, d.d_chimo, d.d_edge_break,
        d.d_chipa, d.d_attach, d.d_mix_hole_particle, d.d_mix_particle, d.d_mix_hole, d.d_mix_false, d.d_empty, d.d_taobao, d.d_outside, d.d_double, d.d_popie, d.d_shift, d.d_ring ? '✓' : '', d.note
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
        updateProgress(40, '術式順轉：解析 Excel 結構', '正在解析試算表標題與內容...');
        setTimeout(() => {
          try {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            updateProgress(70, '術式反轉：動態比對標題欄位', '進行自動表頭 Mapping 與資料對應...');
            setTimeout(() => {
              try {
                const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });
                if (jsonRows.length < 2) {
                  loadingOverlay.classList.add('hidden');
                  alert('Excel 內容為空或無數據行！');
                  e.target.value = '';
                  return;
                }

                // 1. 動態 Excel 表頭對應邏輯 (Header Alignment Algorithm)
                const headerRow = jsonRows[0] || [];
                const fieldIndexMap = {};
                let matchedCount = 0;

                // 針對系統定義的所有 key，逐一比對 Excel 第一列
                for (const [fieldKey, keywords] of Object.entries(fieldMappingRules)) {
                  let foundIndex = -1;
                  for (let colIdx = 0; colIdx < headerRow.length; colIdx++) {
                    const rawHeaderStr = String(headerRow[colIdx] || '').trim().toLowerCase();
                    if (!rawHeaderStr) continue;

                    // 檢查表頭文字是否匹配任何關鍵字
                    const isMatch = keywords.some(kw => rawHeaderStr.includes(kw.toLowerCase()));
                    if (isMatch) {
                      foundIndex = colIdx;
                      break;
                    }
                  }

                  if (foundIndex !== -1) {
                    fieldIndexMap[fieldKey] = foundIndex;
                    matchedCount++;
                  } else {
                    // Fallback 機制：若表頭沒有比對出項目，自動退回預設索引
                    fieldIndexMap[fieldKey] = defaultIndexMap[fieldKey] !== undefined ? defaultIndexMap[fieldKey] : -1;
                  }
                }

                updateProgress(90, '虚式「 茈 」：重構數據模型', `成功對應 ${matchedCount} 個標題欄位，共 ${jsonRows.length - 1} 筆紀錄...`);
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
                        if (num <= 1 && num >= 0) return (num * 100).toFixed(decimals) + '%';
                        else if (num > 1) return num.toFixed(decimals) + '%';
                      }
                      return strVal;
                    }

                    // 輔助函式：透過映射地圖安全的提取欄位數值
                    function getValue(rowArray, key) {
                      const idx = fieldIndexMap[key];
                      if (idx !== undefined && idx >= 0 && idx < rowArray.length) {
                        const val = rowArray[idx];
                        return val !== undefined && val !== null ? val : '';
                      }
                      return '';
                    }

                    const newDataset = [];
                    for (let i = 1; i < jsonRows.length; i++) {
                      const r = jsonRows[i];
                      if (!r || r.length === 0) continue;

                      const d_liugong_val = getValue(r, 'd_liugong');
                      const d_ring_val = getValue(r, 'd_ring');

                      newDataset.push({
                        id: Date.now() + i,
                        shift: getValue(r, 'shift'),
                        preDate: parseExcelDate(getValue(r, 'preDate')),
                        holeNo: getValue(r, 'holeNo'),
                        lineNo: getValue(r, 'lineNo'),
                        date: parseExcelDate(getValue(r, 'date')),
                        machine: getValue(r, 'machine'),
                        startTime: parseExcelTime(getValue(r, 'startTime')),
                        endTime: parseExcelTime(getValue(r, 'endTime')),
                        operator: getValue(r, 'operator'),
                        degree: getValue(r, 'degree'),
                        patternNo: getValue(r, 'patternNo'),
                        runCard: getValue(r, 'runCard'),
                        inputNum: getValue(r, 'inputNum'),
                        yieldVal: parseExcelPercent(getValue(r, 'yieldVal'), 1),
                        inPiece: getValue(r, 'inPiece'),
                        aoiOk: getValue(r, 'aoiOk'),
                        aoiOkIn: getValue(r, 'aoiOkIn'),
                        aoiNg: getValue(r, 'aoiNg'),
                        aoiTotalNgYield: parseExcelPercent(getValue(r, 'aoiTotalNgYield'), 1),
                        fpynVal: parseExcelPercent(getValue(r, 'fpynVal'), 1),
                        totalNum: getValue(r, 'totalNum'),
                        cpk: getValue(r, 'cpk'),
                        ipqcNote: getValue(r, 'ipqcNote'),
                        d_liugong: (d_liugong_val === 1 || String(d_liugong_val).trim() === '✓' || String(d_liugong_val).toLowerCase() === 'true') ? 1 : '',
                        d_yosai: getValue(r, 'd_yosai'),
                        d_pattern_dense: getValue(r, 'd_pattern_dense'),
                        d_darkfield: getValue(r, 'd_darkfield'),
                        d_chimo: getValue(r, 'd_chimo'),
                        d_edge_break: getValue(r, 'd_edge_break'),
                        d_chipa: getValue(r, 'd_chipa'),
                        d_attach: getValue(r, 'd_attach'),
                        d_mix_hole_particle: getValue(r, 'd_mix_hole_particle'),
                        d_mix_particle: getValue(r, 'd_mix_particle'),
                        d_mix_hole: getValue(r, 'd_mix_hole'),
                        d_mix_false: getValue(r, 'd_mix_false'),
                        d_empty: getValue(r, 'd_empty'),
                        d_taobao: getValue(r, 'd_taobao'),
                        d_outside: getValue(r, 'd_outside'),
                        d_double: getValue(r, 'd_double'),
                        d_popie: getValue(r, 'd_popie'),
                        d_shift: getValue(r, 'd_shift'),
                        d_ring: (d_ring_val === 1 || String(d_ring_val).trim() === '✓' || String(d_ring_val).toLowerCase() === 'true') ? 1 : '',
                        note: getValue(r, 'note')
                      });
                    }

                    if (newDataset.length > 0) {
                      dataset = newDataset;
                      auditLogs.push({
                        time: getCurrentDateTimeStr(),
                        action: '匯入',
                        runCard: 'Excel 批次',
                        operator: 'System',
                        detail: `成功匯入 ${newDataset.length} 筆生產數據 (已自動對應 ${matchedCount} 個標題欄位)`
                      });
                      saveDatasetToStorage();
                      updateHoleFilterOptions();
                      resetPageAndRender();
                      updateDashboard();
                      updateProgress(100, '發射完成！', `成功導入 ${newDataset.length} 筆生產數據`);
                      setTimeout(() => {
                        loadingOverlay.classList.add('hidden');
                        showToast(`成功導入 ${newDataset.length} 筆 Excel 生產數據！(對應 ${matchedCount} 欄位)`);
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

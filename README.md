<!DOCTYPE html>
<html lang="th" data-theme="dark" data-layout="standard">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>TMS Dashboard - โรงพยาบาลพระศรีมหาโพธิ์</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script src="https://unpkg.com/lucide@latest"></script>
    <style>
        body { font-family: 'Prompt', sans-serif; transition: background-color 0.3s ease, color 0.3s ease; }
        
        /* THEME 1: Dark Tech Modern (Default) */
        [data-theme="dark"] { --bg-body: #0f172a; --bg-card: rgba(30, 41, 59, 0.85); --border-color: rgba(51, 65, 85, 0.8); --text-main: #cbd5e1; --text-head: #ffffff; --header-bg: linear-gradient(135deg, #093b37 0%, #065f46 50%, #042f2e 100%); --input-bg: #0f172a; --input-border: #334155; --input-text: #f8fafc; }
        
        /* THEME 2: Clean Light Minimal */
        [data-theme="light"] { --bg-body: #f8fafc; --bg-card: rgba(255, 255, 255, 0.95); --border-color: rgba(226, 232, 240, 0.8); --text-main: #334155; --text-head: #0f172a; --header-bg: linear-gradient(135deg, #0f766e 0%, #0d9488 50%, #155e75 100%); --input-bg: #f8fafc; --input-border: #cbd5e1; --input-text: #1e293b; }

        /* THEME 3: Emerald Professional */
        [data-theme="emerald"] { --bg-body: #064e3b; --bg-card: rgba(6, 78, 59, 0.9); --border-color: rgba(4, 120, 87, 0.6); --text-main: #a7f3d0; --text-head: #ffffff; --header-bg: linear-gradient(135deg, #022c22 0%, #064e3b 50%, #047857 100%); --input-bg: #022c22; --input-border: #047857; --input-text: #ecfdf5; }

        /* THEME 4: Sunset Indigo */
        [data-theme="indigo"] { --bg-body: #1e1b4b; --bg-card: rgba(30, 27, 75, 0.88); --border-color: rgba(79, 70, 229, 0.4); --text-main: #c7d2fe; --text-head: #ffffff; --header-bg: linear-gradient(135deg, #312e81 0%, #4338ca 50%, #3730a3 100%); --input-bg: #0f0e26; --input-border: #4f46e5; --input-text: #e0e7ff; }

        /* THEME 5: Royal Amber */
        [data-theme="amber"] { --bg-body: #1c1917; --bg-card: rgba(41, 37, 36, 0.9); --border-color: rgba(120, 113, 108, 0.5); --text-main: #d6d3d1; --text-head: #ffffff; --header-bg: linear-gradient(135deg, #292524 0%, #78716c 50%, #44403c 100%); --input-bg: #1c1917; --input-border: #57534e; --input-text: #f5f5f4; }

        /* THEME 6: Cyber Neon */
        [data-theme="cyber"] { --bg-body: #05050a; --bg-card: rgba(15, 15, 26, 0.9); --border-color: rgba(236, 72, 153, 0.3); --text-main: #e2e8f0; --text-head: #ffffff; --header-bg: linear-gradient(135deg, #831843 0%, #581c87 50%, #1e1b4b 100%); --input-bg: #05050a; --input-border: #8b5cf6; --input-text: #ffffff; }

        body { background-color: var(--bg-body); color: var(--text-main); }
        .glass-card {
            background: var(--bg-card);
            backdrop-filter: blur(16px);
            border: 1px solid var(--border-color);
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.15);
            transition: all 0.3s ease;
        }
        .glass-card:hover {
            box-shadow: 0 20px 35px -10px rgba(0, 0, 0, 0.25);
            transform: translateY(-2px);
        }
        .gradient-header { background: var(--header-bg); }
        .soft-input {
            background-color: var(--input-bg);
            border: 1px solid var(--input-border);
            color: var(--input-text);
            transition: all 0.2s ease;
        }
        .soft-input:focus {
            border-color: #14b8a6;
            box-shadow: 0 0 0 3px rgba(20, 184, 166, 0.2);
            outline: none;
        }
        .soft-input option { background-color: var(--input-bg); color: var(--input-text); }
        
        /* Layout Configurations */
        [data-layout="standard"] .metrics-grid { grid-template-columns: repeat(5, minmax(0, 1fr)); }
        [data-layout="standard"] .charts-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }

        [data-layout="compact"] .metrics-grid { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); }
        [data-layout="compact"] .charts-grid { grid-template-columns: 1fr; }
        
        [data-layout="expanded"] .metrics-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        [data-layout="expanded"] .charts-grid { grid-template-columns: 1fr; }

        @media (max-width: 1024px) {
            .metrics-grid { grid-template-columns: repeat(2, minmax(0, 1fr)) !important; }
            .charts-grid { grid-template-columns: 1fr !important; }
        }

        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: var(--bg-body); }
        ::-webkit-scrollbar-thumb { background: var(--border-color); border-radius: 4px; }
    </style>
</head>
<body class="antialiased min-h-screen selection:bg-teal-500 selection:text-white">
    
    <!-- Top Organized Executive Header -->
    <header class="gradient-header text-white sticky top-0 z-40 shadow-xl border-b border-white/10">
        <div class="max-w-7xl mx-auto px-4 py-4 sm:px-6 lg:px-8 space-y-4">
            
            <!-- Row 1: Brand & Primary Action Buttons -->
            <div class="flex flex-col md:flex-row justify-between items-center gap-4">
                <div class="flex items-center gap-4 w-full md:w-auto justify-between md:justify-start">
                    <div class="flex items-center gap-3.5">
                        <div class="bg-black/25 p-3 rounded-2xl border border-white/20 text-white shadow-inner flex items-center justify-center">
                            <i data-lucide="activity" class="w-7 h-7 text-teal-300"></i>
                        </div>
                        <div>
                            <div class="flex items-center gap-2">
                                <h1 class="text-xl sm:text-2xl font-bold tracking-tight text-white">TMS Dashboard</h1>
                                <span class="text-[10px] font-semibold px-2.5 py-0.5 rounded-full bg-white/20 text-white border border-white/30 flex items-center gap-1">
                                    <i data-lucide="cloud-cog" class="w-3 h-3"></i> Google Sheets Sync
                                </span>
                            </div>
                            <p class="text-white/80 text-xs mt-0.5 flex items-center gap-1.5 font-light">
                                <i data-lucide="hospital" class="w-3.5 h-3.5 text-teal-300"></i> โรงพยาบาลพระศรีมหาโพธิ์ 
                            </p>
                        </div>
                    </div>
                </div>

                <!-- Primary Action Buttons -->
                <div class="flex items-center gap-2.5 w-full md:w-auto justify-end flex-wrap">
                    <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white font-medium px-4 py-2 rounded-xl shadow-md flex items-center gap-2 text-sm transition-all border border-emerald-400/40">
                        <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Export Excel
                    </button>
                    <button onclick="openAddModal()" class="bg-white hover:bg-teal-50 text-slate-900 font-semibold px-4 py-2 rounded-xl shadow-md flex items-center gap-2 text-sm transition-all border border-white/80">
                        <i data-lucide="user-plus" class="w-4 h-4 text-teal-700"></i> บันทึกเคสใหม่
                    </button>
                </div>
            </div>

            <!-- Row 2: Navigation Tabs, Theme & Layout Control Bar -->
            <div class="pt-3 border-t border-white/15 flex flex-col lg:flex-row justify-between items-center gap-3">
                <!-- Navigation Tabs -->
                <div class="bg-black/20 p-1 rounded-2xl border border-white/15 flex gap-1 w-full lg:w-auto">
                    <button onclick="switchTab('dashboard')" id="tab-btn-dashboard" class="flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all bg-white text-slate-900 shadow-md flex items-center justify-center gap-1.5">
                        <i data-lucide="pie-chart" class="w-4 h-4 text-teal-700"></i> ภาพรวม Dashboard
                    </button>
                    <button onclick="switchTab('patients')" id="tab-btn-patients" class="flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all text-white/90 hover:bg-white/15 hover:text-white flex items-center justify-center gap-1.5">
                        <i data-lucide="users" class="w-4 h-4"></i> ทะเบียนผู้ป่วย
                    </button>
                </div>

                <!-- Customization Controls (Cute Theme & Layout Switchers) -->
                <div class="flex items-center gap-3 flex-wrap justify-center lg:justify-end w-full lg:w-auto">
                    <!-- Cute Theme Selector Icons -->
                    <div class="flex items-center gap-1 bg-black/25 px-3 py-1.5 rounded-2xl border border-white/15">
                        <span class="text-xs text-white/70 font-medium mr-1.5 flex items-center gap-1"><i data-lucide="palette" class="w-3.5 h-3.5"></i> ธีมสี:</span>
                        <button onclick="setTheme('dark')" id="theme-btn-dark" title="ดาร์กเทคโมเดิร์น" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-slate-200">
                            <i data-lucide="moon" class="w-4 h-4"></i>
                        </button>
                        <button onclick="setTheme('light')" id="theme-btn-light" title="คลีนไลท์สว่าง" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-amber-300">
                            <i data-lucide="sun" class="w-4 h-4"></i>
                        </button>
                        <button onclick="setTheme('emerald')" id="theme-btn-emerald" title="เอมเมอรัลด์กรีน" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-emerald-300">
                            <i data-lucide="trees" class="w-4 h-4"></i>
                        </button>
                        <button onclick="setTheme('indigo')" id="theme-btn-indigo" title="ซันเซ็ตอินดิโก้" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-indigo-300">
                            <i data-lucide="sparkles" class="w-4 h-4"></i>
                        </button>
                        <button onclick="setTheme('amber')" id="theme-btn-amber" title="รอยัลแอมเบอร์" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-amber-400">
                            <i data-lucide="crown" class="w-4 h-4"></i>
                        </button>
                        <button onclick="setTheme('cyber')" id="theme-btn-cyber" title="ไซเบอร์นีออน" class="p-1.5 rounded-xl hover:bg-white/20 transition-all text-pink-400">
                            <i data-lucide="zap" class="w-4 h-4"></i>
                        </button>
                    </div>

                    <!-- Layout Selector Buttons -->
                    <div class="flex items-center gap-1 bg-black/20 p-1 rounded-2xl border border-white/15">
                        <button onclick="setLayout('standard')" id="layout-btn-standard" title="Standard Grid" class="px-3 py-1.5 rounded-xl text-xs font-medium text-white/80 hover:bg-white/20 transition-all flex items-center gap-1">
                            <i data-lucide="layout-grid" class="w-3.5 h-3.5"></i> มาตรฐาน
                        </button>
                        <button onclick="setLayout('compact')" id="layout-btn-compact" title="Compact Focus" class="px-3 py-1.5 rounded-xl text-xs font-medium text-white/80 hover:bg-white/20 transition-all flex items-center gap-1">
                            <i data-lucide="align-justify" class="w-3.5 h-3.5"></i> กระชับ
                        </button>
                        <button onclick="setLayout('expanded')" id="layout-btn-expanded" title="Expanded Analytics" class="px-3 py-1.5 rounded-xl text-xs font-medium text-white/80 hover:bg-white/20 transition-all flex items-center gap-1">
                            <i data-lucide="maximize-2" class="w-3.5 h-3.5"></i> ขยายกราฟ
                        </button>
                    </div>
                </div>
            </div>

        </div>
    </header>

    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 space-y-8">
        
        <!-- ================= TAB 1: DASHBOARD OVERVIEW ================= -->
        <div id="tab-dashboard" class="space-y-8">
            
            <!-- Welcome Info Banner -->
            <div class="glass-card rounded-3xl p-6 sm:p-8 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
                <div class="space-y-1">
                    <div class="flex items-center gap-2 font-medium text-xs bg-teal-500/20 text-teal-300 w-fit px-3 py-1 rounded-full border border-teal-500/30">
                        <span class="w-2 h-2 rounded-full bg-teal-400 animate-pulse"></span> เชื่อมต่อฐานข้อมูล Real-time
                    </div>
                    <h2 class="text-xl sm:text-2xl font-bold pt-1" style="color: var(--text-head)">ยินดีต้อนรับสู่ศูนย์ข้อมูล TMS</h2>
                    <p class="text-xs opacity-80">ระบบติดตามและประเมินผลการรักษาด้วยเครื่องกระตุ้นแม่เหล็กไฟฟ้าสมอง</p>
                </div>
                <div class="glass-card px-4 py-3 rounded-2xl flex items-center gap-3">
                    <div class="p-2.5 bg-teal-500/20 text-teal-400 rounded-xl">
                        <i data-lucide="calendar-check" class="w-5 h-5"></i>
                    </div>
                    <div>
                        <p class="text-[10px] opacity-60 font-semibold uppercase tracking-wider">สถานะการซิงค์</p>
                        <p id="current-date-display" class="text-xs font-bold">กำลังซิงค์ข้อมูล...</p>
                    </div>
                </div>
            </div>

            <!-- Global Dashboard Filter Panel -->
            <div class="glass-card rounded-3xl p-6 space-y-4">
                <div class="flex items-center justify-between border-b pb-3" style="border-color: var(--border-color)">
                    <div class="flex items-center gap-2.5">
                        <div class="p-2 bg-teal-500/20 text-teal-400 rounded-xl">
                            <i data-lucide="sliders-horizontal" class="w-4 h-4"></i>
                        </div>
                        <h2 class="text-sm font-bold" style="color: var(--text-head)">ตัวกรองข้อมูลภาพรวม (Analytics Filter)</h2>
                    </div>
                    <button onclick="resetDashboardFilters()" class="text-xs font-semibold flex items-center gap-1.5 bg-teal-500/10 hover:bg-teal-500/20 px-3 py-1.5 rounded-xl transition-all text-teal-400">
                        <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> ล้างตัวกรองทั้งหมด
                    </button>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-1">
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">ชนิดเครื่อง TMS</label>
                        <select id="dash-filter-machine" onchange="applyGlobalDashboardFilter()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทั้งหมด (ทุกเครื่อง TMS)</option>
                            <option value="เครื่อง 1">เครื่อง TMS 1</option>
                            <option value="เครื่อง 2">เครื่อง TMS 2</option>
                            <option value="เครื่อง 3">เครื่อง TMS 3</option>
                            <option value="เครื่อง 4">เครื่อง TMS 4</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">ผลการรักษา (Status)</label>
                        <select id="dash-filter-status" onchange="applyGlobalDashboardFilter()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทั้งหมด (ทุกสถานะ)</option>
                            <option value="หาย">หาย (Recovery)</option>
                            <option value="response">Response</option>
                            <option value="no response">No Response</option>
                            <option value="drop out">Drop Out</option>
                            <option value="ไม่ครบ">ไม่ครบ</option>
                            <option value="รอผล">รอผล</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">Protocol (โปรโตคอล)</label>
                        <select id="dash-filter-protocol" onchange="applyGlobalDashboardFilter()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทุก Protocol</option>
                            <option value="itbs">iTBS</option>
                            <option value="standard">Standard</option>
                            <option value="swift">Swift</option>
                        </select>
                    </div>
                </div>
            </div>

            <!-- Summary Metric Cards -->
            <div class="metrics-grid grid gap-5">
                <div class="glass-card rounded-3xl p-6 relative overflow-hidden group">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-teal-500/10 rounded-full opacity-40 group-hover:scale-110 transition-transform"></div>
                    <div class="relative z-10">
                        <div class="flex justify-between items-start">
                            <div>
                                <p class="text-[11px] font-bold opacity-60 uppercase tracking-wider">ผู้ป่วยทั้งหมด</p>
                                <h3 id="stat-total" class="text-3xl font-extrabold mt-2" style="color: var(--text-head)">0</h3>
                            </div>
                            <div class="p-3 bg-teal-500/20 text-teal-400 rounded-2xl shadow-sm">
                                <i data-lucide="users" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <div class="mt-4 pt-3 border-t opacity-80 text-xs flex items-center justify-between" style="border-color: var(--border-color)">
                            <span>สัดส่วนเพศ:</span>
                            <span id="stat-gender-breakdown" class="text-teal-400 font-bold">ชาย 0 | หญิง 0</span>
                        </div>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 relative overflow-hidden group">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-blue-500/10 rounded-full opacity-40 group-hover:scale-110 transition-transform"></div>
                    <div class="relative z-10">
                        <div class="flex justify-between items-start">
                            <div>
                                <p class="text-[11px] font-bold opacity-60 uppercase tracking-wider">ทำครบแล้ว</p>
                                <h3 id="stat-completed" class="text-3xl font-extrabold text-blue-400 mt-2">0</h3>
                            </div>
                            <div class="p-3 bg-blue-500/20 text-blue-400 rounded-2xl shadow-sm">
                                <i data-lucide="check-circle-2" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <div id="stat-completed-desc" class="mt-4 pt-3 border-t opacity-80 text-xs font-medium" style="border-color: var(--border-color)">คิดเป็น 0% ของทั้งหมด</div>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 relative overflow-hidden group">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-emerald-500/10 rounded-full opacity-40 group-hover:scale-110 transition-transform"></div>
                    <div class="relative z-10">
                        <div class="flex justify-between items-start">
                            <div>
                                <p class="text-[11px] font-bold opacity-60 uppercase tracking-wider">หาย (Recovery)</p>
                                <h3 id="stat-recovered" class="text-3xl font-extrabold text-emerald-400 mt-2">0</h3>
                            </div>
                            <div class="p-3 bg-emerald-500/20 text-emerald-400 rounded-2xl shadow-sm">
                                <i data-lucide="award" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <div id="stat-recovered-desc" class="mt-4 pt-3 border-t text-xs text-emerald-400 font-bold" style="border-color: var(--border-color)">0% ของผู้ป่วยที่ทำครบ</div>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 relative overflow-hidden group">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-amber-500/10 rounded-full opacity-40 group-hover:scale-110 transition-transform"></div>
                    <div class="relative z-10">
                        <div class="flex justify-between items-start">
                            <div>
                                <p class="text-[11px] font-bold opacity-60 uppercase tracking-wider">ไม่ตอบสนอง / Drop</p>
                                <h3 id="stat-other-issues" class="text-3xl font-extrabold text-amber-400 mt-2">0</h3>
                            </div>
                            <div class="p-3 bg-amber-500/20 text-amber-400 rounded-2xl shadow-sm">
                                <i data-lucide="alert-triangle" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <div class="mt-4 pt-3 border-t opacity-80 text-xs" style="border-color: var(--border-color)">No response / ไม่ครบ / Drop Out</div>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 relative overflow-hidden group">
                    <div class="absolute -right-4 -bottom-4 w-24 h-24 bg-indigo-500/10 rounded-full opacity-40 group-hover:scale-110 transition-transform"></div>
                    <div class="relative z-10">
                        <div class="flex justify-between items-start">
                            <div>
                                <p class="text-[11px] font-bold opacity-60 uppercase tracking-wider">จำนวนครั้งบำบัดรวม</p>
                                <h3 id="stat-sessions" class="text-3xl font-extrabold text-indigo-400 mt-2">0</h3>
                            </div>
                            <div class="p-3 bg-indigo-500/20 text-indigo-400 rounded-2xl shadow-sm">
                                <i data-lucide="activity" class="w-5 h-5"></i>
                            </div>
                        </div>
                        <div class="mt-4 pt-3 border-t opacity-80 text-xs" style="border-color: var(--border-color)">รวมทุกเครื่อง TMS ในระบบ</div>
                    </div>
                </div>
            </div>

            <!-- Charts Section -->
            <div class="charts-grid grid gap-6">
                <div class="glass-card rounded-3xl p-6 flex flex-col">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="text-sm font-bold flex items-center gap-2" style="color: var(--text-head)">
                            <span class="p-2 bg-teal-500/20 text-teal-400 rounded-xl"><i data-lucide="pie-chart" class="w-4 h-4"></i></span> 
                            สัดส่วนผลการรักษา
                        </h3>
                        <span class="text-[10px] opacity-60 bg-white/10 px-2.5 py-1 rounded-full font-medium">Status</span>
                    </div>
                    <div class="relative flex-grow flex items-center justify-center min-h-[280px]">
                        <canvas id="statusChart"></canvas>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 flex flex-col">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="text-sm font-bold flex items-center gap-2" style="color: var(--text-head)">
                            <span class="p-2 bg-blue-500/20 text-blue-400 rounded-xl"><i data-lucide="bar-chart-3" class="w-4 h-4"></i></span> 
                            ช่วงอายุและเพศผู้ป่วย
                        </h3>
                        <span class="text-[10px] opacity-60 bg-white/10 px-2.5 py-1 rounded-full font-medium">Age & Gender</span>
                    </div>
                    <div class="relative flex-grow flex items-center justify-center min-h-[280px]">
                        <canvas id="ageGenderChart"></canvas>
                    </div>
                </div>

                <div class="glass-card rounded-3xl p-6 flex flex-col">
                    <div class="flex items-center justify-between mb-4">
                        <h3 class="text-sm font-bold flex items-center gap-2" style="color: var(--text-head)">
                            <span class="p-2 bg-indigo-500/20 text-indigo-400 rounded-xl"><i data-lucide="cpu" class="w-4 h-4"></i></span> 
                            ชนิดการใช้งานเครื่อง TMS
                        </h3>
                        <span class="text-[10px] opacity-60 bg-white/10 px-2.5 py-1 rounded-full font-medium">Machine</span>
                    </div>
                    <div class="relative flex-grow flex items-center justify-center min-h-[280px]">
                        <canvas id="machineChart"></canvas>
                    </div>
                </div>
            </div>
        </div>

        <!-- ================= TAB 2: PATIENTS DIRECTORY & HISTORY ================= -->
        <div id="tab-patients" class="space-y-6 hidden">
            <!-- Advanced Filters Panel -->
            <div class="glass-card rounded-3xl p-6 space-y-4">
                <div class="flex items-center justify-between border-b pb-3" style="border-color: var(--border-color)">
                    <div class="flex items-center gap-2.5">
                        <div class="p-2 bg-teal-500/20 text-teal-400 rounded-xl">
                            <i data-lucide="search" class="w-4 h-4"></i>
                        </div>
                        <h2 class="text-sm font-bold" style="color: var(--text-head)">ค้นหาและกรองข้อมูลทะเบียนผู้ป่วย</h2>
                    </div>
                    <button onclick="resetPatientFilters()" class="text-xs font-semibold flex items-center gap-1.5 bg-teal-500/10 hover:bg-teal-500/20 px-3 py-1.5 rounded-xl transition-all text-teal-400">
                        <i data-lucide="rotate-ccw" class="w-3.5 h-3.5"></i> ล้างตัวกรอง
                    </button>
                </div>
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">ชนิดเครื่อง TMS</label>
                        <select id="filter-machine" onchange="applyPatientFilters()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทั้งหมด (ทุกเครื่อง)</option>
                            <option value="เครื่อง 1">เครื่อง TMS 1</option>
                            <option value="เครื่อง 2">เครื่อง TMS 2</option>
                            <option value="เครื่อง 3">เครื่อง TMS 3</option>
                            <option value="เครื่อง 4">เครื่อง TMS 4</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">ผลการรักษา (Status)</label>
                        <select id="filter-status" onchange="applyPatientFilters()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทั้งหมด</option>
                            <option value="หาย">หาย</option>
                            <option value="response">Response</option>
                            <option value="no response">No Response</option>
                            <option value="drop out">Drop Out</option>
                            <option value="ไม่ครบ">ไม่ครบ</option>
                            <option value="รอผล">รอผล</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">Protocol (โปรโตคอล)</label>
                        <select id="filter-protocol" onchange="applyPatientFilters()" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="">ทุก Protocol</option>
                            <option value="itbs">iTBS</option>
                            <option value="standard">Standard</option>
                            <option value="swift">Swift</option>
                        </select>
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1.5 opacity-80">คำค้นหาอิสระ</label>
                        <div class="relative">
                            <input type="text" id="filter-search" oninput="applyPatientFilters()" placeholder="ค้นหาชื่อ, Dx, อาการ..." class="w-full soft-input rounded-2xl pl-10 pr-4 py-2.5 text-sm">
                            <i data-lucide="search" class="w-4 h-4 opacity-50 absolute left-3.5 top-3"></i>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Patient Table Section -->
            <div class="glass-card rounded-3xl overflow-hidden shadow-xl">
                <div class="p-6 border-b flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4" style="border-color: var(--border-color)">
                    <div>
                        <h3 class="text-base font-bold" style="color: var(--text-head)">ทะเบียนรายชื่อผู้ป่วยและประวัติการรักษา TMS</h3>
                        <p class="text-xs opacity-60 mt-0.5">คลิกที่แถวตารางเพื่อดูรายละเอียดประวัติ หรือแก้ไขข้อมูลเคส</p>
                    </div>
                    <div class="text-xs font-bold bg-teal-500/20 text-teal-300 px-4 py-2 rounded-2xl border border-teal-500/30 flex items-center gap-2">
                        <i data-lucide="database" class="w-3.5 h-3.5 text-teal-400"></i> แสดงผล <span id="table-count">0</span> รายการ
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse">
                        <thead>
                            <tr class="text-xs uppercase tracking-wider font-bold border-b opacity-80" style="border-color: var(--border-color); background: rgba(0,0,0,0.1)">
                                <th class="py-4 px-5">ลำดับ</th>
                                <th class="py-4 px-5">เพศ/อายุ</th>
                                <th class="py-4 px-5">Dx & สิทธิ</th>
                                <th class="py-4 px-5">โปรโตคอล</th>
                                <th class="py-4 px-5">เครื่อง TMS</th>
                                <th class="py-4 px-5 text-center">จำนวนครั้ง</th>
                                <th class="py-4 px-5 text-center">สถานะผลลัพธ์</th>
                                <th class="py-4 px-5 text-center">จัดการ</th>
                            </tr>
                        </thead>
                        <tbody id="patient-table-body" class="divide-y text-sm" style="divide-color: var(--border-color)">
                            <!-- Populated by JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </main>

    <!-- Detail Modal -->
    <div id="detail-modal" class="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 hidden flex items-center justify-center p-4">
        <div class="glass-card rounded-3xl max-w-2xl w-full max-h-[90vh] overflow-y-auto shadow-2xl p-6 sm:p-8 space-y-6">
            <div class="flex justify-between items-start border-b pb-4" style="border-color: var(--border-color)">
                <div>
                    <span id="modal-no" class="text-xs font-bold bg-teal-500/20 text-teal-300 px-3 py-1 rounded-full border border-teal-500/30">เคสที่ #1</span>
                    <h3 class="text-lg font-bold mt-2" style="color: var(--text-head)">ข้อมูลรายละเอียดผู้ป่วย</h3>
                </div>
                <button onclick="closeModal()" class="p-2 opacity-70 hover:opacity-100 hover:bg-white/10 rounded-full transition-colors">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>

            <div class="grid grid-cols-2 sm:grid-cols-4 gap-4 p-4 rounded-2xl text-sm border" style="background: rgba(0,0,0,0.15); border-color: var(--border-color)">
                <div>
                    <p class="text-xs opacity-60 font-medium">เพศ / อายุ</p>
                    <p id="modal-gender-age" class="font-bold mt-0.5">-</p>
                </div>
                <div>
                    <p class="text-xs opacity-60 font-medium">การวินิจฉัย (Dx)</p>
                    <p id="modal-dx" class="font-bold mt-0.5">-</p>
                </div>
                <div>
                    <p class="text-xs opacity-60 font-medium">สิทธิการรักษา</p>
                    <p id="modal-rights" class="font-bold mt-0.5">-</p>
                </div>
                <div>
                    <p class="text-xs opacity-60 font-medium">จำนวนครั้งทำ</p>
                    <p id="modal-sessions" class="font-bold text-teal-400 mt-0.5">- ครั้ง</p>
                </div>
            </div>

            <div class="space-y-4">
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider opacity-60 mb-1">โปรโตคอลการทำ TMS</h4>
                    <p id="modal-dtype" class="text-sm p-3.5 rounded-2xl border whitespace-pre-line" style="background: rgba(0,0,0,0.1); border-color: var(--border-color)">-</p>
                </div>
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider text-amber-400 mb-1">อาการก่อนทำ</h4>
                    <p id="modal-sym-before" class="text-sm bg-amber-500/10 border border-amber-500/30 p-3.5 rounded-2xl text-amber-300 whitespace-pre-line">-</p>
                </div>
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider text-emerald-400 mb-1">อาการหลังทำ / ผลลัพธ์</h4>
                    <p id="modal-sym-after" class="text-sm bg-emerald-500/10 border border-emerald-500/30 p-3.5 rounded-2xl text-emerald-300 whitespace-pre-line">-</p>
                </div>
                <div>
                    <h4 class="text-xs font-bold uppercase tracking-wider text-rose-400 mb-1">ผลข้างเคียง (Side Effects / S/E)</h4>
                    <p id="modal-se" class="text-sm bg-rose-500/10 border border-rose-500/30 p-3.5 rounded-2xl text-rose-300 whitespace-pre-line">-</p>
                </div>
            </div>

            <div class="flex justify-between items-center pt-4 border-t" style="border-color: var(--border-color)">
                <div class="flex gap-2">
                    <button id="modal-edit-btn" onclick="" class="bg-amber-600 hover:bg-amber-500 text-white px-4 py-2.5 rounded-2xl font-medium text-sm transition-colors flex items-center gap-2 shadow-sm">
                        <i data-lucide="edit-3" class="w-4 h-4"></i> แก้ไขข้อมูล
                    </button>
                    <button id="modal-delete-btn" onclick="" class="bg-rose-600 hover:bg-rose-500 text-white px-4 py-2.5 rounded-2xl font-medium text-sm transition-colors flex items-center gap-2 shadow-sm">
                        <i data-lucide="trash-2" class="w-4 h-4"></i> ลบเคส
                    </button>
                </div>
                <button onclick="closeModal()" class="bg-white/10 hover:bg-white/20 px-5 py-2.5 rounded-2xl font-medium transition-colors shadow-sm">
                    ปิดหน้าต่าง
                </button>
            </div>
        </div>
    </div>

    <!-- Add / Edit Patient Modal -->
    <div id="add-modal" class="fixed inset-0 bg-black/60 backdrop-blur-sm z-50 hidden flex items-center justify-center p-4">
        <div class="glass-card rounded-3xl max-w-xl w-full max-h-[90vh] overflow-y-auto shadow-2xl p-6 sm:p-8 space-y-6">
            <div class="flex justify-between items-start border-b pb-4" style="border-color: var(--border-color)">
                <div>
                    <h3 id="modal-form-title" class="text-lg font-bold" style="color: var(--text-head)">ลงข้อมูลเคสผู้ป่วยใหม่</h3>
                    <p id="modal-form-subtitle" class="text-xs opacity-60">บันทึกข้อมูลเข้า Google Sheet กลาง</p>
                </div>
                <button onclick="closeAddModal()" class="p-2 opacity-70 hover:opacity-100 hover:bg-white/10 rounded-full transition-colors">
                    <i data-lucide="x" class="w-5 h-5"></i>
                </button>
            </div>

            <form id="add-patient-form" onsubmit="handleFormSubmit(event)" class="space-y-4">
                <input type="hidden" id="form-edit-index" value="-1">
                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">ลำดับเคส (No.)</label>
                        <input type="text" id="form-no" required placeholder="เช่น 34" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">เพศ</label>
                        <select id="form-gender" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="ช">ชาย</option>
                            <option value="ญ">หญิง</option>
                        </select>
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">อายุ (ปี)</label>
                        <input type="number" id="form-age" required placeholder="เช่น 35" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">สิทธิการรักษา</label>
                        <select id="form-rights" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="เบิกได้">เบิกได้</option>
                            <option value="ปสน.">ปสน.</option>
                            <option value="ปกส.">ปกส.</option>
                            <option value="จ่ายเอง">จ่ายเอง</option>
                        </select>
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">การวินิจฉัย (Dx)</label>
                        <input type="text" id="form-dx" required placeholder="เช่น F32.1 with TRD" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">เครื่อง TMS</label>
                        <select id="form-machine" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="เครื่อง 1">เครื่อง 1</option>
                            <option value="เครื่อง 2">เครื่อง 2</option>
                            <option value="เครื่อง 3">เครื่อง 3</option>
                            <option value="เครื่อง 4">เครื่อง 4</option>
                        </select>
                    </div>
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">จำนวนครั้งบำบัด</label>
                        <input type="text" id="form-sessions" required placeholder="เช่น 30" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                    </div>
                    <div>
                        <label class="block text-xs font-semibold mb-1 opacity-80">ผลการรักษา (Status)</label>
                        <select id="form-status" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                            <option value="หาย">หาย</option>
                            <option value="response">Response</option>
                            <option value="no response">No Response</option>
                            <option value="drop out">Drop Out</option>
                            <option value="ไม่ครบ">ไม่ครบ</option>
                            <option value="รอผล">รอผล</option>
                        </select>
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-semibold mb-1 opacity-80">โปรโตคอล / รายละเอียดการทำ (เช่น iTBS, standard, swift)</label>
                    <input type="text" id="form-dtype" placeholder="เช่น iTBS ทำ 10-40-10 นาที" class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm">
                </div>

                <div>
                    <label class="block text-xs font-semibold mb-1 opacity-80">อาการก่อนทำ</label>
                    <textarea id="form-sym-before" rows="2" placeholder="ระบุอาการก่อนรักษา..." class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm"></textarea>
                </div>

                <div>
                    <label class="block text-xs font-semibold mb-1 opacity-80">อาการหลังทำ / ผลลัพธ์</label>
                    <textarea id="form-sym-after" rows="2" placeholder="ระบุอาการหลังรักษา..." class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm"></textarea>
                </div>

                <div>
                    <label class="block text-xs font-semibold mb-1 opacity-80">ผลข้างเคียง (Side Effects / S/E)</label>
                    <textarea id="form-se" rows="2" placeholder="ระบุผลข้างเคียง (ถ้ามี)..." class="w-full soft-input rounded-2xl px-4 py-2.5 text-sm"></textarea>
                </div>

                <div class="flex justify-end gap-3 pt-4 border-t" style="border-color: var(--border-color)">
                    <button type="button" onclick="closeAddModal()" class="bg-white/10 hover:bg-white/20 px-4 py-2.5 rounded-2xl font-medium text-sm transition-colors">
                        ยกเลิก
                    </button>
                    <button type="submit" id="form-submit-btn" class="bg-teal-600 hover:bg-teal-500 text-white px-5 py-2.5 rounded-2xl font-medium text-sm transition-colors shadow-sm">
                        บันทึกข้อมูล
                    </button>
                </div>
            </form>
        </div>
    </div>

    <script>
        // Google Apps Script Web App URL
        const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwN1bRXihT6DEiWojUAX7TRWqW2uZGMmfIchQfvRKR0IhxboLErN4DF7GvHBZx5xQTkTw/exec";

        let patientsData = [];
        let statusChartInstance = null;
        let ageGenderChartInstance = null;
        let machineChartInstance = null;

        // Theme Switcher Logic
        function setTheme(themeName) {
            document.documentElement.setAttribute('data-theme', themeName);
            localStorage.setItem('tms_dashboard_theme', themeName);

            // Update active states for theme buttons
            ['dark', 'light', 'emerald', 'indigo', 'amber', 'cyber'].forEach(t => {
                const btn = document.getElementById(`theme-btn-${t}`);
                if (btn) {
                    if (t === themeName) {
                        btn.className = "p-1.5 rounded-xl bg-white/25 shadow-sm transition-all scale-105";
                    } else {
                        btn.className = "p-1.5 rounded-xl hover:bg-white/20 transition-all opacity-70 hover:opacity-100";
                    }
                }
            });

            if (patientsData.length > 0) {
                applyGlobalDashboardFilter();
            }
        }

        // Layout Switcher Logic
        function setLayout(layoutName) {
            document.documentElement.setAttribute('data-layout', layoutName);
            localStorage.setItem('tms_dashboard_layout', layoutName);

            // Update active states for layout buttons
            ['standard', 'compact', 'expanded'].forEach(l => {
                const btn = document.getElementById(`layout-btn-${l}`);
                if (btn) {
                    if (l === layoutName) {
                        btn.className = "px-3 py-1.5 rounded-xl text-xs font-medium bg-white text-slate-900 shadow-sm transition-all flex items-center gap-1";
                    } else {
                        btn.className = "px-3 py-1.5 rounded-xl text-xs font-medium text-white/80 hover:bg-white/20 transition-all flex items-center gap-1";
                    }
                }
            });
        }

        // Load saved preferences on startup
        const savedTheme = localStorage.getItem('tms_dashboard_theme') || 'dark';
        const savedLayout = localStorage.getItem('tms_dashboard_layout') || 'standard';
        setTheme(savedTheme);
        setLayout(savedLayout);

        async function loadPatientsFromGoogleSheet() {
            try {
                document.getElementById('current-date-display').innerText = "กำลังซิงค์ข้อมูลจาก Google Sheet...";
                const response = await fetch(WEB_APP_URL);
                const data = await response.json();
                
                if (data && data.length > 0) {
                    patientsData = data;
                }
                initDashboard();
            } catch (error) {
                console.error("Error loading data:", error);
                document.getElementById('current-date-display').innerText = "เชื่อมต่อ Google Sheet ไม่สำเร็จ";
            }
        }

        function initDashboard() {
            lucide.createIcons();
            const options = { year: 'numeric', month: 'long', day: 'numeric' };
            document.getElementById('current-date-display').innerText = new Date().toLocaleDateString('th-TH', options) + " (เชื่อมต่อแล้ว)";
            applyGlobalDashboardFilter();
            applyPatientFilters();
        }

        function switchTab(tabName) {
            const dashboardTab = document.getElementById('tab-dashboard');
            const patientsTab = document.getElementById('tab-patients');
            const btnDash = document.getElementById('tab-btn-dashboard');
            const btnPat = document.getElementById('tab-btn-patients');

            dashboardTab.classList.add('hidden');
            patientsTab.classList.add('hidden');
            
            btnDash.className = "flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all text-white/90 hover:bg-white/15 hover:text-white flex items-center justify-center gap-1.5";
            btnPat.className = "flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all text-white/90 hover:bg-white/15 hover:text-white flex items-center justify-center gap-1.5";

            if (tabName === 'dashboard') {
                dashboardTab.classList.remove('hidden');
                btnDash.className = "flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all bg-white text-slate-900 shadow-md flex items-center justify-center gap-1.5";
                applyGlobalDashboardFilter();
            } else if (tabName === 'patients') {
                patientsTab.classList.remove('hidden');
                btnPat.className = "flex-1 lg:flex-none px-4 py-2 rounded-xl text-xs sm:text-sm font-medium transition-all bg-white text-slate-900 shadow-md flex items-center justify-center gap-1.5";
                applyPatientFilters();
            }
        }

        function resetDashboardFilters() {
            document.getElementById('dash-filter-machine').value = '';
            document.getElementById('dash-filter-status').value = '';
            document.getElementById('dash-filter-protocol').value = '';
            applyGlobalDashboardFilter();
        }

        function applyGlobalDashboardFilter() {
            const machineFilter = document.getElementById('dash-filter-machine').value;
            const statusFilter = document.getElementById('dash-filter-status').value;
            const protocolFilter = document.getElementById('dash-filter-protocol').value.toLowerCase();

            const filtered = patientsData.filter(p => {
                if (machineFilter && !p.machines.includes(machineFilter)) return false;
                if (statusFilter && p.status !== statusFilter) return false;
                if (protocolFilter && !p.dtype.toLowerCase().includes(protocolFilter)) return false;
                return true;
            });

            updateDashboardMetrics(filtered);
            renderCharts(filtered);
        }

        function updateDashboardMetrics(data) {
            const total = data.length;
            const maleCount = data.filter(p => p.gender === 'ช').length;
            const femaleCount = data.filter(p => p.gender === 'ญ').length;
            const completed = data.filter(p => ['หาย', 'response', 'no response'].includes(p.status)).length;
            const recovered = data.filter(p => p.status === 'หาย').length;
            const otherIssues = data.filter(p => p.status === 'no response' || p.status === 'ไม่ครบ' || p.status === 'drop out').length;
            const totalSessions = data.reduce((sum, p) => sum + (parseInt(p.total_tx) || 0), 0);

            document.getElementById('stat-total').innerText = total;
            document.getElementById('stat-gender-breakdown').innerText = `ชาย ${maleCount} | หญิง ${femaleCount}`;
            document.getElementById('stat-completed').innerText = completed;
            document.getElementById('stat-completed-desc').innerText = total > 0 ? `คิดเป็น ${((completed/total)*100).toFixed(1)}% ของทั้งหมด` : `0%`;
            document.getElementById('stat-recovered').innerText = recovered;
            document.getElementById('stat-recovered-desc').innerText = completed > 0 ? `${((recovered/completed)*100).toFixed(1)}% ของผู้ป่วยที่ทำครบ` : `0%`;
            document.getElementById('stat-other-issues').innerText = otherIssues;
            document.getElementById('stat-sessions').innerText = totalSessions.toLocaleString();
        }

        function resetPatientFilters() {
            document.getElementById('filter-machine').value = '';
            document.getElementById('filter-status').value = '';
            document.getElementById('filter-protocol').value = '';
            document.getElementById('filter-search').value = '';
            applyPatientFilters();
        }

        function applyPatientFilters() {
            const machineFilter = document.getElementById('filter-machine').value;
            const statusFilter = document.getElementById('filter-status').value;
            const protocolFilter = document.getElementById('filter-protocol').value.toLowerCase();
            const searchKeyword = document.getElementById('filter-search').value.toLowerCase().trim();

            const filtered = patientsData.filter(p => {
                if (machineFilter && !p.machines.includes(machineFilter)) return false;
                if (statusFilter && p.status !== statusFilter) return false;
                if (protocolFilter && !p.dtype.toLowerCase().includes(protocolFilter)) return false;

                if (searchKeyword) {
                    const combined = `${p.no} ${p.dx} ${p.dtype} ${p.symptom_before} ${p.symptom_after} ${p.rights} ${p.gender} ${p.se}`.toLowerCase();
                    if (!combined.includes(searchKeyword)) return false;
                }

                return true;
            });

            renderTable(filtered);
        }

        function renderTable(data) {
            const tbody = document.getElementById('patient-table-body');
            document.getElementById('table-count').innerText = data.length;

            if (data.length === 0) {
                tbody.innerHTML = `<tr><td colspan="8" class="text-center py-12 opacity-60 font-medium">ไม่พบข้อมูลตามเงื่อนไขที่ค้นหา</td></tr>`;
                return;
            }

            tbody.innerHTML = data.map((p) => {
                let badgeClass = 'bg-slate-500/20 text-slate-300 border border-slate-500/30';
                let statusLabel = p.status;
                if (p.status === 'หาย') badgeClass = 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30';
                else if (p.status === 'response') { badgeClass = 'bg-teal-500/20 text-teal-300 border border-teal-500/30'; statusLabel = 'Response'; }
                else if (p.status === 'no response') badgeClass = 'bg-amber-500/20 text-amber-300 border border-amber-500/30';
                else if (p.status === 'drop out') badgeClass = 'bg-rose-500/20 text-rose-300 border border-rose-500/30';
                else if (p.status === 'ไม่ครบ') badgeClass = 'bg-purple-500/20 text-purple-300 border border-purple-500/30';
                else if (p.status === 'รอผล') badgeClass = 'bg-blue-500/20 text-blue-300 border border-blue-500/30';

                const machineBadges = p.machines.map(m => `<span class="bg-teal-500/20 text-teal-300 text-xs px-2.5 py-1 rounded-lg border border-teal-500/30 font-semibold">${m}</span>`).join(' ');
                const originalIndex = patientsData.findIndex(item => item === p);

                return `
                    <tr class="hover:bg-white/5 transition-colors cursor-pointer" onclick='openModal(${originalIndex})'>
                        <td class="py-4 px-5 font-bold">#${p.no}</td>
                        <td class="py-4 px-5"><span class="font-semibold">${p.gender === 'ช' ? 'ชาย' : 'หญิง'}</span>, ${p.age} ปี</td>
                        <td class="py-4 px-5">
                            <div class="font-bold">${p.dx}</div>
                            <div class="text-xs opacity-60 font-medium">สิทธิ: ${p.rights}</div>
                        </td>
                        <td class="py-4 px-5 text-xs opacity-70 max-w-xs truncate">${p.dtype || '-'}</td>
                        <td class="py-4 px-5"><div class="flex flex-wrap gap-1.5">${machineBadges}</div></td>
                        <td class="py-4 px-5 text-center font-extrabold text-teal-400 text-base">${p.total_tx}</td>
                        <td class="py-4 px-5 text-center"><span class="px-3 py-1.5 rounded-full text-xs font-bold ${badgeClass}">${statusLabel}</span></td>
                        <td class="py-4 px-5 text-center">
                            <button class="p-2 bg-white/10 hover:bg-teal-500/20 rounded-xl transition-colors border border-white/10" onclick='event.stopPropagation(); openModal(${originalIndex})'>
                                <i data-lucide="eye" class="w-4 h-4"></i>
                            </button>
                        </td>
                    </tr>
                `;
            }).join('');
            lucide.createIcons();
        }

        function renderCharts(data) {
            const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
            const textColor = isDark ? '#cbd5e1' : '#334155';
            const gridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';

            const statusCounts = {
                'หาย': data.filter(p => p.status === 'หาย').length,
                'Response': data.filter(p => p.status === 'response').length,
                'No Response': data.filter(p => p.status === 'no response').length,
                'Drop Out': data.filter(p => p.status === 'drop out').length,
                'ไม่ครบ': data.filter(p => p.status === 'ไม่ครบ').length,
                'รอผล': data.filter(p => p.status === 'รอผล').length
            };

            const statusCtx = document.getElementById('statusChart').getContext('2d');
            if (statusChartInstance) statusChartInstance.destroy();
            statusChartInstance = new Chart(statusCtx, {
                type: 'doughnut',
                data: {
                    labels: Object.keys(statusCounts),
                    datasets: [{
                        data: Object.values(statusCounts),
                        backgroundColor: ['#10b981', '#14b8a6', '#f59e0b', '#f43f5e', '#8b5cf6', '#3b82f6'],
                        borderWidth: 2,
                        borderColor: isDark ? '#1e293b' : '#ffffff',
                        hoverOffset: 6
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom', labels: { font: { family: 'Prompt', size: 11, weight: '500' }, color: textColor, boxWidth: 12, padding: 15 } } }
                }
            });

            const ageRanges = ['15-20 ปี', '21-30 ปี', '31-40 ปี', '41-50 ปี', '51-60 ปี', '>60 ปี'];
            const maleCounts = [0, 0, 0, 0, 0, 0];
            const femaleCounts = [0, 0, 0, 0, 0, 0];

            data.forEach(p => {
                let idx = 5;
                if (p.age >= 15 && p.age <= 20) idx = 0;
                else if (p.age >= 21 && p.age <= 30) idx = 1;
                else if (p.age >= 31 && p.age <= 40) idx = 2;
                else if (p.age >= 41 && p.age <= 50) idx = 3;
                else if (p.age >= 51 && p.age <= 60) idx = 4;

                if (p.gender === 'ช') maleCounts[idx]++;
                else femaleCounts[idx]++;
            });

            const ageGenderCtx = document.getElementById('ageGenderChart').getContext('2d');
            if (ageGenderChartInstance) ageGenderChartInstance.destroy();
            ageGenderChartInstance = new Chart(ageGenderCtx, {
                type: 'bar',
                data: {
                    labels: ageRanges,
                    datasets: [
                        { label: 'ชาย', data: maleCounts, backgroundColor: '#0ea5e9', borderRadius: 6, barThickness: 16 },
                        { label: 'หญิง', data: femaleCounts, backgroundColor: '#f43f5e', borderRadius: 6, barThickness: 16 }
                    ]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: { stacked: true, grid: { display: false }, ticks: { font: { family: 'Prompt', size: 11 }, color: textColor } },
                        y: { stacked: true, grid: { color: gridColor, borderDash: [4, 4] }, ticks: { stepSize: 2, font: { family: 'Prompt', size: 11 }, color: textColor } }
                    },
                    plugins: { legend: { position: 'bottom', labels: { font: { family: 'Prompt', size: 11, weight: '500' }, color: textColor, boxWidth: 12, padding: 15 } } }
                }
            });

            const machineCounts = { 'เครื่อง 1': 0, 'เครื่อง 2': 0, 'เครื่อง 3': 0, 'เครื่อง 4': 0 };
            data.forEach(p => {
                p.machines.forEach(m => { if (machineCounts[m] !== undefined) machineCounts[m]++; });
            });

            const machineCtx = document.getElementById('machineChart').getContext('2d');
            if (machineChartInstance) machineChartInstance.destroy();
            machineChartInstance = new Chart(machineCtx, {
                type: 'polarArea',
                data: {
                    labels: Object.keys(machineCounts),
                    datasets: [{
                        data: Object.values(machineCounts),
                        backgroundColor: ['rgba(20, 184, 166, 0.85)', 'rgba(14, 165, 233, 0.85)', 'rgba(139, 92, 246, 0.85)', 'rgba(244, 63, 94, 0.85)'],
                        borderWidth: 2,
                        borderColor: isDark ? '#1e293b' : '#ffffff'
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { position: 'bottom', labels: { font: { family: 'Prompt', size: 11, weight: '500' }, color: textColor, boxWidth: 12, padding: 15 } } }
                }
            });
        }

        function exportToExcel() {
            const excelData = patientsData.map((p) => ({
                "ลำดับเคส": p.no,
                "เพศ": p.gender === 'ช' ? 'ชาย' : 'หญิง',
                "อายุ (ปี)": p.age,
                "การวินิจฉัย (Dx)": p.dx,
                "สิทธิการรักษา": p.rights,
                "โปรโตคอล / รายละเอียด": p.dtype,
                "เครื่อง TMS": p.machines.join(', '),
                "จำนวนครั้งบำบัด": p.total_tx,
                "สถานะผลลัพธ์": p.status,
                "อาการก่อนทำ": p.symptom_before,
                "อาการหลังทำ / ผลลัพธ์": p.symptom_after,
                "ผลข้างเคียง (S/E)": p.se
            }));

            const worksheet = XLSX.utils.json_to_sheet(excelData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "TMS Patients Data");
            XLSX.writeFile(workbook, `TMS_Patients_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
        }

        function openModal(index) {
            const p = patientsData[index];
            document.getElementById('modal-no').innerText = `เคสที่ #${p.no}`;
            document.getElementById('modal-gender-age').innerText = `${p.gender === 'ช' ? 'ชาย' : 'หญิง'}, ${p.age} ปี`;
            document.getElementById('modal-dx').innerText = p.dx;
            document.getElementById('modal-rights').innerText = p.rights;
            document.getElementById('modal-sessions').innerText = `${p.total_tx} ครั้ง`;
            document.getElementById('modal-dtype').innerText = p.dtype || '-';
            document.getElementById('modal-sym-before').innerText = p.symptom_before || '-';
            document.getElementById('modal-sym-after').innerText = p.symptom_after || '-';
            document.getElementById('modal-se').innerText = p.se || 'ไม่มี S/E';

            document.getElementById('modal-edit-btn').setAttribute('onclick', `openEditModal(${index})`);
            document.getElementById('modal-delete-btn').setAttribute('onclick', `deletePatient(${index})`);
            document.getElementById('detail-modal').classList.remove('hidden');
        }

        function closeModal() { document.getElementById('detail-modal').classList.add('hidden'); }
        
        function openAddModal() {
            document.getElementById('modal-form-title').innerText = "ลงข้อมูลเคสผู้ป่วยใหม่";
            document.getElementById('modal-form-subtitle').innerText = "บันทึกข้อมูลเข้า Google Sheet กลาง";
            document.getElementById('form-submit-btn').innerText = "บันทึกข้อมูล";
            document.getElementById('form-edit-index').value = "-1";
            document.getElementById('add-patient-form').reset();
            document.getElementById('add-modal').classList.remove('hidden');
        }

        function openEditModal(index) {
            closeModal();
            const p = patientsData[index];
            document.getElementById('modal-form-title').innerText = `แก้ไขข้อมูลเคส #${p.no}`;
            document.getElementById('modal-form-subtitle').innerText = "อัปเดตข้อมูลรายละเอียดการรักษาด้วย TMS";
            document.getElementById('form-submit-btn').innerText = "บันทึกการแก้ไข";
            document.getElementById('form-edit-index').value = index;

            document.getElementById('form-no').value = p.no;
            document.getElementById('form-gender').value = p.gender;
            document.getElementById('form-age').value = p.age;
            document.getElementById('form-rights').value = p.rights;
            document.getElementById('form-dx').value = p.dx;
            document.getElementById('form-machine').value = p.machines[0] || 'เครื่อง 1';
            document.getElementById('form-sessions').value = p.total_tx;
            document.getElementById('form-status').value = p.status;
            document.getElementById('form-dtype').value = p.dtype !== '-' ? p.dtype : '';
            document.getElementById('form-sym-before').value = p.symptom_before !== '-' ? p.symptom_before : '';
            document.getElementById('form-sym-after').value = p.symptom_after !== '-' ? p.symptom_after : '';
            document.getElementById('form-se').value = p.se !== '-' ? p.se : '';

            document.getElementById('add-modal').classList.remove('hidden');
        }

        function closeAddModal() { document.getElementById('add-modal').classList.add('hidden'); }

        async function handleFormSubmit(event) {
            event.preventDefault();
            const editIndex = parseInt(document.getElementById('form-edit-index').value);

            const patientDataObj = {
                action: editIndex >= 0 ? "update" : "add",
                no: document.getElementById('form-no').value,
                gender: document.getElementById('form-gender').value,
                age: parseInt(document.getElementById('form-age').value),
                dx: document.getElementById('form-dx').value,
                rights: document.getElementById('form-rights').value,
                dtype: document.getElementById('form-dtype').value || '-',
                machines: document.getElementById('form-machine').value,
                total_tx: document.getElementById('form-sessions').value,
                status: document.getElementById('form-status').value,
                symptom_before: document.getElementById('form-sym-before').value || '-',
                symptom_after: document.getElementById('form-sym-after').value || '-',
                se: document.getElementById('form-se').value || 'ไม่มี S/E'
            };

            document.getElementById('form-submit-btn').innerText = "กำลังบันทึก...";
            document.getElementById('form-submit-btn').disabled = true;

            try {
                await fetch(WEB_APP_URL, {
                    method: "POST",
                    mode: "no-cors",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(patientDataObj)
                });

                alert('บันทึกข้อมูลลง Google Sheet เรียบร้อยแล้ว!');
                closeAddModal();
                document.getElementById('add-patient-form').reset();
                loadPatientsFromGoogleSheet();
            } catch (error) {
                console.error("Error saving data:", error);
                alert("เกิดข้อผิดพลาดในการบันทึกข้อมูล");
            } finally {
                document.getElementById('form-submit-btn').innerText = "บันทึกข้อมูล";
                document.getElementById('form-submit-btn').disabled = false;
            }
        }

        async function deletePatient(index) {
            const p = patientsData[index];
            if (confirm(`คุณต้องการลบเคส #${p.no} ออกจาก Google Sheet ใช่หรือไม่?`)) {
                const payload = {
                    action: "delete",
                    no: p.no
                };

                try {
                    await fetch(WEB_APP_URL, {
                        method: "POST",
                        mode: "no-cors",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify(payload)
                    });

                    alert('ลบข้อมูลเคสเรียบร้อยแล้ว!');
                    closeModal();
                    loadPatientsFromGoogleSheet();
                } catch (error) {
                    console.error("Error deleting data:", error);
                    alert("เกิดข้อผิดพลาดในการลบข้อมูล");
                }
            }
        }

        window.onload = loadPatientsFromGoogleSheet;
    </script>
</body>
</html>

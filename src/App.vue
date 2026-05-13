<script setup lang="ts">
import { ref, reactive, computed } from 'vue';
import { 
  FileText, 
  FileSpreadsheet, 
  Upload, 
  CheckCircle2, 
  AlertCircle, 
  RefreshCcw, 
  ArrowRightLeft,
  Search,
  FileSearch,
  History,
  Clock,
  Package,
  FolderOpen,
  FileArchive
} from 'lucide-vue-next';
import { Motion } from '@motionone/vue';
import * as mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
type FileType = 'audit_report' | 'shareholder_info' | 'net_assets' | 'profit' | 'balance' | 'trial_balance';

interface FileState {
  file: File | null;
  status: 'idle' | 'parsing' | 'ready' | 'error';
  data: any;
  error?: string;
}

interface BatchProduct {
  id: string;
  name: string;
  files: Record<FileType, FileState>;
  isReady: boolean;
  history: HistoryItem[];
}

interface HistoryItem {
  id: string;
  productName: string;
  compareTime: string;
  mainFile: string;
  excelFiles: string[];
  status: '正常' | '有差异';
  diffCount: number;
  details?: { field: string; message: string; severity: 'warning' | 'error' }[];
  versions?: { 
    id: string; 
    time: string; 
    mainFile: string; 
    excelFiles: string[];
    status: '正常' | '有差异';
    diffCount: number;
  }[];
}

const REQUIRED_EXCELS: { key: FileType; label: string }[] = [
  { key: 'net_assets', label: '净资产变动表' },
  { key: 'profit', label: '利润表' },
  { key: 'balance', label: '资产负债表' },
  { key: 'trial_balance', label: '余额表' },
  { key: 'shareholder_info', label: '持有人历史份额信息' },
];

const MOCK_HISTORY: HistoryItem[] = [
  {
    id: 'H001',
    productName: '某量化策略1号全周期进取型资产管理计划',
    compareTime: '2026-05-07 10:22',
    mainFile: '2025年度审计报告_最终版.docx',
    excelFiles: ['余额表.xlsx', '净资产变动表.xlsx', '利润表.xlsx'],
    status: '有差异',
    diffCount: 3,
    details: [
      { field: '银行存款', message: '审计报告值 3,115,498.20 与余额表不符', severity: 'error' },
      { field: '管理费', message: '计提比例计算结果与报告描述存在 0.01% 差异', severity: 'warning' },
      { field: '净资产', message: '期末余额勾稽关系不平', severity: 'error' }
    ],
    versions: [
      { id: 'V1', time: '2026-05-07 10:22', mainFile: '2025年度审计报告_最终版.docx', excelFiles: ['余额表.xlsx', '利润表.xlsx'], status: '有差异', diffCount: 3 },
      { id: 'V2', time: '2026-05-05 14:10', mainFile: '2025年度审计报告_初稿.docx', excelFiles: ['余额表.xlsx'], status: '有差异', diffCount: 5 }
    ]
  },
  {
    id: 'H002',
    productName: '睿智增长集合资产管理计划01期',
    compareTime: '2026-05-06 15:45',
    mainFile: '2025年度审计报告.docx',
    excelFiles: ['余额表.xlsx', '利润表.xlsx'],
    status: '正常',
    diffCount: 0,
    details: [],
    versions: [
      { id: 'V1', time: '2026-05-06 15:45', mainFile: '2025年度审计报告.docx', excelFiles: ['余额表.xlsx', '利润表.xlsx'], status: '正常', diffCount: 0 }
    ]
  }
];

const demoNetAssetsExcel = [
  ['项目', '本年实收资本', '本年其他综合收益', '本年未分配利润', '本年净资产合计'],
  ['一、上年年末余额', '34,251,624.09', '', '34,251,624.09', '68,503,248.18'],
  ['二、本年年初余额', '34,251,624.09', '', '34,251,624.09', '68,503,248.18'],
  ['三、本期增减变动额', '', '', '-17311759.23', '-17311759.23'],
  ['四、本期期末余额', '34,251,624.09', '', '16939864.86', '51191488.95']
];

const files = reactive<Record<FileType, FileState>>({
  audit_report: { file: null, status: 'idle', data: null },
  shareholder_info: { file: null, status: 'idle', data: null },
  net_assets: { file: null, status: 'idle', data: null },
  profit: { file: null, status: 'idle', data: null },
  balance: { file: null, status: 'idle', data: null },
  trial_balance: { file: null, status: 'idle', data: null },
});

const batchProducts = ref<BatchProduct[]>([]);
const uploadMode = ref<'batch' | 'manual'>('batch');
const globalShareholder = reactive<FileState>({ file: null, status: 'idle', data: null });
const isProcessing = ref(false);
const selectedProductId = ref<string | null>(null);
const previewContent = ref<{ type: 'word' | 'excel'; title: string; data: any } | null>(null);
const selectedItems = ref<string[]>([]);
const view = ref<'upload' | 'comparison' | 'history'>('upload');
const activeTab = ref<string>('net_assets');
const wordHtml = ref<string>('');
const versionHistoryItem = ref<HistoryItem | null>(null);

const currentProduct = computed(() => {
  return batchProducts.value.find(p => p.id === selectedProductId.value) || batchProducts.value[0];
});

const isAllFilesReady = computed(() => {
  return batchProducts.value.length > 0 || globalShareholder.status === 'ready';
});

const tableData = computed(() => {
  let excelData: any[][] = [[]];
  let wordData: any[][] = [[]];
  const currentFiles = currentProduct.value?.files || files;

  if (activeTab.value === 'net_assets') {
    excelData = currentFiles.net_assets.data as any[][] || demoNetAssetsExcel;
    wordData = currentFiles.audit_report.data as any[][] || demoNetAssetsExcel;
  } else if (activeTab.value === 'shareholder_info') {
    excelData = globalShareholder.data as any[][] || [
      ['内部分码', '持有人', '2025年12月31日', '2024年12月31日'],
      ['A01', '某1账户', '992,459,596.35', '992,459,596.35'],
      ['', '合计', '1,389,443,434.88', '1,389,443,434.88']
    ];
    wordData = [
      ['持有人', '2025年12月31日', '2024年12月31日'],
      ['某1账户', '992,459,596.35', '992,459,596.35'],
      ['合计', '1,389,443,434.88', '1,389,443,434.88']
    ];
  } else {
    const customData = currentFiles[activeTab.value as FileType]?.data;
    excelData = Array.isArray(customData) ? customData : [['暂无对照数据', '']];
    wordData = [['未提取', '']];
  }

  return { excel: excelData, word: wordData };
});

const sortedHistory = computed(() => {
  const sessionHistory = batchProducts.value.map(p => ({
    id: p.id,
    productName: p.name,
    compareTime: '刚刚',
    mainFile: p.files.audit_report.file?.name || '未上传报告',
    excelFiles: Object.keys(p.files).filter(k => k !== 'audit_report' && p.files[k as FileType].status === 'ready' && p.files[k as FileType].file !== null).map(k => p.files[k as FileType].file?.name || ''),
    status: (p.isReady ? '正常' : '有差异') as '正常' | '有差异',
    diffCount: 2,
    versions: [
      { id: 'V1', time: '刚刚', mainFile: p.files.audit_report.file?.name || '', excelFiles: [], status: '正常' as const, diffCount: 0 }
    ]
  }));
  return [...sessionHistory, ...MOCK_HISTORY].sort((a, b) => a.status === '有差异' ? -1 : 1);
});

const parseFileContent = async (type: FileType, file: File): Promise<any> => {
  const arrayBuffer = await file.arrayBuffer();
  if (type === 'audit_report') {
    const result = await mammoth.convertToHtml({ arrayBuffer });
    return result.value; 
  } else {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
  }
};

const handleManualFileUpload = async (type: FileType, file: File) => {
  try {
    const data = await parseFileContent(type, file);
    files[type] = { file, status: 'ready', data };
  } catch (err) {
    files[type] = { file, status: 'error', data: null, error: '解析失败' };
  }
};

const handleBatchUpload = async (e: Event) => {
  const target = e.target as HTMLInputElement;
  const uploadedFiles = target.files;
  if (!uploadedFiles || uploadedFiles.length === 0) return;

  isProcessing.value = true;
  const newProducts: Record<string, BatchProduct> = {};

  const processSingleFile = async (path: string, file: File) => {
    const parts = path.split('/');
    const productName = parts.length > 1 ? parts[parts.length - 2] : '未命名产品';
    const fileName = file.name.toLowerCase();

    if (!newProducts[productName]) {
      newProducts[productName] = {
        id: Math.random().toString(36).substr(2, 9),
        name: productName,
        files: {
          audit_report: { file: null, status: 'idle', data: null },
          shareholder_info: { file: null, status: 'idle', data: null },
          net_assets: { file: null, status: 'idle', data: null },
          profit: { file: null, status: 'idle', data: null },
          balance: { file: null, status: 'idle', data: null },
          trial_balance: { file: null, status: 'idle', data: null },
        },
        isReady: false,
        history: []
      };
    }

    let type: FileType | null = null;
    if (fileName.includes('审计报告') || fileName.endsWith('.docx')) type = 'audit_report';
    else if (fileName.includes('净资产') || fileName.includes('净值变动')) type = 'net_assets';
    else if (fileName.includes('利润')) type = 'profit';
    else if (fileName.includes('资产负债')) type = 'balance';
    else if (fileName.includes('余额表')) type = 'trial_balance';

    if (type) {
      try {
        const data = await parseFileContent(type, file);
        newProducts[productName].files[type] = { file, status: 'ready', data };
      } catch (err) {
        newProducts[productName].files[type] = { file, status: 'error', data: null };
      }
    }
  };

  for (let i = 0; i < uploadedFiles.length; i++) {
    const file = uploadedFiles[i];
    await processSingleFile((file as any).webkitRelativePath || file.name, file);
  }

  batchProducts.value = Object.values(newProducts).map(p => {
    const hasCritical = p.files.audit_report.status === 'ready' && 
                        (p.files.net_assets.status === 'ready' || p.files.balance.status === 'ready');
    return { ...p, isReady: hasCritical };
  });
  isProcessing.value = false;
};

const loadDemoData = () => {
  wordHtml.value = '<h3>演示报告</h3><p>银行存款合计：3,115,498.20 元</p>';
  files.audit_report = { file: new File([], "演示报表.docx"), status: 'ready', data: demoNetAssetsExcel };
  view.value = 'comparison';
};
</script>

<template>
  <div class="min-h-screen bg-slate-100 flex flex-col font-sans">
    <!-- Navbar -->
    <nav class="bg-slate-800 text-white px-6 py-3 flex items-center justify-between shadow-md">
      <div class="flex items-center gap-2">
        <div class="w-8 h-8 bg-yellow-400 rounded-full flex items-center justify-center">
          <Search class="w-5 h-5 text-slate-800" />
        </div>
        <span class="font-bold tracking-tight">审计数据对比工具 (Vue 重构版)</span>
      </div>
    </nav>

    <!-- Main -->
    <div class="flex-1 flex overflow-hidden">
      <!-- Sidebar -->
      <aside class="w-64 bg-white border-r border-slate-200 hidden md:flex flex-col">
        <div class="p-4 border-b border-slate-100 font-bold text-slate-400 text-xs uppercase tracking-widest">功能菜单</div>
        <nav class="p-2 space-y-1">
          <button @click="view = 'upload'" :class="cn('w-full flex items-center gap-3 px-4 py-3 text-sm font-bold rounded-lg transition-colors', (view === 'upload' || view === 'comparison') ? 'text-primary bg-red-50' : 'text-slate-600 hover:bg-slate-50')">
            <FileSearch class="w-5 h-5" />
            产品财务审计
          </button>
          <button @click="view = 'history'" :class="cn('w-full flex items-center gap-3 px-4 py-3 text-sm font-bold rounded-lg transition-colors', view === 'history' ? 'text-primary bg-red-50' : 'text-slate-600 hover:bg-slate-50')">
            <History class="w-5 h-5" />
            产品对比历史
          </button>
        </nav>
      </aside>

      <main class="flex-1 overflow-auto bg-slate-50 relative flex flex-col">
        <div v-if="view === 'upload' || view === 'comparison'" class="bg-white border-b border-slate-200 px-6 flex items-center gap-2">
          <button @click="view = 'upload'" :class="cn('px-6 py-3 text-sm font-bold border-b-2 transition-all', view === 'upload' ? 'border-primary text-primary' : 'border-transparent text-slate-400 hover:text-slate-600')">数据采集上传</button>
          <button 
            v-if="view === 'comparison' || isAllFilesReady"
            @click="view = 'comparison'" 
            :class="cn('px-6 py-3 text-sm font-bold border-b-2 transition-all flex items-center gap-2', view === 'comparison' ? 'border-primary text-primary' : 'border-transparent text-slate-400 hover:text-slate-600')"
          >
            审计比对报告
          </button>
        </div>

        <div class="flex-1 overflow-auto p-8">
          <!-- Upload View -->
          <div v-if="view === 'upload'" class="max-w-5xl mx-auto space-y-10">
            <header class="flex flex-col md:flex-row md:items-end justify-between gap-6 mb-12">
              <div>
                <h1 class="text-4xl font-extrabold text-slate-900 tracking-tight mb-2">审计报告核对工具</h1>
                <p class="text-slate-500 text-lg font-medium">支持批量自动化处理。</p>
              </div>
              <div class="flex bg-slate-200 p-1 rounded-2xl">
                <button @click="uploadMode = 'batch'" :class="cn('px-6 py-2 rounded-xl text-sm font-black transition-all', uploadMode === 'batch' ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-700')">批量上传</button>
                <button @click="uploadMode = 'manual'" :class="cn('px-6 py-2 rounded-xl text-sm font-black transition-all', uploadMode === 'manual' ? 'bg-white text-primary shadow-sm' : 'text-slate-500 hover:text-slate-700')">单产品上传</button>
              </div>
            </header>

            <div class="grid grid-cols-1 md:grid-cols-3 gap-8">
              <section v-if="uploadMode === 'manual'" class="md:col-span-2 bg-white p-8 rounded-3xl border border-slate-200 shadow-xl">
                 <h2 class="text-xl font-bold mb-6">单产品文件选取</h2>
                 <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div v-for="item in [{key:'audit_report',label:'审计报告(Word)'}, ...REQUIRED_EXCELS]" :key="item.key" :class="cn('relative border-2 border-dashed rounded-2xl p-4 transition-all', files[item.key as FileType].status === 'ready' ? 'bg-green-50 border-green-200' : 'bg-slate-50 border-slate-200')">
                       <input type="file" @change="(e: any) => handleManualFileUpload(item.key as FileType, e.target.files[0])" class="absolute inset-0 opacity-0 cursor-pointer" />
                       <div class="flex items-center gap-3">
                         <component :is="files[item.key as FileType].status === 'ready' ? CheckCircle2 : Upload" :class="cn('w-6 h-6', files[item.key as FileType].status === 'ready' ? 'text-green-500' : 'text-slate-400')" />
                         <div><div class="font-bold text-xs">{{ item.label }}</div><div class="text-[10px]">{{ files[item.key as FileType].file?.name || '点击上传' }}</div></div>
                       </div>
                    </div>
                 </div>
              </section>
              <section v-else class="md:col-span-2 bg-white p-8 rounded-3xl border border-slate-200 shadow-xl">
                 <h2 class="text-xl font-bold mb-6">批量目录上传</h2>
                 <div class="relative min-h-[200px] border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center bg-slate-50">
                    <input type="file" multiple webkitdirectory class="absolute inset-0 opacity-0 cursor-pointer" @change="handleBatchUpload" />
                    <RefreshCcw v-if="isProcessing" class="w-8 h-8 animate-spin" />
                    <div v-else-if="batchProducts.length > 0" class="w-full space-y-2">
                       <div v-for="p in batchProducts" :key="p.id" class="p-2 bg-white border rounded shadow-sm text-sm font-bold flex justify-between">
                          <span>{{ p.name }}</span>
                          <button @click="selectedProductId = p.id; view = 'comparison'" class="text-primary text-xs">查看</button>
                       </div>
                    </div>
                    <div v-else class="text-center"><Upload class="w-8 h-8 mx-auto mb-2 text-slate-300" /><span class="text-sm font-bold">拖拽文件夹至此</span></div>
                 </div>
              </section>

              <section class="bg-white p-8 rounded-3xl border border-slate-200 shadow-xl">
                <h2 class="text-xl font-bold mb-6">全局持有人信息</h2>
                <div class="border-2 border-dashed rounded-xl p-4 text-center bg-slate-50">
                  <Upload class="w-6 h-6 mx-auto mb-2 text-slate-300" />
                  <p class="text-xs font-bold">上传份额流水表</p>
                </div>
              </section>
            </div>

            <div class="text-center mt-8">
              <button @click="loadDemoData" class="py-4 px-12 bg-primary text-white rounded-xl font-bold text-xl hover:scale-105 transition-all shadow-lg">开始核对</button>
            </div>
          </div>

          <!-- Comparison View -->
          <div v-else-if="view === 'comparison'" class="h-full flex flex-col bg-white rounded-2xl shadow-xl border border-slate-200">
             <div class="p-4 border-b bg-slate-50 flex justify-between items-center">
                <div class="flex gap-4">
                   <div class="flex flex-col"><span class="text-[10px] uppercase font-bold text-slate-400">当前产品</span><span class="font-bold">{{ currentProduct?.name || '演示数据' }}</span></div>
                   <div class="flex gap-2 ml-4">
                      <button v-for="t in [{id:'net_assets',l:'净值变动表'},{id:'profit',l:'利润表'}]" :key="t.id" @click="activeTab = t.id" :class="cn('px-3 py-1 rounded text-xs font-bold', activeTab === t.id ? 'bg-primary text-white' : 'bg-slate-200')">{{ t.l }}</button>
                   </div>
                </div>
                <button @click="view = 'upload'" class="text-xs font-bold text-slate-400 hover:text-primary">重核</button>
             </div>
             <div class="flex-1 flex overflow-hidden p-4 gap-4">
                <div class="flex-1 border rounded p-4 overflow-auto"><h3 class="font-bold mb-2">审计数据</h3><div v-html="wordHtml" class="prose prose-sm"></div></div>
                <div class="flex-1 border rounded p-4 overflow-auto">
                   <h3 class="font-bold mb-2">报表数据</h3>
                   <table class="w-full text-xs text-left border-collapse">
                      <thead class="bg-slate-100"><tr><th v-for="h in tableData.excel[0]" :key="h" class="p-2 border">{{ h }}</th></tr></thead>
                      <tbody><tr v-for="(row, i) in tableData.excel.slice(1)" :key="i"><td v-for="(c, ci) in row" :key="ci" class="p-2 border">{{ c }}</td></tr></tbody>
                   </table>
                </div>
             </div>
          </div>

          <!-- History View -->
          <div v-else-if="view === 'history'" class="max-w-6xl mx-auto">
             <h1 class="text-2xl font-bold mb-6">比对历史</h1>
             <div class="bg-white rounded-xl border shadow-sm">
                <table class="w-full text-left">
                  <thead class="bg-slate-50 border-b"><tr><th class="p-4">产品名称</th><th class="p-4">状态</th><th class="p-4 text-right">时间</th></tr></thead>
                  <tbody>
                    <tr v-for="item in sortedHistory" :key="item.id" class="border-b last:border-0 hover:bg-slate-50 cursor-pointer" @click="selectedProductId = item.id; view = 'comparison'">
                      <td class="p-4 font-bold">{{ item.productName }}</td>
                      <td class="p-4"><span :class="cn('px-2 py-0.5 rounded text-[10px] font-bold', item.status === '正常' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700')">{{ item.status }}</span></td>
                      <td class="p-4 text-right text-xs text-slate-400">{{ item.compareTime }}</td>
                    </tr>
                  </tbody>
                </table>
             </div>
          </div>
        </div>
      </main>
    </div>
  </div>
</template>

<style>
@import "tailwindcss";
</style>

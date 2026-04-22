/**
 * AI 導入專案儀表板 — Google Sheet 後端
 *
 * 部署步驟：
 * 1) 開一份新的 Google Sheet（建議命名：AI 導入專案資料庫）
 * 2) 擴充功能 → Apps Script，把 Code.gs 內容整份貼進去
 * 3) 把下方 EDIT_TOKEN 改成只有你知道的隨機字串（之後儀表板寫回要比對這個）
 * 4) 回到 Sheet，重新整理，就會出現「AI 導入」選單 → 點「初次設定」
 *    → 會自動建立 5 個分頁（project / phases / tasks / metrics / risks）並灌入預設資料
 * 5) 部署 → 新增部署作業 → 類型：網頁應用程式
 *    執行身分：我（你的帳號）
 *    存取權：「任何人」
 *    → 複製 Web App URL
 * 6) 把 URL 貼到 index.html 最上方的 WEBAPP_URL 常數
 *
 * 日後若要改資料：
 *   A. 在儀表板上開編輯模式（輸入 EDIT_TOKEN）→ 點擊編輯即自動存回
 *   B. 直接在 Sheet 手改也可以，儀表板重新整理就會讀到
 */

const EDIT_TOKEN = '請改成你自己的隨機字串_abc123';

const SHEETS = {
  project: ['key', 'value'],
  phases: ['order', 'id', 'name', 'start', 'end', 'team', 'color', 'expanded', 'objective'],
  tasks: ['phase_id', 'task_id', 'name', 'status', 'note', 'due'],
  metrics: ['phase_id', 'order', 'name', 'target'],
  risks: ['phase_id', 'order', 'text']
};

// ============ Web App endpoints ============

function doGet() {
  try {
    return jsonOut(readAll());
  } catch (e) {
    return jsonOut({ error: String(e) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== EDIT_TOKEN) {
      return jsonOut({ ok: false, error: 'Invalid token' });
    }
    if (!body.data || !Array.isArray(body.data.phases)) {
      return jsonOut({ ok: false, error: 'Invalid data' });
    }
    writeAll(body.data);
    return jsonOut({ ok: true, savedAt: new Date().toISOString() });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ Read ============

function readAll() {
  const ss = SpreadsheetApp.getActive();

  const project = {};
  const projectRows = ss.getSheetByName('project').getDataRange().getValues().slice(1);
  projectRows.forEach(r => { if (r[0]) project[String(r[0])] = String(r[1] == null ? '' : r[1]); });

  const phaseRows = rowsAsObjects(ss, 'phases');
  const taskRows = rowsAsObjects(ss, 'tasks');
  const metricRows = rowsAsObjects(ss, 'metrics');
  const riskRows = rowsAsObjects(ss, 'risks');

  const phases = phaseRows
    .slice()
    .sort((a, b) => Number(a.order || 0) - Number(b.order || 0))
    .map(p => ({
      id: String(p.id),
      name: String(p.name),
      start: formatMonth(p.start),
      end: formatMonth(p.end),
      team: String(p.team),
      color: String(p.color),
      expanded: p.expanded === true || String(p.expanded).toLowerCase() === 'true',
      objective: String(p.objective || ''),
      tasks: taskRows
        .filter(t => String(t.phase_id) === String(p.id))
        .map(t => ({
          id: Number(t.task_id),
          name: String(t.name),
          status: String(t.status),
          note: String(t.note == null ? '' : t.note),
          due: formatDueDate(t.due)
        })),
      metrics: metricRows
        .filter(m => String(m.phase_id) === String(p.id))
        .sort((a, b) => Number(a.order || 0) - Number(b.order || 0))
        .map(m => ({ name: String(m.name), target: String(m.target) })),
      risks: riskRows
        .filter(r => String(r.phase_id) === String(p.id))
        .sort((a, b) => Number(a.order || 0) - Number(b.order || 0))
        .map(r => String(r.text))
    }));

  return { project, phases };
}

// 把 due date 統一成 "YYYY-MM-DD"（給前端 <input type="date"> 用）
function formatDueDate(val) {
  if (val == null || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    var mm = m[2].length === 1 ? '0' + m[2] : m[2];
    var dd = m[3].length === 1 ? '0' + m[3] : m[3];
    return m[1] + '-' + mm + '-' + dd;
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
  }
  return '';
}

// 把 start/end 統一成 "YYYY/MM"，避免 Sheet 把 2026/04 自動轉成 Date 之後變成
// "Wed Apr 01 2026 00:00:00 GMT+0800 (台北標準時間)" 這種長字串
function formatMonth(val) {
  if (val == null || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, 'Asia/Taipei', 'yyyy/MM');
  }
  var s = String(val).trim();
  var m = s.match(/(\d{4})[\/\-](\d{1,2})/);
  if (m) {
    var mm = m[2].length === 1 ? '0' + m[2] : m[2];
    return m[1] + '/' + mm;
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy/MM');
  }
  return s;
}

function rowsAsObjects(ss, name) {
  const values = ss.getSheetByName(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

// ============ Write ============

function writeAll(data) {
  const ss = SpreadsheetApp.getActive();
  writeProject(ss, data.project || {});
  writePhaseData(ss, data.phases || []);
}

function writeProject(ss, project) {
  const sheet = ss.getSheetByName('project');
  sheet.clearContents();
  const rows = [SHEETS.project];
  Object.keys(project).forEach(k => rows.push([k, project[k]]));
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
}

function writePhaseData(ss, phases) {
  const phaseSheet = ss.getSheetByName('phases');
  const taskSheet = ss.getSheetByName('tasks');
  const metricSheet = ss.getSheetByName('metrics');
  const riskSheet = ss.getSheetByName('risks');

  [phaseSheet, taskSheet, metricSheet, riskSheet].forEach(s => s.clearContents());

  const phaseRows = [SHEETS.phases];
  const taskRows = [SHEETS.tasks];
  const metricRows = [SHEETS.metrics];
  const riskRows = [SHEETS.risks];

  phases.forEach((p, pi) => {
    phaseRows.push([
      pi + 1,
      p.id,
      p.name,
      "'" + formatMonth(p.start),
      "'" + formatMonth(p.end),
      p.team,
      p.color,
      p.expanded === true,
      p.objective || ''
    ]);
    (p.tasks || []).forEach(t => {
      taskRows.push([p.id, t.id, t.name, t.status, t.note || '', t.due ? "'" + formatDueDate(t.due) : '']);
    });
    (p.metrics || []).forEach((m, mi) => {
      metricRows.push([p.id, mi + 1, m.name, m.target]);
    });
    (p.risks || []).forEach((r, ri) => {
      riskRows.push([p.id, ri + 1, r]);
    });
  });

  writeMatrix(phaseSheet, phaseRows);
  writeMatrix(taskSheet, taskRows);
  writeMatrix(metricSheet, metricRows);
  writeMatrix(riskSheet, riskRows);
}

function writeMatrix(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
}

// ============ First-time setup ============

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
  // 刪掉預設的 Sheet1 / 工作表1（如果還是空的）
  ['Sheet1', '工作表1'].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s && s.getLastRow() === 0) ss.deleteSheet(s);
  });
  writeAll(DEFAULT_DATA);
  try {
    SpreadsheetApp.getUi().alert('✅ 完成！\n\n下一步：部署 → 新增部署作業 → 網頁應用程式（執行身分：我，存取權：任何人），然後把 URL 貼進 index.html 的 WEBAPP_URL。');
  } catch (e) { /* ignore if not in UI context */ }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI 導入')
    .addItem('初次設定（建立分頁 + 預設資料）', 'setup')
    .addSeparator()
    .addItem('重新填入預設資料（覆蓋現有）', 'resetToDefault')
    .addToUi();
}

function resetToDefault() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert('確定要覆蓋？', '所有 Sheet 現有資料會被清空並改回預設值。', ui.ButtonSet.YES_NO);
  if (res === ui.Button.YES) {
    writeAll(DEFAULT_DATA);
    ui.alert('已還原為預設資料。');
  }
}

// ============ Default seed data ============

const DEFAULT_DATA = {
  project: {
    name: 'AI 導入專案',
    owner: 'Penny',
    timeline: '2026/04 - 2026/12（9 個月）',
    mode: '自建為主',
    vision: '分七階段循序導入 AI，以自動化取代重複勞動，讓團隊專注在需要人判斷的工作上。'
  },
  phases: [
    {
      id: 'cs', name: '階段 1：客服',
      start: '2026/04', end: '2026/05',
      team: '0 人（AI 全自動）',
      color: '#3b82f6',
      expanded: true,
      objective: '客服離職後由 AI 完全接手，不補人力。驗證 AI 基礎建設（RAG / Query Log / API 整合）能穩定運作。',
      metrics: [
        { name: '答對率', target: '≥ 90%' },
        { name: 'Fallback rate（無法回答率）', target: '< 10%' },
        { name: '平均首次回覆時間', target: '< 30 秒' },
        { name: '客訴數', target: '不超過離職前基準' }
      ],
      tasks: [
        { id: 1, name: '情境擬定及訪談', status: 'done', note: '' },
        { id: 2, name: 'LINE 客服自動回覆串接', status: 'done', note: '' },
        { id: 3, name: 'RAG 知識庫建置', status: 'done', note: '' },
        { id: 4, name: '電商 MCP & 物流 API 整合', status: 'in-progress', note: '本週完成' },
        { id: 5, name: 'Query Log 系統', status: 'todo', note: '下週完成' },
        { id: 6, name: '答案品質評估機制（每週抽查 + fallback 監控）', status: 'todo', note: '' },
        { id: 7, name: 'Escalation 流程：AI 處理不了時通知你介入', status: 'todo', note: '' },
        { id: 8, name: '知識庫持續更新 SOP（商品異動、新 FAQ）', status: 'todo', note: '防止 AI 劣化' }
      ],
      risks: [
        '知識庫未跟上商品/政策異動 → AI 給錯答案',
        'Edge case 沒有 escalation 會惡化客訴',
        'AI 品質下降無人察覺（沒有評估機制的話）'
      ]
    },
    {
      id: 'channel', name: '階段 2：通路（Line 禮物 + 團媽）',
      start: '2026/05', end: '2026/06',
      team: '3 人',
      color: '#8b5cf6',
      expanded: false,
      objective: '減少訂單處理、客戶訊息、對帳等重複性工作。3 人仍在，但釋出時間專注於選品與客戶關係。',
      metrics: [
        { name: '訂單處理時間', target: '減少 ≥ 60%' },
        { name: '人工對帳工時', target: '減少 ≥ 80%' },
        { name: '客戶訊息首次回覆時間', target: '< 5 分鐘' },
        { name: '月營收（觀察指標）', target: '不下滑，期望提升' }
      ],
      tasks: [
        { id: 1, name: '通路現況訪談：訪談 3 位同事找出前 10 大重複任務', status: 'todo', note: '' },
        { id: 2, name: '流程 SOP 梳理與優先級排序', status: 'todo', note: '' },
        { id: 3, name: '訂單資料結構與系統整合規劃', status: 'todo', note: '' },
        { id: 4, name: 'Line Gift 訂單自動擷取與入庫', status: 'todo', note: '' },
        { id: 5, name: '團媽訊息範本 + AI 智能推薦商品', status: 'todo', note: '' },
        { id: 6, name: '團媽 CRM（客戶分層、歷史訂單）', status: 'todo', note: '' },
        { id: 7, name: '自動化對帳機制', status: 'todo', note: '' },
        { id: 8, name: '2 週試運轉 → 成效量測 → 正式上線', status: 'todo', note: '' }
      ],
      risks: [
        'Line Gift 平台 API 限制可能卡住某些自動化',
        '團媽個性化需求難以完全標準化（保留人工介面）',
        '同事擔心「被取代」的抗拒心理（需溝通）'
      ]
    },
    {
      id: 'purchase', name: '階段 3：採購',
      start: '2026/07', end: '2026/07',
      team: '1 人（管理部內）',
      color: '#10b981',
      expanded: false,
      objective: '自動化比價、進貨單、庫存預測，讓採購專注於供應商關係與策略議價。',
      metrics: [
        { name: '進貨單產生工時', target: '減少 ≥ 70%' },
        { name: '缺貨次數 / 月', target: '減少 ≥ 50%' },
        { name: '比價覆蓋率', target: '100% 主要品項' },
        { name: '庫存週轉天數', target: '優化 ≥ 15%' }
      ],
      tasks: [
        { id: 1, name: '採購流程訪談與 SOP 梳理', status: 'todo', note: '' },
        { id: 2, name: '供應商資料庫建立（聯絡/條件/歷史價格）', status: 'todo', note: '' },
        { id: 3, name: 'ERP / 庫存系統串接（取得即時庫存）', status: 'todo', note: '' },
        { id: 4, name: '自動比價機制（至少 2 家供應商）', status: 'todo', note: '' },
        { id: 5, name: '庫存預測模型（銷售歷史 + 季節性）', status: 'todo', note: '需 ≥ 6 個月歷史資料' },
        { id: 6, name: 'PO 自動生成 + 人工審核介面', status: 'todo', note: '' },
        { id: 7, name: '成效量測與迭代', status: 'todo', note: '' }
      ],
      risks: [
        '供應商資料不完整（要補齊要花時間）',
        '庫存預測初期準確率低，需持續調校',
        'ERP 系統若封閉可能無法串接'
      ]
    },
    {
      id: 'accounting', name: '階段 4：會計',
      start: '2026/08', end: '2026/08',
      team: '1 人（管理部內）',
      color: '#f59e0b',
      expanded: false,
      objective: '自動發票辨識、分類記帳、對帳、月結報表。會計從「做帳」轉向「審核與分析」。',
      metrics: [
        { name: '發票處理工時', target: '減少 ≥ 80%' },
        { name: '記帳準確率', target: '≥ 99%' },
        { name: '月結完成時間', target: '縮短 ≥ 50%' },
        { name: '異常交易捕捉率', target: '≥ 90%' }
      ],
      tasks: [
        { id: 1, name: '會計流程訪談與會計科目梳理', status: 'todo', note: '' },
        { id: 2, name: '發票 OCR + AI 分類系統', status: 'todo', note: '二聯/三聯/電子發票' },
        { id: 3, name: '記帳規則引擎（自動對應科目）', status: 'todo', note: '' },
        { id: 4, name: '銀行對帳自動化（串接銀行 API / CSV 匯入）', status: 'todo', note: '' },
        { id: 5, name: '月結報表範本（損益、資產負債）', status: 'todo', note: '' },
        { id: 6, name: '異常交易偵測（金額異常、重複入帳）', status: 'todo', note: '' },
        { id: 7, name: '稅務申報資料自動產生', status: 'todo', note: '' }
      ],
      risks: [
        '發票格式多樣，OCR 初期準確率約 85-90%',
        '稅務法規變動需持續維護',
        '金額敏感，錯誤容忍度低，需審核流程'
      ]
    },
    {
      id: 'hr', name: '階段 5：HR + 全專案回顧',
      start: '2026/09', end: '2026/10',
      team: '1 人（管理部內）',
      color: '#ec4899',
      expanded: false,
      objective: 'HR 自動化著重員工問答、教育訓練、表單流程。最後一個月做全專案 ROI 回顧。',
      metrics: [
        { name: 'HR Q&A 自助解決率', target: '≥ 70%' },
        { name: '訓練素材覆蓋率', target: '100% 主要職務' },
        { name: '表單流程工時', target: '減少 ≥ 60%' },
        { name: '全專案節省工時 / 月', target: '明確量化' }
      ],
      tasks: [
        { id: 1, name: '員工手冊 / 規章知識庫建置', status: 'todo', note: '' },
        { id: 2, name: '員工 FAQ 機器人（薪資、假勤、福利）', status: 'todo', note: '' },
        { id: 3, name: '教育訓練素材庫 + AI 測驗', status: 'todo', note: '' },
        { id: 4, name: '請假 / 費用申請自動化流程', status: 'todo', note: '' },
        { id: 5, name: '績效與考核流程數位化', status: 'todo', note: '' },
        { id: 6, name: '全專案 ROI 回顧：節省工時 / 投入成本 / 下一階段規劃', status: 'todo', note: '10 月中旬' },
        { id: 7, name: '知識庫年度維護 SOP 建立', status: 'todo', note: '' }
      ],
      risks: [
        '員工隱私與資料權限分級',
        '主觀判斷類事務（考核、勞資）難以完全自動化',
        'HR 事務量小，ROI 可能不如前面階段'
      ]
    },
    {
      id: 'production', name: '階段 6：生產部',
      start: '2026/11', end: '2026/11',
      team: '待定',
      color: '#06b6d4',
      expanded: false,
      objective: '（待填寫）生產部 AI 導入目標：自動化排程、品管記錄、產能預測，讓生產端專注於品質與改善。',
      metrics: [
        { name: '（待填寫）指標 1', target: '—' },
        { name: '（待填寫）指標 2', target: '—' },
        { name: '（待填寫）指標 3', target: '—' }
      ],
      tasks: [
        { id: 1, name: '生產部流程訪談與 SOP 梳理', status: 'todo', note: '' },
        { id: 2, name: '生產排程 / 產能預測模型', status: 'todo', note: '' },
        { id: 3, name: '品管記錄數位化與異常偵測', status: 'todo', note: '' },
        { id: 4, name: '成效量測與迭代', status: 'todo', note: '' }
      ],
      risks: [
        '（待填寫）風險項目'
      ]
    },
    {
      id: 'retail', name: '階段 7：直營部',
      start: '2026/12', end: '2026/12',
      team: '待定',
      color: '#84cc16',
      expanded: false,
      objective: '（待填寫）直營部 AI 導入目標：門市營運、排班、顧客關係、銷售分析自動化，讓店長專注於現場服務與團隊帶領。',
      metrics: [
        { name: '（待填寫）指標 1', target: '—' },
        { name: '（待填寫）指標 2', target: '—' },
        { name: '（待填寫）指標 3', target: '—' }
      ],
      tasks: [
        { id: 1, name: '直營部流程訪談與 SOP 梳理', status: 'todo', note: '' },
        { id: 2, name: '門市銷售與庫存分析儀表板', status: 'todo', note: '' },
        { id: 3, name: '排班 / 出勤自動化', status: 'todo', note: '' },
        { id: 4, name: '顧客 CRM 與回訪提醒', status: 'todo', note: '' },
        { id: 5, name: '成效量測與迭代', status: 'todo', note: '' }
      ],
      risks: [
        '（待填寫）風險項目'
      ]
    }
  ]
};

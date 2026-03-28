/**
* exportTrialPlBalancesMonthly_
*
* 目的：
* - trial_pl.balances を「月=列」の横持ちTBとして出力する
*
* 出力シート：raw_trial_pl_balances_pivot
*
* 使い方：
* 1) 単独実行（おすすめ）
* exportTrialPlBalancesMonthly_()
* → config の company_id/from_date/to_date を使って月ループ→TB作成
*
* 2) 月1回分として呼ぶ（後方互換）
* exportTrialPlBalancesMonthly_({ ... })
*
* 行キー（固定列）：
* - line_type (line/total)
* - account_category_name
* - account_group_name
* - account_item_name
* - hierarchy_level
*
* 月列（動的に追加）：
* - {YYYY-MM}_debit
* - {YYYY-MM}_credit
* - {YYYY-MM}_pl (= credit - debit)
*/

// =============================
// 公開関数：単独実行もOK
// =============================
function exportTrialPlBalancesMonthly_(options) {
// 引数なしなら「単独実行モード」
if (options === undefined) {
return exportTrialPlBalancesMonthly__runFromConfig_();
}

// 引数ありなら「1回分の追記モード（後方互換）」
return exportTrialPlBalancesMonthly__appendOneMonth_(options);
}

// =============================
// 単独実行：configから期間を取って月ループ
// =============================
function exportTrialPlBalancesMonthly__runFromConfig_() {
const ss = SpreadsheetApp.getActive();

const companyId = getConfig_('company_id');
if (!companyId) throw new Error('config に company_id がありません');

const fromDateStr = String(getConfig_('from_date') || '').trim();
const toDateStr = String(getConfig_('to_date') || '').trim();
if (!fromDateStr || !toDateStr) throw new Error('config に from_date / to_date がありません');

const from = parseYmd_(fromDateStr);
const to = parseYmd_(toDateStr);
if (from > to) throw new Error('from_date が to_date より後です');

const periodLabel = `${fromDateStr}..${toDateStr}`;

let resetDone = false;

let cur = new Date(from.getFullYear(), from.getMonth(), 1);

while (cur <= to) {
const monthStart = new Date(cur.getFullYear(), cur.getMonth(), 1);
const monthEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);

const start = (monthStart < from) ? from : monthStart;
const end = (monthEnd > to) ? to : monthEnd;

const startStr = formatYmd_(start);
const endStr = formatYmd_(end);
const monthKey = Utilities.formatDate(monthStart, 'Asia/Tokyo', 'yyyy-MM');

const res = apiGet_('/api/1/reports/trial_pl', {
company_id: companyId,
start_date: startStr,
end_date: endStr,
});

const balancesRaw = (res.trial_pl && Array.isArray(res.trial_pl.balances)) ? res.trial_pl.balances : [];

exportTrialPlBalancesMonthly__appendOneMonth_({
periodLabel,
companyId,
fromDateStr,
toDateStr,
monthKey,
balancesRaw,
reset: !resetDone,
});

resetDone = true;
log_('TB_BUILD', `month=${monthKey} range=${startStr}..${endStr} rows=${balancesRaw.length}`);

cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
}

// 最後にTBシートを表示（好み）
const tbSh = ss.getSheetByName('raw_trial_pl_balances_pivot');
if (tbSh) {
ss.setActiveSheet(tbSh);
tbSh.activate();
}

log_('TB_BUILD', `done period=${periodLabel}`);
return true;
}

// =============================
// 1ヶ月分を追記（横持ちTB）
// =============================
function exportTrialPlBalancesMonthly__appendOneMonth_(options) {
const {
periodLabel,
companyId,
fromDateStr,
toDateStr,
monthKey,
balancesRaw,
reset,
} = options;

const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('raw_trial_pl_balances_pivot') || ss.insertSheet('raw_trial_pl_balances_pivot');

// 固定列（キー）
const KEY_COLS = ['line_type','account_category_name','account_group_name','account_item_name','hierarchy_level'];

// reset時：初期化＋メタ＋固定ヘッダ
if (reset) {
sh.clearContents();

sh.getRange(1, 1, 1, 2).setValues([['meta_key','meta_value']]);
sh.getRange(2, 1, 4, 2).setValues([
['period', periodLabel],
['company_id', companyId],
['from_date', fromDateStr],
['to_date', toDateStr],
]);

sh.getRange(7, 1, 1, KEY_COLS.length).setValues([KEY_COLS]);
sh.setFrozenRows(7);
}

// ヘッダ取得（ヘッダ行は7固定）
const headerRow = 7;
const header = sh.getRange(headerRow, 1, 1, sh.getLastColumn()).getValues()[0].map(v => String(v || ''));

// 月列（3本）を追加
const monthCols = ensureMonthColumns_(sh, headerRow, header, monthKey);

// 今月分をキー→金額にまとめる
const monthMap = new Map();
const balances = normalizeBalancesCarryForward_(balancesRaw || []);

for (const b of balances) {
const cat = b.account_category_name || '';
const grp = b.account_group_name || '';
const item = b.account_item_name || '';
const lvl = (b.hierarchy_level === undefined || b.hierarchy_level === null) ? '' : String(b.hierarchy_level);

const isTotal = (b.total_line === true || b.total_line === 'true');
const lineType = isTotal ? 'total' : 'line';

const keyStr = [lineType, cat, grp, item, lvl].join('||');

const debit = Number(b.debit_amount || 0);
const credit = Number(b.credit_amount || 0);

if (!monthMap.has(keyStr)) {
monthMap.set(keyStr, { lineType, cat, grp, item, lvl, debit: 0, credit: 0 });
}
const r = monthMap.get(keyStr);
r.debit += debit;
r.credit += credit;
}

// 既存キー行を読み込み → key→rowIndex
const dataStartRow = headerRow + 1;
const lastRow = sh.getLastRow();

const keyToRow = new Map();
if (lastRow >= dataStartRow) {
const existing = sh.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, KEY_COLS.length).getValues();
for (let i = 0; i < existing.length; i++) {
const rr = existing[i].map(v => String(v || ''));
const k = rr.join('||');
if (k.replace(/\|/g, '').trim() === '') continue;
keyToRow.set(k, dataStartRow + i);
}
}

// 未存在キーは末尾に追加
const toInsert = [];
for (const r of monthMap.values()) {
const keyLine = [r.lineType, r.cat, r.grp, r.item, r.lvl];
const k = keyLine.join('||');
if (!keyToRow.has(k)) {
toInsert.push(keyLine);
}
}

if (toInsert.length) {
const insertRow = sh.getLastRow() + 1;
sh.getRange(insertRow, 1, toInsert.length, KEY_COLS.length).setValues(toInsert);
for (let i = 0; i < toInsert.length; i++) {
keyToRow.set(toInsert[i].join('||'), insertRow + i);
}
}

// 今月列に値を書き込む（debit/credit/pl）を一括
const targets = [];
for (const r of monthMap.values()) {
const k = [r.lineType, r.cat, r.grp, r.item, r.lvl].join('||');
const rowNum = keyToRow.get(k);
if (rowNum) targets.push({ rowNum, r });
}
targets.sort((a, b) => a.rowNum - b.rowNum);

if (targets.length) {
const firstRow = targets[0].rowNum;
const lastRow2 = targets[targets.length - 1].rowNum;
const height = lastRow2 - firstRow + 1;

const mat = Array.from({ length: height }, () => ['', '', '']);
for (const t of targets) {
const idx = t.rowNum - firstRow;
const debit = t.r.debit;
const credit = t.r.credit;
const pl = credit - debit;
mat[idx] = [debit, credit, pl];
}

sh.getRange(firstRow, monthCols.debitCol, height, 3).setValues(mat);
sh.getRange(firstRow, monthCols.debitCol, height, 3).setNumberFormat('#,##0');
}

return true;
}

// =============================
// helper：月列を確保
// =============================
function ensureMonthColumns_(sheet, headerRow, header, monthKey) {
const want = [`${monthKey}_debit`, `${monthKey}_credit`, `${monthKey}_pl`];
let lastCol = header.length;

const idxs = {};
for (const colName of want) {
const i = header.indexOf(colName);
if (i >= 0) {
idxs[colName] = i + 1; // 1-based
} else {
lastCol += 1;
sheet.getRange(headerRow, lastCol).setValue(colName);
header.push(colName);
idxs[colName] = lastCol;
}
}

return {
debitCol: idxs[want[0]],
creditCol: idxs[want[1]],
plCol: idxs[want[2]],
};
}

// =============================
// helper：carry-forward
// =============================
function normalizeBalancesCarryForward_(balances) {
let lastCat = '';
let lastGrp = '';
const out = [];

for (const b of (balances || [])) {
const x = Object.assign({}, b);

if (x.account_category_name) lastCat = x.account_category_name;
else x.account_category_name = lastCat;

if (x.account_group_name) lastGrp = x.account_group_name;
else x.account_group_name = lastGrp;

out.push(x);
}
return out;
}
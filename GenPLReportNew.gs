/**
* buildSingleMonthPLFromTrialPL（TB出力なし・高速版）
*
* 目的：
* - config の from_date〜to_date の期間で trial_pl を取得し、PL表示用に整形して出力
* - 期間が複数月にまたがる場合は「月ごとに trial_pl を取得」して合算
* - balances 内でカテゴリ/グループが空の行が混ざっても carry-forward で正規化
* - 重複lineを落とす（dropDeeperDuplicateLines_）
* - 出力後に account_item_id を付与（enrichPLWithAccountItemId）
*
* 出力シート：report_pl_single_month_api
*/
function buildSingleMonthPLFromTrialPL() {
const ss = SpreadsheetApp.getActive();

// （任意）開始時のアクティブシートを控える
const originalSheet = ss.getActiveSheet();

const companyId = getConfig_('company_id');
if (!companyId) throw new Error('config に company_id がありません');

const fromDateStr = String(getConfig_('from_date') || '').trim();
const toDateStr = String(getConfig_('to_date') || '').trim();
if (!fromDateStr || !toDateStr) throw new Error('config に from_date / to_date がありません');

const from = parseYmd_(fromDateStr);
const to = parseYmd_(toDateStr);
if (from > to) throw new Error('from_date が to_date より後です');

const periodLabel = `${fromDateStr}..${toDateStr}`;

// PL出力先
const plSh = ss.getSheetByName('report_pl_single_month_api') || ss.insertSheet('report_pl_single_month_api');

try {
// PLシート初期化（※ activate しない）
plSh.clearContents();
plSh.appendRow([
'period',
'line_type',
'account_category_name',
'account_group_name',
'account_item_name',
'hierarchy_level',
'debit_amount',
'credit_amount',
'pl_amount',
]);

const agg = new Map();
let orderCounter = 0;

// 月次ループ（単月でもこのループで統一）
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
const balances = normalizeBalancesCarryForward_(balancesRaw);

// 合算
for (const b of balances) {
const cat = b.account_category_name || '';
const grp = b.account_group_name || '';
const item = b.account_item_name || '';
const lvl = (b.hierarchy_level === undefined || b.hierarchy_level === null) ? '' : String(b.hierarchy_level);

const isTotal = (b.total_line === true || b.total_line === 'true');
const lineType = isTotal ? 'total' : 'line';

const k = [lineType, cat, grp, item, lvl].join('||');

const debit = Number(b.debit_amount || 0);
const credit = Number(b.credit_amount || 0);

if (!agg.has(k)) {
agg.set(k, {
orderIndex: orderCounter++,
lineType,
cat,
grp,
item,
lvl,
debit: 0,
credit: 0,
});
}

const r = agg.get(k);
r.debit += debit;
r.credit += credit;

// 保険：後から非空の名称が来たら埋める
if (!r.cat && cat) r.cat = cat;
if (!r.grp && grp) r.grp = grp;
if (!r.item && item) r.item = item;
}

log_('PL_PERIOD_TRIAL_PL', `month=${monthKey} range=${startStr}..${endStr} rows=${balancesRaw.length}`);
cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
}

// 重複除去
const deduped = dropDeeperDuplicateLines_(Array.from(agg.values()));

// 出力
const rows = deduped
.sort((a, b) => a.orderIndex - b.orderIndex)
.map(r => {
const plAmount = r.credit - r.debit;
return [
periodLabel,
r.lineType,
r.cat,
r.grp,
r.item,
r.lvl,
r.debit,
r.credit,
plAmount,
];
});

if (rows.length) {
plSh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
plSh.getRange(2, 7, rows.length, 3).setNumberFormat('#,##0');
}

// account_item_id 付与
try {
const updated = enrichPLWithAccountItemId();
log_('PL_ENRICH', `account_item_id filled rows=${updated}`);
} catch (e) {
log_('PL_ENRICH_ERR', String(e));
}

log_('PL_PERIOD_DONE', `period=${periodLabel} rows=${rows.length}`);
return rows.length;

} finally {
// ★ 最後に必ずPLをアクティブにする
ss.setActiveSheet(plSh);
plSh.activate();

// 失敗時に元シートへ戻したいならこちら（好み）
// ss.setActiveSheet(originalSheet);
// originalSheet.activate();
}
}

// function parseYmd_(ymd) {
// const m = String(ymd).trim().match(/^(\d{4})-(\d{2})-(\d{2})$/);
// if (!m) throw new Error(`日付形式が不正です: ${ymd}（YYYY-MM-DD）`);
// return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
// }
/**
* parseYmd_
* - Date / 'YYYY-MM-DD' / ISO文字列(例 2012-12-31T15:00:00.000Z) を受け取り
* - Asia/Tokyo 기준の「日付」(00:00) に正規化した Date を返す
*/
function parseYmd_(v) {
// 1) Dateなら、その日付だけに丸める
if (v instanceof Date) {
return new Date(v.getFullYear(), v.getMonth(), v.getDate());
}

const s = String(v || '').trim();
if (!s) throw new Error('日付が空です');

// 2) まず 'YYYY-MM-DD' を許可
let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
if (m) {
return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// 3) ISO文字列 'YYYY-MM-DDTHH:mm:ss.sssZ' 等を許可
// → Dateにして、JSTで見た日付に丸める
const d = new Date(s);
if (!isNaN(d.getTime())) {
// JSTの日付として切り出す（Z→JSTのズレ吸収）
const y = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy');
const mo = Utilities.formatDate(d, 'Asia/Tokyo', 'MM');
const da = Utilities.formatDate(d, 'Asia/Tokyo', 'dd');
return new Date(Number(y), Number(mo) - 1, Number(da));
}

throw new Error(`日付形式が不正です: ${s}（YYYY-MM-DD or ISO）`);
}

function formatYmd_(d) {
return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function dropDeeperDuplicateLines_(items) {
// 同じ (lineType, category, group, item) が複数あるとき
// 先に出た方（orderIndexが小さい方）だけ残す
const best = new Map();

for (const r of items) {
const baseKey = [r.lineType, r.cat, r.grp, r.item].join('||');

if (!best.has(baseKey)) {
best.set(baseKey, r);
continue;
}

const cur = best.get(baseKey);

// orderIndex が小さい方を残す
if ((r.orderIndex ?? 999999999) < (cur.orderIndex ?? 999999999)) {
best.set(baseKey, r);
}
}

return Array.from(best.values());
}
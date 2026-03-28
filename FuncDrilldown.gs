/**
* PLシート（report_pl_single_month_api）に account_item_id 列を追加し、
* master_account_items を参照して account_item_name → account_item_id を埋める。
*
* - line_type が "line" の行だけ対象（totalは空のまま）
* - 既に列があれば再利用（上書き更新）
*/
function enrichPLWithAccountItemId() {
const ss = SpreadsheetApp.getActive();
const pl = ss.getSheetByName('report_pl_single_month_api');
const master = ss.getSheetByName('master_account_items');
if (!pl) throw new Error('report_pl_single_month_api シートが見つかりません');
if (!master) throw new Error('master_account_items シートが見つかりません');

// masterから name→id の辞書を作る
const m = master.getDataRange().getValues();
if (m.length < 2) throw new Error('master_account_items にデータがありません');

const mh = m[0].map(String);
const idCol = findHeaderIndex_(mh, ['account_item_id', 'id']);
const nameCol = findHeaderIndex_(mh, ['account_item_name', 'name']);
if (idCol < 0 || nameCol < 0) {
throw new Error('master_account_items のヘッダに id/name が見つかりません（account_item_id or id / account_item_name or name）');
}

const nameToId = new Map();
for (let r = 1; r < m.length; r++) {
const id = m[r][idCol];
const name = String(m[r][nameCol] || '').trim();
if (!name) continue;
nameToId.set(name, id);
}

// PLシートのヘッダ
const pr = pl.getDataRange().getValues();
if (pr.length < 2) throw new Error('report_pl_single_month_api にデータがありません');

const ph = pr[0].map(String);

// PLの必要列位置（あなたの現状の列に合わせる）
const periodCol = findHeaderIndex_(ph, ['period']);
const lineTypeCol = findHeaderIndex_(ph, ['line_type']);
const itemNameCol = findHeaderIndex_(ph, ['account_item_name']);
if (periodCol < 0 || lineTypeCol < 0 || itemNameCol < 0) {
throw new Error('PLシートのヘッダに period / line_type / account_item_name が見つかりません');
}

// account_item_id列が無ければ末尾に追加
let itemIdCol = findHeaderIndex_(ph, ['account_item_id']);
if (itemIdCol < 0) {
itemIdCol = ph.length;
pl.getRange(1, itemIdCol + 1).setValue('account_item_id');
}

// 埋める（2行目から）
const out = [];
for (let r = 1; r < pr.length; r++) {
const lineType = String(pr[r][lineTypeCol] || '').trim();
const itemName = String(pr[r][itemNameCol] || '').trim();

if (lineType !== 'line' || !itemName) {
out.push(['']); // total行などは空
continue;
}

// マスタに無い場合は空（後で気づけるようにしておく）
const id = nameToId.get(itemName);
out.push([id === undefined ? '' : id]);
}

pl.getRange(2, itemIdCol + 1, out.length, 1).setValues(out);
return out.length;
}

/** ヘッダ探索（候補名を順に探す） */
function findHeaderIndex_(headerRow, candidates) {
const lower = headerRow.map(h => String(h).trim().toLowerCase());
for (const c of candidates) {
const idx = lower.indexOf(String(c).trim().toLowerCase());
if (idx >= 0) return idx;
}
return -1;
}

/**
* 使い方：
* 1) report_pl_single_month_api のドリルダウンしたい「line」行を選択
* 2) メニュー/関数実行でこれを呼ぶ
* 3) drilldown_journal_lines に明細が出る
*/
function drilldownSelectedPLRowToJournalLines() {
const ss = SpreadsheetApp.getActive();
const pl = ss.getSheetByName('report_pl_single_month_api');
const fact = ss.getSheetByName('fact_journal_lines');
if (!pl) throw new Error('report_pl_single_month_api が見つかりません');
if (!fact) throw new Error('fact_journal_lines が見つかりません');

const active = ss.getActiveSheet();
if (active.getName() !== 'report_pl_single_month_api') {
throw new Error('report_pl_single_month_api シート上で行を選択して実行してください');
}

const row = pl.getActiveRange().getRow();
if (row < 2) throw new Error('ヘッダ行は選択できません（2行目以降を選択してください）');

const plData = pl.getDataRange().getValues();
const ph = plData[0].map(String);

const periodCol = findHeaderIndex_(ph, ['period']);
const lineTypeCol = findHeaderIndex_(ph, ['line_type']);
const itemNameCol = findHeaderIndex_(ph, ['account_item_name']);
const itemIdCol = findHeaderIndex_(ph, ['account_item_id']);

if (periodCol < 0 || lineTypeCol < 0) throw new Error('PLヘッダに period / line_type が見つかりません');
if (itemIdCol < 0) throw new Error('PLに account_item_id 列がありません。先に enrichPLWithAccountItemId() を実行してください');

const r = plData[row - 1];
const lineType = String(r[lineTypeCol] || '').trim();
if (lineType !== 'line') throw new Error('total行はドリルダウン対象外です（line行を選択してください）');

const period = String(r[periodCol] || '').trim();
const accountItemId = String(r[itemIdCol] || '').trim();
const accountItemName = String(r[itemNameCol] || '').trim();

if (!accountItemId) {
throw new Error(`account_item_id が空です（科目名="${accountItemName}" がマスタに無い可能性）。enrichPLWithAccountItemId() を見直してください`);
}

const range = parsePeriodLabel_(period); // {from:Date,to:Date}

// fact_journal_lines のヘッダ
const fd = fact.getDataRange().getValues();
if (fd.length < 2) throw new Error('fact_journal_lines にデータがありません');
const fh = fd[0].map(String);

const issueDateCol = findHeaderIndex_(fh, ['issue_date', 'date']);
const factItemIdCol = findHeaderIndex_(fh, ['account_item_id']);
if (issueDateCol < 0 || factItemIdCol < 0) {
throw new Error('fact_journal_lines のヘッダに issue_date / account_item_id が見つかりません');
}

// 抽出
const outRows = [];
for (let i = 1; i < fd.length; i++) {
const rr = fd[i];

const id = String(rr[factItemIdCol] || '').trim();
if (id !== accountItemId) continue;

const d = toDate_(rr[issueDateCol]);
if (!d) continue;

if (d >= range.from && d <= range.to) {
outRows.push(rr);
}
}

// 出力シート（ヘッダはfactのヘッダそのまま）
const out = ss.getSheetByName('drilldown_journal_lines') || ss.insertSheet('drilldown_journal_lines');
out.clearContents();

// 上部にメタ情報
out.getRange(1, 1, 1, 2).setValues([['meta_key', 'meta_value']]);
out.getRange(2, 1, 4, 2).setValues([
['period', period],
['account_item_id', accountItemId],
['account_item_name', accountItemName],
['rows', outRows.length],
]);

const headerRow = 7;
out.getRange(headerRow, 1, 1, fh.length).setValues([fh]);

if (outRows.length) {
out.getRange(headerRow + 1, 1, outRows.length, fh.length).setValues(outRows);
}

out.setFrozenRows(headerRow);

// ★ ドリルダウンシートをアクティブにする
ss.setActiveSheet(out);
out.activate();
if (outRows.length > 0) {
out.setActiveRange(out.getRange(headerRow + 1, 1));
}
return outRows.length;
}
/** "YYYY-MM-DD..YYYY-MM-DD" をDate範囲にする */
function parsePeriodLabel_(s) {
const m = String(s).trim().match(/^(\d{4}-\d{2}-\d{2})\.\.(\d{4}-\d{2}-\d{2})$/);
if (!m) throw new Error(`period形式が想定外です: ${s}（YYYY-MM-DD..YYYY-MM-DD）`);
const from = parseYmd_(m[1]);
const to = parseYmd_(m[2]);
// toは当日終端として扱いたいので 23:59:59 相当の判定にする（Date比較用に日付のみでOK）
return { from, to };
}

/** セル値が Date/文字列 どちらでも Date に寄せる */
function toDate_(v) {
if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
const s = String(v || '').trim();
if (!s) return null;

// まず YYYY-MM-DD を期待
const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

return null;
}
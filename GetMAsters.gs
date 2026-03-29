/**
* 20_master.gs（完全版）
*
* 【できること】
* 1) 会社一覧（companies）を company_id シートへ
* 2) 勘定科目（account items）を master_account_items へ
* - 標準：selectables から「カテゴリ/収支区分等の属性」を付けて作成
* - フォールバック：/account_items 直叩き（必要なら）
* 3) 部門（sections）を master_sections へ
* 4) 取引先（partners）を master_partners へ
*
* 【シート名（推奨）】
* - company_id
* - master_account_items
* - master_sections
* - master_partners
*
* 【設計方針】
* - “ボタン1発” で全部更新できる（syncAllMasters）
* - 個別更新もできる（syncCompanies / syncAccountItems / syncSections / syncPartners）
*/


/* =========================
* まとめ実行（メニューから呼ぶ）
* ========================= */
function syncAllMasters() {
const results = [];

// company_id が無くても取れる
results.push(['companies', syncCompanies()]);

// ここから先は company_id が必要
// （あなたの方針どおり：configに手入力を前提）
results.push(['account_items', syncAccountItems()]); // 属性付きが標準
results.push(['sections', syncSections()]);
results.push(['partners', syncPartners()]);

log_('MASTER_ALL', JSON.stringify(results));
return results;
}

/* =========================
* 会社一覧（companies）をシートへ
* ========================= */
function syncCompanies() {
const res = apiGet_('/api/1/companies', {});
const companies = res.companies || [];
if (companies.length === 0) throw new Error('companies が0件でした');

const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('company_id') || ss.insertSheet('company_id');

sh.clearContents();
sh.appendRow(['id', 'display_name', 'company_number', 'role']);

const rows = companies.map(c => [
c.id,
c.display_name || '',
c.company_number || '',
c.role || '',
]);

sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
log_('MASTER_COMPANIES', `rows=${rows.length}`);
return rows.length;
}

/* =========================
* 勘定科目（属性付き）をシートへ（標準）
* =========================
*
* A案：/api/1/forms/selectables?includes=account_item を使う
* - 科目カテゴリ（title / balance等）と、カテゴリ配下の account_items をまとめて取れる
* - 科目を “属性つき” で master_account_items に持てる
*/
function syncAccountItems() {
const companyId = mustCompanyId_();

// まずは属性付き（selectables）で試す
try {
const count = syncAccountItemsWithAttributes_(companyId);
return count;
} catch (e) {
// フォールバック（必要なら）
log_('MASTER_ACCOUNT_ITEMS_WARN', `selectables failed -> fallback: ${String(e)}`);
const count = syncAccountItemsFallback_(companyId);
return count;
}
}

/**
* selectables を使って勘定科目＋属性を作る（本体）
*
* master_account_items への出力列（この順で作る）
* - account_item_id
* - name
* - account_category_id
* - account_category_title
* - account_category_role
* - balance （例：income / expense など。PLの収益/費用に使える）
* - group_name （決算書表示名のヒント：取れれば埋める）
* - tax_code （取れれば）
* - available （取れれば）
*/
function syncAccountItemsWithAttributes_(companyId) {
const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('master_account_items') || ss.insertSheet('master_account_items');

// 1) selectables を取得（account_item を含める）
const res = apiGet_('/api/1/forms/selectables', {
company_id: companyId,
includes: 'account_item',
});

// 2) account_groups（小カテゴリ）から category_id → group_name を引けるようにする
// （レスポンスに無ければ空のまま）
const groupNameByCategoryId = new Map();
if (Array.isArray(res.account_groups)) {
for (const g of res.account_groups) {
if (g && g.account_category_id != null && g.name) {
groupNameByCategoryId.set(String(g.account_category_id), String(g.name));
}
}
}

// 3) account_categories 配下の account_items をフラット化し、科目行を作る
const rows = [];
const categories = Array.isArray(res.account_categories) ? res.account_categories : [];

for (const cat of categories) {
if (!cat) continue;

const catId = (cat.id != null) ? String(cat.id) : ''; // 取れない場合もある
const catTitle = String(cat.title || '');
const catRole = String(cat.role || '');
const balance = String(cat.balance || ''); // 収益/費用などの区分に利用

const groupName = catId ? (groupNameByCategoryId.get(catId) || '') : '';

const items = Array.isArray(cat.account_items) ? cat.account_items : [];
for (const it of items) {
if (!it || it.id == null) continue;

rows.push([
Number(it.id), // account_item_id
String(it.name || ''), // name
catId, // account_category_id（空あり）
catTitle, // account_category_title
catRole, // account_category_role
balance, // balance（income/expense等）
groupName, // group_name（取れれば）
String(it.tax_code || ''), // tax_code（取れれば）
(it.available === undefined ? '' : it.available), // available（取れれば）
]);
}
}

// 4) シートへ書く
sh.clearContents();
sh.appendRow([
'account_item_id',
'name',
'account_category_id',
'account_category_title',
'account_category_role',
'balance',
'group_name',
'tax_code',
'available',
]);

if (rows.length > 0) {
sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

log_('MASTER_ACCOUNT_ITEMS', `mode=selectables rows=${rows.length} company_id=${companyId}`);
return rows.length;
}

/**
* フォールバック：/api/1/account_items 直叩き
* - もし selectables でうまく取れない／仕様差が出る場合に備えて残す
* - 属性は最低限（type/tax_code/available）程度
*/
function syncAccountItemsFallback_(companyId) {
const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('master_account_items') || ss.insertSheet('master_account_items');

const res = apiGet_('/api/1/account_items', { company_id: companyId });
const items = res.account_items || [];

sh.clearContents();
sh.appendRow([
'account_item_id',
'name',
'type',
'tax_code',
'available',
]);

const rows = items.map(it => [
Number(it.id),
String(it.name || ''),
String(it.type || ''),
String(it.tax_code || ''),
(it.available === undefined ? '' : it.available),
]);

if (rows.length > 0) {
sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

log_('MASTER_ACCOUNT_ITEMS', `mode=fallback rows=${rows.length} company_id=${companyId}`);
return rows.length;
}

/* =========================
* 部門（sections）
* ========================= */
function syncSections() {
const companyId = mustCompanyId_();

const res = apiGet_('/api/1/sections', { company_id: companyId });
const sections = res.sections || [];

const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('master_sections') || ss.insertSheet('master_sections');

sh.clearContents();
sh.appendRow(['id', 'name', 'shortcut1', 'shortcut2', 'available']);

if (sections.length === 0) {
log_('MASTER_SECTIONS', `rows=0 company_id=${companyId}`);
return 0;
}

const rows = sections.map(s => [
s.id,
s.name,
s.shortcut1 || '',
s.shortcut2 || '',
(s.available === undefined ? '' : s.available),
]);

sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
log_('MASTER_SECTIONS', `rows=${rows.length} company_id=${companyId}`);
return rows.length;
}

/* =========================
* 取引先（partners）
* ========================= */
function syncPartners() {
const companyId = mustCompanyId_();

const res = apiGet_('/api/1/partners', { company_id: companyId });
const partners = res.partners || [];

const ss = SpreadsheetApp.getActive();
const sh = ss.getSheetByName('master_partners') || ss.insertSheet('master_partners');

sh.clearContents();
sh.appendRow(['id', 'name', 'code', 'available']);

if (partners.length === 0) {
log_('MASTER_PARTNERS', `rows=0 company_id=${companyId}`);
return 0;
}

const rows = partners.map(p => [
p.id,
p.name,
p.code || '',
(p.available === undefined ? '' : p.available),
]);

sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
log_('MASTER_PARTNERS', `rows=${rows.length} company_id=${companyId}`);
return rows.length;
}

/* =========================
* 内部ヘルパ
* ========================= */
function mustCompanyId_() {
const companyId = getConfig_('company_id');
if (!companyId) throw new Error('config に company_id を入力してください');
return String(companyId);
}

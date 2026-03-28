/**
* 30_fact_deals.gs
* - deals を取得
* - deal.details を 1行 = 1明細 として fact_journal_lines に展開
*/

function syncDealsToFact() {
  const companyId = getConfig_('company_id');
  if (!companyId) throw new Error('config に company_id がありません');

  // --- 【修正】日付取得とフォーマット強制 ---
  const fromRaw = getConfig_('from_date');
  const toRaw = getConfig_('to_date');
  
  if (!fromRaw || !toRaw) throw new Error('config に from_date / to_date がありません');

  // Dateオブジェクトであっても確実に "yyyy-MM-dd" 形式の文字列に変換する
  const fromDateStr = (fromRaw instanceof Date) ? formatYmdForFact_(fromRaw) : String(fromRaw).trim();
  const toDateStr = (toRaw instanceof Date) ? formatYmdForFact_(toRaw) : String(toRaw).trim();

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('fact_journal_lines') || ss.insertSheet('fact_journal_lines');

  // ヘッダがなければ作る
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'source_type',
      'source_id',
      'line_no',
      'issue_date',
      'account_item_id',
      'amount',
      'partner_id',
      'section_id',
      'description',
    ]);
  }

  const limit = 100;
  let offset = 0;
  let totalLines = 0;

  while (true) {
    // --- 【重要】APIへ渡す日付を文字列(fromDateStr)に変更 ---
    const res = apiGet_('/api/1/deals', {
      company_id: companyId,
      start_issue_date: fromDateStr, 
      end_issue_date: toDateStr,
      limit: limit,
      offset: offset,
    });

    const deals = res.deals || [];
    if (deals.length === 0) break;

    const rows = [];

    for (const d of deals) {
      const issueDate = d.issue_date;
      const partnerId = d.partner_id || '';
      const sectionId = d.section_id || '';

      (d.details || []).forEach((line, idx) => {
        rows.push([
          'deal',                // source_type
          d.id,                  // source_id
          idx + 1,               // line_no
          issueDate,             // issue_date
          line.account_item_id,  // 勘定科目
          line.amount,           // 金額
          partnerId,             // 取引先
          sectionId,             // 部門
          line.description || '',
        ]);
      });
    }

    if (rows.length > 0) {
      sh.getRange(
        sh.getLastRow() + 1,
        1,
        rows.length,
        rows[0].length
      ).setValues(rows);

      totalLines += rows.length;
    }

    offset += limit;
    if (deals.length < limit) break;
  }

  log_('FACT_DEALS', `inserted_lines=${totalLines}`);
  return totalLines;
}

/**
 * 日付オブジェクトを yyyy-MM-dd 形式の文字列に変換する補助関数
 * 他の関数と重複しないよう名称を変更しています
 */
function formatYmdForFact_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}
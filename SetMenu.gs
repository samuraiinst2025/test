function onOpen() {
  try {
SpreadsheetApp.getUi()
.createMenu('freee PoC')
.addItem('① Redirect URI を表示', 'logRedirectUri')
.addItem('② 認可する', 'authorize')
.addItem('③ 疎通テスト（users/me）', 'testPing')
.addItem('④ 会社情報テスト（companies）', 'testCompanies')
.addItem('⑤ マスタを取得', 'syncAllMasters')
.addItem('⑥ 明細取得（deals）', 'syncDealsToFact')
.addItem('⑦ 単月PL作成（trial_pl）', 'buildSingleMonthPLFromTrialPL')
.addItem('⑧ ドリルダウン（仕訳明細）', 'drilldownSelectedPLRowToJournalLines')
.addItem('⑨ 月別TB作成）', 'exportTrialPlBalancesMonthly_')
.addItem('⑩実績分析を実行', 'main') // BudgetSupport.gs の main関数を呼び出す
.addToUi();
} catch(e) {
    Logger.log('カスタムメニューの作成に失敗しました: ' + e.toString());
}
} 
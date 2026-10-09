/**
 * 理事ポータル：講座準備の8段階(企画決定→講師関係→募集準備→募集開始→申込管理→開催準備→開催→終了処理)の判定ロジック。
 * 講座準備進捗確認(course-progress.html)とトップ画面の「進捗状況」カード(index.html)の両方で使う共通ファイル。
 * 判定の基準を1か所にまとめ、2つの画面で完了・滞留の判定がずれないようにしている。
 * 「不要」の選択肢(2026-09-27追加)：参加者募集フォーム・講師資料PDF・理事資料を使わない講座では、phaseJsonの
 *   naForm / naLecturerPdf / naDirectorMaterials にチェック日を入れ、その項目を完了(不要)として扱う。
 *   不要のチェック欄そのもの(optional:true)は、段階の完了判定には含めない。
 * 使い方: PrepPhases.evaluate(講座, { invoiceStates, materialCounts, reportedCourses, checklistItems, phases, formSummaries })
 *         (invoiceStates等はGASの getPrepOverview の戻り値。formSummariesは講座ID→参加者募集フォームの集計結果)
 */
(function () {
  'use strict';

  var WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];
  function parseYmd(s) {
    var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
    return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
  }
  function todayDate() { var d = new Date(); d.setHours(0, 0, 0, 0); return d; }
  function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function diffDays(a, b) { return Math.round((a.getTime() - b.getTime()) / 86400000); }
  function daysUntil(dateStr) { var d = parseYmd(dateStr); return d ? diffDays(d, todayDate()) : null; }
  function jpDate(dateStr, withWeekday) {
    var d = parseYmd(dateStr);
    if (!d) return dateStr || '';
    var base = d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日';
    return withWeekday ? base + '（' + WEEKDAY_LABELS[d.getDay()] + '）' : base;
  }
  function jpMonth(ym) {
    var mm = String(ym || '').match(/^(\d{4})-(\d{1,2})/);
    return mm ? (Number(mm[1]) + '年' + Number(mm[2]) + '月度') : String(ym || '');
  }
  function safeJson(s) {
    try { var v = JSON.parse(s || '{}'); return (v && typeof v === 'object') ? v : {}; } catch (e) { return {}; }
  }

  function evaluate(c, data) {
    var invoiceStates = data.invoiceStates || {}, materialCounts = data.materialCounts || {}, reportedCourses = data.reportedCourses || {};
    var checklistItems = data.checklistItems || [], phases = data.phases || [], formSummaries = data.formSummaries || {};
    var inv = invoiceStates[c.id] || null;
    var invActive = inv && inv.status !== '取消' ? inv : null;
    var checklist = safeJson(c.prepChecklistJson);
    var phaseState = safeJson(c.phaseJson);
    var form = formSummaries[c.id];
    var mat = materialCounts[c.id] || { director: 0, lecturer: 0 };
    var eventDate = parseYmd(c.eventDate);
    var naForm = !!phaseState.naForm, naLecturerPdf = !!phaseState.naLecturerPdf, naDirectorMaterials = !!phaseState.naDirectorMaterials;
    var naLecturer = !!phaseState.naLecturer, naCoursePdf = !!phaseState.naCoursePdf; // 講師なし(勉強会など)／講座案内PDFなし
    // 謝礼0円の請書(支払いなし)は、講師が確定した時点で支払いまで完了とみなす
    var noPayment = !!(invActive && invActive.noPayment && (invActive.status === '確定済み' || invActive.status === '支払い済み'));
    function naItem(key, on, label) {
      return { type: 'check', kind: 'phase', key: key, optional: true, na: true, done: on, label: label, note: on ? (jpDate(phaseState[key]) + ' に「不要」に設定') : '' };
    }

    var groups = [];

    // 1 企画決定
    groups.push([
      { type: 'auto', done: !!c.eventDate, label: '開催日が決まっている', note: c.eventDate ? jpDate(c.eventDate, true) : '講座編集で開催日を登録してください' },
      { type: 'auto', done: !!c.venue, label: '開催場所が決まっている', note: c.venue || '講座編集で開催場所を登録してください' },
      { type: 'auto', done: !!c.directorName, label: '担当理事が決まっている', note: c.directorName || '講座編集で担当理事を登録してください' }
    ]);

    // 2 講師関係(①②③)
    var invNote = invActive ? ('請求書のステータス：' + invActive.status) : (inv ? '請求書は取り消されています' : '請求書管理簿で作成してください');
    var NO_LECT = '講師なしの講座のため不要';
    groups.push(naLecturer ? [
      { type: 'auto', done: true, label: '講師関係（打ち合わせ・請書兼請求書）', note: NO_LECT },
      naItem('naLecturer', true, '講師がいない講座（勉強会など）')
    ] : [
      { type: 'check', kind: 'lecturerMeetingDone', done: !!c.lecturerMeetingDone, label: '① 講師打ち合わせ済み', note: '打ち合わせメモは講座進捗ページで記入できます' },
      { type: 'auto', done: !!invActive, label: '② 請書兼請求書を送付した', note: invNote + (noPayment ? '（謝礼0円の請書）' : '') },
      { type: 'auto', done: !!(invActive && (invActive.status === '確定済み' || invActive.status === '支払い済み')), label: '③ 講師から返信があった（請求書が確定済み）', note: invActive ? '' : '請求書の発行後、講師が内容を確定すると完了になります' },
      naItem('naLecturer', false, '講師がいない講座（勉強会など）')
    ]);

    // 3 募集準備
    groups.push([
      { type: 'auto', done: !!c.pdfFileId || naCoursePdf, label: '講座案内PDFを登録した', note: c.pdfFileId ? (c.pdfFileName || '') : (naCoursePdf ? '不要（この講座では講座案内PDFはありません）' : '講座進捗ページの「PDF機能」から登録できます') },
      naItem('naCoursePdf', naCoursePdf, '講座案内PDFはない（不要）'),
      { type: 'auto', done: !!c.formSheetUrl || naForm, label: '参加者募集フォームを登録した', note: c.formSheetUrl ? '' : (naForm ? '不要（この講座では参加者募集フォームを使いません）' : '講座進捗ページの「⑤ 参加者募集フォーム」に回答用スプレッドシートのURLを登録してください') },
      naItem('naForm', naForm, '参加者募集フォームは使わない（不要）')
    ]);

    // 4 募集開始
    groups.push([
      { type: 'check', kind: 'phase', key: 'recruitStarted', done: !!phaseState.recruitStarted, label: '告知・募集を開始した', note: phaseState.recruitStarted ? (jpDate(phaseState.recruitStarted) + ' にチェック') : '' }
    ]);

    // 5 申込管理
    var appNote = '';
    if (!c.formSheetUrl && naForm) appNote = '参加者募集フォームを使わない講座です（申込は別の方法で管理）';
    else if (!c.formSheetUrl) appNote = '参加者募集フォームが未登録のため、申込数は表示できません';
    else if (!form || form.loading) appNote = '申込数を取得中…';
    else if (form.error) appNote = '申込数を取得できませんでした';
    else appNote = '合計 ' + form.total + '件（会場 ' + form.venueCount + '件／ZOOM ' + form.zoomCount + '件）';
    groups.push([
      { type: 'info', label: '現在の申込数', note: appNote },
      { type: 'check', kind: 'phase', key: 'applicationsClosed', done: !!phaseState.applicationsClosed, label: '申込を締め切り、参加者が確定した', note: phaseState.applicationsClosed ? (jpDate(phaseState.applicationsClosed) + ' にチェック') : '' }
    ]);

    // 6 開催準備(④)
    var prepItems = checklistItems.map(function (item) {
      return { type: 'check', kind: 'checklist', key: item, done: !!checklist[item], label: item };
    });
    if (naLecturer) {
      prepItems.push({ type: 'auto', done: true, label: '講師資料PDF', note: NO_LECT });
    } else {
      prepItems.push({ type: 'auto', done: !!c.lecturerMaterialFileId || naLecturerPdf, label: '講師資料PDFを登録した', note: c.lecturerMaterialFileId ? (c.lecturerMaterialFileName || '') : (naLecturerPdf ? '不要（この講座では講師資料PDFはありません）' : '講座進捗ページの「④ 講師資料PDF」から登録できます') });
      prepItems.push(naItem('naLecturerPdf', naLecturerPdf, '講師資料PDFはない（不要）'));
    }
    prepItems.push({ type: 'info', label: '資料庫の資料', note: '理事 ' + (naDirectorMaterials && !mat.director ? '不要' : mat.director + '件') + '／講師 ' + mat.lecturer + '件' });
    prepItems.push(naItem('naDirectorMaterials', naDirectorMaterials, '理事資料はない（不要）'));
    groups.push(prepItems);

    // 7 開催(⑤)
    groups.push([
      { type: 'check', kind: 'held', done: c.status === '実施済み', label: '⑤ 講座を開催した（ステータスを「実施済み」にする）', note: '現在のステータス：' + (c.status || '準備中') }
    ]);

    // 8 終了処理(⑥)
    groups.push([
      naLecturer
        ? { type: 'auto', done: true, label: '⑥ 講師への支払い', note: NO_LECT }
        : { type: 'auto', done: !!(invActive && (invActive.status === '支払い済み' || noPayment)), label: '⑥ 講師への支払いが済んだ',
            note: noPayment ? '謝礼0円の請書のため、支払いはありません' : (invActive ? ('請求書のステータス：' + invActive.status + (invActive.paymentMethod === '現金' ? '（現金・当日持参）' : '')) : '請求書が未作成です') },
      { type: 'auto', done: !!reportedCourses[c.id], label: '月次報告書に記載した', note: reportedCourses[c.id] ? (jpMonth(reportedCourses[c.id]) + 'の報告書（提出済み）に記載') : '提出済みの月次報告書の「直近の講座開催」に、この講座名で記載されると完了になります' }
    ]);

    var today = todayDate();
    var cancelled = c.status === '中止';
    var held = c.status === '実施済み';
    var eventIdx = phases.map(function (p) { return p.key; }).indexOf('event');
    var result = phases.map(function (ph, i) {
      var items = groups[i] || [];
      var judged = items.filter(function (it) { return it.type !== 'info' && !it.optional; });
      var done = judged.length > 0 && judged.every(function (it) { return it.done; });
      var target = eventDate ? addDays(eventDate, ph.offsetDays) : null;
      // 開催済みの講座では、開催より前の段階が未チェックでも今から対応する意味はないため、
      // 滞留(赤)にはせず「記録なし」として扱う(終了処理だけは開催後の作業なので引き続き判定する)。
      var skipped = !done && held && eventIdx !== -1 && i < eventIdx;
      var overdue = !done && !skipped && !cancelled && !!target && diffDays(today, target) > 0;
      return { key: ph.key, label: ph.label, offsetDays: ph.offsetDays, items: items, done: done, skipped: skipped, target: target, overdue: overdue };
    });
    var currentIdx = -1;
    for (var i = 0; i < result.length; i++) { if (!result[i].done && !result[i].skipped) { currentIdx = i; break; } }
    return { phases: result, currentIdx: currentIdx, cancelled: cancelled, invActive: invActive,
      na: { form: naForm, lecturerPdf: naLecturerPdf || naLecturer, directorMaterials: naDirectorMaterials, lecturer: naLecturer, coursePdf: naCoursePdf } };
  }

  /** 請求書由来のアラート(ダッシュボードと同じ基準) */
  function invoiceAlerts(c, invActive, alertSettings) {
    alertSettings = alertSettings || {};
    var alerts = [];
    if (!invActive) return alerts;
    if (invActive.status === '作成済み') {
      var du = daysUntil(invActive.eventDate || c.eventDate);
      if (du !== null && du <= Number(alertSettings.unrespondedAlertDays || 14)) alerts.push({ cls: 'chip-red', label: '講師側未返信' });
    }
    if (invActive.status === '確定済み' && invActive.paymentDueDate && !invActive.noPayment) {
      var dd = daysUntil(invActive.paymentDueDate);
      if (dd !== null && dd < 0) alerts.push({ cls: 'chip-red', label: '支払期限超過' });
      else if (dd !== null && dd <= Number(alertSettings.paymentWarningDays || 2)) alerts.push({ cls: 'chip-yellow', label: '支払期限接近' });
    }
    return alerts;
  }


  window.PrepPhases = {
    evaluate: evaluate,
    invoiceAlerts: invoiceAlerts,
    parseYmd: parseYmd, todayDate: todayDate, addDays: addDays, diffDays: diffDays, daysUntil: daysUntil
  };
})();

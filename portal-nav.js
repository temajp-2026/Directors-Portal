/**
 * 理事ポータル：画面間の移動(ナビゲーション)の共通ファイル。
 * ・各画面のヘッダーのすぐ下に、6つの画面を行き来できる「画面切り替えバー」を表示する(今いる画面は強調表示)。
 * ・講座ごとに、その講座の「講座進捗／準備進捗／請求書／資料庫／SNS」へ直接移動するリンクを作る
 *   (PortalNav.courseLinksHtml)。移動先の画面は ?course=<講座ID> を受け取り、その講座を開いた状態で表示する。
 * このファイルが読み込めなくても、各画面は今まで通り動く(バーとリンクが出ないだけ)。
 */
(function () {
  'use strict';

  var PAGES = [
    { file: 'index.html',           label: 'トップ',     color: '#5f6368' },
    { file: 'course-progress.html', label: '準備進捗',   color: '#1e8e3e' },
    { file: 'invoice-ledger.html',  label: '請求書',     color: '#1a73e8' },
    { file: 'reports.html',         label: '月次報告書', color: '#e8710a' },
    { file: 'materials.html',       label: '資料庫',     color: '#8430ce' },
    { file: 'sns.html',             label: 'SNS',        color: '#d01884' }
  ];

  // 講座ごとのリンク(key は exclude 指定に使う)
  var COURSE_LINKS = [
    { key: 'course',    label: '講座進捗', href: function (id) { return 'index.html?openCourse=' + encodeURIComponent(id); } },
    { key: 'progress',  label: '準備進捗', href: function (id) { return 'course-progress.html?course=' + encodeURIComponent(id); } },
    { key: 'invoice',   label: '請求書',   href: function (id) { return 'invoice-ledger.html?course=' + encodeURIComponent(id); } },
    { key: 'materials', label: '資料庫',   href: function (id) { return 'materials.html?course=' + encodeURIComponent(id); } },
    { key: 'sns',       label: 'SNS',      href: function (id) { return 'sns.html?course=' + encodeURIComponent(id); } }
  ];

  function currentFile() {
    var f = String(window.location.pathname || '').split('/').pop() || 'index.html';
    if (f === '請求書管理簿.html') f = 'invoice-ledger.html';
    return f;
  }

  function injectStyle() {
    if (document.getElementById('portalNavStyle')) return;
    var css =
      '.portal-nav{background:var(--gw-surface,#fff);border-bottom:1px solid var(--gw-border,#dadce0);}' +
      '.portal-nav-inner{max-width:1100px;margin:0 auto;padding:6px 16px;display:flex;gap:6px;overflow-x:auto;-webkit-overflow-scrolling:touch;scrollbar-width:none;}' +
      '.portal-nav-inner::-webkit-scrollbar{display:none;}' +
      '.portal-nav a{flex-shrink:0;display:inline-flex;align-items:center;gap:6px;padding:6px 12px;border-radius:999px;text-decoration:none;white-space:nowrap;' +
        'font-size:calc(var(--fs-scale,1)*12.5px);color:var(--gw-text,#202124);border:1px solid transparent;}' +
      '.portal-nav a:hover{background:var(--gw-bg,#f8f9fa);}' +
      '.portal-nav a .pn-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}' +
      '.portal-nav a.pn-current{font-weight:600;background:var(--gw-bg,#f8f9fa);border-color:var(--gw-border,#dadce0);cursor:default;}' +
      '.course-links{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-top:6px;}' +
      '.course-links-label{font-size:calc(var(--fs-scale,1)*11px);color:var(--gw-text-secondary,#5f6368);margin-right:2px;}' +
      '.course-links a{display:inline-block;padding:3px 10px;border-radius:999px;border:1px solid var(--gw-border,#dadce0);background:var(--gw-surface,#fff);' +
        'font-size:calc(var(--fs-scale,1)*11.5px);color:var(--gw-blue,#1a73e8);text-decoration:none;white-space:nowrap;}' +
      '.course-links a:hover{background:var(--gw-blue-tint,#e8f0fe);border-color:var(--gw-blue,#1a73e8);}' +
      '@media print{.portal-nav,.course-links{display:none !important;}}';
    var st = document.createElement('style');
    st.id = 'portalNavStyle';
    st.textContent = css;
    document.head.appendChild(st);
  }

  function buildNav() {
    var header = document.querySelector('.app-header');
    if (!header || document.getElementById('portalNav')) return;
    var cur = currentFile();
    var nav = document.createElement('nav');
    nav.className = 'portal-nav no-print';
    nav.id = 'portalNav';
    nav.setAttribute('aria-label', '画面の切り替え');
    var inner = document.createElement('div');
    inner.className = 'portal-nav-inner';
    PAGES.forEach(function (p) {
      var a = document.createElement('a');
      a.href = p.file;
      if (p.file === cur) { a.className = 'pn-current'; a.setAttribute('aria-current', 'page'); }
      var dot = document.createElement('span');
      dot.className = 'pn-dot';
      dot.style.background = p.color;
      a.appendChild(dot);
      a.appendChild(document.createTextNode(p.label));
      inner.appendChild(a);
    });
    nav.appendChild(inner);
    header.parentNode.insertBefore(nav, header.nextSibling);
    var curLink = inner.querySelector('.pn-current');
    if (curLink && curLink.scrollIntoView && inner.scrollWidth > inner.clientWidth) {
      inner.scrollLeft = Math.max(0, curLink.offsetLeft - 16);
    }
  }

  /**
   * 講座ごとのリンク(HTML文字列)。opts.exclude: 表示しないリンクのkey(今いる画面)、opts.newTab: 新しいタブで開く、
   * opts.label: 先頭の見出し(既定「この講座の：」)
   */
  function courseLinksHtml(courseId, opts) {
    opts = opts || {};
    if (!courseId) return '';
    var ex = [].concat(opts.exclude || []);
    var target = opts.newTab ? ' target="_blank" rel="noopener"' : '';
    var html = '<div class="course-links"><span class="course-links-label">' + (opts.label || 'この講座の：') + '</span>';
    COURSE_LINKS.forEach(function (l) {
      if (ex.indexOf(l.key) !== -1) return;
      html += '<a href="' + l.href(courseId).replace(/"/g, '&quot;') + '"' + target + '>' + l.label + ' ›</a>';
    });
    return html + '</div>';
  }

  /** URLの ?course=<講座ID> を読む(無ければ空文字) */
  function courseParam() {
    try { return new URLSearchParams(window.location.search).get('course') || ''; } catch (e) { return ''; }
  }

  injectStyle();
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', buildNav);
  else buildNav();

  window.PortalNav = { courseLinksHtml: courseLinksHtml, courseParam: courseParam };
})();

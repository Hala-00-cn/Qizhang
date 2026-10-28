import {
  loadState, saveState, resetState, typeLabels, formatMoney, localDateKey,
  getAccountBalance, accountDisplayValue, getAssetSummary, monthTransactions,
  incomeExpenseSummary, uid, cloneCategoriesForLedger, roundMoney
} from './data.js?v=15';
import { buildExcelWorkbook, readExcelWorkbook } from './excel.js?v=17';

const pageEl = document.querySelector('#page');
const headerEl = document.querySelector('#appHeader');
const navEl = document.querySelector('#bottomNav');
const sheetRoot = document.querySelector('#sheetRoot');
const toastEl = document.querySelector('#toast');

let state = loadState();
let activePage = 'home';
let viewingMonth = new Date();
let searchTerm = '';
let editingTransactionId = null;
let entryDraft = null;
let editingAccountId = null;
let editingLedgerId = null;
let editingCategoryId = null;
let statsPeriod = 'month';
let reorderDrag = null;
let pendingExcelImport = null;

const iconPaths = {
  home: '<path d="m3 11 9-8 9 8"/><path d="M5 10v10h14V10"/><path d="M9 20v-6h6v6"/>',
  list: '<path d="M8 6h13M8 12h13M8 18h13"/><path d="M3 6h.01M3 12h.01M3 18h.01"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  chart: '<path d="M4 19V9M10 19V5M16 19v-7M22 19V2"/>',
  user: '<circle cx="12" cy="8" r="4"/><path d="M4 21a8 8 0 0 1 16 0"/>',
  bell: '<path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9"/><path d="M10 21h4"/>',
  eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  eyeOff: '<path d="m3 3 18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.9 4.2A10.5 10.5 0 0 1 22 12a16 16 0 0 1-2 3M6.6 6.6A15.8 15.8 0 0 0 2 12s3.5 7 10 7a9.7 9.7 0 0 0 3.4-.6"/>',
  down: '<path d="m6 9 6 6 6-6"/>',
  right: '<path d="m9 18 6-6-6-6"/>',
  search: '<circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/>',
  filter: '<path d="M4 6h16M7 12h10M10 18h4"/>',
  arrowUp: '<path d="m12 19V5m-5 5 5-5 5 5"/>',
  arrowDown: '<path d="m12 5v14m-5-5 5 5 5-5"/>',
  transfer: '<path d="M7 7h12l-3-3m3 3-3 3M17 17H5l3 3m-3-3 3-3"/>',
  wallet: '<path d="M4 6h15a2 2 0 0 1 2 2v10H4a2 2 0 0 1-2-2V6a3 3 0 0 1 3-3h13"/><path d="M16 12h5"/>',
  close: '<path d="m6 6 12 12M18 6 6 18"/>',
  backspace: '<path d="m21 5-3-3H8L2 12l6 10h10l3-3V5Z"/><path d="m11 9 5 6m0-6-5 6"/>',
  calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M16 3v4M8 3v4M3 11h18"/>',
  book: '<path d="M4 19.5A2.5 2.5 0 0 1 6.5 17H20V3H6.5A2.5 2.5 0 0 0 4 5.5v14Z"/><path d="M8 7h8"/>',
  layers: '<path d="m12 2 9 5-9 5-9-5 9-5Z"/><path d="m3 12 9 5 9-5M3 17l9 5 9-5"/>',
  tag: '<path d="M20.6 13.6 11 4H4v7l9.6 9.6a2 2 0 0 0 2.8 0l4.2-4.2a2 2 0 0 0 0-2.8Z"/><circle cx="7.5" cy="7.5" r=".5"/>',
  credit: '<rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20"/>',
  handshake: '<path d="m11 17 2 2a2 2 0 0 0 3-3l-5-5M14 14l1 1a2 2 0 0 0 3-3l-4-4-3 3a2 2 0 0 1-3-3l3-3 5 1 5 5M2 12l4-6 3 2M22 12l-3-6-3 2M6 17l2-2M8 19l2-2"/>',
  settings: '<circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-2.8 2.8-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.6v.2h-4V21a1.7 1.7 0 0 0-1-1.6 1.7 1.7 0 0 0-1.9.3l-.1.1L4.2 17l.1-.1a1.7 1.7 0 0 0 .3-1.9A1.7 1.7 0 0 0 3 14H2.8v-4H3a1.7 1.7 0 0 0 1.6-1 1.7 1.7 0 0 0-.3-1.9L4.2 7 7 4.2l.1.1a1.7 1.7 0 0 0 1.9.3A1.7 1.7 0 0 0 10 3V2.8h4V3a1.7 1.7 0 0 0 1 1.6 1.7 1.7 0 0 0 1.9-.3l.1-.1L19.8 7l-.1.1a1.7 1.7 0 0 0-.3 1.9A1.7 1.7 0 0 0 21 10h.2v4H21a1.7 1.7 0 0 0-1.6 1Z"/>',
  moon: '<path d="M21 12.8A9 9 0 1 1 11.2 3 7 7 0 0 0 21 12.8Z"/>',
  info: '<circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/>',
  trash: '<path d="M3 6h18M8 6V4h8v2M19 6l-1 15H6L5 6M10 11v5M14 11v5"/>',
  edit: '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4 11.5-11.5Z"/>',
  grip: '<path d="M8 7h8M8 12h8M8 17h8"/>'
};

const categoryIconChoices = [
  '🍚','🥣','🍜','🍱','🍰','☕','🧋','🍎','🛒','🛍️','👕','👟',
  '🏠','💡','📱','💻','🎧','🎮','🎬','🎵','📚','✏️','💊','🏥',
  '🚇','🚌','🚕','🚲','🚗','⛽','✈️','🚄','🏨','🌴','💼','💰',
  '🎁','🧧','📈','💳','🔧','🐾','👶','❤️','⭐','🍀','🏷️','➕'
];
const assetTagOptions = ['活钱管理','保险保障','长期投资','稳健理财'];
const assetTagColors = ['#24a66f','#5f8fe8','#8a78d6','#f0a844'];

function icon(name, cls = '') {
  return `<svg class="${cls}" aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${iconPaths[name] || iconPaths.info}</svg>`;
}

function persist() {
  saveState(state);
  document.body.classList.toggle('dark', Boolean(state.settings.dark));
}

function escapeHtml(value = '') {
  return String(value).replace(/[&<>'"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[ch]));
}

function showToast(message) {
  toastEl.textContent = message;
  toastEl.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toastEl.classList.remove('show'), 1800);
}

function activeLedger() {
  return state.ledgers.find(l => l.id === state.settings.activeLedgerId) || state.ledgers[0];
}

function categoriesForLedger(ledgerId = state.settings.activeLedgerId) {
  return state.categories.filter(category => category.ledgerId === ledgerId);
}

function header() {
  const ledger = activeLedger();
  const titles = { transactions: '搜索账目', stats: '洞察', assets: '资产管理', mine: '我的' };
  const hasPendingReminder = currentReminders().length > 0 || state.messages.some(m => !m.read);
  headerEl.innerHTML = `
    <div class="header-title">
      ${activePage === 'home' ? `<button class="ledger-switch" data-action="ledger"><span class="ledger-dot" style="background:${ledger.color}"></span>${escapeHtml(ledger.name)}${icon('down')}</button>` : `<span>${titles[activePage] || '栖账'}</span>`}
    </div>
    <div class="header-actions">
      ${activePage === 'home' ? `<button class="icon-button" data-action="calendar" aria-label="日历账目">${icon('calendar')}</button><button class="icon-button" data-action="search" aria-label="搜索账目">${icon('search')}</button>` : ''}
      <button class="icon-button" data-action="messages" aria-label="消息中心">${icon('bell')}${hasPendingReminder ? '<span class="notification-dot"></span>' : ''}</button>
    </div>`;
}

function nav() {
  const items = [
    ['home', '明细', 'list'], ['stats', '洞察', 'chart'], ['add', '记账', 'plus'], ['assets', '资产', 'wallet'], ['mine', '我的', 'user']
  ];
  navEl.innerHTML = items.map(([id, label, iconName]) => id === 'add'
    ? `<button class="nav-item add" data-page="add"><span class="add-orb">${icon(iconName)}</span><span>${label}</span></button>`
    : `<button class="nav-item ${activePage === id ? 'active' : ''}" data-page="${id}">${icon(iconName)}<span>${label}</span></button>`
  ).join('');
}

function categoryById(id) { return state.categories.find(c => c.id === id); }
function accountById(id) { return state.accounts.find(a => a.id === id); }

function txLabel(tx) {
  const cat = categoryById(tx.categoryId);
  if (cat) return cat.name;
  if (tx.type === 'transfer') return '账户转账';
  if (tx.type === 'lend') return `借出${tx.counterparty ? ` · ${tx.counterparty}` : ''}`;
  if (tx.type === 'borrow') return `借入${tx.counterparty ? ` · ${tx.counterparty}` : ''}`;
  if (tx.type === 'repay') return '归还借款';
  if (tx.type === 'collect') return '收回借款';
  return typeLabels[tx.type] || '账目';
}

function txIcon(tx) { return categoryById(tx.categoryId)?.icon || ({ transfer: '⇄', lend: '↗', borrow: '↙', repay: '↗', collect: '↙', adjustment: '≈' }[tx.type] || '记'); }
function txColor(tx) { return categoryById(tx.categoryId)?.parentId === 'food' ? '#fff0ed' : tx.type === 'income' ? '#e9f8f3' : tx.type === 'transfer' ? '#eeebff' : '#f3f1f8'; }

function txAmountHtml(tx) {
  const hidden = state.settings.amountsHidden;
  const special = !['income', 'expense'].includes(tx.type);
  const sign = tx.type === 'income' ? '+' : tx.type === 'expense' ? '-' : '';
  return `<span class="tx-amount ${tx.type === 'income' ? 'income' : special ? 'special' : ''}">${sign}${formatMoney(tx.amount, hidden)}</span>`;
}

function transactionRows(transactions, limit = Infinity) {
  const items = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, limit);
  if (!items.length) return `<div class="empty"><div class="empty-orb">${icon('list')}</div><strong>还没有账目</strong><p>点击底部“记账”，从今天的第一笔开始。</p></div>`;
  const groups = Object.groupBy ? Object.groupBy(items, tx => localDateKey(tx.date)) : items.reduce((acc, tx) => ((acc[localDateKey(tx.date)] ||= []).push(tx), acc), {});
  return `<div class="transaction-list">${Object.entries(groups).map(([date, rows]) => {
    const summary = incomeExpenseSummary(rows);
    const d = new Date(`${date}T00:00:00`);
    const today = localDateKey(new Date()) === date;
    const yesterday = localDateKey(new Date(Date.now() - 86400000)) === date;
    const label = today ? '今天' : yesterday ? '昨天' : `${d.getMonth() + 1}月${d.getDate()}日`;
    return `<section class="day-group"><div class="day-head"><span>${label} · ${['日','一','二','三','四','五','六'][d.getDay()]}</span><span class="day-totals"><span>收 ${formatMoney(summary.income, state.settings.amountsHidden)}</span><span>支 ${formatMoney(summary.expense, state.settings.amountsHidden)}</span></span></div>${rows.map(tx => {
      const acc = accountById(tx.accountId);
      const time = new Date(tx.date).toLocaleTimeString('zh-CN', { hour: '2-digit', minute: '2-digit' });
      return `<button class="tx-row" data-tx-id="${tx.id}"><span class="tx-icon" style="background:${txColor(tx)}">${txIcon(tx)}</span><span class="tx-info"><span class="tx-name"><strong>${escapeHtml(txLabel(tx))}</strong><span>${time}</span></span><span class="tx-meta">${escapeHtml(acc?.name || '已删除账户')}${tx.note ? ` · ${escapeHtml(tx.note)}` : ''}</span></span>${txAmountHtml(tx)}</button>`;
    }).join('')}</section>`;
  }).join('')}</div>`;
}

function currentReminders() {
  const result = [];
  const now = new Date();
  state.accounts.filter(a => a.kind === 'liability').forEach(account => {
    const debt = getAccountBalance(state, account.id);
    if (debt <= 0 || !account.dueDay) return;
    let due = new Date(now.getFullYear(), now.getMonth(), account.dueDay, 9);
    if (due < now) due = new Date(now.getFullYear(), now.getMonth() + 1, account.dueDay, 9);
    result.push({ id: account.id, kind: 'account', title: `${account.name}还款`, amount: debt, due, subtitle: `${account.dueDay}日还款` });
  });
  return result.sort((a, b) => a.due - b.due);
}

function renderHome() {
  const txs = monthTransactions(state, state.settings.activeLedgerId, viewingMonth);
  const summary = incomeExpenseSummary(txs);
  const hidden = state.settings.amountsHidden;
  const usage = summary.income > 0 ? Math.min(100, summary.expense / summary.income * 100) : 0;
  const budgets = getBudgetProgress(state.settings.activeLedgerId, viewingMonth);
  const budgetTotal = budgets.reduce((sum, item) => sum + item.amount, 0);
  const budgetUsed = budgets.reduce((sum, item) => sum + item.used, 0);
  pageEl.innerHTML = `
    <section class="month-overview">
      <button class="month-block" data-action="pick-home-month" aria-label="选择明细月份"><span>${viewingMonth.getFullYear()}年 ${icon('down')}</span><strong>${viewingMonth.getMonth() + 1}<small>月</small></strong></button>
      <div class="month-metric"><span>收入</span><strong>${formatMoney(summary.income, hidden)}</strong></div>
      <div class="month-metric"><span>支出</span><strong>${formatMoney(summary.expense, hidden)}</strong></div>
      <div class="month-metric"><span>结余</span><strong>${formatMoney(summary.income - summary.expense, hidden)}</strong></div>
      <div class="month-progress"><i style="width:${usage}%"></i></div>
    </section>
    ${budgets.length ? `<section class="budget-card"><div class="section-head"><div><span class="section-title">本月预算</span><span class="budget-caption">${escapeHtml(activeLedger().name)} · 分类预算</span></div><button class="text-button" data-action="budgets">管理</button></div><div class="budget-summary"><strong>${formatMoney(budgetUsed,hidden)} <small>/ ${formatMoney(budgetTotal,hidden)}</small></strong><span>剩余 ${formatMoney(Math.max(0,budgetTotal-budgetUsed),hidden)}</span></div><div class="budget-list">${budgets.map(item=>`<div class="budget-line"><div class="budget-line-head"><span>${escapeHtml(item.category.name)}</span><span>${formatMoney(item.used,hidden)} / ${formatMoney(item.amount,hidden)}</span></div><div class="budget-track ${item.percent>=100?'over':item.percent>=80?'near':''}"><i style="width:${Math.min(100,item.percent)}%"></i></div></div>`).join('')}</div></section>` : `<section class="budget-empty"><span>还没有设置本月预算</span><button class="text-button" data-action="budgets">设置预算</button></section>`}
    <section class="detail-section"><div class="section-head"><span class="section-title">账目明细</span><button class="text-button" data-page="transactions">筛选</button></div>${transactionRows(txs, 80)}</section>`;
}

function getBudgetProgress(ledgerId = state.settings.activeLedgerId, monthDate = new Date()) {
  const budgets = (state.budgets || []).filter(item => item.ledgerId === ledgerId && Number(item.amount) > 0);
  const txs = monthTransactions(state, ledgerId, monthDate).filter(tx => tx.type === 'expense');
  return budgets.map(budget => {
    const category = categoryById(budget.categoryId);
    if (!category || category.parentId || category.type !== 'expense') return null;
    const childIds = new Set(categoriesForLedger(ledgerId).filter(item => item.id === category.id || item.parentId === category.id).map(item => item.id));
    const used = roundMoney(txs.filter(tx => childIds.has(tx.categoryId)).reduce((sum, tx) => sum + Number(tx.amount || 0), 0));
    return { ...budget, category, used, percent: budget.amount ? used / budget.amount * 100 : 0 };
  }).filter(Boolean).sort((a,b)=> (a.category.order||0)-(b.category.order||0));
}

function renderTransactions() {
  const current = new Date(viewingMonth.getFullYear(), viewingMonth.getMonth(), 1);
  let txs = monthTransactions(state, state.settings.activeLedgerId, current);
  if (searchTerm) {
    const query = searchTerm.toLowerCase();
    txs = txs.filter(tx => [tx.note, txLabel(tx), accountById(tx.accountId)?.name, tx.amount].some(v => String(v || '').toLowerCase().includes(query)));
  }
  pageEl.innerHTML = `
    <div class="month-switch"><button data-month="-1">${icon('right', 'flip')}</button><div class="month-label"><strong>${current.getFullYear()}年 ${current.getMonth()+1}月</strong><span>${txs.length} 笔账目</span></div><button data-month="1">${icon('right')}</button></div>
    <div class="toolbar"><label class="search-box">${icon('search')}<input id="txSearch" value="${escapeHtml(searchTerm)}" placeholder="搜索金额、分类、账户或备注" /></label><button class="filter-button" data-action="filter">${icon('filter')}</button></div>
    ${transactionRows(txs)}`;
  const left = pageEl.querySelector('[data-month="-1"] svg');
  if (left) left.style.transform = 'rotate(180deg)';
}

function spendingByCategory(txs, type = 'expense') {
  const map = {};
  txs.filter(tx => tx.type === type).forEach(tx => {
    const child = categoryById(tx.categoryId);
    const parent = child?.parentId ? categoryById(child.parentId) : child;
    const key = parent?.id || 'other';
    map[key] ||= { id: key, name: parent?.name || '其他', value: 0 };
    map[key].value += tx.amount;
  });
  return Object.values(map).sort((a,b) => b.value-a.value);
}

function categoryDonutStyle(items, total, palette) {
  if (!items.length || !total) return 'var(--line)';
  let cursor = 0;
  const stops = items.map(({value},index) => {
    const start = cursor;
    cursor += value / total * 100;
    return `${palette[index % palette.length]} ${start.toFixed(2)}% ${cursor.toFixed(2)}%`;
  });
  return `conic-gradient(${stops.join(',')})`;
}

function renderCategoryAnalytics(title, items, total, palette) {
  const emptyText = title.startsWith('收入') ? '当前周期暂无收入分类数据。' : '当前周期暂无支出分类数据。';
  return `<section class="section card chart-card"><div class="chart-head"><strong>${title}</strong><span>一级分类占比</span></div>${items.length ? `<div class="donut-row"><div class="donut" data-label="${title}" style="background:${categoryDonutStyle(items,total,palette)}"></div><div class="legend">${items.slice(0,5).map(({name,value},i)=>`<div class="legend-item"><span class="legend-dot" style="background:${palette[i%palette.length]}"></span><span>${escapeHtml(name)}</span><strong>${total ? Math.round(value/total*100) : 0}%</strong></div>`).join('')}</div></div>` : `<div class="empty"><strong>暂无${title}</strong><p>${emptyText}</p></div>`}</section>
    <section class="section"><div class="section-head"><span class="section-title">${title}明细</span><span class="subtle">点击查看二级分类</span></div><div class="card menu-card">${items.length ? items.map(({id,name,value},i)=>`<button class="menu-row" data-category-breakdown="${id}" data-category-type="${title.startsWith('收入')?'income':'expense'}"><span class="menu-icon" style="color:${palette[i%palette.length]};background:${palette[i%palette.length]}18">${i+1}</span><span class="menu-text"><strong>${escapeHtml(name)}</strong><span>${total ? (value/total*100).toFixed(1) : 0}%</span></span><strong>${formatMoney(value, state.settings.amountsHidden)}</strong>${icon('right','chevron')}</button>`).join('') : `<div class="empty"><p>${emptyText}</p></div>`}</div></section>`;
}

function statsTransactions() {
  const anchor = new Date(viewingMonth);
  return state.transactions.filter(tx => {
    if (tx.ledgerId !== state.settings.activeLedgerId) return false;
    const date = new Date(tx.date);
    if (statsPeriod === 'day') return localDateKey(date) === localDateKey(anchor);
    if (statsPeriod === 'year') return date.getFullYear() === anchor.getFullYear();
    return date.getFullYear() === anchor.getFullYear() && date.getMonth() === anchor.getMonth();
  });
}

function openCategoryBreakdown(parentId, type) {
  const parent = categoryById(parentId);
  if (!parent) return;
  const txs = statsTransactions().filter(tx => {
    if (tx.type !== type) return false;
    const category = categoryById(tx.categoryId);
    return category?.id === parentId || category?.parentId === parentId;
  });
  const children = state.categories.filter(category => category.parentId === parentId).map(category => {
    const rows = txs.filter(tx => tx.categoryId === category.id);
    return { category, rows, total: rows.reduce((sum, tx) => sum + tx.amount, 0) };
  }).filter(item => item.rows.length).sort((a,b) => b.total-a.total);
  const total = txs.reduce((sum, tx) => sum + tx.amount, 0);
  openSheet(`<div class="sheet-head"><h2>${escapeHtml(parent.name)}明细</h2><button class="close-button" data-action="close">${icon('close')}</button></div>
    <div class="card stats-summary breakdown-summary"><div><span>${type==='income'?'收入':'支出'}合计</span><strong>${formatMoney(total, state.settings.amountsHidden)}</strong></div><div><span>二级分类</span><strong>${children.length}</strong></div><div><span>账目数量</span><strong>${txs.length}</strong></div></div>
    <section class="section"><div class="section-head"><span class="section-title">二级分类明细</span></div><div class="card menu-card">${children.length ? children.map(({category,rows,total:childTotal})=>`<div class="menu-row breakdown-row"><span class="menu-icon">${category.icon}</span><span class="menu-text"><strong>${escapeHtml(category.name)}</strong><span>${rows.length} 笔 · ${total ? (childTotal/total*100).toFixed(1) : 0}%</span></span><strong>${formatMoney(childTotal, state.settings.amountsHidden)}</strong></div>`).join('') : '<div class="empty"><p>暂无二级分类数据</p></div>'}</div></section>
    <section class="section"><div class="section-head"><span class="section-title">相关账目</span></div>${transactionRows(txs)}</section>`, true);
}

function renderStats() {
  const anchor = new Date(viewingMonth);
  const txs = statsTransactions();
  const summary = incomeExpenseSummary(txs);
  const expenseCategories = spendingByCategory(txs, 'expense');
  const incomeCategories = spendingByCategory(txs, 'income');
  let points;
  if (statsPeriod === 'year') {
    points = Array.from({length:12},(_,month)=>({ label:`${month+1}月`, value:txs.filter(tx=>new Date(tx.date).getMonth()===month && tx.type==='expense').reduce((s,tx)=>s+tx.amount,0) }));
  } else if (statsPeriod === 'day') {
    points = Array.from({length:6},(_,slot)=>({ label:`${slot*4}时`, value:txs.filter(tx=>new Date(tx.date).getHours()>=slot*4 && new Date(tx.date).getHours()<(slot+1)*4 && tx.type==='expense').reduce((s,tx)=>s+tx.amount,0) }));
  } else {
    points = Array.from({length:6},(_,slot)=>({ label:`${slot*5+1}日`, value:txs.filter(tx=>{const day=new Date(tx.date).getDate();return day>=slot*5+1&&day<=Math.min(slot*5+5,31)&&tx.type==='expense'}).reduce((s,tx)=>s+tx.amount,0) }));
  }
  const max = Math.max(...points.map(point=>point.value), 1);
  const periodTitle = statsPeriod === 'day' ? '当日支出' : statsPeriod === 'year' ? `${anchor.getFullYear()}年支出` : `${anchor.getMonth()+1}月支出`;
  const palette = ['#24a66f','#fb806f','#f0a844','#5f8fe8','#8a78d6'];
  pageEl.innerHTML = `
    <div class="segmented"><button data-stats-period="day" class="${statsPeriod==='day'?'active':''}">日</button><button data-stats-period="month" class="${statsPeriod==='month'?'active':''}">月</button><button data-stats-period="year" class="${statsPeriod==='year'?'active':''}">年</button></div>
    <div class="card stats-summary"><div><span>收入</span><strong>${formatMoney(summary.income, state.settings.amountsHidden)}</strong></div><div><span>支出</span><strong>${formatMoney(summary.expense, state.settings.amountsHidden)}</strong></div><div><span>结余</span><strong>${formatMoney(summary.income-summary.expense, state.settings.amountsHidden)}</strong></div></div>
    <section class="section card chart-card"><div class="chart-head"><div><strong>${periodTitle}</strong><span> · ${escapeHtml(activeLedger().name)}</span></div><span>${formatMoney(summary.expense, state.settings.amountsHidden)}</span></div><div class="bar-chart">${points.map(point=>`<div class="bar-wrap"><div class="bar" style="height:${Math.max(3,point.value/max*100)}%"></div><span class="bar-label">${point.label}</span></div>`).join('')}</div></section>
    ${renderCategoryAnalytics('支出分类', expenseCategories, summary.expense, palette)}
    ${renderCategoryAnalytics('收入分类', incomeCategories, summary.income, [...palette].reverse())}`;
}

function renderAssets() {
  const summary = getAssetSummary(state);
  const hidden = state.settings.amountsHidden;
  const assetAccounts = state.accounts.filter(account => account.kind === 'asset');
  const debtAccounts = state.accounts.filter(account => account.kind === 'liability');
  const assetTotal = Math.max(summary.assets, 1);
  const taggedAccounts = assetAccounts.filter(account=>account.includeInAssets!==false);
  const taggedTotal = taggedAccounts.reduce((sum,account)=>sum+getAccountBalance(state,account.id),0);
  const tagStats = assetTagOptions.map((name,index)=>{
    const amount=taggedAccounts.filter(account=>(account.assetTag||assetTagOptions[0])===name).reduce((sum,account)=>sum+getAccountBalance(state,account.id),0);
    return {name,amount,ratio:taggedTotal>0?Math.max(0,amount)/taggedTotal*100:0,color:assetTagColors[index]};
  });
  pageEl.innerHTML = `
    <section class="net-worth-card">
      <button class="privacy-float" data-action="privacy" aria-label="隐藏金额">${icon(state.settings.amountsHidden ? 'eyeOff' : 'eye')}</button>
      <span>净资产</span><strong>${formatMoney(summary.net, hidden)}</strong>
      <div class="worth-split"><div><span>总资产</span><b>${formatMoney(summary.assets, hidden)}</b></div><div><span>总负债</span><b>${formatMoney(summary.debts, hidden)}</b></div></div>
      <div class="asset-composition"><div class="asset-ring" style="background:conic-gradient(var(--primary) 0 ${Math.max(8,100-summary.debts/assetTotal*100)}%, var(--coral) 0 100%)"><i><span>资产构成</span><b>${hidden?'••••':Math.round(summary.assets).toLocaleString('zh-CN')}</b></i></div><div class="asset-legend"><span><i style="background:var(--primary)"></i>资产账户</span><span><i style="background:var(--coral)"></i>负债账户</span></div></div>
    </section>
    <section class="section"><div class="section-head"><div><span class="section-title">资产配置</span><span class="asset-tag-caption">按账户用途分类 · 占可统计账户资产</span></div></div><div class="asset-tag-grid">${tagStats.map(item=>`<div class="asset-tag-card"><div class="asset-tag-head"><span><i style="background:${item.color}"></i>${item.name}</span><b>${item.ratio.toFixed(1)}%</b></div><strong>${formatMoney(item.amount,hidden)}</strong><div class="asset-tag-track"><i style="width:${Math.min(100,item.ratio)}%;background:${item.color}"></i></div></div>`).join('')}</div></section>
    <section class="section"><div class="section-head"><span class="section-title">资产账户</span><strong class="asset-total">${formatMoney(summary.assets, hidden)}</strong></div><div class="account-card-list">${assetAccounts.map(account => accountCard(account, hidden)).join('')}</div></section>
    ${debtAccounts.length ? `<section class="section"><div class="section-head"><span class="section-title">负债账户</span><strong class="debt-total">${formatMoney(summary.debts, hidden)}</strong></div><div class="account-card-list debt-list">${debtAccounts.map(account => accountCard(account, hidden)).join('')}</div></section>` : ''}`;
}

function accountCard(account, hidden) {
  const value = accountDisplayValue(state, account);
  const availableCredit = Math.max(0, Number(account.limit || 0) - Math.abs(value));
  const txs = state.transactions.filter(tx => tx.accountId === account.id || tx.toAccountId === account.id);
  const income = txs.filter(tx => tx.type === 'income' || (tx.type === 'transfer' && tx.toAccountId === account.id)).reduce((sum,tx)=>sum+tx.amount,0);
  const expense = txs.filter(tx => tx.type === 'expense' || (tx.type === 'transfer' && tx.accountId === account.id)).reduce((sum,tx)=>sum+tx.amount,0);
  return `<button class="account-wide ${account.kind==='liability'?'debt':''}" data-account-id="${account.id}"><span class="account-symbol" style="background:${account.color}">${escapeHtml(account.name.slice(0,1))}</span><span class="account-wide-main"><span>${escapeHtml(account.name)}</span><strong>${formatMoney(value, hidden)}</strong></span><span class="account-flow"><small>${account.kind==='liability'?'额度':'流入'} ${account.kind==='liability'?formatMoney(account.limit||0,hidden):formatMoney(income,hidden)}</small><small>${account.kind==='liability'?'可用余额':'流出'} ${account.kind==='liability'?formatMoney(availableCredit,hidden):formatMoney(expense,hidden)}</small></span>${icon('right','chevron')}</button>`;
}

function openAccountTransactions(accountId) {
  const account = accountById(accountId);
  if (!account) return;
  const txs = state.transactions.filter(tx => tx.accountId === accountId || tx.toAccountId === accountId);
  const inflow = txs.filter(tx => (tx.accountId === accountId && ['income','borrow','collect'].includes(tx.type)) || (tx.toAccountId === accountId && tx.type === 'transfer')).reduce((sum,tx)=>sum+tx.amount,0);
  const outflow = txs.filter(tx => tx.accountId === accountId && ['expense','lend','repay','transfer'].includes(tx.type)).reduce((sum,tx)=>sum+tx.amount,0);
  openSheet(`<div class="sheet-head"><h2>${escapeHtml(account.name)}流水</h2><button class="close-button" data-action="close">${icon('close')}</button></div>
    <div class="account-detail-hero"><span class="account-symbol large" style="background:${account.color}">${escapeHtml(account.name.slice(0,1))}</span><span>当前${account.kind==='liability'?'欠款':'余额'}</span><strong>${formatMoney(account.kind==='liability'?Math.abs(accountDisplayValue(state,account)):accountDisplayValue(state,account), state.settings.amountsHidden)}</strong><small>账户编辑与新增请前往“我的 → 账户管理”</small></div>
    <div class="card stats-summary"><div><span>流入</span><strong>${formatMoney(inflow, state.settings.amountsHidden)}</strong></div><div><span>流出</span><strong>${formatMoney(outflow, state.settings.amountsHidden)}</strong></div><div><span>账目</span><strong>${txs.length} 笔</strong></div></div>
    <section class="section"><div class="section-head"><span class="section-title">消费与资金明细</span></div>${transactionRows(txs)}</section>`, true);
}

function menuRow(iconName, title, subtitle, action, suffix = icon('right','chevron')) {
  return `<button class="menu-row" data-action="${action}"><span class="menu-icon">${icon(iconName)}</span><span class="menu-text"><strong>${title}</strong><span>${subtitle}</span></span>${suffix}</button>`;
}

function renderMine() {
  const assets = getAssetSummary(state);
  pageEl.innerHTML = `
    <div class="card profile-card"><div class="profile-mark">栖</div><div class="profile-copy"><strong>本地账本</strong><span>数据仅保存在这台设备 · ${state.transactions.length} 笔账目</span></div><span>${formatMoney(assets.net, state.settings.amountsHidden)}</span></div>
    <section class="section"><div class="section-head"><span class="section-title">账本、账户与分类</span></div><div class="card menu-card">
      ${menuRow('book','账本管理',`${state.ledgers.length} 个账本`,'ledgers')}
      ${menuRow('wallet','账户管理',`${state.accounts.length} 个账户`,'accounts')}
      ${menuRow('tag','分类管理',`${activeLedger().name} · ${categoriesForLedger().filter(c=>c.parentId).length} 个二级分类`,'categories')}
      ${menuRow('chart','预算管理',`${(state.budgets||[]).filter(b=>b.ledgerId===state.settings.activeLedgerId).length} 项 · 按账本独立设置`,'budgets')}
    </div></section>
    <section class="section"><div class="section-head"><span class="section-title">通用设置</span></div><div class="card menu-card">
      ${menuRow('moon','深色模式','手动切换应用外观','dark',`<span class="toggle ${state.settings.dark?'on':''}"></span>`)}
      ${menuRow('eye','金额隐私','首页与资产页一键遮挡','privacy')}
      ${menuRow('settings','本地数据','重置数据或载入体验数据','data')}
      ${menuRow('info','关于栖账','离线记账 MVP · V1.0','about')}
    </div></section>`;
}

function render() {
  header(); nav();
  if (activePage === 'home') renderHome();
  if (activePage === 'transactions') renderTransactions();
  if (activePage === 'stats') renderStats();
  if (activePage === 'assets') renderAssets();
  if (activePage === 'mine') renderMine();
  window.scrollTo({ top: 0, behavior: 'instant' });
}

function openSheet(html, full = false) {
  sheetRoot.innerHTML = `<div class="overlay" data-action="close-overlay"><div class="sheet ${full ? 'full' : ''}" role="dialog" aria-modal="true"><div class="sheet-handle"></div>${html}</div></div>`;
}

function closeSheet() { sheetRoot.innerHTML = ''; entryDraft = null; editingTransactionId = null; }

function openLedgerSheet() {
  openSheet(`<div class="sheet-head"><h2>切换账本</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="card menu-card">${state.ledgers.map(l=>`<button class="menu-row" data-ledger-id="${l.id}"><span class="account-symbol" style="background:${l.color}">${l.icon}</span><span class="menu-text"><strong>${escapeHtml(l.name)}</strong><span>${state.transactions.filter(tx=>tx.ledgerId===l.id).length} 笔账目</span></span>${l.id===state.settings.activeLedgerId?'✓':icon('right')}</button>`).join('')}</div><button class="action-button" style="width:100%;margin-top:14px" data-action="new-ledger">创建新账本</button>`);
}

function dateInputValue(value = new Date()) {
  const d = new Date(value); const offset = d.getTimezoneOffset();
  return new Date(d.getTime() - offset*60000).toISOString().slice(0,16);
}

function defaultDraft(type = 'expense', tx = null) {
  if (tx) {
    const cat = categoryById(tx.categoryId);
    return { ...tx, expression: String(tx.amount), primaryCategoryId: cat?.parentId || '', categoryId: tx.categoryId || '' };
  }
  const available = state.accounts.filter(a => a.kind === (type === 'expense' ? 'asset' : 'asset'));
  const primary = categoriesForLedger().find(c => c.type === (type === 'income' ? 'income' : 'expense') && !c.parentId);
  return { type, expression: '', amount: 0, accountId: available[0]?.id || state.accounts[0]?.id || '', toAccountId: '', primaryCategoryId: primary?.id || '', categoryId: '', note: '', counterparty: '', date: dateInputValue(), dueDate: '', ledgerId: state.settings.activeLedgerId, includeInAssets: true };
}

function evaluateExpression(expression) {
  if (!expression) return 0;
  const clean = expression.replace(/×/g,'*').replace(/÷/g,'/');
  if (!/^[0-9.+\-*/() ]+$/.test(clean) || /\/\s*0(?:\.0+)?(?:\D|$)/.test(clean)) return NaN;
  try { const value = Function(`"use strict";return (${clean})`)(); return Number.isFinite(value) ? Math.round(value*100)/100 : NaN; } catch { return NaN; }
}

function updateCalculatorDisplay() {
  if (!entryDraft) return;
  const expressionEl = sheetRoot.querySelector('.expression');
  const amountEl = sheetRoot.querySelector('.amount-display');
  const value = evaluateExpression(entryDraft.expression);
  if (expressionEl) expressionEl.textContent = entryDraft.expression || '请输入金额';
  if (amountEl) amountEl.textContent = Number.isFinite(value) && value ? formatMoney(value) : '¥0';
}

function childCategoryButtons(parentId, selectedId = '') {
  return categoriesForLedger(entryDraft?.ledgerId).filter(category => category.parentId === parentId).map(category => `<button class="child-cat ${selectedId===category.id?'active':''}" data-category-id="${category.id}"><span class="cat-orb">${category.icon}</span><span>${escapeHtml(category.name)}</span></button>`).join('');
}

function updatePrimaryCategorySelection() {
  if (!entryDraft) return;
  sheetRoot.querySelectorAll('[data-primary-cat]').forEach(button => button.classList.toggle('active', button.dataset.primaryCat === entryDraft.primaryCategoryId));
  const childList = sheetRoot.querySelector('.child-cats');
  if (childList) childList.innerHTML = childCategoryButtons(entryDraft.primaryCategoryId, entryDraft.categoryId);
}

function updateChildCategorySelection() {
  if (!entryDraft) return;
  sheetRoot.querySelectorAll('[data-category-id]').forEach(button => button.classList.toggle('active', button.dataset.categoryId === entryDraft.categoryId));
}

function entryContentHtml() {
  const draft = entryDraft;
  const categoryType = draft.type === 'income' ? 'income' : 'expense';
  const primary = categoriesForLedger(draft.ledgerId).filter(c => c.type === categoryType && !c.parentId);
  if (!primary.some(c => c.id === draft.primaryCategoryId)) draft.primaryCategoryId = primary[0]?.id || '';
  const needsCategory = ['expense','income'].includes(draft.type);
  const needsToAccount = draft.type === 'transfer';
  const needsCounterparty = ['lend','borrow'].includes(draft.type);
  const value = evaluateExpression(draft.expression);
  return `${needsCategory ? `<div class="category-panel"><div class="primary-cats">${primary.map(cat=>`<button class="primary-cat ${draft.primaryCategoryId===cat.id?'active':''}" data-primary-cat="${cat.id}">${cat.name}</button>`).join('')}</div><div class="child-cats">${childCategoryButtons(draft.primaryCategoryId, draft.categoryId)}</div></div>` : ''}
    <div class="card form-card">
      <div class="form-row"><label>账户</label><select id="draftAccount">${state.accounts.filter(a=>draft.type==='expense'||draft.type==='income' ? true : a.kind==='asset').map(a=>`<option value="${a.id}" ${draft.accountId===a.id?'selected':''}>${escapeHtml(a.name)}</option>`).join('')}</select></div>
      ${needsToAccount ? `<div class="form-row"><label>转入</label><select id="draftToAccount"><option value="">请选择账户</option>${state.accounts.filter(a=>a.id!==draft.accountId).map(a=>`<option value="${a.id}" ${draft.toAccountId===a.id?'selected':''}>${escapeHtml(a.name)}</option>`).join('')}</select></div>`:''}
      ${needsCounterparty ? `<div class="form-row"><label>对象</label><input id="draftCounterparty" value="${escapeHtml(draft.counterparty)}" placeholder="可选填写" /></div><div class="form-row"><label>到期日</label><input id="draftDueDate" type="date" value="${draft.dueDate ? localDateKey(draft.dueDate):''}" /></div>`:''}
      <div class="form-row"><label>日期</label><input id="draftDate" type="datetime-local" value="${dateInputValue(draft.date)}" /></div>
      <div class="form-row"><label>备注</label><input id="draftNote" value="${escapeHtml(draft.note)}" placeholder="写点说明（可选）" /></div>
    </div>
    <div class="amount-panel"><div class="expression">${escapeHtml(draft.expression || '请输入金额')}</div><div class="amount-display">${Number.isFinite(value) && value ? formatMoney(value) : '¥0'}</div>
      <div class="calculator">
        ${['7','8','9','÷','4','5','6','×','1','2','3','-','0','.','⌫','+'].map(key=>`<button class="key ${['÷','×','-','+'].includes(key)?'op':''}" data-key="${key}">${key==='⌫'?icon('backspace'):key}</button>`).join('')}
        <button class="key wide" data-key="clear">清空</button><button class="key submit wide" data-action="save-entry">${editingTransactionId?'保存修改':'保存账目'}</button>
      </div>
    </div>`;
}

function renderEntrySheet() {
  const draft = entryDraft;
  const html = `<div class="sheet-head"><h2>${editingTransactionId ? '编辑账目' : '记一笔'}</h2><button class="close-button" data-action="close">${icon('close')}</button></div>
    <div class="type-tabs three">${['expense','income','transfer'].map(type=>`<button class="type-tab ${draft.type===type?'active':''}" data-draft-type="${type}">${typeLabels[type]}</button>`).join('')}</div>
    <div class="entry-content">${entryContentHtml()}</div>`;
  openSheet(html, true);
}

function updateEntryTypeSelection() {
  if (!entryDraft) return;
  sheetRoot.querySelectorAll('[data-draft-type]').forEach(button => button.classList.toggle('active', button.dataset.draftType === entryDraft.type));
  const content = sheetRoot.querySelector('.entry-content');
  if (content) content.innerHTML = entryContentHtml();
}

function openEntry(type = 'expense', transaction = null) {
  editingTransactionId = transaction?.id || null;
  entryDraft = defaultDraft(type, transaction);
  renderEntrySheet();
}

function syncDraftFields() {
  if (!entryDraft) return;
  const val = id => sheetRoot.querySelector(`#${id}`)?.value;
  if (val('draftAccount') !== undefined) entryDraft.accountId = val('draftAccount');
  if (val('draftToAccount') !== undefined) entryDraft.toAccountId = val('draftToAccount');
  if (val('draftCounterparty') !== undefined) entryDraft.counterparty = val('draftCounterparty');
  if (val('draftDueDate') !== undefined) entryDraft.dueDate = val('draftDueDate');
  if (val('draftDate')) entryDraft.date = new Date(val('draftDate')).toISOString();
  if (val('draftNote') !== undefined) entryDraft.note = val('draftNote');
}

function saveEntry() {
  syncDraftFields();
  const amount = evaluateExpression(entryDraft.expression);
  if (!Number.isFinite(amount) || amount <= 0) return showToast('请输入有效金额');
  if (['expense','income'].includes(entryDraft.type) && !entryDraft.categoryId) return showToast('请选择二级分类');
  if (entryDraft.type === 'transfer' && !entryDraft.toAccountId) return showToast('请选择转入账户');
  const tx = { id: editingTransactionId || uid('tx'), type: entryDraft.type, amount, accountId: entryDraft.accountId, toAccountId: entryDraft.toAccountId || undefined, categoryId: entryDraft.categoryId || undefined, note: entryDraft.note?.trim() || '', counterparty: entryDraft.counterparty?.trim() || '', date: entryDraft.date || new Date().toISOString(), dueDate: entryDraft.dueDate || undefined, ledgerId: entryDraft.ledgerId, includeInAssets: Boolean(entryDraft.includeInAssets) };
  const wasEditing = Boolean(editingTransactionId);
  if (editingTransactionId) {
    state.transactions = state.transactions.map(item => item.id === editingTransactionId ? tx : item);
  } else {
    state.transactions.push(tx);
  }
  if (['lend','borrow'].includes(tx.type)) {
    const existing = state.loans.find(l => l.transactionId === tx.id);
    const loan = { id: existing?.id || uid('loan'), transactionId: tx.id, type: tx.type, principal: tx.amount, settled: existing?.settled || 0, counterparty: tx.counterparty, dueDate: tx.dueDate, ledgerId: tx.ledgerId, includeInAssets: tx.includeInAssets, status: 'open' };
    state.loans = existing ? state.loans.map(l => l.id === existing.id ? loan : l) : [...state.loans, loan];
  }
  persist(); closeSheet(); activePage = 'home'; viewingMonth = new Date(tx.date); render(); showToast(wasEditing ? '修改已保存' : '账目已保存');
}

function openTransactionDetail(id) {
  const tx = state.transactions.find(item => item.id === id); if (!tx) return;
  const account = accountById(tx.accountId); const target = accountById(tx.toAccountId);
  openSheet(`<div class="sheet-head"><h2>账目详情</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="detail-amount"><div class="orb">${txIcon(tx)}</div><strong>${tx.type==='income'?'+':tx.type==='expense'?'-':''}${formatMoney(tx.amount)}</strong><span>${typeLabels[tx.type] || '账目'}</span></div><div class="card detail-list">
    <div class="detail-row"><span>分类</span><strong>${escapeHtml(txLabel(tx))}</strong></div>
    <div class="detail-row"><span>账户</span><strong>${escapeHtml(account?.name || '已删除账户')}${target ? ` → ${escapeHtml(target.name)}`:''}</strong></div>
    <div class="detail-row"><span>日期</span><strong>${new Date(tx.date).toLocaleString('zh-CN',{month:'long',day:'numeric',hour:'2-digit',minute:'2-digit'})}</strong></div>
    <div class="detail-row"><span>账本</span><strong>${escapeHtml(state.ledgers.find(l=>l.id===tx.ledgerId)?.name || '')}</strong></div>
    <div class="detail-row"><span>备注</span><strong>${escapeHtml(tx.note || '—')}</strong></div>
  </div><div class="detail-actions"><button class="action-button secondary" data-edit-tx="${tx.id}">${icon('edit')} 编辑</button><button class="action-button danger" data-delete-tx="${tx.id}">${icon('trash')} 删除</button></div>`);
}

function deleteTransaction(id) {
  const tx = state.transactions.find(t=>t.id===id); if (!tx) return;
  if (!confirm('确认删除这笔账目吗？删除后无法恢复。')) return;
  state.transactions = state.transactions.filter(t=>t.id!==id);
  state.loans = state.loans.filter(l=>l.transactionId!==id);
  persist(); closeSheet(); render(); showToast('账目已永久删除');
}

function openAccounts() {
  editingAccountId = null;
  openSheet(`<div class="sheet-head"><h2>账户管理</h2><button class="close-button" data-action="close">${icon('close')}</button></div><p class="sort-tip">按住左侧手柄拖动排序，资产页和记账账户顺序会自动同步。</p><div class="manager-list reorder-list" data-reorder-container="accounts">${state.accounts.map(account=>`<div class="manager-row sortable-row" data-reorder-id="${account.id}"><button class="drag-handle" data-drag-handle aria-label="拖动${escapeHtml(account.name)}排序">${icon('grip')}</button><button class="manager-row-main" data-edit-account="${account.id}"><span class="account-symbol" style="background:${account.color}">${escapeHtml(account.name.slice(0,1))}</span><span class="menu-text"><strong>${escapeHtml(account.name)}</strong><span>${account.kind==='liability'?'负债账户':`资产账户 · ${escapeHtml(account.assetTag||assetTagOptions[0])}`} · ${state.transactions.filter(tx=>tx.accountId===account.id||tx.toAccountId===account.id).length} 笔关联账目</span></span><strong>${formatMoney(accountDisplayValue(state,account))}</strong>${icon('right','chevron')}</button></div>`).join('')}</div><button class="action-button manager-add" data-action="new-account">${icon('plus')} 新增账户</button>`);
}

function openAccountForm(id = null) {
  editingAccountId = id;
  const account = state.accounts.find(item=>item.id===id);
  const isLiability = account?.kind === 'liability';
  const currentValue = account ? getAccountBalance(state, account.id) : 0;
  const availableBalance = roundMoney(account ? (isLiability ? Math.max(0, Number(account.limit || 0) - currentValue) : currentValue) : 0);
  openSheet(`<div class="sheet-head"><h2>${account?'编辑账户':'新增账户'}</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="card form-card">
    <div class="form-row"><label>名称</label><input id="accountName" value="${escapeHtml(account?.name||'')}" placeholder="例如：工资卡" /></div>
    <div class="form-row"><label>类型</label><select id="accountKind" ${account?'disabled':''}><option value="asset" ${!isLiability?'selected':''}>资产账户</option><option value="liability" ${isLiability?'selected':''}>负债账户</option></select></div>
    <div class="form-row"><label>可用余额</label><input id="accountAvailable" type="number" inputmode="decimal" step="0.01" value="${availableBalance.toFixed(2)}" /></div>
    <div class="form-row asset-tag-row ${isLiability?'hidden':''}"><label>资产标签</label><select id="accountAssetTag">${assetTagOptions.map(tag=>`<option value="${tag}" ${(account?.assetTag||assetTagOptions[0])===tag?'selected':''}>${tag}</option>`).join('')}</select></div>
    <div class="form-row liability-only ${!isLiability?'hidden':''}"><label>总额度</label><input id="accountLimit" type="number" inputmode="decimal" value="${account?.limit||10000}" /></div>
    <div class="form-row liability-only ${!isLiability?'hidden':''}"><label>账单日</label><input id="accountBillDay" type="number" min="1" max="28" value="${account?.billDay||1}" /></div>
    <div class="form-row liability-only ${!isLiability?'hidden':''}"><label>还款日</label><input id="accountDueDay" type="number" min="1" max="28" value="${account?.dueDay||9}" /></div>
    <div class="form-row"><label>资产统计</label><select id="accountInclude"><option value="yes" ${account?.includeInAssets!==false?'selected':''}>计入</option><option value="no" ${account?.includeInAssets===false?'selected':''}>不计入</option></select></div>
  </div>${account?'<p class="subtle" style="line-height:1.7;padding:4px 6px">可用余额可随时核对修改；保存差额会生成余额校准记录，不会改动历史账目。</p>':''}<div class="detail-actions"><button class="action-button" data-action="save-account">保存账户</button>${account?`<button class="action-button danger" data-delete-account="${account.id}">删除账户</button>`:'<button class="action-button secondary" data-action="accounts">返回列表</button>'}</div>`);
}

function saveAccount() {
  const name = sheetRoot.querySelector('#accountName')?.value.trim();
  const kind = sheetRoot.querySelector('#accountKind')?.value;
  const available = roundMoney(Number(sheetRoot.querySelector('#accountAvailable')?.value || 0));
  if (!name) return showToast('请填写账户名称');
  if (!Number.isFinite(available)) return showToast('请填写有效的可用余额');
  const existing = state.accounts.find(item=>item.id===editingAccountId);
  const limit = kind==='liability'?roundMoney(Math.max(0,Number(sheetRoot.querySelector('#accountLimit')?.value||0))):undefined;
  if (kind === 'liability' && (available < 0 || available > limit)) return showToast('可用余额需在 0 与总额度之间');
  const desiredBalance = roundMoney(kind === 'liability' ? Math.max(0, limit - available) : available);
  const account = { id: existing?.id||uid('account'), name, type: kind==='liability'?'credit':existing?.type||'custom', kind, assetTag: kind==='asset'?(sheetRoot.querySelector('#accountAssetTag')?.value||assetTagOptions[0]):undefined, initialBalance: existing ? existing.initialBalance : kind==='asset'?desiredBalance:0, initialDebt: existing ? existing.initialDebt : kind==='liability'?desiredBalance:0, limit, billDay: kind==='liability'?Math.min(28,Math.max(1,Number(sheetRoot.querySelector('#accountBillDay')?.value||1))):undefined, dueDay: kind==='liability'?Math.min(28,Math.max(1,Number(sheetRoot.querySelector('#accountDueDay')?.value||9))):undefined, color: existing?.color||['#24a66f','#3d8fd1','#f08070','#eaa13f'][state.accounts.length%4], includeInAssets: sheetRoot.querySelector('#accountInclude')?.value!=='no', order: existing?.order||state.accounts.length+1 };
  state.accounts = existing ? state.accounts.map(item=>item.id===existing.id?account:item) : [...state.accounts, account];
  if (existing) {
    const currentBalance = getAccountBalance(state, account.id);
    const adjustment = roundMoney(desiredBalance - currentBalance);
    if (Math.abs(adjustment) >= 0.01) state.transactions.push({ id: uid('tx'), type: 'adjustment', amount: adjustment, accountId: account.id, note: '账户可用余额校准', date: new Date().toISOString(), ledgerId: state.settings.activeLedgerId });
  }
  persist(); openAccounts(); render(); showToast(existing?'账户与可用余额已更新':'账户已创建');
}

function deleteAccount(id) {
  const account = state.accounts.find(item=>item.id===id); if(!account) return;
  const used = state.transactions.some(tx=>tx.accountId===id||tx.toAccountId===id);
  if (used) return showToast('该账户存在历史账目，请先迁移账目');
  if (state.accounts.length <= 1) return showToast('至少保留一个账户');
  if (!confirm(`确认永久删除“${account.name}”吗？`)) return;
  state.accounts = state.accounts.filter(item=>item.id!==id); persist(); openAccounts(); render(); showToast('账户已删除');
}

function openMessages() {
  state.messages.forEach(m=>m.read=true); persist();
  const reminders = currentReminders();
  const reminderRows = reminders.map(reminder => {
    const days = Math.ceil((reminder.due - new Date()) / 86400000);
    return `<div class="reminder-row"><span class="reminder-icon">${icon('bell')}</span><span class="reminder-content"><strong>${escapeHtml(reminder.title)}</strong><span>${escapeHtml(reminder.subtitle)} · ${reminder.due.getMonth()+1}月${reminder.due.getDate()}日 · ${days < 0 ? `逾期 ${Math.abs(days)} 天` : `${days} 天后`}</span></span><span class="reminder-amount"><strong>${formatMoney(reminder.amount, state.settings.amountsHidden)}</strong><span>待处理</span></span></div>`;
  }).join('');
  const storedRows = state.messages.filter(message => !reminders.some(reminder => reminder.id === message.relatedId)).map(m=>`<div class="reminder-row"><span class="reminder-icon">${icon('bell')}</span><span class="reminder-content"><strong>${escapeHtml(m.title)}</strong><span>${escapeHtml(m.body)} · ${new Date(m.createdAt).toLocaleDateString('zh-CN')}</span></span><span class="subtle">${m.handled?'已处理':'待处理'}</span></div>`).join('');
  openSheet(`<div class="sheet-head"><h2>消息中心</h2><button class="close-button" data-action="close">${icon('close')}</button></div>${reminderRows || storedRows ? `<div class="card reminder-card">${reminderRows}${storedRows}</div>` : '<div class="empty"><div class="empty-orb">'+icon('bell')+'</div><strong>暂无消息</strong><p>还款提醒会统一显示在这里。</p></div>'}`);
  header();
}

function openCalendarSheet() {
  const today = new Date();
  const year = today.getFullYear();
  const month = today.getMonth();
  const firstDay = new Date(year, month, 1).getDay();
  const dayCount = new Date(year, month + 1, 0).getDate();
  const monthTxs = monthTransactions(state, state.settings.activeLedgerId, today);
  const cells = [];
  for (let i = 0; i < firstDay; i++) cells.push('<span class="calendar-day blank"></span>');
  for (let day = 1; day <= dayCount; day++) {
    const dayTxs = monthTxs.filter(tx => new Date(tx.date).getDate() === day);
    const totals = incomeExpenseSummary(dayTxs);
    const isToday = day === today.getDate();
    cells.push(`<button class="calendar-day ${isToday?'today':''}" data-calendar-day="${day}"><strong>${day}</strong>${totals.expense ? `<small>-${Math.round(totals.expense)}</small>` : totals.income ? `<small class="income">+${Math.round(totals.income)}</small>` : '<i></i>'}</button>`);
  }
  openSheet(`<div class="sheet-head"><h2>${year}年 ${month+1}月</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="calendar-week"><span>日</span><span>一</span><span>二</span><span>三</span><span>四</span><span>五</span><span>六</span></div><div class="calendar-grid">${cells.join('')}</div><div class="card calendar-summary"><span>本月共 ${monthTxs.length} 笔账目</span><button class="text-button" data-action="search">查看全部明细</button></div>`);
}

function openMonthPicker(year = viewingMonth.getFullYear()) {
  const months = Array.from({length:12}, (_,index) => {
    const date = new Date(year, index, 1);
    const count = monthTransactions(state, state.settings.activeLedgerId, date).length;
    const active = viewingMonth.getFullYear() === year && viewingMonth.getMonth() === index;
    return `<button class="month-choice ${active?'active':''}" data-pick-month="${year}-${String(index+1).padStart(2,'0')}"><strong>${index+1}月</strong><span>${count} 笔</span></button>`;
  }).join('');
  openSheet(`<div class="sheet-head"><h2>选择明细月份</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="month-picker-head"><button data-picker-year="${year-1}" aria-label="上一年">${icon('right','flip')}</button><strong>${year}年</strong><button data-picker-year="${year+1}" aria-label="下一年">${icon('right')}</button></div><div class="month-picker-grid">${months}</div>`);
  const previousIcon = sheetRoot.querySelector('.month-picker-head .flip');
  if (previousIcon) previousIcon.style.transform = 'rotate(180deg)';
}

function openDataSheet() {
  openSheet(`<div class="sheet-head"><h2>本地数据</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="card menu-card"><button class="menu-row" data-action="excel-export"><span class="menu-icon">⇩</span><span class="menu-text"><strong>导出 Excel 文件</strong><span>将账本、账目、账户、分类和预算保存为 .xlsx</span></span>${icon('right')}</button><button class="menu-row" data-action="excel-import"><span class="menu-icon">⇧</span><span class="menu-text"><strong>导入 Excel 文件</strong><span>选择 .xlsx 工作簿，预览后合并到本机</span></span>${icon('right')}</button><button class="menu-row" data-action="load-demo"><span class="menu-icon">${icon('layers')}</span><span class="menu-text"><strong>载入体验数据</strong><span>覆盖当前数据，用于体验完整界面</span></span>${icon('right')}</button><button class="menu-row" data-action="clear-data"><span class="menu-icon" style="color:var(--danger);background:var(--coral-soft)">${icon('trash')}</span><span class="menu-text"><strong>清空全部数据</strong><span>永久删除当前设备内的所有账务数据</span></span>${icon('right')}</button></div><input id="excelFileInput" type="file" accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" hidden /><p class="subtle" style="line-height:1.7;padding:4px 6px">数据仅保存在本机。Excel 文件包含完整财务信息，请妥善保管；导入时采用合并方式，不会清空其他本地数据。</p>`);
}

function openBudgetManager() {
  const ledgerId = state.settings.activeLedgerId;
  const categories = categoriesForLedger(ledgerId).filter(item => !item.parentId && item.type === 'expense').sort((a,b)=>(a.order||0)-(b.order||0));
  const current = new Map((state.budgets||[]).filter(item=>item.ledgerId===ledgerId).map(item=>[item.categoryId,item.amount]));
  openSheet(`<div class="sheet-head"><h2>分类预算 · ${escapeHtml(activeLedger().name)}</h2><button class="close-button" data-action="close">${icon('close')}</button></div><p class="subtle budget-note">设置每月一级支出分类额度；各账本预算独立，按自然月统计，留空或填 0 表示不设预算。</p><div class="card budget-form-list">${categories.map(category=>`<label class="budget-form-row"><span class="budget-category-icon">${escapeHtml(category.icon||'•')}</span><span class="budget-category-name">${escapeHtml(category.name)}</span><span class="budget-input-wrap"><small>¥</small><input inputmode="decimal" type="number" min="0" step="0.01" placeholder="不设" value="${current.has(category.id)?escapeHtml(current.get(category.id)):''}" data-budget-category="${category.id}" /></span></label>`).join('')}</div><button class="action-button budget-save" data-action="save-budgets">保存预算</button>`);
}

function saveBudgets() {
  const ledgerId = state.settings.activeLedgerId;
  const amounts = new Map([...sheetRoot.querySelectorAll('[data-budget-category]')].map(input=>[input.dataset.budgetCategory, Number(input.value || 0)]));
  if ([...amounts.values()].some(value=>!Number.isFinite(value)||value<0)) return showToast('预算金额不能小于 0');
  const retained = (state.budgets||[]).filter(item=>item.ledgerId!==ledgerId);
  for (const [categoryId, rawAmount] of amounts) {
    const amount = roundMoney(rawAmount);
    if (!amount) continue;
    const previous = (state.budgets||[]).find(item=>item.ledgerId===ledgerId&&item.categoryId===categoryId);
    retained.push({ id: previous?.id || uid('budget'), ledgerId, categoryId, amount });
  }
  state.budgets = retained; persist(); closeSheet(); render(); showToast('预算已保存');
}

function openExcelImportPreview(fileName, plan) {
  pendingExcelImport = plan;
  const records = plan.counts.map(item=>`<div class="detail-row"><span>${escapeHtml(item.name)}</span><strong>${item.count} 条</strong></div>`).join('');
  const warnings = plan.warnings.length ? `<div class="import-warnings"><strong>需要留意</strong>${plan.warnings.slice(0,6).map(item=>`<p>${escapeHtml(item)}</p>`).join('')}${plan.warnings.length>6?`<p>另有 ${plan.warnings.length-6} 条...</p>`:''}</div>` : '';
  openSheet(`<div class="sheet-head"><h2>导入预览</h2><button class="close-button" data-action="close-import">${icon('close')}</button></div><p class="subtle">${escapeHtml(fileName)} · 将合并到当前设备</p><div class="card detail-list">${records||'<div class="empty"><p>没有识别到可导入的数据</p></div>'}</div>${warnings}<p class="subtle import-hint">同编号记录将以文件内容更新；其他本地记录保留。确认后立即写入本机。</p><div class="detail-actions"><button class="action-button" data-action="confirm-excel-import" ${plan.total?'':'disabled'}>确认合并导入</button><button class="action-button secondary" data-action="close-import">取消</button></div>`);
}

function prepareExcelImport(workbook) {
  const aliases = { '账目':'transactions','明细':'transactions','交易记录':'transactions','Transactions':'transactions','账户':'accounts','账户列表':'accounts','Accounts':'accounts','账本':'ledgers','账本列表':'ledgers','Ledgers':'ledgers','分类':'categories','分类列表':'categories','Categories':'categories','预算':'budgets','Budgets':'budgets','借贷':'loans','Loans':'loans','消息':'messages','Messages':'messages' };
  const incoming = {};
  for (const [sheet, rows] of Object.entries(workbook.tables)) {
    const key = aliases[sheet]; if (!key) continue;
    incoming[key] = [...(incoming[key]||[]), ...rows];
  }
  for (const [key, rows] of Object.entries(incoming)) rows.forEach(item=>{ if(!item.id) item.id=uid(`import-${key}`); });
  const warnings = [];
  const mergedLedgers = [...state.ledgers, ...(incoming.ledgers||[])];
  const mergedAccounts = [...state.accounts, ...(incoming.accounts||[])];
  const mergedCategories = [...state.categories, ...(incoming.categories||[])];
  const resolveByName = (list,name,ledgerId) => list.find(item=>item.name===name&&(!ledgerId||!item.ledgerId||item.ledgerId===ledgerId));
  const normalized = {};
  for (const [key, rows] of Object.entries(incoming)) {
    normalized[key] = rows.map((row,index)=>{
      const item = { ...row, id: row.id || uid(`import-${key}`) };
      if (key === 'transactions') {
        if (!item.ledgerId) item.ledgerId = resolveByName(mergedLedgers,row.ledgerName)?.id || state.settings.activeLedgerId;
        if (!item.accountId) item.accountId = resolveByName(mergedAccounts,row.accountName)?.id;
        if (!item.toAccountId && row.toAccountName) item.toAccountId = resolveByName(mergedAccounts,row.toAccountName)?.id;
        if (!item.categoryId && row.categoryName) item.categoryId = resolveByName(mergedCategories,row.categoryName,item.ledgerId)?.id;
        if (typeof item.type === 'string') item.type = ({'支出':'expense','收入':'income','转账':'transfer','expense':'expense','income':'income','transfer':'transfer'})[item.type] || item.type;
        if (!item.date && row.date) item.date = row.date;
        const validTypes = ['expense','income','transfer','lend','borrow','repay','collect','adjustment'];
        const accountExists = id => mergedAccounts.some(account=>account.id===id);
        const categoryExists = id => mergedCategories.some(category=>category.id===id && category.ledgerId===item.ledgerId);
        const needsCategory = ['expense','income'].includes(item.type);
        if (!validTypes.includes(item.type) || !Number.isFinite(Number(item.amount)) || Number(item.amount)<=0 || !item.accountId || !accountExists(item.accountId) || !item.ledgerId || !mergedLedgers.some(ledger=>ledger.id===item.ledgerId) || !item.date || (needsCategory&&(!item.categoryId||!categoryExists(item.categoryId))) || (item.type==='transfer'&&(!item.toAccountId||!accountExists(item.toAccountId)))) {
          warnings.push(`账目第 ${index+1} 行缺少有效类型、金额、账户、日期或分类，已跳过。`); return null;
        }
        item.amount = roundMoney(item.amount);
      }
      if (key === 'budgets') {
        item.ledgerId ||= resolveByName(mergedLedgers,row.ledgerName)?.id || state.settings.activeLedgerId;
        item.categoryId ||= resolveByName(mergedCategories,row.categoryName,item.ledgerId)?.id;
        if (!item.categoryId || !Number.isFinite(Number(item.amount)) || Number(item.amount)<=0) { warnings.push(`预算第 ${index+1} 行信息不完整，已跳过。`); return null; }
        item.amount = roundMoney(item.amount);
      }
      if (key === 'ledgers' && !item.name) { warnings.push(`账本第 ${index+1} 行没有名称，已跳过。`); return null; }
      if (key === 'accounts' && (!item.name || !item.kind)) { warnings.push(`账户第 ${index+1} 行缺少名称或账户性质，已跳过。`); return null; }
      if (key === 'categories') {
        item.ledgerId ||= resolveByName(mergedLedgers,row.ledgerName)?.id || state.settings.activeLedgerId;
        if (!item.name || !item.type || !item.ledgerId) { warnings.push(`分类第 ${index+1} 行缺少名称、类型或所属账本，已跳过。`); return null; }
      }
      if (key === 'accounts' && item.kind === 'asset' && !assetTagOptions.includes(item.assetTag)) item.assetTag = assetTagOptions[0];
      return item;
    }).filter(Boolean);
  }
  if (workbook.settings && Object.keys(workbook.settings).length) normalized.settings = workbook.settings;
  const labelsByKey = { transactions:'账目',accounts:'账户',ledgers:'账本',categories:'分类',budgets:'预算',loans:'借贷',messages:'消息' };
  const counts = Object.entries(normalized).filter(([key,value])=>key!=='settings'&&value.length).map(([key,value])=>({name:labelsByKey[key]||key,count:value.length}));
  return { data: normalized, counts, warnings, total: counts.reduce((sum,item)=>sum+item.count,0) };
}

async function exportExcel() {
  try {
    const bytes = await buildExcelWorkbook(state);
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const filename = `栖账数据-${new Date().toISOString().slice(0,10)}.xlsx`;
    const file = new File([blob], filename, { type: blob.type });
    const nativeExcelFile = window.Capacitor?.Plugins?.ExcelFile;
    if (window.Capacitor?.isNativePlatform?.()) {
      if (!nativeExcelFile?.save) throw new Error('当前 App 版本暂不支持保存 Excel，请更新后重试。');
      let binary='';
      for(let offset=0;offset<bytes.length;offset+=0x8000) binary+=String.fromCharCode(...bytes.subarray(offset,offset+0x8000));
      const result=await nativeExcelFile.save({ fileName: filename, mimeType: blob.type, data: btoa(binary) });
      if (!result?.saved || Number(result.bytesWritten)!==bytes.length) throw new Error('文件未完整写入，请重新导出。');
      showToast('Excel 文件已保存');
      return;
    }
    if (navigator.canShare?.({ files: [file] }) && navigator.share) await navigator.share({ files: [file], title: '栖账 Excel 数据' });
    else { const url=URL.createObjectURL(blob); const link=document.createElement('a'); link.href=url; link.download=filename; link.click(); setTimeout(()=>URL.revokeObjectURL(url),1500); }
    showToast('Excel 文件已准备好');
  } catch (error) { console.warn('Excel export failed',error); showToast('导出失败，请稍后重试'); }
}

async function readExcelFile(file) {
  try {
    const workbook = await readExcelWorkbook(file);
    const plan = prepareExcelImport(workbook);
    openExcelImportPreview(file.name, plan);
  } catch (error) { console.warn('Excel import failed',error); showToast(error.message || '无法读取这个 Excel 文件'); }
}

function commitExcelImport() {
  if (!pendingExcelImport?.total) return;
  const merge = (existing=[], incoming=[], keyField='id') => {
    const byId = new Map(existing.map((item,index)=>[item.id,index]));
    const result = [...existing];
    incoming.forEach(item=>{
      const key=keyField==='budget'?`${item.ledgerId}|${item.categoryId}`:item.id;
      const index=keyField==='budget'?result.findIndex(existingItem=>`${existingItem.ledgerId}|${existingItem.categoryId}`===key):byId.get(key);
      if(index===undefined||index<0){ byId.set(item.id,result.length); result.push(item); } else result[index]=item;
    });
    return result;
  };
  for (const [key,rows] of Object.entries(pendingExcelImport.data)) {
    if (key==='settings') state.settings = { ...state.settings, ...rows };
    else if (Array.isArray(rows)) state[key] = merge(state[key]||[],rows,key==='budgets'?'budget':'id');
  }
  state.budgets ||= [];
  if (!state.ledgers.some(item=>item.id===state.settings.activeLedgerId)) state.settings.activeLedgerId=state.ledgers[0]?.id;
  persist(); pendingExcelImport=null; closeSheet(); render(); showToast('Excel 数据已合并导入');
}

function openAbout() {
  openSheet(`<div class="sheet-head"><h2>关于栖账</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="detail-amount"><div class="profile-mark" style="margin:auto">栖</div><strong style="font-size:22px">栖账</strong><span>本地离线记账 · MVP V1.0</span></div><div class="card detail-list"><div class="detail-row"><span>数据位置</span><strong>仅当前设备</strong></div><div class="detail-row"><span>账号与云端</span><strong>无需账号 · 不上传</strong></div><div class="detail-row"><span>网络依赖</span><strong>无</strong></div></div>`);
}

function openLedgerManager() {
  editingLedgerId = null;
  openSheet(`<div class="sheet-head"><h2>账本管理</h2><button class="close-button" data-action="close">${icon('close')}</button></div><p class="sort-tip">按住左侧手柄拖动排序，账本切换列表会自动同步。</p><div class="manager-list reorder-list" data-reorder-container="ledgers">${state.ledgers.map(ledger=>`<div class="manager-row sortable-row" data-reorder-id="${ledger.id}"><button class="drag-handle" data-drag-handle aria-label="拖动${escapeHtml(ledger.name)}排序">${icon('grip')}</button><button class="manager-row-main" data-edit-ledger="${ledger.id}"><span class="account-symbol" style="background:${ledger.color}">${escapeHtml(ledger.icon)}</span><span class="menu-text"><strong>${escapeHtml(ledger.name)}</strong><span>${state.transactions.filter(tx=>tx.ledgerId===ledger.id).length} 笔账目${ledger.id===state.settings.activeLedgerId?' · 当前使用':''}</span></span>${icon('right','chevron')}</button></div>`).join('')}</div><button class="action-button manager-add" data-action="new-ledger">${icon('plus')} 新增账本</button>`);
}

function openLedgerForm(id = null) {
  editingLedgerId = id;
  const ledger = state.ledgers.find(item=>item.id===id);
  openSheet(`<div class="sheet-head"><h2>${ledger?'编辑账本':'新增账本'}</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="card form-card"><div class="form-row"><label>名称</label><input id="ledgerName" value="${escapeHtml(ledger?.name||'')}" placeholder="例如：日常生活" /></div><div class="form-row"><label>图标文字</label><input id="ledgerIcon" maxlength="2" value="${escapeHtml(ledger?.icon||'')}" placeholder="日" /></div><div class="form-row"><label>主题颜色</label><input id="ledgerColor" type="color" value="${ledger?.color||'#24a66f'}" /></div></div><div class="detail-actions"><button class="action-button" data-action="save-ledger">保存账本</button>${ledger?`<button class="action-button danger" data-delete-ledger="${ledger.id}">删除账本</button>`:'<button class="action-button secondary" data-action="ledgers">返回列表</button>'}</div>`);
}

function saveLedger() {
  const name = sheetRoot.querySelector('#ledgerName')?.value.trim();
  if(!name) return showToast('请填写账本名称');
  const existing = state.ledgers.find(item=>item.id===editingLedgerId);
  const sourceCategories = categoriesForLedger();
  const ledger = { id:existing?.id||uid('ledger'), name, icon:sheetRoot.querySelector('#ledgerIcon')?.value.trim()||name.slice(0,1), color:sheetRoot.querySelector('#ledgerColor')?.value||'#24a66f' };
  state.ledgers = existing ? state.ledgers.map(item=>item.id===existing.id?ledger:item) : [...state.ledgers,ledger];
  if(!existing) {
    state.categories.push(...cloneCategoriesForLedger(sourceCategories, ledger.id));
    state.settings.activeLedgerId=ledger.id;
  }
  persist(); openLedgerManager(); render(); showToast(existing?'账本已更新':'账本已创建');
}

function deleteLedger(id) {
  const ledger = state.ledgers.find(item=>item.id===id); if(!ledger) return;
  if(state.ledgers.length<=1) return showToast('至少保留一个账本');
  if(state.transactions.some(tx=>tx.ledgerId===id)) return showToast('该账本存在账目，请先处理账目');
  if(!confirm(`确认永久删除“${ledger.name}”吗？`)) return;
  state.ledgers=state.ledgers.filter(item=>item.id!==id);
  state.categories=state.categories.filter(category=>category.ledgerId!==id);
  if(state.settings.activeLedgerId===id) state.settings.activeLedgerId=state.ledgers[0].id;
  persist(); openLedgerManager(); render(); showToast('账本已删除');
}

function openCategories() {
  editingCategoryId = null;
  const ledgerCategories = categoriesForLedger();
  const groups = ledgerCategories.filter(cat=>!cat.parentId).map(parent=>`<section class="category-manage-group sortable-category" data-reorder-id="${parent.id}"><div class="category-manage-head"><button class="drag-handle compact" data-drag-handle aria-label="拖动${escapeHtml(parent.name)}排序">${icon('grip')}</button><span class="account-symbol category-symbol">${escapeHtml(parent.icon)}</span><span><strong>${escapeHtml(parent.name)}</strong><small>${parent.type==='income'?'收入':'支出'}一级分类</small></span><button data-new-child="${parent.id}" aria-label="新增${escapeHtml(parent.name)}二级分类">${icon('plus')}</button><button data-edit-category="${parent.id}" aria-label="编辑${escapeHtml(parent.name)}">${icon('edit')}</button></div><div class="category-child-list" data-reorder-container="category-children" data-parent-id="${parent.id}">${ledgerCategories.filter(cat=>cat.parentId===parent.id).map(child=>`<div class="category-child-row" data-reorder-id="${child.id}"><button class="drag-handle compact" data-drag-handle aria-label="拖动${escapeHtml(child.name)}排序">${icon('grip')}</button><button class="category-child-edit" data-edit-category="${child.id}"><span>${child.icon}</span>${escapeHtml(child.name)}${icon('right','chevron')}</button></div>`).join('')}</div></section>`).join('');
  openSheet(`<div class="sheet-head"><h2>分类管理</h2><button class="close-button" data-action="close">${icon('close')}</button></div><p class="sort-tip">当前账本：${escapeHtml(activeLedger().name)}。分类仅用于这个账本，拖动排序会同步到该账本的记账页。</p><div class="category-parent-list" data-reorder-container="category-parents">${groups}</div><button class="action-button manager-add" data-action="new-category">${icon('plus')} 新增一级分类</button>`);
}

function openCategoryForm(id = null, parentId = '') {
  editingCategoryId = id;
  const ledgerCategories = categoriesForLedger();
  const category = ledgerCategories.find(item=>item.id===id);
  const selectedParent = category?.parentId || parentId;
  const currentType = category?.type || ledgerCategories.find(item=>item.id===selectedParent)?.type || 'expense';
  const selectedIcon = category?.icon || '🍀';
  openSheet(`<div class="sheet-head"><h2>${category?'编辑分类':selectedParent?'新增二级分类':'新增一级分类'}</h2><button class="close-button" data-action="close">${icon('close')}</button></div><div class="card form-card"><div class="form-row"><label>名称</label><input id="categoryName" value="${escapeHtml(category?.name||'')}" placeholder="分类名称" /></div><div class="form-row"><label>图标</label><input id="categoryIcon" value="${escapeHtml(selectedIcon)}" maxlength="4" placeholder="🍀" /></div><div class="form-row"><label>收支类型</label><select id="categoryType"><option value="expense" ${currentType==='expense'?'selected':''}>支出</option><option value="income" ${currentType==='income'?'selected':''}>收入</option></select></div><div class="form-row"><label>上级分类</label><select id="categoryParent"><option value="">无（一级分类）</option>${ledgerCategories.filter(item=>!item.parentId).map(parent=>`<option value="${parent.id}" ${selectedParent===parent.id?'selected':''}>${escapeHtml(parent.name)}</option>`).join('')}</select></div></div><div class="icon-picker"><div class="icon-picker-title"><strong>选择图标</strong><span>也可以在上方直接输入</span></div><div class="icon-picker-grid">${categoryIconChoices.map(choice=>`<button class="${choice===selectedIcon?'active':''}" data-category-icon="${choice}" aria-label="选择图标 ${choice}">${choice}</button>`).join('')}</div></div><div class="detail-actions"><button class="action-button" data-action="save-category">保存分类</button>${category?`<button class="action-button danger" data-delete-category="${category.id}">删除分类</button>`:'<button class="action-button secondary" data-action="categories">返回列表</button>'}</div>`);
}

function saveCategory() {
  const name=sheetRoot.querySelector('#categoryName')?.value.trim(); if(!name) return showToast('请填写分类名称');
  const parentId=sheetRoot.querySelector('#categoryParent')?.value||undefined;
  const type=parentId ? state.categories.find(item=>item.id===parentId)?.type : sheetRoot.querySelector('#categoryType')?.value;
  const existing=state.categories.find(item=>item.id===editingCategoryId);
  if(existing && !existing.parentId && parentId) return showToast('已有一级分类不能直接改为二级分类');
  const category={ id:existing?.id||uid('category'), ledgerId:existing?.ledgerId||state.settings.activeLedgerId, name, icon:sheetRoot.querySelector('#categoryIcon')?.value.trim()||'●', type, ...(parentId?{parentId}:{}) };
  state.categories=existing?state.categories.map(item=>item.id===existing.id?category:item):[...state.categories,category];
  persist(); openCategories(); render(); showToast(existing?'分类已更新':'分类已创建');
}

function deleteCategory(id) {
  const category=state.categories.find(item=>item.id===id); if(!category) return;
  if(state.transactions.some(tx=>tx.categoryId===id)) return showToast('该分类存在历史账目，请先迁移账目');
  if(state.categories.some(item=>item.parentId===id)) return showToast('请先删除或迁移二级分类');
  if(!confirm(`确认永久删除“${category.name}”吗？`)) return;
  state.categories=state.categories.filter(item=>item.id!==id); persist(); openCategories(); render(); showToast('分类已删除');
}

function reorderItems(items, orderedIds) {
  const byId = new Map(items.map(item => [item.id, item]));
  const ordered = orderedIds.map(id => byId.get(id)).filter(Boolean);
  const remaining = items.filter(item => !orderedIds.includes(item.id));
  return [...ordered, ...remaining].map((item, index) => ({ ...item, order: index + 1 }));
}

function commitReorder(kind, orderedIds, parentId = '') {
  if (kind === 'accounts') state.accounts = reorderItems(state.accounts, orderedIds);
  if (kind === 'ledgers') state.ledgers = reorderItems(state.ledgers, orderedIds);
  if (kind === 'category-parents') {
    const ledgerId = state.settings.activeLedgerId;
    const ledgerCategories = categoriesForLedger(ledgerId);
    const parents = reorderItems(ledgerCategories.filter(category => !category.parentId), orderedIds);
    const knownIds = new Set();
    const reorderedCategories = parents.flatMap(parent => {
      knownIds.add(parent.id);
      const children = ledgerCategories.filter(category => category.parentId === parent.id);
      children.forEach(child => knownIds.add(child.id));
      return [parent, ...children];
    }).concat(ledgerCategories.filter(category => !knownIds.has(category.id)));
    state.categories = state.categories.filter(category => category.ledgerId !== ledgerId).concat(reorderedCategories);
  }
  if (kind === 'category-children' && parentId) {
    const children = reorderItems(state.categories.filter(category => category.parentId === parentId), orderedIds);
    let childIndex = 0;
    state.categories = state.categories.map(category => category.parentId === parentId ? children[childIndex++] : category);
  }
  persist();
}

function startReorder(event) {
  const handle = event.target.closest('[data-drag-handle]');
  if (!handle) return;
  const item = handle.closest('[data-reorder-id]');
  const container = item?.parentElement?.closest('[data-reorder-container]');
  if (!item || !container) return;
  reorderDrag = { pointerId: event.pointerId, item, container, handle, startY: event.clientY, moved: false };
  item.classList.add('dragging');
  document.body.classList.add('sorting');
  handle.setPointerCapture?.(event.pointerId);
  event.preventDefault();
}

function moveReorder(event) {
  if (!reorderDrag || event.pointerId !== reorderDrag.pointerId) return;
  const { item, container } = reorderDrag;
  if (Math.abs(event.clientY - reorderDrag.startY) > 4) reorderDrag.moved = true;
  if (!reorderDrag.moved) return;
  const siblings = [...container.children].filter(child => child !== item && child.matches('[data-reorder-id]'));
  const before = siblings.find(child => event.clientY < child.getBoundingClientRect().top + child.getBoundingClientRect().height / 2);
  if (before) container.insertBefore(item, before); else container.appendChild(item);
  event.preventDefault();
}

function finishReorder(event) {
  if (!reorderDrag || event.pointerId !== reorderDrag.pointerId) return;
  const { item, container, handle, moved } = reorderDrag;
  item.classList.remove('dragging');
  document.body.classList.remove('sorting');
  handle.releasePointerCapture?.(event.pointerId);
  if (moved) {
    const orderedIds = [...container.children].filter(child => child.matches('[data-reorder-id]')).map(child => child.dataset.reorderId);
    commitReorder(container.dataset.reorderContainer, orderedIds, container.dataset.parentId || '');
    showToast('排序已保存');
  }
  reorderDrag = null;
}

function onClick(event) {
  const pageButton = event.target.closest('[data-page]');
  if (pageButton) {
    const next = pageButton.dataset.page;
    if (next === 'add') return openEntry('expense');
    activePage = next; render(); return;
  }
  const entry = event.target.closest('[data-entry-type]'); if (entry) return openEntry(entry.dataset.entryType);
  const txButton = event.target.closest('[data-tx-id]'); if (txButton) return openTransactionDetail(txButton.dataset.txId);
  const accountButton = event.target.closest('[data-account-id]'); if (accountButton) return openAccountTransactions(accountButton.dataset.accountId);
  const categoryBreakdown = event.target.closest('[data-category-breakdown]'); if (categoryBreakdown) return openCategoryBreakdown(categoryBreakdown.dataset.categoryBreakdown, categoryBreakdown.dataset.categoryType);
  const txEdit = event.target.closest('[data-edit-tx]'); if (txEdit) { const tx=state.transactions.find(t=>t.id===txEdit.dataset.editTx); return openEntry(tx.type, tx); }
  const txDelete = event.target.closest('[data-delete-tx]'); if (txDelete) return deleteTransaction(txDelete.dataset.deleteTx);
  const statsPeriodButton = event.target.closest('[data-stats-period]'); if (statsPeriodButton) { statsPeriod=statsPeriodButton.dataset.statsPeriod; renderStats(); return; }
  const editAccount = event.target.closest('[data-edit-account]'); if (editAccount) return openAccountForm(editAccount.dataset.editAccount);
  const deleteAccountButton = event.target.closest('[data-delete-account]'); if (deleteAccountButton) return deleteAccount(deleteAccountButton.dataset.deleteAccount);
  const editLedger = event.target.closest('[data-edit-ledger]'); if (editLedger) return openLedgerForm(editLedger.dataset.editLedger);
  const deleteLedgerButton = event.target.closest('[data-delete-ledger]'); if (deleteLedgerButton) return deleteLedger(deleteLedgerButton.dataset.deleteLedger);
  const editCategory = event.target.closest('[data-edit-category]'); if (editCategory) return openCategoryForm(editCategory.dataset.editCategory);
  const newChild = event.target.closest('[data-new-child]'); if (newChild) return openCategoryForm(null,newChild.dataset.newChild);
  const deleteCategoryButton = event.target.closest('[data-delete-category]'); if (deleteCategoryButton) return deleteCategory(deleteCategoryButton.dataset.deleteCategory);
  const categoryIconButton = event.target.closest('[data-category-icon]'); if (categoryIconButton) { const input=sheetRoot.querySelector('#categoryIcon'); if(input) input.value=categoryIconButton.dataset.categoryIcon; sheetRoot.querySelectorAll('[data-category-icon]').forEach(button=>button.classList.toggle('active',button===categoryIconButton)); return; }
  const month = event.target.closest('[data-month]'); if (month) { viewingMonth = new Date(viewingMonth.getFullYear(),viewingMonth.getMonth()+Number(month.dataset.month),1); render(); return; }
  const pickerYear = event.target.closest('[data-picker-year]'); if (pickerYear) return openMonthPicker(Number(pickerYear.dataset.pickerYear));
  const pickedMonth = event.target.closest('[data-pick-month]'); if (pickedMonth) { const [year,monthNumber]=pickedMonth.dataset.pickMonth.split('-').map(Number); viewingMonth=new Date(year,monthNumber-1,1); closeSheet(); activePage='home'; render(); return; }
  const draftType = event.target.closest('[data-draft-type]'); if (draftType) { syncDraftFields(); entryDraft={...defaultDraft(draftType.dataset.draftType), date:entryDraft.date,note:entryDraft.note,ledgerId:entryDraft.ledgerId}; updateEntryTypeSelection(); return; }
  const primaryCat = event.target.closest('[data-primary-cat]'); if (primaryCat) { syncDraftFields(); entryDraft.primaryCategoryId=primaryCat.dataset.primaryCat; entryDraft.categoryId=''; updatePrimaryCategorySelection(); return; }
  const childCat = event.target.closest('[data-category-id]'); if (childCat) { syncDraftFields(); entryDraft.categoryId=childCat.dataset.categoryId; updateChildCategorySelection(); return; }
  const key = event.target.closest('[data-key]'); if (key) { syncDraftFields(); const value=key.dataset.key; if(value==='clear')entryDraft.expression=''; else if(value==='⌫')entryDraft.expression=entryDraft.expression.slice(0,-1); else if(entryDraft.expression.length<24)entryDraft.expression+=value; updateCalculatorDisplay(); return; }
  const ledger = event.target.closest('[data-ledger-id]'); if (ledger) { state.settings.activeLedgerId=ledger.dataset.ledgerId; persist(); closeSheet(); render(); return; }
  const action = event.target.closest('[data-action]')?.dataset.action;
  if (!action) return;
  if (action === 'close') return closeSheet();
  if (action === 'close-overlay' && event.target.classList.contains('overlay')) return closeSheet();
  if (action === 'ledger') return openLedgerSheet();
  if (action === 'privacy') { state.settings.amountsHidden=!state.settings.amountsHidden; persist(); render(); return; }
  if (action === 'messages') return openMessages();
  if (action === 'calendar') return openCalendarSheet();
  if (action === 'pick-home-month') return openMonthPicker();
  if (action === 'search') { closeSheet(); activePage='transactions'; searchTerm=''; render(); return; }
  if (action === 'accounts') return openAccounts();
  if (action === 'new-account') return openAccountForm();
  if (action === 'save-account') return saveAccount();
  if (action === 'ledgers') return openLedgerManager();
  if (action === 'new-ledger') return openLedgerForm();
  if (action === 'save-ledger') return saveLedger();
  if (action === 'categories') return openCategories();
  if (action === 'budgets') return openBudgetManager();
  if (action === 'save-budgets') return saveBudgets();
  if (action === 'new-category') return openCategoryForm();
  if (action === 'save-category') return saveCategory();
  if (action === 'save-entry') return saveEntry();
  if (action === 'dark') { state.settings.dark=!state.settings.dark; persist(); render(); return; }
  if (action === 'data') return openDataSheet();
  if (action === 'excel-export') return exportExcel();
  if (action === 'excel-import') { sheetRoot.querySelector('#excelFileInput')?.click(); return; }
  if (action === 'confirm-excel-import') return commitExcelImport();
  if (action === 'close-import') { pendingExcelImport=null; return closeSheet(); }
  if (action === 'load-demo') { if(confirm('载入体验数据会覆盖当前内容，是否继续？')) { state=resetState(true); closeSheet(); activePage='home'; render(); showToast('体验数据已载入'); } return; }
  if (action === 'clear-data') { if(confirm('确认永久清空全部数据吗？此操作无法恢复。')) { state=resetState(false); closeSheet(); activePage='home'; render(); showToast('本地数据已清空'); } return; }
  if (action === 'about') return openAbout();
  if (action === 'filter') return showToast('可直接使用上方搜索框筛选账目');
}

function onInput(event) {
  if (event.target.id === 'txSearch') { searchTerm = event.target.value; renderTransactions(); const input=pageEl.querySelector('#txSearch'); input?.focus(); input?.setSelectionRange(searchTerm.length,searchTerm.length); }
  if (event.target.id === 'accountKind') {
    const isLiability = event.target.value === 'liability';
    sheetRoot.querySelectorAll('.liability-only').forEach(row=>row.classList.toggle('hidden',!isLiability));
    if (!editingAccountId) {
      const availableInput = sheetRoot.querySelector('#accountAvailable');
      if (availableInput) availableInput.value = isLiability ? roundMoney(sheetRoot.querySelector('#accountLimit')?.value || 0).toFixed(2) : '0.00';
    }
  }
  if (event.target.id === 'accountLimit' && editingAccountId) {
    const existing = accountById(editingAccountId);
    const availableInput = sheetRoot.querySelector('#accountAvailable');
    if (existing?.kind === 'liability' && availableInput) availableInput.value = roundMoney(Math.max(0, Number(event.target.value || 0) - getAccountBalance(state, existing.id))).toFixed(2);
  }
}

document.addEventListener('click', onClick);
document.addEventListener('input', onInput);
document.addEventListener('change', event => {
  if (event.target.id !== 'excelFileInput') return;
  const file=event.target.files?.[0]; event.target.value='';
  if(file) readExcelFile(file);
});
document.addEventListener('pointerdown', startReorder);
document.addEventListener('pointermove', moveReorder, { passive: false });
document.addEventListener('pointerup', finishReorder);
document.addEventListener('pointercancel', finishReorder);
document.body.classList.toggle('dark', Boolean(state.settings.dark));
render();

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) navigator.serviceWorker.register('./sw.js').catch(console.warn);

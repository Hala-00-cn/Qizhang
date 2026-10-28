export const STORAGE_KEY = 'qizhang.local.v1';

export const categorySeed = [
  { id: 'food', name: '餐饮', type: 'expense', icon: '餐' },
  { id: 'breakfast', parentId: 'food', name: '早餐', type: 'expense', icon: '🥣' },
  { id: 'lunch', parentId: 'food', name: '午餐', type: 'expense', icon: '🍱' },
  { id: 'dinner', parentId: 'food', name: '晚餐', type: 'expense', icon: '🍲' },
  { id: 'snack', parentId: 'food', name: '零食饮品', type: 'expense', icon: '🧋' },
  { id: 'transport', name: '交通', type: 'expense', icon: '行' },
  { id: 'metro', parentId: 'transport', name: '公交地铁', type: 'expense', icon: '🚇' },
  { id: 'taxi', parentId: 'transport', name: '打车', type: 'expense', icon: '🚕' },
  { id: 'fuel', parentId: 'transport', name: '加油', type: 'expense', icon: '⛽' },
  { id: 'train', parentId: 'transport', name: '火车机票', type: 'expense', icon: '✈️' },
  { id: 'shopping', name: '购物', type: 'expense', icon: '购' },
  { id: 'daily', parentId: 'shopping', name: '日用百货', type: 'expense', icon: '🛍️' },
  { id: 'clothes', parentId: 'shopping', name: '服饰', type: 'expense', icon: '👕' },
  { id: 'digital', parentId: 'shopping', name: '数码', type: 'expense', icon: '🎧' },
  { id: 'home', parentId: 'shopping', name: '家居', type: 'expense', icon: '🛋️' },
  { id: 'life', name: '生活', type: 'expense', icon: '家' },
  { id: 'rent', parentId: 'life', name: '房租', type: 'expense', icon: '🏠' },
  { id: 'utilities', parentId: 'life', name: '水电燃气', type: 'expense', icon: '💡' },
  { id: 'medical', parentId: 'life', name: '医疗', type: 'expense', icon: '💊' },
  { id: 'entertainment', parentId: 'life', name: '娱乐', type: 'expense', icon: '🎬' },
  { id: 'income_work', name: '工作收入', type: 'income', icon: '薪' },
  { id: 'salary', parentId: 'income_work', name: '工资', type: 'income', icon: '💼' },
  { id: 'bonus', parentId: 'income_work', name: '奖金', type: 'income', icon: '🎁' },
  { id: 'sidejob', parentId: 'income_work', name: '兼职', type: 'income', icon: '🧩' },
  { id: 'income_other', name: '其他收入', type: 'income', icon: '收' },
  { id: 'refund', parentId: 'income_other', name: '退款报销', type: 'income', icon: '↩️' },
  { id: 'interest', parentId: 'income_other', name: '利息收益', type: 'income', icon: '📈' },
  { id: 'gift', parentId: 'income_other', name: '红包礼金', type: 'income', icon: '🧧' }
];

function isoOffset(days = 0, hour = 12, minute = 0) {
  const d = new Date();
  d.setDate(d.getDate() + days);
  d.setHours(hour, minute, 0, 0);
  return d.toISOString();
}

function baseState() {
  return {
    version: 2,
    settings: { dark: false, amountsHidden: false, activeLedgerId: 'ledger-daily', reminderRange: 7 },
    ledgers: [
      { id: 'ledger-daily', name: '日常生活', color: '#6857d9', icon: '日' }
    ],
    categories: categorySeed.map(category => ({ ...category, ledgerId: 'ledger-daily' })),
    accounts: [
      { id: 'cash', name: '现金', type: 'cash', kind: 'asset', initialBalance: 0, color: '#ee9b55', includeInAssets: true, order: 1 },
      { id: 'bank', name: '储蓄卡', type: 'bank', kind: 'asset', initialBalance: 0, color: '#5a77d7', includeInAssets: true, order: 2 },
      { id: 'wechat', name: '微信', type: 'wallet', kind: 'asset', initialBalance: 0, color: '#3db795', includeInAssets: true, order: 3 },
      { id: 'alipay', name: '支付宝', type: 'wallet', kind: 'asset', initialBalance: 0, color: '#4198eb', includeInAssets: true, order: 4 }
    ],
    transactions: [],
    budgets: [],
    loans: [],
    messages: []
  };
}

export function demoState() {
  const state = baseState();
  state.ledgers.push({ id: 'ledger-trip', name: '旅行计划', color: '#f08070', icon: '旅' });
  state.categories.push(...cloneCategoriesForLedger(state.categories, 'ledger-trip'));
  state.accounts = [
    { id: 'cash', name: '现金', type: 'cash', kind: 'asset', initialBalance: 860, color: '#ee9b55', includeInAssets: true, order: 1 },
    { id: 'bank', name: '招商银行', type: 'bank', kind: 'asset', initialBalance: 18260, color: '#5a77d7', includeInAssets: true, order: 2 },
    { id: 'wechat', name: '微信', type: 'wallet', kind: 'asset', initialBalance: 2340, color: '#3db795', includeInAssets: true, order: 3 },
    { id: 'alipay', name: '支付宝', type: 'wallet', kind: 'asset', initialBalance: 1280, color: '#4198eb', includeInAssets: true, order: 4 },
    { id: 'huabei', name: '花呗', type: 'credit', kind: 'liability', initialDebt: 1268, limit: 12000, billDay: 1, dueDay: 9, color: '#f08070', includeInAssets: true, order: 5 }
  ];
  state.transactions = [
    { id: 't1', type: 'expense', amount: 26, accountId: 'wechat', categoryId: 'breakfast', note: '咖啡和三明治', date: isoOffset(0, 8, 26), ledgerId: 'ledger-daily' },
    { id: 't2', type: 'expense', amount: 18, accountId: 'alipay', categoryId: 'metro', note: '通勤', date: isoOffset(0, 8, 3), ledgerId: 'ledger-daily' },
    { id: 't3', type: 'expense', amount: 68.5, accountId: 'wechat', categoryId: 'lunch', note: '和同事午餐', date: isoOffset(-1, 12, 18), ledgerId: 'ledger-daily' },
    { id: 't4', type: 'expense', amount: 236, accountId: 'huabei', categoryId: 'daily', note: '生活用品', date: isoOffset(-1, 19, 45), ledgerId: 'ledger-daily' },
    { id: 't5', type: 'income', amount: 8500, accountId: 'bank', categoryId: 'salary', note: '本月工资', date: isoOffset(0, 9, 30), ledgerId: 'ledger-daily' },
    { id: 't6', type: 'expense', amount: 42, accountId: 'alipay', categoryId: 'taxi', note: '雨天打车', date: isoOffset(-3, 21, 2), ledgerId: 'ledger-daily' },
    { id: 't7', type: 'expense', amount: 109, accountId: 'wechat', categoryId: 'dinner', note: '朋友聚餐', date: isoOffset(-5, 20, 10), ledgerId: 'ledger-daily' },
    { id: 't8', type: 'transfer', amount: 1000, accountId: 'bank', toAccountId: 'alipay', note: '日常备用金', date: isoOffset(-8), ledgerId: 'ledger-daily' }
  ];
  state.budgets = [
    { id: 'budget-food', ledgerId: 'ledger-daily', categoryId: 'food', amount: 1800 },
    { id: 'budget-transport', ledgerId: 'ledger-daily', categoryId: 'transport', amount: 600 },
    { id: 'budget-shopping', ledgerId: 'ledger-daily', categoryId: 'shopping', amount: 1200 }
  ];
  state.loans = [];
  state.messages = [
    { id: 'm1', type: 'debt', title: '花呗还款提醒', body: '当前欠款待处理', relatedId: 'huabei', createdAt: isoOffset(-1), handled: false, read: false }
  ];
  return state;
}

export function loadState() {
  try {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
      const parsed = JSON.parse(stored);
      let migrated = migrateLedgerCategories(parsed);
      if (!Array.isArray(migrated.budgets)) migrated = { ...migrated, budgets: [] };
      if (migrated !== parsed) saveState(migrated);
      return migrated;
    }
  } catch (error) {
    console.warn('Unable to read local data', error);
  }
  const useDemo = new URLSearchParams(location.search).get('demo') === '1';
  return useDemo ? demoState() : baseState();
}

export function saveState(state) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

export function resetState(demo = false) {
  const state = demo ? demoState() : baseState();
  saveState(state);
  return state;
}

export const typeLabels = {
  expense: '支出', income: '收入', transfer: '转账', lend: '借出', borrow: '借入', repay: '还款', collect: '收款', adjustment: '余额校准'
};

export function formatMoney(value, hidden = false) {
  if (hidden) return '••••';
  const number = roundMoney(value || 0);
  return `¥${number.toLocaleString('zh-CN', { minimumFractionDigits: Number.isInteger(number) ? 0 : 2, maximumFractionDigits: 2 })}`;
}

export function roundMoney(value) {
  const number = Number(value || 0);
  return Number.isFinite(number) ? Math.round((number + Number.EPSILON) * 100) / 100 : 0;
}

export function localDateKey(value) {
  const d = new Date(value);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

export function getAccountBalance(state, accountId) {
  const account = state.accounts.find(item => item.id === accountId);
  if (!account) return 0;
  if (account.kind === 'liability') {
    let debt = Number(account.initialDebt || 0);
    state.transactions.forEach(tx => {
      if (tx.accountId === accountId && tx.type === 'expense') debt += tx.amount;
      if (tx.toAccountId === accountId && ['transfer', 'repay'].includes(tx.type)) debt -= tx.amount;
      if (tx.accountId === accountId && tx.type === 'income') debt -= tx.amount;
      if (tx.accountId === accountId && tx.type === 'adjustment') debt += tx.amount;
    });
    return roundMoney(Math.max(0, debt));
  }
  let balance = Number(account.initialBalance || 0);
  state.transactions.forEach(tx => {
    if (tx.accountId === accountId) {
      if (['income', 'borrow', 'collect'].includes(tx.type)) balance += tx.amount;
      if (['expense', 'lend', 'repay', 'transfer'].includes(tx.type)) balance -= tx.amount;
      if (tx.type === 'adjustment') balance += tx.amount;
    }
    if (tx.toAccountId === accountId && tx.type === 'transfer') balance += tx.amount;
  });
  return roundMoney(balance);
}

export function accountDisplayValue(state, account) {
  const value = getAccountBalance(state, account.id);
  return account.kind === 'liability' ? -value : value;
}

export function getAssetSummary(state) {
  const assets = state.accounts.filter(a => a.kind === 'asset' && a.includeInAssets).reduce((sum, a) => sum + getAccountBalance(state, a.id), 0);
  const debts = state.accounts.filter(a => a.kind === 'liability' && a.includeInAssets).reduce((sum, a) => sum + getAccountBalance(state, a.id), 0);
  const receivable = state.loans.filter(l => l.type === 'lend' && l.includeInAssets && l.status !== 'closed').reduce((sum, l) => sum + Math.max(0, l.principal - l.settled), 0);
  const payable = state.loans.filter(l => l.type === 'borrow' && l.includeInAssets && l.status !== 'closed').reduce((sum, l) => sum + Math.max(0, l.principal - l.settled), 0);
  return { assets: assets + receivable, debts: debts + payable, net: assets + receivable - debts - payable };
}

export function monthTransactions(state, ledgerId = state.settings.activeLedgerId, monthDate = new Date()) {
  return state.transactions.filter(tx => {
    const d = new Date(tx.date);
    return tx.ledgerId === ledgerId && d.getFullYear() === monthDate.getFullYear() && d.getMonth() === monthDate.getMonth();
  });
}

export function incomeExpenseSummary(transactions) {
  return transactions.reduce((acc, tx) => {
    if (tx.type === 'income') acc.income += tx.amount;
    if (tx.type === 'expense') acc.expense += tx.amount;
    return acc;
  }, { income: 0, expense: 0 });
}

export function uid(prefix = 'id') {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

export function cloneCategoriesForLedger(categories, ledgerId) {
  const idMap = new Map(categories.map(category => [category.id, uid('category')]));
  return categories.map(category => ({
    ...category,
    id: idMap.get(category.id),
    ledgerId,
    parentId: category.parentId ? idMap.get(category.parentId) : undefined
  }));
}

function migrateLedgerCategories(state) {
  const categories = Array.isArray(state.categories) ? state.categories : [];
  const unscoped = categories.filter(category => !category.ledgerId);
  if (!unscoped.length || !state.ledgers?.length) return state;

  const primaryLedgerId = state.ledgers.some(ledger => ledger.id === state.settings?.activeLedgerId)
    ? state.settings.activeLedgerId
    : state.ledgers[0].id;
  const legacyIds = new Set(unscoped.map(category => category.id));
  const migratedCategories = [
    ...categories.filter(category => category.ledgerId),
    ...unscoped.map(category => ({ ...category, ledgerId: primaryLedgerId }))
  ];
  const categoryMaps = new Map();

  state.ledgers.filter(ledger => ledger.id !== primaryLedgerId).forEach(ledger => {
    const copies = cloneCategoriesForLedger(unscoped, ledger.id);
    categoryMaps.set(ledger.id, new Map(unscoped.map((category, index) => [category.id, copies[index].id])));
    migratedCategories.push(...copies);
  });

  const transactions = (state.transactions || []).map(transaction => {
    const mappedId = categoryMaps.get(transaction.ledgerId)?.get(transaction.categoryId);
    return mappedId && legacyIds.has(transaction.categoryId) ? { ...transaction, categoryId: mappedId } : transaction;
  });

  return { ...state, version: Math.max(2, Number(state.version || 1)), categories: migratedCategories, transactions };
}

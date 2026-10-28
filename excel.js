const encoder = new TextEncoder();
const decoder = new TextDecoder();

const sheetSchemas = [
  ['账目', 'transactions', ['id','type','amount','accountId','toAccountId','categoryId','ledgerId','date','note','counterparty','dueDate','includeInAssets']],
  ['账户', 'accounts', ['id','name','type','kind','assetTag','initialBalance','initialDebt','limit','billDay','dueDay','color','includeInAssets','order']],
  ['账本', 'ledgers', ['id','name','icon','color']],
  ['分类', 'categories', ['id','ledgerId','parentId','name','type','icon']],
  ['预算', 'budgets', ['id','ledgerId','categoryId','amount']],
  ['借贷', 'loans', ['id','transactionId','type','principal','settled','counterparty','dueDate','ledgerId','includeInAssets','status']],
  ['消息', 'messages', ['id','type','title','body','relatedId','createdAt','handled','read']]
];

const labels = {
  id: '编号', type: '类型', amount: '金额', accountId: '账户编号', toAccountId: '转入账户编号', categoryId: '分类编号',
  ledgerId: '账本编号', date: '日期', note: '备注', counterparty: '对象', dueDate: '到期日', includeInAssets: '计入资产',
  name: '名称', kind: '账户性质', initialBalance: '初始余额', initialDebt: '初始欠款', limit: '总额度', billDay: '账单日',
  dueDay: '还款日', color: '颜色', order: '排序', icon: '图标', assetTag: '资产标签', parentId: '上级分类编号', amountBudget: '预算金额',
  transactionId: '关联账目编号', principal: '本金', settled: '已结清', status: '状态', title: '标题', body: '内容',
  relatedId: '关联编号', createdAt: '创建时间', handled: '已处理', read: '已读', amountMonthly: '每月预算'
};

const headerKeys = Object.fromEntries(Object.entries(labels).map(([key, label]) => [label, key]));
for (const key of Object.keys(labels)) headerKeys[key] = key;
Object.assign(headerKeys, {
  '账户': 'accountName', '账户名称': 'accountName', '转入账户': 'toAccountName', '转入账户名称': 'toAccountName',
  '分类': 'categoryName', '分类名称': 'categoryName', '账本': 'ledgerName', '账本名称': 'ledgerName',
  '时间': 'date', '金额(元)': 'amount', '收支类型': 'type', '备注说明': 'note', '交易对象': 'counterparty',
  '每月预算': 'amount', '每月额度': 'amount', '数据JSON': 'data',
  account: 'accountName', account_name: 'accountName', target_account: 'toAccountName',
  category: 'categoryName', category_name: 'categoryName', ledger: 'ledgerName', ledger_name: 'ledgerName',
  date_time: 'date', description: 'note', transaction_type: 'type', amount_yuan: 'amount'
});

const numberFields = new Set(['amount','initialBalance','initialDebt','limit','billDay','dueDay','order','principal','settled']);
const booleanFields = new Set(['includeInAssets','handled','read']);
const typeNames = { '支出':'expense','收入':'income','转账':'transfer','借出':'lend','借入':'borrow','还款':'repay','收款':'collect','余额校准':'adjustment' };

export async function buildExcelWorkbook(state) {
  const files = [];
  const sheets = sheetSchemas.map(([title, key, fields], index) => {
    const rows = [fields.map(field => labels[field] || field)];
    if (index === 0) rows[0].push('其他字段');
    for (const record of state[key] || []) {
      const base = fields.map(field => record[field] ?? '');
      const known = new Set(fields);
      const extras = Object.fromEntries(Object.entries(record).filter(([field]) => !known.has(field)));
      if (index === 0) base.push(Object.keys(extras).length ? JSON.stringify(extras) : '');
      rows.push(base);
    }
    return { title, index, xml: worksheetXml(rows) };
  });
  const settingsRows = [['名称','数据JSON'], ...Object.entries(state.settings || {}).map(([name, value]) => [name, JSON.stringify(value)])];
  sheets.push({ title: '设置', index: sheets.length, xml: worksheetXml(settingsRows) });

  files.push({ name: '[Content_Types].xml', data: xmlBytes(contentTypesXml(sheets.length)) });
  files.push({ name: '_rels/.rels', data: xmlBytes(rootRelsXml()) });
  files.push({ name: 'xl/workbook.xml', data: xmlBytes(workbookXml(sheets)) });
  files.push({ name: 'xl/_rels/workbook.xml.rels', data: xmlBytes(workbookRelsXml(sheets.length)) });
  sheets.forEach((sheet, index) => files.push({ name: `xl/worksheets/sheet${index + 1}.xml`, data: xmlBytes(sheet.xml) }));
  return zipStore(files);
}

function xmlBytes(value) { return encoder.encode(value); }
function xmlEscape(value) { return String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[char])); }
function columnName(index) {
  let name = '';
  for (let n = index + 1; n; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + (n - 1) % 26) + name;
  return name;
}
function cellXml(value, ref) {
  if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
  if (typeof value === 'boolean') return `<c r="${ref}" t="b"><v>${value ? 1 : 0}</v></c>`;
  const text = String(value ?? '');
  return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xmlEscape(text)}</t></is></c>`;
}
function worksheetXml(rows) {
  const body = rows.map((row, r) => `<row r="${r + 1}">${row.map((value, c) => cellXml(value, `${columnName(c)}${r + 1}`)).join('')}</row>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${body}</sheetData></worksheet>`;
}
function contentTypesXml(count) {
  const overrides = Array.from({length:count}, (_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides}</Types>`;
}
function rootRelsXml() { return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`; }
function workbookXml(sheets) {
  const entries = sheets.map((sheet, i) => `<sheet name="${xmlEscape(sheet.title)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${entries}</sheets></workbook>`;
}
function workbookRelsXml(count) {
  const rels = Array.from({length:count}, (_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

function crc32(bytes) {
  let crc = -1;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ -1) >>> 0;
}
function put16(view, offset, value) { view.setUint16(offset, value, true); }
function put32(view, offset, value) { view.setUint32(offset, value >>> 0, true); }
function zipStore(files) {
  const localParts = [], centralParts = [];
  let offset = 0;
  for (const file of files) {
    const name = encoder.encode(file.name), data = file.data, crc = crc32(data);
    const local = new Uint8Array(30 + name.length + data.length), lv = new DataView(local.buffer);
    put32(lv,0,0x04034b50); put16(lv,4,20); put16(lv,6,0x0800); put16(lv,8,0); put32(lv,14,crc); put32(lv,18,data.length); put32(lv,22,data.length); put16(lv,26,name.length); put16(lv,28,0);
    local.set(name,30); local.set(data,30+name.length); localParts.push(local);
    const central = new Uint8Array(46 + name.length), cv = new DataView(central.buffer);
    put32(cv,0,0x02014b50); put16(cv,4,20); put16(cv,6,20); put16(cv,8,0x0800); put16(cv,10,0); put32(cv,16,crc); put32(cv,20,data.length); put32(cv,24,data.length); put16(cv,28,name.length); put16(cv,30,0); put16(cv,32,0); put32(cv,42,offset);
    central.set(name,46); centralParts.push(central); offset += local.length;
  }
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  put32(ev,0,0x06054b50); put16(ev,8,files.length); put16(ev,10,files.length); put32(ev,12,centralSize); put32(ev,16,offset);
  return concatBytes([...localParts, ...centralParts, end]);
}
function concatBytes(parts) {
  const result = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) { result.set(part, offset); offset += part.length; }
  return result;
}

async function inflateRaw(bytes) {
  if (typeof DecompressionStream !== 'function') throw new Error('当前系统不支持读取压缩 Excel，请更新系统 WebView 后重试。');
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  } catch { throw new Error('无法解压此 Excel 文件，请确认文件未损坏。'); }
}
async function unzip(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let endOffset = -1;
  for (let p = Math.max(0, bytes.length - 65557); p <= bytes.length - 22; p++) if (view.getUint32(p, true) === 0x06054b50) endOffset = p;
  if (endOffset < 0) throw new Error('不是有效的 .xlsx 文件。');
  const count = view.getUint16(endOffset + 10, true), centralOffset = view.getUint32(endOffset + 16, true), entries = new Map();
  let p = centralOffset;
  for (let i = 0; i < count; i++) {
    if (view.getUint32(p, true) !== 0x02014b50) throw new Error('Excel 压缩包目录损坏。');
    const method = view.getUint16(p + 10, true), size = view.getUint32(p + 20, true), nameLength = view.getUint16(p + 28, true), extraLength = view.getUint16(p + 30, true), commentLength = view.getUint16(p + 32, true), localOffset = view.getUint32(p + 42, true);
    const name = decoder.decode(bytes.slice(p + 46, p + 46 + nameLength));
    const localNameLength = view.getUint16(localOffset + 26, true), localExtraLength = view.getUint16(localOffset + 28, true), dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.slice(dataStart, dataStart + size);
    entries.set(name, method === 0 ? compressed : method === 8 ? await inflateRaw(compressed) : null);
    if (!entries.get(name)) throw new Error(`Excel 文件使用了不支持的压缩方式：${name}`);
    p += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}
function xmlDocument(text) {
  const doc = new DOMParser().parseFromString(text, 'application/xml');
  if (doc.querySelector('parsererror')) throw new Error('Excel 工作簿 XML 格式错误。');
  return doc;
}
function textContent(node, selector) { return node.getElementsByTagNameNS('*', selector)[0]?.textContent || ''; }
function cellColumn(ref) { return [...ref.matchAll(/[A-Z]+/g)][0]?.[0] || 'A'; }
function columnIndex(name) { let value=0; for (const char of name) value=value*26+char.charCodeAt(0)-64; return value-1; }
function parseWorksheet(xml, shared) {
  const doc = xmlDocument(xml), rows = [];
  for (const row of doc.getElementsByTagNameNS('*','row')) {
    const values = [];
    for (const cell of row.getElementsByTagNameNS('*','c')) {
      const index = columnIndex(cellColumn(cell.getAttribute('r'))), type = cell.getAttribute('t');
      let value = '';
      if (type === 'inlineStr') value = textContent(cell,'t');
      else {
        const raw = textContent(cell,'v');
        if (type === 's') value = shared[Number(raw)] ?? '';
        else if (type === 'b') value = raw === '1';
        else if (raw !== '') value = Number.isFinite(Number(raw)) ? Number(raw) : raw;
      }
      values[index] = value;
    }
    rows.push(values.map(value => value ?? ''));
  }
  return rows;
}
function fromExcelSerial(value) {
  if (typeof value !== 'number' || value < 1) return value;
  return new Date(Date.UTC(1899, 11, 30) + Math.round(value * 86400000)).toISOString();
}
function normalizeValue(key, value) {
  if (value === '' || value === null || value === undefined) return undefined;
  if (numberFields.has(key)) return Number(value);
  if (booleanFields.has(key)) return value === true || value === 1 || String(value).toLowerCase() === 'true' || String(value) === '是';
  if (key === 'type' && typeNames[value]) return typeNames[value];
  if (['date','dueDate','createdAt'].includes(key) && typeof value === 'number') return fromExcelSerial(value);
  return String(value);
}
function parseTable(rows) {
  if (!rows?.length) return [];
  const headers = rows[0].map(value => headerKeys[String(value).trim()] || String(value).trim());
  return rows.slice(1).filter(row => row.some(value => value !== '')).map(row => {
    const record = {};
    headers.forEach((key, index) => {
      const value = row[index];
      if (!key || value === '' || value === undefined) return;
      if (key === '其他字段') {
        try { Object.assign(record, JSON.parse(String(value))); } catch { }
        return;
      }
      const normalized = normalizeValue(key, value);
      if (normalized !== undefined) record[key] = normalized;
    });
    return record;
  });
}

export async function readExcelWorkbook(file) {
  const entries = await unzip(new Uint8Array(await file.arrayBuffer()));
  const workbookBytes = entries.get('xl/workbook.xml');
  if (!workbookBytes) throw new Error('Excel 文件中没有工作簿。');
  const workbook = xmlDocument(decoder.decode(workbookBytes));
  const rels = xmlDocument(decoder.decode(entries.get('xl/_rels/workbook.xml.rels') || new Uint8Array()));
  const targets = new Map([...rels.getElementsByTagNameNS('*','Relationship')].map(rel => [rel.getAttribute('Id'), rel.getAttribute('Target')]));
  const sharedBytes = entries.get('xl/sharedStrings.xml');
  const shared = sharedBytes ? [...xmlDocument(decoder.decode(sharedBytes)).getElementsByTagNameNS('*','si')].map(si => [...si.getElementsByTagNameNS('*','t')].map(t => t.textContent).join('')) : [];
  const tables = {};
  for (const sheet of workbook.getElementsByTagNameNS('*','sheet')) {
    const title = sheet.getAttribute('name'), relId = sheet.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id') || sheet.getAttribute('r:id');
    const target = targets.get(relId);
    if (!target) continue;
    const path = target.startsWith('/') ? target.slice(1) : `xl/${target.replace(/^\.\//,'')}`;
    const content = entries.get(path);
    if (content) tables[title] = parseTable(parseWorksheet(decoder.decode(content), shared));
  }
  const settings = {};
  for (const row of tables['设置'] || []) {
    if (!row.name) continue;
    try { settings[row.name] = JSON.parse(row.data); } catch { settings[row.name] = row.data; }
  }
  return { tables, settings };
}

export function excelSchemas() { return sheetSchemas; }

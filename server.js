/**
 * Gastos Vini App — Poupança Baby Dashboard Server
 * Integrado em tempo real com a planilha Google Sheets (2ª Aba: "Poupança Baby", Colunas F, G, H, I, J)
 */

const http = require('http');
const https = require('https');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const PORT = process.env.PORT || 3075;
const SPREADSHEET_ID = '1Z60EkXO4zn6JEtLMeQ7HvTrIpqeuTgQEt73m2DZX1wY';
const SHEET_EXPORT_URL = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=xlsx`;
const DATA_DIR = path.join(__dirname, 'data');
const CACHE_FILE = path.join(DATA_DIR, 'expenses_cache.json');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// Faz download seguindo redirecionamentos (301, 302, 307, 308)
function fetchBuffer(url, maxRedirects = 6) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error('Too many redirects'));
    https.get(url, { headers: { 'User-Agent': 'Mozilla/5.0 (GastosViniApp/1.0)' } }, (res) => {
      if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
        res.resume();
        return resolve(fetchBuffer(res.headers.location, maxRedirects - 1));
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error(`HTTP status ${res.statusCode}`));
      }
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => resolve(Buffer.concat(chunks)));
      res.on('error', reject);
    }).on('error', reject);
  });
}

// Leitor ZIP em memória via Central Directory (nativo Node.js com zlib)
function unzipEntries(buf) {
  const entries = {};
  let eocdOffset = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset === -1) throw new Error('Invalid XLSX/ZIP buffer (EOCD not found)');

  const cdCount = buf.readUInt16LE(eocdOffset + 10);
  const cdOffset = buf.readUInt32LE(eocdOffset + 16);

  let ptr = cdOffset;
  for (let i = 0; i < cdCount; i++) {
    if (buf.readUInt32LE(ptr) !== 0x02014b50) break;
    const compression = buf.readUInt16LE(ptr + 10);
    const compSize = buf.readUInt32LE(ptr + 20);
    const fileNameLen = buf.readUInt16LE(ptr + 28);
    const extraLen = buf.readUInt16LE(ptr + 30);
    const commentLen = buf.readUInt16LE(ptr + 32);
    const localHeaderOffset = buf.readUInt32LE(ptr + 42);
    const fileName = buf.toString('utf8', ptr + 46, ptr + 46 + fileNameLen);

    const localNameLen = buf.readUInt16LE(localHeaderOffset + 26);
    const localExtraLen = buf.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
    const compData = buf.subarray(dataStart, dataStart + compSize);

    if (compression === 0) {
      entries[fileName] = compData;
    } else if (compression === 8) {
      entries[fileName] = zlib.inflateRawSync(compData);
    }
    ptr += 46 + fileNameLen + extraLen + commentLen;
  }
  return entries;
}

function decodeXmlEntities(str) {
  if (!str) return '';
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(parseInt(code, 10)))
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function excelSerialToDate(serial) {
  const days = Math.floor(serial);
  const ms = Date.UTC(1899, 11, 30) + days * 86400000;
  const dt = new Date(ms);
  const dd = String(dt.getUTCDate()).padStart(2, '0');
  const mm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = dt.getUTCFullYear();
  return { dd, mm: parseInt(mm, 10), mmStr: mm, yyyy };
}

function inferDateMonth(rIdx, hVal, jVal, prevMonth) {
  // Se H tiver padrão dd/mm no texto
  const mMatch = hVal.match(/(?:^|\D)(\d{1,2})\/(\d{2})(?:\D|$)/);
  if (mMatch) {
    const m = parseInt(mMatch[2], 10);
    if (m >= 1 && m <= 12) return m;
  }
  const mTextMatch = (hVal + ' ' + jVal).toLowerCase();
  const ptMonths = {
    janeiro: 1, fevereiro: 2, 'março': 3, marco: 3, abril: 4, maio: 5, junho: 6,
    julho: 7, agosto: 8, setembro: 9, outubro: 10, novembro: 11, dezembro: 12
  };
  if (rIdx === 129) return 3; // 23/01 até mês 25/05 fica no bloco de março
  if (rIdx === 160 || rIdx === 161) return 6;
  if (rIdx === 207) return 9;
  const numMonthMatch = jVal.match(/m[êe]s\s*(\d{1,2})/i) || jVal.match(/(\d{2})\/\d{2}/);
  if (numMonthMatch) {
    const m = parseInt(numMonthMatch[1], 10);
    if (m >= 1 && m <= 12) return m;
  }
  for (const [name, num] of Object.entries(ptMonths)) {
    if (mTextMatch.includes(name)) return num;
  }
  return prevMonth || 1;
}

function inferBlockMonth(rIdx, dateMonth, jVal) {
  // Mapeamento fiel aos blocos mensais da planilha Poupança Baby
  if ([2, 3, 4, 6].includes(rIdx)) return 4;
  if ([5, 7, 8, 9, 10, 11].includes(rIdx)) return 5;
  if (rIdx >= 12 && rIdx <= 16) return 6;
  if (rIdx >= 17 && rIdx <= 22) return 7;
  if ((rIdx >= 23 && rIdx <= 30) || [32, 33, 34, 36].includes(rIdx)) return 8;
  if ([31, 35].includes(rIdx) || (rIdx >= 37 && rIdx <= 43) || rIdx === 45) return 9;
  if ([44, 46, 47, 48, 49].includes(rIdx)) return 10;
  if (rIdx >= 50 && rIdx <= 61) return 11;
  if (rIdx >= 62 && rIdx <= 81) return 12;
  if ((rIdx >= 82 && rIdx <= 90) || (rIdx >= 92 && rIdx <= 102)) return 1;
  if (rIdx === 91 || (rIdx >= 103 && rIdx <= 120)) return 2;
  if ((rIdx >= 121 && rIdx <= 130) || rIdx === 139) return 3;
  if ((rIdx >= 131 && rIdx <= 133) || (rIdx >= 135 && rIdx <= 138) || rIdx === 140) return 4;
  if (rIdx === 134 || (rIdx >= 141 && rIdx <= 153)) return 5;
  if (rIdx >= 154 && rIdx <= 168) return 6;
  if (rIdx >= 169 && rIdx <= 182) return 7;
  if (rIdx >= 183 && rIdx <= 200) return 8;
  if (rIdx >= 201 && rIdx <= 215) return 9;
  if (rIdx >= 216 && rIdx <= 223) return 10;
  if (rIdx >= 224 && rIdx <= 226) return 11;
  if (rIdx >= 227 && rIdx <= 229) return 12;

  // Para novas linhas adicionadas após a linha 229, verifica se o texto menciona "mês XX" ou "MM/AA"
  const mRef = jVal.match(/m[êe]s\s*(\d{1,2})/i) || jVal.match(/\b(\d{2})\/2[5-9]\b/);
  if (mRef) {
    const m = parseInt(mRef[1], 10);
    if (m >= 1 && m <= 12) return m;
  }
  return dateMonth;
}

function parsePoupancaBabyXlsx(xlsxBuf) {
  const entries = unzipEntries(xlsxBuf);

  // 1. Parse sharedStrings.xml
  const sharedStrings = [];
  if (entries['xl/sharedStrings.xml']) {
    const ssXml = entries['xl/sharedStrings.xml'].toString('utf8');
    const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let siMatch;
    while ((siMatch = siRegex.exec(ssXml)) !== null) {
      const inner = siMatch[1];
      const tRegex = /<t(?:\s+[^>]*)?>([\s\S]*?)<\/t>/g;
      let tMatch;
      let combined = '';
      while ((tMatch = tRegex.exec(inner)) !== null) {
        combined += decodeXmlEntities(tMatch[1]);
      }
      sharedStrings.push(combined);
    }
  }

  // 2. Encontrar a aba "Poupança Baby" (ou 2ª aba visível)
  let sheetTarget = 'xl/worksheets/sheet6.xml';
  if (entries['xl/workbook.xml'] && entries['xl/_rels/workbook.xml.rels']) {
    const wbXml = entries['xl/workbook.xml'].toString('utf8');
    const relsXml = entries['xl/_rels/workbook.xml.rels'].toString('utf8');
    const relMap = {};
    const relRegex = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    let rMatch;
    while ((rMatch = relRegex.exec(relsXml)) !== null) {
      relMap[rMatch[1]] = 'xl/' + rMatch[2].replace(/^\/+/, '');
    }

    const sheetRegex = /<sheet\b([^>]*)\/?>/g;
    let sMatch;
    const visibleSheets = [];
    while ((sMatch = sheetRegex.exec(wbXml)) !== null) {
      const attrs = sMatch[1];
      const nameM = attrs.match(/name="([^"]+)"/);
      const stateM = attrs.match(/state="([^"]+)"/);
      const rIdM = attrs.match(/r:id="([^"]+)"/);
      if (nameM && rIdM) {
        const name = decodeXmlEntities(nameM[1]);
        const state = stateM ? stateM[1] : 'visible';
        if (state !== 'hidden') visibleSheets.push({ name, rId: rIdM[1] });
        if (name.toLowerCase().includes('poupan') && name.toLowerCase().includes('baby')) {
          sheetTarget = relMap[rIdM[1]] || sheetTarget;
        }
      }
    }
    if (!entries[sheetTarget] && visibleSheets.length >= 2) {
      sheetTarget = relMap[visibleSheets[1].rId] || sheetTarget;
    }
  }

  const sheetXml = entries[sheetTarget].toString('utf8');
  const rowRegex = /<row\b[^>]*r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
  let rowMatch;

  const records = [];
  let sheetFormulaTotal = null;
  let lastMonth = 4;

  while ((rowMatch = rowRegex.exec(sheetXml)) !== null) {
    const rIdx = parseInt(rowMatch[1], 10);
    if (rIdx === 1) continue; // Linha 1 é o cabeçalho

    const rowInner = rowMatch[2];
    const cols = {};
    const cellRegex = /<c\b([^>\/]*)(?:\/>|>([\s\S]*?)<\/c>)/g;
    let cMatch;
    while ((cMatch = cellRegex.exec(rowInner)) !== null) {
      const cAttrs = cMatch[1];
      const cBody = cMatch[2] || '';
      const refM = cAttrs.match(/r="([A-Z]+)\d+"/);
      if (!refM) continue;
      const colLetter = refM[1];
      if (!['F', 'G', 'H', 'I', 'J'].includes(colLetter)) continue;

      const typeM = cAttrs.match(/t="([^"]+)"/);
      const cellType = typeM ? typeM[1] : '';
      const vM = cBody.match(/<v>([\s\S]*?)<\/v>/);
      let val = vM ? decodeXmlEntities(vM[1]) : '';
      if (cellType === 's' && val !== '') {
        val = sharedStrings[parseInt(val, 10)] || '';
      }
      cols[colLetter] = val.trim();
    }

    const fVal = cols.F || '';
    const gVal = cols.G || '';
    const hVal = cols.H || '';
    const iVal = (cols.I || '').replace(/\.0$/, '');
    const jVal = cols.J || '';

    // Verifica se é a linha de "Total Gastos" (ex: Linha 231)
    if (hVal.toLowerCase().includes('total gastos') || jVal.toLowerCase().includes('total gastos')) {
      sheetFormulaTotal = parseFloat(gVal) || 0;
      continue;
    }

    // Ignora linhas totalmente vazias em G e J
    if (!gVal && !jVal) continue;

    const amount = gVal ? parseFloat(gVal) : 0;
    const year = /^\d{4}$/.test(iVal) ? parseInt(iVal, 10) : 2026;

    let dateDisplay = hVal;
    let dateMonth = null;
    const numH = Number(hVal);
    if (hVal !== '' && !Number.isNaN(numH) && numH > 40000 && numH < 55000) {
      const parsed = excelSerialToDate(numH);
      dateMonth = parsed.mm;
      dateDisplay = `${parsed.dd}/${parsed.mmStr}/${year}`;
    } else {
      dateMonth = inferDateMonth(rIdx, hVal, jVal, lastMonth);
      dateDisplay = hVal ? `${hVal} (${year})` : `Sem dia (${year})`;
    }
    lastMonth = dateMonth;

    const blockMonth = inferBlockMonth(rIdx, dateMonth, jVal);
    const situacaoClean = fVal === '3550.0' ? '3550' : fVal;

    records.push([
      rIdx,
      situacaoClean,
      Math.round(amount * 100) / 100,
      hVal,
      dateDisplay,
      year,
      `${year}-${String(dateMonth).padStart(2, '0')}`,
      `${year}-${String(blockMonth).padStart(2, '0')}`,
      jVal || '(Sem descrição)'
    ]);
  }

  const totalAmount = Math.round(records.reduce((acc, r) => acc + r[2], 0) * 100) / 100;
  return {
    spreadsheetId: SPREADSHEET_ID,
    sheetName: 'Poupança Baby',
    columns: ['F (Situação)', 'G (Saídas poupança)', 'H (Dia)', 'I (Ano)', 'J (Item)'],
    lastSync: new Date().toISOString(),
    totalItems: records.length,
    totalAmount,
    sheetFormulaTotal: sheetFormulaTotal !== null ? Math.round(sheetFormulaTotal * 100) / 100 : totalAmount,
    records
  };
}

async function syncFromGoogleSheets() {
  const xlsxBuf = await fetchBuffer(SHEET_EXPORT_URL);
  const payload = parsePoupancaBabyXlsx(xlsxBuf);
  fs.writeFileSync(CACHE_FILE, JSON.stringify(payload, null, 2), 'utf8');
  return payload;
}

async function getExpensesData(forceRefresh = false) {
  if (!forceRefresh && fs.existsSync(CACHE_FILE)) {
    try {
      return JSON.parse(fs.readFileSync(CACHE_FILE, 'utf8'));
    } catch (e) {
      console.warn('[WARN] Cache corrompido, buscando novamente da planilha...');
    }
  }
  return await syncFromGoogleSheets();
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname;

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    return res.end();
  }

  // Endpoint de leitura / atualização
  if (pathname === '/api/expenses' && req.method === 'GET') {
    try {
      const force = parsedUrl.searchParams.get('refresh') === '1';
      const data = await getExpensesData(force);
      res.writeHead(200, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-cache, no-store, must-revalidate'
      });
      return res.end(JSON.stringify({ ok: true, ...data }));
    } catch (err) {
      console.error('[ERRO /api/expenses]:', err.message);
      res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
      return res.end(JSON.stringify({ ok: false, error: err.message }));
    }
  }

  // Endpoint do botão "Atualizar Planilha"
  if (pathname === '/api/sync' && (req.method === 'POST' || req.method === 'GET')) {
    try {
      const data = await syncFromGoogleSheets();
      console.log(`[SYNC OK] ${data.totalItems} itens sincronizados — Total: R$ ${data.totalAmount}`);
      res.writeHead(200, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-cache, no-store, must-revalidate'
      });
      return res.end(JSON.stringify({ ok: true, ...data }));
    } catch (err) {
      console.error('[ERRO /api/sync]:', err.message);
      res.writeHead(500, { 'Content-Type': 'application/json; charset=utf-8' });
      return res.end(JSON.stringify({ ok: false, error: err.message }));
    }
  }

  // Servir arquivos estáticos do App
  let filePath = pathname === '/' ? path.join(__dirname, 'index.html') : path.join(__dirname, pathname);
  if (!filePath.startsWith(__dirname)) {
    res.writeHead(403);
    return res.end('Forbidden');
  }

  if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
    const ext = path.extname(filePath).toLowerCase();
    const mimeTypes = {
      '.html': 'text/html; charset=utf-8',
      '.js': 'application/javascript; charset=utf-8',
      '.css': 'text/css; charset=utf-8',
      '.json': 'application/json; charset=utf-8',
      '.png': 'image/png',
      '.svg': 'image/svg+xml'
    };
    res.writeHead(200, {
      'Content-Type': mimeTypes[ext] || 'text/plain; charset=utf-8',
      'Cache-Control': 'no-cache, no-store, must-revalidate'
    });
    return fs.createReadStream(filePath).pipe(res);
  }

  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Not Found');
});

// Faz uma sincronização inicial ao subir o servidor
syncFromGoogleSheets()
  .then((d) => console.log(`[INIT SYNC] Planilha Poupança Baby carregada (${d.totalItems} itens, R$ ${d.totalAmount})`))
  .catch((err) => console.warn('[INIT SYNC WARN] Não foi possível sincronizar no boot:', err.message));

server.listen(PORT, () => {
  console.log(`🚀 Gastos Vini App rodando em http://localhost:${PORT}`);
});

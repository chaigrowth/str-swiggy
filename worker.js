/* ═══════════════════════════════════════════════════════════════
   STR WORKER — Parse → Flag → Write to IndexedDB
   Architecture:
   - Parse Excel with dense mode (array-of-arrays, no property enum)
   - Flag each row inline
   - Write to IndexedDB in 2000-row batches
   - Main thread never holds raw rows — only queries IndexedDB on demand
═══════════════════════════════════════════════════════════════ */

importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

const NEG_PAT = [
  /\bac\b|air conditioner|split ac|window ac|refrigerat|fridge|freezer|washing machine|dishwasher|microwave oven|\btelevision\b|\btv\b|led tv|smart tv|\blaptop\b|\bcomputer\b|desktop pc|mobile phone|\bphone\b|smartphone|tablet|iphone|samsung|oneplus/i,
  /\bjob\b|career|vacancy|hiring|salary|recruitment|internship|fresher jobs/i,
  /free download|apk\b|torrent|\bcrack\b|\bhack\b|cheat code|mod apk/i,
  /stock price|share price|\bsensex\b|\bnifty\b|mutual fund|\bipo\b|\btrading\b|\bdemat\b/i,
  /\bhotel\b|\bresort\b|flight ticket|holiday package|tour package|travel package|book cab|cab booking/i,
  /\bschool\b|\bcollege\b|\buniversity\b|\bexam\b|\bsyllabus\b|admit card/i,
  /\bporn\b|\bxxx\b|\bsex\b|adult content|\bescort\b/i,
  /breaking news|today news|cricket score|match score|ipl score|football score/i,
  /petrol price|diesel price|gas cylinder|lpg rate/i,
  /real estate|flat for sale|house for sale|pg accommodation/i,
];
const NEG_R = [
  'Electronics/appliance — zero FMCG relevance',
  'Job/recruitment — no purchase intent',
  'Piracy/software — no commercial relevance',
  'Finance/stock — zero product relevance',
  'Travel query — not relevant to FMCG',
  'Education — no purchase intent',
  'Adult content — negate immediately',
  'News/sports — no product intent',
  'Fuel/utility — completely unrelated',
  'Real estate — no FMCG relevance',
];
const BLEED_PAT = [
  /\bcheap\b|low cost|lowest price|\bdiscount\b|\bcoupon\b|promo code|\bcashback\b|best price|price list/i,
  /\breviews?\b|\bvs\b|\bversus\b|\bcompare\b|which is better|difference between|alternative to/i,
  /\bdiy\b|home.?made|how to make|how to prepare|make at home/i,
  /\bwholesale\b|bulk order|\bb2b\b|\bdistributor\b|\bsupplier\b|\bmanufacturer\b/i,
  /side effects|harmful\b|dangerous\b|\bbanned\b/i,
];
const BLEED_R = [
  'Deal-seeking — low purchase intent',
  'Research/comparison — not ready to convert',
  'DIY intent — not looking to purchase',
  'B2B/wholesale — wrong audience',
  'Negative sentiment — brand safety risk',
];

function flagRow(sq, kw, matchType, meta, burnt, clicks, conv, roi) {
  const s = (sq || '').toLowerCase().trim();
  const k = (kw || '').toLowerCase().trim();
  for (let i = 0; i < NEG_PAT.length; i++)
    if (NEG_PAT[i].test(s)) return { flag: 'NEGATE', reason: NEG_R[i] };
  if (matchType.includes('BROAD') && k && s) {
    const kt = k.split(/\s+/).filter(w => w.length > 2);
    const st = s.split(/\s+/);
    if (kt.length > 0 && st.length > 1 && !kt.some(t => st.some(w => w.includes(t) || t.includes(w))))
      return { flag: 'NEGATE', reason: 'Broad match: zero token overlap with keyword' };
  }
  for (let i = 0; i < BLEED_PAT.length; i++)
    if (BLEED_PAT[i].test(s)) return { flag: 'BLEED', reason: BLEED_R[i] };
  if (burnt > 500 && clicks === 0 && conv === 0)
    return { flag: 'BLEED', reason: `₹${burnt.toFixed(0)} spent — 0 clicks, 0 conversions` };
  if (burnt > 100 && roi > 0 && roi < 0.3)
    return { flag: 'BLEED', reason: `ROI ${roi.toFixed(2)}x — far below break-even` };
  if (meta && s.split(/\s+/).length > 1) {
    const cw = ((meta.category||'')+' '+(meta.franchise||'')+' '+(meta.brand||''))
      .toLowerCase().split(/[\s\/\+\-]+/).filter(w => w.length > 2);
    const sw = s.split(/\s+/);
    if (cw.length > 0 && !cw.some(t => sw.some(w => w.includes(t)||t.includes(w))))
      return { flag: 'IRR', reason: `No overlap with "${meta.category||''}" / "${meta.franchise||''}"` };
  }
  return { flag: 'OK', reason: 'Relevant to campaign context' };
}

/* ── IndexedDB ── */
let db = null;
function openDB() {
  return new Promise((res, rej) => {
    const req = indexedDB.open('STR_DB', 2);
    req.onupgradeneeded = ev => {
      const d = ev.target.result;
      ['rows','meta'].forEach(s => { if (d.objectStoreNames.contains(s)) d.deleteObjectStore(s); });
      const store = d.createObjectStore('rows', { keyPath: 'id' });
      store.createIndex('flag',     'flag',     { unique: false });
      store.createIndex('campaign', 'campaign', { unique: false });
      store.createIndex('category', 'category', { unique: false });
      d.createObjectStore('meta', { keyPath: 'key' });
    };
    req.onsuccess = ev => { db = ev.target.result; res(); };
    req.onerror   = () => rej(req.error);
  });
}
function clearStores() {
  return new Promise((res, rej) => {
    const tx = db.transaction(['rows','meta'], 'readwrite');
    tx.objectStore('rows').clear();
    tx.objectStore('meta').clear();
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}
function writeBatch(rows) {
  return new Promise((res, rej) => {
    const tx = db.transaction('rows', 'readwrite');
    const st = tx.objectStore('rows');
    for (const r of rows) st.put(r);
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}
function putMeta(key, value) {
  return new Promise((res, rej) => {
    const tx = db.transaction('meta', 'readwrite');
    tx.objectStore('meta').put({ key, value });
    tx.oncomplete = res; tx.onerror = () => rej(tx.error);
  });
}

/* ── Main ── */
self.onmessage = async function(e) {
  if (e.data.type !== 'PARSE_STR') return;
  const { buffer, mapData } = e.data;

  const lookup = {};
  (mapData||[]).forEach(r => {
    const n = (r['Campaign Name']||r['CAMPAIGN_NAME']||'').trim().toLowerCase();
    if (n) lookup[n] = {
      brand:     r.Brand     ||r.brand     ||'',
      category:  r.Category  ||r.category  ||'',
      franchise: r.Franchise ||r.franchise ||'',
    };
  });

  self.postMessage({ type: 'STATUS', msg: 'Opening database...' });
  try { await openDB(); } catch(err) { self.postMessage({ type:'ERROR', msg:'IndexedDB: '+err.message }); return; }
  self.postMessage({ type: 'STATUS', msg: 'Clearing old data...' });
  await clearStores();
  self.postMessage({ type: 'STATUS', msg: 'Parsing Excel file...' });

  let wsData;
  try {
    const wb = XLSX.read(buffer, { type:'array', dense:true, cellDates:false, cellNF:false, cellStyles:false, cellHTML:false });
    wsData = wb.Sheets[wb.SheetNames[0]]['!data'];
  } catch(err) { self.postMessage({ type:'ERROR', msg:'Excel parse: '+err.message }); return; }

  if (!wsData || wsData.length < 2) { self.postMessage({ type:'ERROR', msg:'Sheet empty or unreadable' }); return; }

  // Column index map
  const ci = {};
  (wsData[0]||[]).forEach((cell,i) => { if (cell) ci[String(cell.v||'').trim().toUpperCase()] = i; });
  function cv(row, col) { const i=ci[col]; if(i===undefined||!row[i]) return ''; return row[i].v!==undefined?row[i].v:''; }

  const totalRows = wsData.length - 1;
  self.postMessage({ type:'STATUS', msg:`Analyzing ${totalRows.toLocaleString()} rows...` });

  const stats = { total:0, OK:0, BLEED:0, IRR:0, NEGATE:0, totalSpend:0, wastedSpend:0, totalImpr:0, totalClicks:0, totalConv:0 };
  const campMap = new Map();
  const catSet  = new Set();
  const campSet = new Set();
  let batch = [];
  let processed = 0;

  for (let r = 1; r < wsData.length; r++) {
    const row = wsData[r] || [];
    const sq        = String(cv(row,'SEARCH_QUERY')     ||'');
    const kw        = String(cv(row,'KEYWORD')           ||'');
    const matchType = String(cv(row,'MATCH_TYPE')        ||'');
    const cname     = String(cv(row,'CAMPAIGN_NAME')     ||'').trim();
    const burnt     = parseFloat(cv(row,'TOTAL_BUDGET_BURNT'))  ||0;
    const clicks    = parseFloat(cv(row,'TOTAL_CLICKS'))        ||0;
    const conv      = parseFloat(cv(row,'TOTAL_CONVERSIONS'))   ||0;
    const roi       = parseFloat(cv(row,'TOTAL_ROI'))           ||0;
    const impr      = parseFloat(cv(row,'TOTAL_IMPRESSIONS'))   ||0;
    const date      = String(cv(row,'METRICS_DATE')      ||'');

    const meta = lookup[cname.toLowerCase()]||null;
    const { flag:f, reason } = flagRow(sq, kw, matchType, meta, burnt, clicks, conv, roi);

    stats.total++;  stats[f]++;
    stats.totalSpend  += burnt;  stats.totalImpr  += impr;
    stats.totalClicks += clicks; stats.totalConv  += conv;
    if (f !== 'OK') stats.wastedSpend += burnt;

    const cat = meta?.category||'';
    if (cat) catSet.add(cat);
    if (cname) {
      campSet.add(cname);
      if (!campMap.has(cname)) campMap.set(cname, { name:cname, spend:0, OK:0, BLEED:0, IRR:0, NEGATE:0, impr:0, clicks:0, category:cat });
      const cb = campMap.get(cname);
      cb.spend+=burnt; cb[f]++; cb.impr+=impr; cb.clicks+=clicks;
    }

    processed++;
    batch.push({ id:processed, date, sq, kw,
      matchType: matchType.replace('KEYWORD_MATCH_TYPE_',''),
      campaign:cname, category:cat, franchise:meta?.franchise||'', brand:meta?.brand||'',
      burnt, impr, clicks, conv, roi, flag:f, reason });

    if (batch.length >= 2000) {
      await writeBatch(batch); batch=[];
      self.postMessage({ type:'PROGRESS', processed, total:totalRows, pct:Math.round(processed/totalRows*100) });
      await new Promise(res=>setTimeout(res,0));
    }
  }

  if (batch.length) await writeBatch(batch);

  const campArr = Array.from(campMap.values()).sort((a,b)=>b.spend-a.spend);
  await putMeta('stats',      stats);
  await putMeta('campaigns',  campArr);
  await putMeta('categories', Array.from(catSet).sort());
  await putMeta('campNames',  Array.from(campSet).sort());
  await putMeta('totalRows',  processed);

  self.postMessage({ type:'DONE', stats, campaigns:campArr, categories:Array.from(catSet).sort(), campNames:Array.from(campSet).sort(), total:processed });
};

const XLSX = require('xlsx-js-style');
const wb = XLSX.utils.book_new();

// ===== STYLE DEFINITIONS =====
const S = {
  input: { fill: { fgColor: { rgb: 'FFFFF2CC' } }, font: { color: { rgb: 'FF4472C4' }, sz: 10, name: 'Calibri' }, numFmt: '#,##0', border: { bottom: { style: 'thin', color: { rgb: 'FFD9D9D9' } } } },
  inputPct: { fill: { fgColor: { rgb: 'FFFFF2CC' } }, font: { color: { rgb: 'FF4472C4' }, sz: 10, name: 'Calibri' }, numFmt: '0.0%', border: { bottom: { style: 'thin', color: { rgb: 'FFD9D9D9' } } } },
  inputDec: { fill: { fgColor: { rgb: 'FFFFF2CC' } }, font: { color: { rgb: 'FF4472C4' }, sz: 10, name: 'Calibri' }, numFmt: '$#,##0.000', border: { bottom: { style: 'thin', color: { rgb: 'FFD9D9D9' } } } },
  inputCurrency: { fill: { fgColor: { rgb: 'FFFFF2CC' } }, font: { color: { rgb: 'FF4472C4' }, sz: 10, name: 'Calibri' }, numFmt: '$#,##0', border: { bottom: { style: 'thin', color: { rgb: 'FFD9D9D9' } } } },
  inputCurrencyDec: { fill: { fgColor: { rgb: 'FFFFF2CC' } }, font: { color: { rgb: 'FF4472C4' }, sz: 10, name: 'Calibri' }, numFmt: '$#,##0.00', border: { bottom: { style: 'thin', color: { rgb: 'FFD9D9D9' } } } },
  calc: { font: { color: { rgb: 'FF000000' }, sz: 10, name: 'Calibri' }, numFmt: '$#,##0', border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  calcPct: { font: { color: { rgb: 'FF000000' }, sz: 10, name: 'Calibri' }, numFmt: '0.0%', border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  calcNum: { font: { color: { rgb: 'FF000000' }, sz: 10, name: 'Calibri' }, numFmt: '#,##0', border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  calcDec: { font: { color: { rgb: 'FF000000' }, sz: 10, name: 'Calibri' }, numFmt: '#,##0.0', border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  title: { font: { bold: true, sz: 14, name: 'Calibri', color: { rgb: 'FF1A1A1A' } } },
  subtitle: { font: { sz: 10, name: 'Calibri', color: { rgb: 'FF888888' }, italic: true } },
  section: { font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: 'FFFFFFFF' } }, fill: { fgColor: { rgb: 'FF404040' } }, border: { bottom: { style: 'medium', color: { rgb: 'FF404040' } } } },
  subsection: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF333333' } }, fill: { fgColor: { rgb: 'FFF2F2F2' } }, border: { bottom: { style: 'thin', color: { rgb: 'FFD0D0D0' } } } },
  colHeader: { font: { bold: true, sz: 9, name: 'Calibri', color: { rgb: 'FFFFFFFF' } }, fill: { fgColor: { rgb: 'FF4472C4' } }, alignment: { horizontal: 'center', wrapText: true } },
  label: { font: { sz: 10, name: 'Calibri', color: { rgb: 'FF333333' } }, border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  totalLabel: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF1A1A1A' } }, border: { top: { style: 'medium', color: { rgb: 'FF333333' } }, bottom: { style: 'double', color: { rgb: 'FF333333' } } } },
  totalCalc: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF1A1A1A' } }, numFmt: '$#,##0', border: { top: { style: 'medium', color: { rgb: 'FF333333' } }, bottom: { style: 'double', color: { rgb: 'FF333333' } } } },
  totalCalcPct: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF1A1A1A' } }, numFmt: '0.0%', border: { top: { style: 'medium', color: { rgb: 'FF333333' } }, bottom: { style: 'double', color: { rgb: 'FF333333' } } } },
  unit: { font: { sz: 9, name: 'Calibri', color: { rgb: 'FF999999' }, italic: true }, border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  good: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF2D6A4F' } }, fill: { fgColor: { rgb: 'FFD8F3DC' } }, numFmt: '$#,##0', border: { bottom: { style: 'thin', color: { rgb: 'FF52B788' } } } },
  goodText: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FF2D6A4F' } }, fill: { fgColor: { rgb: 'FFD8F3DC' } }, border: { bottom: { style: 'thin', color: { rgb: 'FF52B788' } } } },
  badText: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FFD00000' } }, fill: { fgColor: { rgb: 'FFFDE8E8' } }, border: { bottom: { style: 'thin', color: { rgb: 'FFE57373' } } } },
  recommend: { font: { bold: true, sz: 10, name: 'Calibri', color: { rgb: 'FFFFFFFF' } }, fill: { fgColor: { rgb: 'FF2D6A4F' } } },
  recommendLabel: { font: { bold: true, sz: 11, name: 'Calibri', color: { rgb: 'FFFFFFFF' } }, fill: { fgColor: { rgb: 'FF2D6A4F' } } },
  dash: { font: { sz: 10, name: 'Calibri', color: { rgb: 'FFCCCCCC' } }, alignment: { horizontal: 'center' }, border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
  calcLabel: { font: { sz: 10, name: 'Calibri', color: { rgb: 'FF666666' }, italic: true }, border: { bottom: { style: 'thin', color: { rgb: 'FFE0E0E0' } } } },
};

// ===== HELPERS =====
// Place a hardcoded value
function sv(ws, r, c, val, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = { v: val, t: typeof val === 'number' ? 'n' : 's', s: style };
}
// Place a formula
function sf(ws, r, c, formula, style) {
  const addr = XLSX.utils.encode_cell({ r, c });
  ws[addr] = { f: formula, t: 'n', s: style };
}
// Excel cell ref from 0-indexed row/col
function ref(r, c) { return XLSX.utils.encode_cell({ r, c }); }
// Cross-sheet ref
function iRef(r) { return "Inputs!C" + (r + 1); } // 0-indexed code row -> 1-indexed Excel row, column C
function dRef(r) { return "Inputs!C" + (r + 1); } // derived calcs also on Inputs sheet col C
// Remove gridlines
function noGrid(ws) {
  if (!ws['!sheetViews']) ws['!sheetViews'] = [{}];
  ws['!sheetViews'][0].showGridLines = false;
}
// Set sheet range
function setRange(ws, maxR, maxC) {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } });
}

// ===== INPUT VALUES =====
const I = {
  leads: 90695, contactRate: 0.35, convRate: 0.0196, commPerSale: 13.59,
  bonusPerPeriod: 9693, periodsPerYear: 10, agents: 8, agentCost: 1000,
  smsPerLead: 2, smsCost: 0.018, aht: 1.5, ahtCost: 0.15,
  dialCost: 0.00, dialsPerLead: 6, vtComm: 0.05, buildCost: 50000,
  targetMargin: 0.50, platformFee: 5000, perLeadFee: 0.15,
  expansionPayment: 25000, gracePeriod: 6,
  careCalls: 650000, careChats: 325000, carePrice: 0.75,
  careChatPrice: 0.35, careContainment: 0.50,
};

// ============================================================
// SHEET 1: INPUTS + DERIVED CALCULATIONS
// ============================================================
const ws1 = XLSX.utils.aoa_to_sheet([[]]);
ws1['!cols'] = [{ wch: 50 }, { wch: 4 }, { wch: 18 }, { wch: 14 }];
noGrid(ws1);

// Track row positions for references (0-indexed code rows)
const R = {}; // R.leads = code row where leads value lives, in column C (index 2)
let r = 0;

sv(ws1, r, 0, 'CLARO DR — COMMERCIAL OPTIONS MODEL', S.title); r++;
sv(ws1, r, 0, 'Versa AI | Confidential | April 2026', S.subtitle); r++;
r++;

// --- PILOT PARAMETERS ---
sv(ws1, r, 0, 'PILOT PARAMETERS', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Unit', S.section); r++;
const inputRows1 = [
  ['Leads per period (5 weeks)', I.leads, 'leads', S.input, 'leads'],
  ['Contact rate (VERSA automated)', I.contactRate, '%', S.inputPct, 'contactRate'],
  ['Blended conversion rate', I.convRate, '%', S.inputPct, 'convRate'],
  ['Blended commission per sale', I.commPerSale, 'USD', S.inputCurrencyDec, 'commPerSale'],
  ['Bonus revenue per period', I.bonusPerPeriod, 'USD', S.inputCurrency, 'bonusPerPeriod'],
  ['Campaign periods per year', I.periodsPerYear, 'periods', S.input, 'periodsPerYear'],
  ['Agent HC (post-VERSA)', I.agents, 'heads', S.input, 'agents'],
  ['Agent cost per period', I.agentCost, 'USD/head', S.inputCurrency, 'agentCost'],
];
inputRows1.forEach(([label, val, unit, style, key]) => {
  sv(ws1, r, 0, label, S.label); sv(ws1, r, 2, val, style); sv(ws1, r, 3, unit, S.unit);
  R[key] = r; r++;
});
r++;

// --- VERSA PLATFORM COSTS ---
sv(ws1, r, 0, 'VERSA PLATFORM COSTS', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Unit', S.section); r++;
const inputRows2 = [
  ['SMS per lead', I.smsPerLead, 'SMS', S.input, 'smsPerLead'],
  ['SMS cost', I.smsCost, 'USD/SMS', S.inputDec, 'smsCost'],
  ['AHT screening per contact', I.aht, 'min', S.inputDec, 'aht'],
  ['AHT cost per minute', I.ahtCost, 'USD/min', S.inputCurrencyDec, 'ahtCost'],
  ['Dial cost per dial', I.dialCost, 'USD', S.inputCurrencyDec, 'dialCost'],
  ['Dials per lead', I.dialsPerLead, 'dials', S.input, 'dialsPerLead'],
];
inputRows2.forEach(([label, val, unit, style, key]) => {
  sv(ws1, r, 0, label, S.label); sv(ws1, r, 2, val, style); sv(ws1, r, 3, unit, S.unit);
  R[key] = r; r++;
});
r++;

// --- COMMERCIAL LEVERS ---
sv(ws1, r, 0, 'COMMERCIAL LEVERS', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Unit', S.section); r++;
const inputRows3 = [
  ['VT commission (% of topline)', I.vtComm, '%', S.inputPct, 'vtComm'],
  ['VERSA build cost (one-off)', I.buildCost, 'USD', S.inputCurrency, 'buildCost'],
  ['Target gross margin (VERSA)', I.targetMargin, '%', S.inputPct, 'targetMargin'],
];
inputRows3.forEach(([label, val, unit, style, key]) => {
  sv(ws1, r, 0, label, S.label); sv(ws1, r, 2, val, style); sv(ws1, r, 3, unit, S.unit);
  R[key] = r; r++;
});
r++;

// --- OPTION-SPECIFIC ---
sv(ws1, r, 0, 'OPTION-SPECIFIC INPUTS', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Unit', S.section); r++;
const inputRows4 = [
  ['Option B: Monthly platform fee', I.platformFee, 'USD/month', S.inputCurrency, 'platformFee'],
  ['Option C: Per-lead processing fee', I.perLeadFee, 'USD/lead', S.inputCurrencyDec, 'perLeadFee'],
  ['Option D: Monthly pre-payment (if no care expansion)', I.expansionPayment, 'USD/month', S.inputCurrency, 'expansionPayment'],
  ['Option D: Grace period before pre-payment', I.gracePeriod, 'months', S.input, 'gracePeriod'],
];
inputRows4.forEach(([label, val, unit, style, key]) => {
  sv(ws1, r, 0, label, S.label); sv(ws1, r, 2, val, style); sv(ws1, r, 3, unit, S.unit);
  R[key] = r; r++;
});
r++;

// --- INBOUND CARE ---
sv(ws1, r, 0, 'INBOUND CARE ESTIMATES', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Unit', S.section); r++;
const inputRows5 = [
  ['Est. Claro DR inbound calls/month', I.careCalls, 'calls', S.input, 'careCalls'],
  ['Est. Claro DR inbound chats/month', I.careChats, 'chats', S.input, 'careChats'],
  ['Price per billable care call', I.carePrice, 'USD', S.inputCurrencyDec, 'carePrice'],
  ['Price per billable chat', I.careChatPrice, 'USD', S.inputCurrencyDec, 'careChatPrice'],
  ['Est. containment rate', I.careContainment, '%', S.inputPct, 'careContainment'],
];
inputRows5.forEach(([label, val, unit, style, key]) => {
  sv(ws1, r, 0, label, S.label); sv(ws1, r, 2, val, style); sv(ws1, r, 3, unit, S.unit);
  R[key] = r; r++;
});
r++; r++;

// === DERIVED CALCULATIONS (formulas on Inputs sheet) ===
sv(ws1, r, 0, 'DERIVED CALCULATIONS', S.section); sv(ws1, r, 2, 'Value', S.section); sv(ws1, r, 3, 'Formula', S.section); r++;

// Helper: iR(key) returns "C{excel_row}" for referencing an input on this sheet
function iR(key) { return 'C' + (R[key] + 1); }

const derivedRows = [
  ['Contacts per period', `${iR('leads')}*${iR('contactRate')}`, '#,##0', 'contacts'],
  ['Sales per period', `${iR('contacts')}*${iR('convRate')}`, '#,##0', 'sales'],
  ['Base commission revenue', `${iR('sales')}*${iR('commPerSale')}`, '$#,##0', 'baseRevenue'],
  ['Total revenue (incl. bonuses)', `${iR('baseRevenue')}+${iR('bonusPerPeriod')}`, '$#,##0', 'totalRevenue'],
  ['Agent cost total', `${iR('agents')}*${iR('agentCost')}`, '$#,##0', 'agentCostTotal'],
  ['SMS cost total', `${iR('leads')}*${iR('smsPerLead')}*${iR('smsCost')}`, '$#,##0', 'smsCostTotal'],
  ['Dial cost total', `${iR('leads')}*${iR('dialsPerLead')}*${iR('dialCost')}`, '$#,##0', 'dialCostTotal'],
  ['AHT screening cost total', `${iR('contacts')}*${iR('aht')}*${iR('ahtCost')}`, '$#,##0', 'ahtCostTotal'],
  ['VERSA platform costs', `${iR('smsCostTotal')}+${iR('dialCostTotal')}+${iR('ahtCostTotal')}`, '$#,##0', 'versaPlatformCosts'],
  ['VT commission fee', `${iR('totalRevenue')}*${iR('vtComm')}`, '$#,##0', 'vtFee'],
  ['VT total take (agents + fee)', `${iR('agentCostTotal')}+${iR('vtFee')}`, '$#,##0', 'vtTotal'],
  ['Platform fee per period (5wk)', `${iR('platformFee')}*(5*7/30)`, '$#,##0', 'platformFeePeriod'],
  ['Lead fee revenue per period', `${iR('leads')}*${iR('perLeadFee')}`, '$#,##0', 'leadFeeRevenue'],
  ['Months of pre-payment in Y1', `MAX(0,12-${iR('gracePeriod')})`, '#,##0', 'monthsPrePay'],
  ['Expansion guarantee Y1', `${iR('expansionPayment')}*${iR('monthsPrePay')}`, '$#,##0', 'expansionY1'],
  ['Billable care calls/month', `${iR('careCalls')}*${iR('careContainment')}`, '#,##0', 'billableCare'],
  ['Billable chats/month', `${iR('careChats')}*${iR('careContainment')}`, '#,##0', 'billableChats'],
  ['Care call revenue/month', `${iR('billableCare')}*${iR('carePrice')}`, '$#,##0', 'careCallRevM'],
  ['Care chat revenue/month', `${iR('billableChats')}*${iR('careChatPrice')}`, '$#,##0', 'careChatRevM'],
  ['Care total revenue/month', `${iR('careCallRevM')}+${iR('careChatRevM')}`, '$#,##0', 'careRevM'],
  ['Care total revenue/year', `${iR('careRevM')}*12`, '$#,##0', 'careRevA'],
  ['Care margin/year (65%)', `${iR('careRevA')}*0.65`, '$#,##0', 'careMarginA'],
  ['Care build cost (3.5x)', `${iR('buildCost')}*3.5`, '$#,##0', 'careBuildCost'],
  // Option A per-period
  ['Opt A: VERSA revenue/period', `${iR('totalRevenue')}-${iR('vtTotal')}`, '$#,##0', 'oA_rev'],
  ['Opt A: VERSA margin/period', `${iR('oA_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oA_margin'],
  ['Opt A: Margin %', `IF(${iR('oA_rev')}>0,${iR('oA_margin')}/${iR('oA_rev')},0)`, '0.0%', 'oA_marginPct'],
  ['Opt A: Annual revenue', `${iR('oA_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_annRev'],
  ['Opt A: Annual margin', `${iR('oA_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_annMargin'],
  ['Opt A: Year 1 net', `${iR('oA_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oA_y1'],
  ['Opt A: Payback (periods)', `IF(${iR('oA_margin')}>0,${iR('buildCost')}/${iR('oA_margin')},"Never")`, '#,##0.0', 'oA_payback'],
  // Option B
  ['Opt B: VERSA revenue/period', `${iR('totalRevenue')}-${iR('vtTotal')}+${iR('platformFeePeriod')}`, '$#,##0', 'oB_rev'],
  ['Opt B: VERSA margin/period', `${iR('oB_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oB_margin'],
  ['Opt B: Margin %', `IF(${iR('oB_rev')}>0,${iR('oB_margin')}/${iR('oB_rev')},0)`, '0.0%', 'oB_marginPct'],
  ['Opt B: Annual revenue', `${iR('oB_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oB_annRev'],
  ['Opt B: Annual margin', `${iR('oB_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oB_annMargin'],
  ['Opt B: Year 1 net', `${iR('oB_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oB_y1'],
  ['Opt B: Payback (periods)', `IF(${iR('oB_margin')}>0,${iR('buildCost')}/${iR('oB_margin')},"Never")`, '#,##0.0', 'oB_payback'],
  // Option C
  ['Opt C: VERSA revenue/period', `${iR('totalRevenue')}-${iR('vtTotal')}+${iR('leadFeeRevenue')}`, '$#,##0', 'oC_rev'],
  ['Opt C: VERSA margin/period', `${iR('oC_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oC_margin'],
  ['Opt C: Margin %', `IF(${iR('oC_rev')}>0,${iR('oC_margin')}/${iR('oC_rev')},0)`, '0.0%', 'oC_marginPct'],
  ['Opt C: Annual revenue', `${iR('oC_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oC_annRev'],
  ['Opt C: Annual margin', `${iR('oC_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oC_annMargin'],
  ['Opt C: Year 1 net', `${iR('oC_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oC_y1'],
  ['Opt C: Payback (periods)', `IF(${iR('oC_margin')}>0,${iR('buildCost')}/${iR('oC_margin')},"Never")`, '#,##0.0', 'oC_payback'],
  // Option D (same per-period as A, but annual includes expansion guarantee)
  ['Opt D: VERSA revenue/period', `${iR('oA_rev')}`, '$#,##0', 'oD_rev'],
  ['Opt D: VERSA margin/period', `${iR('oA_margin')}`, '$#,##0', 'oD_margin'],
  ['Opt D: Margin %', `${iR('oA_marginPct')}`, '0.0%', 'oD_marginPct'],
  ['Opt D: Annual revenue', `${iR('oA_annRev')}+${iR('expansionY1')}`, '$#,##0', 'oD_annRev'],
  ['Opt D: Annual margin', `${iR('oA_annMargin')}+${iR('expansionY1')}`, '$#,##0', 'oD_annMargin'],
  ['Opt D: Blended margin %', `IF(${iR('oD_annRev')}>0,${iR('oD_annMargin')}/${iR('oD_annRev')},0)`, '0.0%', 'oD_annMarginPct'],
  ['Opt D: Year 1 net', `${iR('oD_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oD_y1'],
  ['Opt D: Payback (periods)', `${iR('oA_payback')}`, '#,##0.0', 'oD_payback'],
  // Client costs
  ['Opt A: Claro annual cost', `${iR('totalRevenue')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_claroCost'],
  ['Opt B: Claro annual cost', `(${iR('totalRevenue')}+${iR('platformFeePeriod')})*${iR('periodsPerYear')}`, '$#,##0', 'oB_claroCost'],
  ['Opt C: Claro annual cost', `(${iR('totalRevenue')}+${iR('leadFeeRevenue')})*${iR('periodsPerYear')}`, '$#,##0', 'oC_claroCost'],
  ['Opt D: Claro annual cost', `${iR('oA_claroCost')}+${iR('expansionY1')}`, '$#,##0', 'oD_claroCost'],
  // Scale factor
  ['Scale factor (200 seats)', `200/${iR('agents')}/(17/${iR('agents')})`, '#,##0.0', 'sf200'],
];

// We need to register the row for each derived key BEFORE building formulas
// since some formulas reference other derived rows. So first pass: assign rows.
let derivedStart = r;
derivedRows.forEach(([label, formula, fmt, key], i) => {
  R[key] = derivedStart + i;
});

// Second pass: now iR(key) works for derived keys too, rebuild formulas
derivedRows.forEach(([label, formulaTemplate, fmt, key], i) => {
  const cr = derivedStart + i;
  // Rebuild formula with correct references (iR uses R which is now populated)
  sv(ws1, cr, 0, label, S.calcLabel);
  // For the formula, we need to re-evaluate iR calls. Since JS already interpolated them
  // during array creation above, and R was populated for inputs but not derived at that point,
  // we need a different approach. Let's just re-generate the formula string.
});

// Actually, the issue is that when the derivedRows array was created, iR() was called
// for derived keys that didn't exist yet in R. Let me fix this by building formulas lazily.

// Clear and rebuild properly
const derivedDefs = [
  ['Contacts per period', (iR) => `${iR('leads')}*${iR('contactRate')}`, '#,##0', 'contacts'],
  ['Sales per period', (iR) => `${iR('contacts')}*${iR('convRate')}`, '#,##0', 'sales'],
  ['Base commission revenue', (iR) => `${iR('sales')}*${iR('commPerSale')}`, '$#,##0', 'baseRevenue'],
  ['Total revenue (incl. bonuses)', (iR) => `${iR('baseRevenue')}+${iR('bonusPerPeriod')}`, '$#,##0', 'totalRevenue'],
  ['Agent cost total', (iR) => `${iR('agents')}*${iR('agentCost')}`, '$#,##0', 'agentCostTotal'],
  ['SMS cost total', (iR) => `${iR('leads')}*${iR('smsPerLead')}*${iR('smsCost')}`, '$#,##0', 'smsCostTotal'],
  ['Dial cost total', (iR) => `${iR('leads')}*${iR('dialsPerLead')}*${iR('dialCost')}`, '$#,##0', 'dialCostTotal'],
  ['AHT screening cost total', (iR) => `${iR('contacts')}*${iR('aht')}*${iR('ahtCost')}`, '$#,##0', 'ahtCostTotal'],
  ['VERSA platform costs', (iR) => `${iR('smsCostTotal')}+${iR('dialCostTotal')}+${iR('ahtCostTotal')}`, '$#,##0', 'versaPlatformCosts'],
  ['VT commission fee', (iR) => `${iR('totalRevenue')}*${iR('vtComm')}`, '$#,##0', 'vtFee'],
  ['VT total take (agents + fee)', (iR) => `${iR('agentCostTotal')}+${iR('vtFee')}`, '$#,##0', 'vtTotal'],
  ['Platform fee per period (5wk)', (iR) => `${iR('platformFee')}*(5*7/30)`, '$#,##0', 'platformFeePeriod'],
  ['Lead fee revenue per period', (iR) => `${iR('leads')}*${iR('perLeadFee')}`, '$#,##0', 'leadFeeRevenue'],
  ['Months of pre-payment in Y1', (iR) => `MAX(0,12-${iR('gracePeriod')})`, '#,##0', 'monthsPrePay'],
  ['Expansion guarantee Y1', (iR) => `${iR('expansionPayment')}*${iR('monthsPrePay')}`, '$#,##0', 'expansionY1'],
  ['Billable care calls/month', (iR) => `${iR('careCalls')}*${iR('careContainment')}`, '#,##0', 'billableCare'],
  ['Billable chats/month', (iR) => `${iR('careChats')}*${iR('careContainment')}`, '#,##0', 'billableChats'],
  ['Care call revenue/month', (iR) => `${iR('billableCare')}*${iR('carePrice')}`, '$#,##0', 'careCallRevM'],
  ['Care chat revenue/month', (iR) => `${iR('billableChats')}*${iR('careChatPrice')}`, '$#,##0', 'careChatRevM'],
  ['Care total revenue/month', (iR) => `${iR('careCallRevM')}+${iR('careChatRevM')}`, '$#,##0', 'careRevM'],
  ['Care total revenue/year', (iR) => `${iR('careRevM')}*12`, '$#,##0', 'careRevA'],
  ['Care margin/year (65%)', (iR) => `${iR('careRevA')}*0.65`, '$#,##0', 'careMarginA'],
  ['Care build cost (3.5x)', (iR) => `${iR('buildCost')}*3.5`, '$#,##0', 'careBuildCost'],
  // Option A
  ['Opt A: VERSA revenue/period', (iR) => `${iR('totalRevenue')}-${iR('vtTotal')}`, '$#,##0', 'oA_rev'],
  ['Opt A: VERSA margin/period', (iR) => `${iR('oA_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oA_margin'],
  ['Opt A: Margin %', (iR) => `IF(${iR('oA_rev')}>0,${iR('oA_margin')}/${iR('oA_rev')},0)`, '0.0%', 'oA_marginPct'],
  ['Opt A: Annual revenue', (iR) => `${iR('oA_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_annRev'],
  ['Opt A: Annual margin', (iR) => `${iR('oA_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_annMargin'],
  ['Opt A: Year 1 net', (iR) => `${iR('oA_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oA_y1'],
  ['Opt A: Payback (periods)', (iR) => `IF(${iR('oA_margin')}>0,${iR('buildCost')}/${iR('oA_margin')},"Never")`, '#,##0.0', 'oA_payback'],
  // Option B
  ['Opt B: VERSA revenue/period', (iR) => `${iR('totalRevenue')}-${iR('vtTotal')}+${iR('platformFeePeriod')}`, '$#,##0', 'oB_rev'],
  ['Opt B: VERSA margin/period', (iR) => `${iR('oB_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oB_margin'],
  ['Opt B: Margin %', (iR) => `IF(${iR('oB_rev')}>0,${iR('oB_margin')}/${iR('oB_rev')},0)`, '0.0%', 'oB_marginPct'],
  ['Opt B: Annual revenue', (iR) => `${iR('oB_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oB_annRev'],
  ['Opt B: Annual margin', (iR) => `${iR('oB_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oB_annMargin'],
  ['Opt B: Year 1 net', (iR) => `${iR('oB_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oB_y1'],
  ['Opt B: Payback (periods)', (iR) => `IF(${iR('oB_margin')}>0,${iR('buildCost')}/${iR('oB_margin')},"Never")`, '#,##0.0', 'oB_payback'],
  // Option C
  ['Opt C: VERSA revenue/period', (iR) => `${iR('totalRevenue')}-${iR('vtTotal')}+${iR('leadFeeRevenue')}`, '$#,##0', 'oC_rev'],
  ['Opt C: VERSA margin/period', (iR) => `${iR('oC_rev')}-${iR('versaPlatformCosts')}`, '$#,##0', 'oC_margin'],
  ['Opt C: Margin %', (iR) => `IF(${iR('oC_rev')}>0,${iR('oC_margin')}/${iR('oC_rev')},0)`, '0.0%', 'oC_marginPct'],
  ['Opt C: Annual revenue', (iR) => `${iR('oC_rev')}*${iR('periodsPerYear')}`, '$#,##0', 'oC_annRev'],
  ['Opt C: Annual margin', (iR) => `${iR('oC_margin')}*${iR('periodsPerYear')}`, '$#,##0', 'oC_annMargin'],
  ['Opt C: Year 1 net', (iR) => `${iR('oC_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oC_y1'],
  ['Opt C: Payback (periods)', (iR) => `IF(${iR('oC_margin')}>0,${iR('buildCost')}/${iR('oC_margin')},"Never")`, '#,##0.0', 'oC_payback'],
  // Option D
  ['Opt D: VERSA revenue/period', (iR) => `${iR('oA_rev')}`, '$#,##0', 'oD_rev'],
  ['Opt D: VERSA margin/period', (iR) => `${iR('oA_margin')}`, '$#,##0', 'oD_margin'],
  ['Opt D: Annual revenue', (iR) => `${iR('oA_annRev')}+${iR('expansionY1')}`, '$#,##0', 'oD_annRev'],
  ['Opt D: Annual margin', (iR) => `${iR('oA_annMargin')}+${iR('expansionY1')}`, '$#,##0', 'oD_annMargin'],
  ['Opt D: Blended margin %', (iR) => `IF(${iR('oD_annRev')}>0,${iR('oD_annMargin')}/${iR('oD_annRev')},0)`, '0.0%', 'oD_annMarginPct'],
  ['Opt D: Year 1 net', (iR) => `${iR('oD_annMargin')}-${iR('buildCost')}`, '$#,##0', 'oD_y1'],
  ['Opt D: Payback (periods)', (iR) => `${iR('oA_payback')}`, '#,##0.0', 'oD_payback'],
  // Client costs
  ['Opt A: Claro annual cost', (iR) => `${iR('totalRevenue')}*${iR('periodsPerYear')}`, '$#,##0', 'oA_claroCost'],
  ['Opt B: Claro annual cost', (iR) => `(${iR('totalRevenue')}+${iR('platformFeePeriod')})*${iR('periodsPerYear')}`, '$#,##0', 'oB_claroCost'],
  ['Opt C: Claro annual cost', (iR) => `(${iR('totalRevenue')}+${iR('leadFeeRevenue')})*${iR('periodsPerYear')}`, '$#,##0', 'oC_claroCost'],
  ['Opt D: Claro annual cost', (iR) => `${iR('oA_claroCost')}+${iR('expansionY1')}`, '$#,##0', 'oD_claroCost'],
  // Hybrid B+D
  ['Hybrid B+D: Annual margin', (iR) => `${iR('oB_annMargin')}+${iR('expansionY1')}`, '$#,##0', 'hybrid_annMargin'],
  ['Hybrid B+D: Year 1 net', (iR) => `${iR('hybrid_annMargin')}-${iR('buildCost')}`, '$#,##0', 'hybrid_y1'],
  ['Hybrid B+D: Year 2+ potential', (iR) => `${iR('oB_annMargin')}+${iR('careRevA')}`, '$#,##0', 'hybrid_y2'],
  // Scale
  ['Scale factor (200/17 seats)', (iR) => `200/17`, '#,##0.0', 'sf200'],
];

// First pass: assign row numbers
derivedDefs.forEach(([, , , key], i) => { R[key] = derivedStart + i; });

// Second pass: write formulas with correct iR references
derivedDefs.forEach(([label, formulaFn, fmt, key], i) => {
  const cr = derivedStart + i;
  const formulaStr = formulaFn(iR);
  const isPct = fmt.includes('%');
  sv(ws1, cr, 0, label, S.calcLabel);
  const style = isPct
    ? { ...S.calcPct, numFmt: fmt }
    : { ...S.calc, numFmt: fmt };
  sf(ws1, cr, 2, formulaStr, style);
});
r = derivedStart + derivedDefs.length;

setRange(ws1, r, 3);
XLSX.utils.book_append_sheet(wb, ws1, 'Inputs');

// ============================================================
// Helper for cross-sheet formula referencing Inputs!C{row}
// ============================================================
function iF(key) { return 'Inputs!C' + (R[key] + 1); }

// ============================================================
// SHEET 2: OPTION COMPARISON (all formulas)
// ============================================================
const ws2 = XLSX.utils.aoa_to_sheet([[]]);
ws2['!cols'] = [{ wch: 40 }, { wch: 22 }, { wch: 26 }, { wch: 26 }, { wch: 32 }];
noGrid(ws2);
r = 0;
sv(ws2, r, 0, 'SIDE-BY-SIDE OPTION COMPARISON', S.title); r++;
sv(ws2, r, 0, 'All figures USD unless stated', S.subtitle); r++;
r++;

['Metric', 'A: Pure Commission', 'B: Commission +\nPlatform Fee', 'C: Per-Lead Fee +\nCommission', 'D: Commission +\nExpansion Guarantee'].forEach((h, c) => {
  sv(ws2, r, c, h, c === 0 ? { ...S.colHeader, alignment: { horizontal: 'left' } } : S.colHeader);
});
r++;

sv(ws2, r, 0, 'PER PERIOD (5 WEEKS)', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws2, r, c, '', S.subsection);
r++;

// Row builder with formulas
function compRow(ws, row, label, keys, style, labelStyle) {
  sv(ws, row, 0, label, labelStyle || S.label);
  keys.forEach((k, c) => sf(ws, row, c + 1, iF(k), style));
}

compRow(ws2, r, 'VERSA Revenue', ['oA_rev', 'oB_rev', 'oC_rev', 'oD_rev'], S.calc); r++;
compRow(ws2, r, 'VERSA Platform Costs', ['versaPlatformCosts', 'versaPlatformCosts', 'versaPlatformCosts', 'versaPlatformCosts'], S.calc); r++;
compRow(ws2, r, 'VERSA Margin', ['oA_margin', 'oB_margin', 'oC_margin', 'oD_margin'], S.totalCalc, S.totalLabel); r++;
compRow(ws2, r, 'VERSA Gross Margin %', ['oA_marginPct', 'oB_marginPct', 'oC_marginPct', 'oA_marginPct'], S.calcPct); r++;
r++;

sv(ws2, r, 0, 'ANNUAL', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws2, r, c, '', S.subsection);
r++;

compRow(ws2, r, 'VERSA Annual Revenue', ['oA_annRev', 'oB_annRev', 'oC_annRev', 'oD_annRev'], S.calc); r++;
compRow(ws2, r, 'VERSA Annual Margin (before build)', ['oA_annMargin', 'oB_annMargin', 'oC_annMargin', 'oD_annMargin'], S.calc); r++;
compRow(ws2, r, 'Build Cost Recovery (periods)', ['oA_payback', 'oB_payback', 'oC_payback', 'oD_payback'], S.calcDec); r++;
compRow(ws2, r, 'Year 1 Net Margin (after build)', ['oA_y1', 'oB_y1', 'oC_y1', 'oD_y1'], S.totalCalc, S.totalLabel); r++;
compRow(ws2, r, 'Blended Annual Margin %', ['oA_marginPct', 'oB_marginPct', 'oC_marginPct', 'oD_annMarginPct'], S.totalCalcPct, S.totalLabel); r++;
r++;

sv(ws2, r, 0, 'CLARO COST', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws2, r, c, '', S.subsection);
r++;
compRow(ws2, r, 'Total Cost to Claro (annual)', ['oA_claroCost', 'oB_claroCost', 'oC_claroCost', 'oD_claroCost'], S.calc); r++;

// Meets target - formula-based
sv(ws2, r, 0, 'Meets 50% Margin Target?', S.totalLabel);
['oA_marginPct', 'oB_marginPct', 'oC_marginPct', 'oD_annMarginPct'].forEach((k, c) => {
  sf(ws2, r, c + 1, `IF(${iF(k)}>=${iF('targetMargin')},"YES","NO")`, S.calc);
});
r++;

setRange(ws2, r, 4);
XLSX.utils.book_append_sheet(wb, ws2, 'Option Comparison');

// ============================================================
// SHEET 3: OPTION DETAILS (all formulas)
// ============================================================
const ws3 = XLSX.utils.aoa_to_sheet([[]]);
ws3['!cols'] = [{ wch: 44 }, { wch: 22 }, { wch: 22 }, { wch: 22 }, { wch: 28 }];
noGrid(ws3);
r = 0;
sv(ws3, r, 0, 'OPTION DETAIL BREAKDOWNS', S.title); r++;
sv(ws3, r, 0, 'Per 5-week campaign period', S.subtitle); r++;
r++;

['', 'A: Pure Commission', 'B: + Platform Fee', 'C: + Per-Lead Fee', 'D: + Expansion\nGuarantee'].forEach((h, c) => {
  sv(ws3, r, c, h, c === 0 ? S.label : S.colHeader);
});
r++;

function detailRow(ws, row, label, keys, style, labelStyle, inputStyle) {
  sv(ws, row, 0, label, labelStyle || S.label);
  keys.forEach((k, c) => {
    if (typeof k === 'string') sf(ws, row, c + 1, iF(k), inputStyle || style);
    else sf(ws, row, c + 1, k.f, inputStyle || style); // custom formula
  });
}

sv(ws3, r, 0, 'REVENUE COMPONENTS', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws3, r, c, '', S.subsection); r++;

detailRow(ws3, r, 'Total commission revenue (incl. bonuses)', ['totalRevenue', 'totalRevenue', 'totalRevenue', 'totalRevenue'], S.calc, S.label, S.inputCurrency); r++;
sv(ws3, r, 0, 'Platform fee (per period)', S.label);
sv(ws3, r, 1, 0, S.calc); sf(ws3, r, 2, iF('platformFeePeriod'), S.inputCurrency); sv(ws3, r, 3, 0, S.calc); sv(ws3, r, 4, 0, S.calc); r++;
sv(ws3, r, 0, 'Lead processing fee (per period)', S.label);
sv(ws3, r, 1, 0, S.calc); sv(ws3, r, 2, 0, S.calc); sf(ws3, r, 3, iF('leadFeeRevenue'), S.inputCurrency); sv(ws3, r, 4, 0, S.calc); r++;
sv(ws3, r, 0, 'Expansion guarantee (annualised / 10)', S.label);
sv(ws3, r, 1, 0, S.calc); sv(ws3, r, 2, 0, S.calc); sv(ws3, r, 3, 0, S.calc); sf(ws3, r, 4, `${iF('expansionY1')}/${iF('periodsPerYear')}`, S.inputCurrency); r++;

sv(ws3, r, 0, 'Gross revenue', S.totalLabel);
sf(ws3, r, 1, iF('totalRevenue'), S.totalCalc);
sf(ws3, r, 2, `${iF('totalRevenue')}+${iF('platformFeePeriod')}`, S.totalCalc);
sf(ws3, r, 3, `${iF('totalRevenue')}+${iF('leadFeeRevenue')}`, S.totalCalc);
sf(ws3, r, 4, `${iF('totalRevenue')}+${iF('expansionY1')}/${iF('periodsPerYear')}`, S.totalCalc);
r++; r++;

sv(ws3, r, 0, 'VT TAKE', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws3, r, c, '', S.subsection); r++;
detailRow(ws3, r, 'VT agent cost recovery', ['agentCostTotal', 'agentCostTotal', 'agentCostTotal', 'agentCostTotal'], S.calc); r++;
detailRow(ws3, r, 'VT commission fee', ['vtFee', 'vtFee', 'vtFee', 'vtFee'], S.calc); r++;
detailRow(ws3, r, 'Total VT take', ['vtTotal', 'vtTotal', 'vtTotal', 'vtTotal'], S.totalCalc, S.totalLabel); r++;
r++;

sv(ws3, r, 0, 'VERSA P&L', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws3, r, c, '', S.subsection); r++;
compRow(ws3, r, 'VERSA revenue (after VT)', ['oA_rev', 'oB_rev', 'oC_rev', 'oA_rev'], S.calc); r++;
detailRow(ws3, r, 'SMS costs', ['smsCostTotal', 'smsCostTotal', 'smsCostTotal', 'smsCostTotal'], S.calc); r++;
detailRow(ws3, r, 'Dial costs', ['dialCostTotal', 'dialCostTotal', 'dialCostTotal', 'dialCostTotal'], S.calc); r++;
detailRow(ws3, r, 'AHT screening costs', ['ahtCostTotal', 'ahtCostTotal', 'ahtCostTotal', 'ahtCostTotal'], S.calc); r++;
detailRow(ws3, r, 'Total VERSA platform costs', ['versaPlatformCosts', 'versaPlatformCosts', 'versaPlatformCosts', 'versaPlatformCosts'], S.calc); r++;
compRow(ws3, r, 'VERSA margin (per period)', ['oA_margin', 'oB_margin', 'oC_margin', 'oA_margin'], S.totalCalc, S.totalLabel); r++;
compRow(ws3, r, 'VERSA margin %', ['oA_marginPct', 'oB_marginPct', 'oC_marginPct', 'oA_marginPct'], S.totalCalcPct, S.totalLabel); r++;
r++;

sv(ws3, r, 0, 'ANNUAL', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws3, r, c, '', S.subsection); r++;
compRow(ws3, r, 'VERSA annual revenue', ['oA_annRev', 'oB_annRev', 'oC_annRev', 'oD_annRev'], S.calc); r++;
compRow(ws3, r, 'VERSA annual margin', ['oA_annMargin', 'oB_annMargin', 'oC_annMargin', 'oD_annMargin'], S.calc); r++;
sv(ws3, r, 0, 'Less: build cost', S.label);
for (let c = 1; c <= 4; c++) sf(ws3, r, c, `-${iF('buildCost')}`, S.calc); r++;
compRow(ws3, r, 'Year 1 net margin', ['oA_y1', 'oB_y1', 'oC_y1', 'oD_y1'], S.totalCalc, S.totalLabel); r++;

setRange(ws3, r, 4);
XLSX.utils.book_append_sheet(wb, ws3, 'Option Details');

// ============================================================
// SHEET 4: SCALE SCENARIOS (all formulas)
// ============================================================
const ws4 = XLSX.utils.aoa_to_sheet([[]]);
ws4['!cols'] = [{ wch: 48 }, { wch: 22 }, { wch: 24 }, { wch: 24 }];
noGrid(ws4);
r = 0;
sv(ws4, r, 0, 'SCALE SCENARIOS (ANNUAL)', S.title); r++;
sv(ws4, r, 0, 'Pilot to 200-seat expansion with inbound care', S.subtitle); r++;
r++;

['Metric', 'Pilot (17 seats)', '200-Seat Expansion', '+ Inbound Care'].forEach((h, c) => {
  sv(ws4, r, c, h, c === 0 ? { ...S.colHeader, alignment: { horizontal: 'left' } } : S.colHeader);
});
r++;

sv(ws4, r, 0, 'OUTBOUND SALES', S.subsection);
for (let c = 1; c <= 3; c++) sv(ws4, r, c, '', S.subsection); r++;

sv(ws4, r, 0, 'Leads per period', S.label);
sf(ws4, r, 1, iF('leads'), S.calcNum);
sf(ws4, r, 2, `${iF('leads')}*${iF('sf200')}`, S.calcNum);
sf(ws4, r, 3, `${iF('leads')}*${iF('sf200')}`, S.calcNum); r++;

sv(ws4, r, 0, 'Sales revenue per period', S.label);
sf(ws4, r, 1, iF('totalRevenue'), S.calc);
sf(ws4, r, 2, `${iF('totalRevenue')}*${iF('sf200')}`, S.calc);
sf(ws4, r, 3, `${iF('totalRevenue')}*${iF('sf200')}`, S.calc); r++;

sv(ws4, r, 0, 'VERSA margin per period (sales)', S.label);
sf(ws4, r, 1, iF('oA_margin'), S.calc);
sf(ws4, r, 2, `${iF('oA_margin')}*${iF('sf200')}`, S.calc);
sf(ws4, r, 3, `${iF('oA_margin')}*${iF('sf200')}`, S.calc); r++;

sv(ws4, r, 0, 'Annual VERSA margin (sales)', S.totalLabel);
sf(ws4, r, 1, iF('oA_annMargin'), S.totalCalc);
sf(ws4, r, 2, `${iF('oA_annMargin')}*${iF('sf200')}`, S.totalCalc);
sf(ws4, r, 3, `${iF('oA_annMargin')}*${iF('sf200')}`, S.totalCalc); r++;
r++;

sv(ws4, r, 0, 'INBOUND CARE (VOICE + CHAT)', S.subsection);
for (let c = 1; c <= 3; c++) sv(ws4, r, c, '', S.subsection); r++;

const careRows = [
  ['Monthly inbound calls', 'careCalls', S.calcNum],
  ['Monthly inbound chats', 'careChats', S.calcNum],
  ['Billable calls (50% containment)', 'billableCare', S.calcNum],
  ['Billable chats (50% containment)', 'billableChats', S.calcNum],
  ['Care revenue / month', 'careRevM', S.calc],
  ['Care revenue / year', 'careRevA', S.calc],
];
careRows.forEach(([label, key, style]) => {
  sv(ws4, r, 0, label, S.label);
  sv(ws4, r, 1, '—', S.dash); sv(ws4, r, 2, '—', S.dash);
  sf(ws4, r, 3, iF(key), style); r++;
});
sv(ws4, r, 0, 'Care margin / year (est. 65%)', S.totalLabel);
sv(ws4, r, 1, '—', S.dash); sv(ws4, r, 2, '—', S.dash);
sf(ws4, r, 3, iF('careMarginA'), S.totalCalc); r++;
r++;

sv(ws4, r, 0, 'COMBINED', S.subsection);
for (let c = 1; c <= 3; c++) sv(ws4, r, c, '', S.subsection); r++;

sv(ws4, r, 0, 'Total VERSA annual revenue', S.totalLabel);
sf(ws4, r, 1, iF('oA_annRev'), S.totalCalc);
sf(ws4, r, 2, `${iF('oA_annRev')}*${iF('sf200')}`, S.totalCalc);
sf(ws4, r, 3, `${iF('oA_annRev')}*${iF('sf200')}+${iF('careRevA')}`, S.totalCalc); r++;

sv(ws4, r, 0, 'Total VERSA annual margin', S.totalLabel);
sf(ws4, r, 1, iF('oA_annMargin'), S.totalCalc);
sf(ws4, r, 2, `${iF('oA_annMargin')}*${iF('sf200')}`, S.totalCalc);
sf(ws4, r, 3, `${iF('oA_annMargin')}*${iF('sf200')}+${iF('careMarginA')}`, S.totalCalc); r++;

setRange(ws4, r, 3);
XLSX.utils.book_append_sheet(wb, ws4, 'Scale Scenarios');

// ============================================================
// SHEET 5: INBOUND CARE (all formulas)
// ============================================================
const ws5 = XLSX.utils.aoa_to_sheet([[]]);
ws5['!cols'] = [{ wch: 48 }, { wch: 30 }, { wch: 24 }];
noGrid(ws5);
r = 0;
sv(ws5, r, 0, 'INBOUND CARE — THE REAL PRIZE', S.title); r++;
sv(ws5, r, 0, 'Based on ~6.5M subscribers, ~15% monthly contact rate', S.subtitle); r++;
r++;

sv(ws5, r, 0, 'VOICE', S.subsection); sv(ws5, r, 1, '', S.subsection); sv(ws5, r, 2, '', S.subsection); r++;
sv(ws5, r, 0, 'Estimated inbound calls/month', S.label); sf(ws5, r, 1, iF('careCalls'), S.inputCurrency); sv(ws5, r, 2, 'calls', S.unit); r++;
sv(ws5, r, 0, 'Billable calls (containment)', S.label); sf(ws5, r, 1, iF('billableCare'), S.calcNum); sv(ws5, r, 2, 'calls', S.unit); r++;
sv(ws5, r, 0, 'Price per billable call', S.label); sf(ws5, r, 1, iF('carePrice'), S.inputCurrencyDec); sv(ws5, r, 2, 'USD', S.unit); r++;
sv(ws5, r, 0, 'Voice revenue / month', S.label); sf(ws5, r, 1, iF('careCallRevM'), S.calc); sv(ws5, r, 2, 'USD', S.unit); r++;
r++;

sv(ws5, r, 0, 'CHAT', S.subsection); sv(ws5, r, 1, '', S.subsection); sv(ws5, r, 2, '', S.subsection); r++;
sv(ws5, r, 0, 'Estimated inbound chats/month', S.label); sf(ws5, r, 1, iF('careChats'), S.inputCurrency); sv(ws5, r, 2, 'chats', S.unit); r++;
sv(ws5, r, 0, 'Billable chats (containment)', S.label); sf(ws5, r, 1, iF('billableChats'), S.calcNum); sv(ws5, r, 2, 'chats', S.unit); r++;
sv(ws5, r, 0, 'Price per billable chat', S.label); sf(ws5, r, 1, iF('careChatPrice'), S.inputCurrencyDec); sv(ws5, r, 2, 'USD', S.unit); r++;
sv(ws5, r, 0, 'Chat revenue / month', S.label); sf(ws5, r, 1, iF('careChatRevM'), S.calc); sv(ws5, r, 2, 'USD', S.unit); r++;
r++;

sv(ws5, r, 0, 'COMBINED', S.subsection); sv(ws5, r, 1, '', S.subsection); sv(ws5, r, 2, '', S.subsection); r++;
sv(ws5, r, 0, 'Total monthly revenue', S.totalLabel); sf(ws5, r, 1, iF('careRevM'), S.totalCalc); r++;
sv(ws5, r, 0, 'Total annual revenue', S.totalLabel); sf(ws5, r, 1, iF('careRevA'), S.totalCalc); r++;
sv(ws5, r, 0, 'Est. gross margin', S.label); sv(ws5, r, 1, 0.65, S.calcPct); r++;
sv(ws5, r, 0, 'Annual margin', S.totalLabel); sf(ws5, r, 1, iF('careMarginA'), S.totalCalc); r++;
sv(ws5, r, 0, 'Build cost (est. 3.5x sales build)', S.label); sf(ws5, r, 1, iF('careBuildCost'), S.inputCurrency); r++;
r++; r++;

sv(ws5, r, 0, 'SALES vs CARE COMPARISON', S.section); sv(ws5, r, 1, 'Outbound Sales', S.section); sv(ws5, r, 2, 'Inbound Care', S.section); r++;
sv(ws5, r, 0, 'Revenue model', S.label); sv(ws5, r, 1, 'Commission (outcome-dependent)', S.badText); sv(ws5, r, 2, 'Per-call (outcome-independent)', S.goodText); r++;
sv(ws5, r, 0, 'Annual VERSA revenue', S.label); sf(ws5, r, 1, iF('oA_annRev'), S.calc); sf(ws5, r, 2, iF('careRevA'), S.good); r++;
sv(ws5, r, 0, 'Annual VERSA margin', S.label); sf(ws5, r, 1, iF('oA_annMargin'), S.calc); sf(ws5, r, 2, iF('careMarginA'), S.good); r++;
sv(ws5, r, 0, 'Revenue predictability', S.label); sv(ws5, r, 1, 'LOW', S.badText); sv(ws5, r, 2, 'HIGH', S.goodText); r++;
sv(ws5, r, 0, 'Chargeback risk', S.label); sv(ws5, r, 1, 'YES (180-day clawback)', S.badText); sv(ws5, r, 2, 'NONE', S.goodText); r++;
r++;

sv(ws5, r, 0, 'Care is the strategic prize — accept sales pilot only with expansion guarantee', S.recommendLabel);
sv(ws5, r, 1, '', S.recommend); sv(ws5, r, 2, '', S.recommend);

setRange(ws5, r, 2);
XLSX.utils.book_append_sheet(wb, ws5, 'Inbound Care Opportunity');

// ============================================================
// SHEET 6: BOARD DECISION + RECOMMENDATION (all formulas for economics)
// ============================================================
const ws6 = XLSX.utils.aoa_to_sheet([[]]);
ws6['!cols'] = [{ wch: 40 }, { wch: 22 }, { wch: 24 }, { wch: 24 }, { wch: 30 }];
noGrid(ws6);
r = 0;
sv(ws6, r, 0, 'BOARD DECISION MATRIX', S.title); r++;
r++;

['Criterion', 'A: Pure Commission', 'B: + Platform Fee', 'C: + Per-Lead Fee', 'D: + Expansion Guarantee'].forEach((h, c) => {
  sv(ws6, r, c, h, c === 0 ? { ...S.colHeader, alignment: { horizontal: 'left' } } : S.colHeader);
});
r++;

const boardRows = [
  ['Meets margin target', 'NO', 'PARTIAL', 'PARTIAL', 'YES (blended)'],
  ['Downside protection', 'NONE', 'MODERATE', 'MODERATE', 'STRONG'],
  ['Path to care volumes', 'NONE', 'NONE', 'NONE', 'CONTRACTUAL'],
  ['Ease of sell to Claro', 'EASY', 'MODERATE', 'HARD', 'MODERATE'],
  ['Strategic value', 'LOW', 'LOW', 'LOW', 'HIGH'],
  ['Recurring revenue quality', 'LOW', 'MODERATE', 'MODERATE', 'HIGH'],
];
const ratingStyle = v => {
  if (['YES (blended)', 'STRONG', 'CONTRACTUAL', 'HIGH', 'EASY'].includes(v)) return S.goodText;
  if (['NO', 'NONE', 'LOW', 'HARD'].includes(v)) return S.badText;
  return S.label;
};
boardRows.forEach(([l, ...vals]) => {
  sv(ws6, r, 0, l, S.label);
  vals.forEach((v, c) => sv(ws6, r, c + 1, v, ratingStyle(v)));
  r++;
});
r++; r++;

sv(ws6, r, 0, 'RECOMMENDED: OPTION B + D HYBRID', S.recommendLabel);
for (let c = 1; c <= 4; c++) sv(ws6, r, c, '', S.recommend); r++;
sv(ws6, r, 0, 'Combine platform fee (B) with expansion guarantee into care (D)', S.subtitle); r++;
r++;

sv(ws6, r, 0, 'Component', S.colHeader); sv(ws6, r, 1, 'Detail', S.colHeader);
for (let c = 2; c <= 4; c++) sv(ws6, r, c, '', S.colHeader); r++;
[
  ['1. Technology Platform Fee', '$5,000/month — positions VERSA as tech provider'],
  ['2. Commission Pass-Through', 'Standard commission, VT recovers agents + 5% fee'],
  ['3. Expansion Guarantee', 'Care within 6mo, or $25K/mo pre-payment (credited against future care)'],
  ['4. Minimum Lead Supply SLA', '>=90% of pilot volumes per period'],
  ['5. Minimum Commitment', '6 months / 3 periods, early exit = unamortised build'],
  ['6. Right of First Refusal', 'First right on AI expansion across Claro DR / America Movil'],
].forEach(([comp, detail]) => {
  sv(ws6, r, 0, comp, S.label); sv(ws6, r, 1, detail, S.label); r++;
});
r++;

sv(ws6, r, 0, 'YEAR 1 ECONOMICS (B+D HYBRID)', S.subsection);
for (let c = 1; c <= 4; c++) sv(ws6, r, c, '', S.subsection); r++;

sv(ws6, r, 0, 'Sales pilot annual margin', S.label); sf(ws6, r, 1, iF('oB_annMargin'), S.calc); r++;
sv(ws6, r, 0, 'Expansion guarantee (if care not delivered)', S.label); sf(ws6, r, 1, iF('expansionY1'), S.inputCurrency); r++;
sv(ws6, r, 0, 'Total Year 1 VERSA margin (before build)', S.totalLabel); sf(ws6, r, 1, iF('hybrid_annMargin'), S.totalCalc); r++;
sv(ws6, r, 0, 'Build cost', S.label); sf(ws6, r, 1, `0-${iF('buildCost')}`, S.calc); r++;
sv(ws6, r, 0, 'Year 1 net margin', S.totalLabel); sf(ws6, r, 1, iF('hybrid_y1'), S.totalCalc); r++;
sv(ws6, r, 0, 'If care IS delivered: additional annual revenue', S.label); sf(ws6, r, 1, iF('careRevA'), S.good); r++;
sv(ws6, r, 0, 'Combined Year 2+ potential', S.totalLabel); sf(ws6, r, 1, iF('hybrid_y2'), S.good); r++;

setRange(ws6, r, 4);
XLSX.utils.book_append_sheet(wb, ws6, 'Board Decision Matrix');

// ===== WRITE =====
XLSX.writeFile(wb, './Claro_DR_Commercial_Options_Model.xlsx');
console.log('Done — formulas + formatting + no gridlines');

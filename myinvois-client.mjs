/**
 * MyInvois Batch Doctor — Public Client-Side CSV Diagnostic & Diff Preview Module
 * Note: Official 11-Sheet SpreadsheetML (.xls) & full PEPPOL UBL 2.1 batch compilation
 * execute in stateless Cloud RAM gated by license verification.
 */

export const RESTRICTED_MSIC_PREFIXES = [
  { prefix: '451', sector: 'Automotive — Sale of motor vehicles (Table 3.6: Consolidation Prohibited)' },
  { prefix: '511', sector: 'Aviation — Passenger air transport / flight tickets (Table 3.6: Consolidation Prohibited)' },
  { prefix: '410', sector: 'Construction — Buildings & construction contracts (Table 3.6: Consolidation Prohibited)' },
  { prefix: '421', sector: 'Construction — Civil engineering (Table 3.6: Consolidation Prohibited)' },
  { prefix: '4663', sector: 'Construction Materials Wholesale (Table 3.6: Consolidation Prohibited)' },
  { prefix: '4752', sector: 'Hardware / Construction Materials Retail (Table 3.6: Consolidation Prohibited)' },
  { prefix: '351', sector: 'Electricity Supply Provider (Table 3.6: Consolidation Prohibited)' },
  { prefix: '611', sector: 'Telecommunications — Wired/Postpaid (Table 3.6: Consolidation Prohibited)' },
  { prefix: '612', sector: 'Telecommunications — Wireless/Postpaid (Table 3.6: Consolidation Prohibited)' },
  { prefix: '920', sector: 'Licensed Betting & Gaming (Table 3.6: Consolidation Prohibited)' },
];

export function parseCsv(csvText) {
  const lines = [];
  let cur = '';
  let row = [];
  let inQuotes = false;
  const text = String(csvText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') {
        inQuotes = false;
      } else {
        cur += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',' || ch === '\t') {
      row.push(cur.trim());
      cur = '';
    } else if (ch === '\n') {
      row.push(cur.trim());
      if (row.some((cell) => cell !== '')) lines.push(row);
      row = [];
      cur = '';
    } else {
      cur += ch;
    }
  }
  if (cur !== '' || row.length > 0) {
    row.push(cur.trim());
    if (row.some((cell) => cell !== '')) lines.push(row);
  }
  if (lines.length < 2) return { headers: [], rows: [] };
  const headers = lines[0];
  const rows = lines.slice(1).map((vals) => {
    const obj = {};
    headers.forEach((h, idx) => {
      obj[h] = vals[idx] ?? '';
    });
    return obj;
  });
  return { headers, rows };
}

function pickField(row, candidates) {
  const keys = Object.keys(row);
  for (const cand of candidates) {
    const exact = keys.find((k) => k.toLowerCase().trim() === cand.toLowerCase());
    if (exact && row[exact] !== '') return row[exact];
  }
  for (const cand of candidates) {
    const partial = keys.find((k) => k.toLowerCase().includes(cand.toLowerCase()));
    if (partial && row[partial] !== '') return row[partial];
  }
  return '';
}

export function normalizeIsoDate(raw) {
  const s = String(raw || '').trim();
  if (!s) return { date: new Date().toISOString().slice(0, 10), fixed: true };
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return { date: s.slice(0, 10), fixed: false };
  const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (dmy) {
    return { date: `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`, fixed: true };
  }
  const parsed = new Date(s);
  if (!Number.isNaN(parsed.getTime())) {
    return { date: parsed.toISOString().slice(0, 10), fixed: true };
  }
  return { date: new Date().toISOString().slice(0, 10), fixed: true };
}

export function normalizeUom(raw) {
  const u = String(raw || '').trim().toUpperCase();
  if (['C62', 'XUN', 'KGM', 'MTR', 'LTR', 'HUR', 'DAY', 'MON'].includes(u)) return { uom: u, fixed: false };
  if (!u || ['PCS', 'PC', 'PIECE', 'PIECES', 'EA', 'EACH', 'ITEM', 'NOS'].includes(u)) return { uom: 'C62', fixed: u !== 'C62' };
  return { uom: 'XUN', fixed: true };
}

export function cleanMoney(raw) {
  const cleaned = String(raw ?? '0').replace(/[^0-9.\-]/g, '');
  const n = Number(cleaned);
  return Number.isFinite(n) ? Number(n.toFixed(2)) : 0;
}

export function diagnoseAndRepairBatch(rawRows, supplierProfile = {}) {
  const supplier = {
    name: supplierProfile.name || 'SAMPLE MALAYSIA SME SDN BHD',
    tin: supplierProfile.tin || 'C25849102030',
    brn: supplierProfile.brn || '202401009988',
    msic: String(supplierProfile.msic || '47912').padStart(5, '0'),
    mode: supplierProfile.mode || 'CONSOLIDATED',
  };

  const diagnostics = {
    totalInputRows: rawRows.length,
    err236Fixes: 0,
    myiL10TaxFixes: 0,
    dateFormatFixes: 0,
    uomFixes: 0,
    exemptionReasonFixes: 0,
    roundingFixes: 0,
    over10kSplitCount: 0,
    restrictedMsicWarning: null,
    batchChunksCount: 1,
  };

  if (supplier.msic === '00000') {
    diagnostics.restrictedMsicWarning = 'Supplier MSIC cannot be dummy 00000 — select a valid 5-digit MSIC (e.g. 47912 for online retail).';
  } else {
    const hit = RESTRICTED_MSIC_PREFIXES.find((r) => supplier.msic.startsWith(r.prefix));
    if (hit && supplier.mode === 'CONSOLIDATED') {
      diagnostics.restrictedMsicWarning = `LHDN Table 3.6 Alert for MSIC ${supplier.msic}: ${hit.sector}. Switch mode to Individual e-Invoices.`;
    }
  }

  const validConsolidatedRows = [];
  const splitIndividualRows = [];

  rawRows.forEach((row, idx) => {
    const docNo =
      pickField(row, ['Order ID', 'Invoice No', 'Document Number', 'Receipt No', 'Order Number', 'DocNo']) ||
      `DOC-${String(idx + 1).padStart(4, '0')}`;
    const rawDate = pickField(row, ['Order Creation Date', 'Issue Date', 'Date', 'Transaction Date', 'Created Time']);
    const dateNorm = normalizeIsoDate(rawDate);
    if (dateNorm.fixed) diagnostics.dateFormatFixes++;

    const desc =
      pickField(row, ['Product Name', 'Description', 'Item Name', 'Item Description', 'Variation Name']) ||
      `Retail Receipt ${docNo}`;
    const qtyRaw = Number(pickField(row, ['Quantity', 'Qty']) || 1);
    const qty = Number.isFinite(qtyRaw) && qtyRaw > 0 ? qtyRaw : 1;

    const unitPriceRaw = pickField(row, ['Deal Price', 'Unit Price', 'Item Price', 'Price']);
    const totalRaw = pickField(row, ['Total Amount', 'Estimated Total Amount', 'Grand Total', 'Subtotal', 'Amount']);

    let lineSubtotal = 0;
    let unitPrice = 0;
    if (unitPriceRaw !== '') {
      unitPrice = cleanMoney(unitPriceRaw);
      lineSubtotal = Number((unitPrice * qty).toFixed(2));
    } else {
      lineSubtotal = cleanMoney(totalRaw);
      unitPrice = Number((lineSubtotal / qty).toFixed(2));
    }

    const rawUom = pickField(row, ['UOM', 'UnitCode', 'Unit']);
    const uomNorm = normalizeUom(rawUom);
    if (uomNorm.fixed) diagnostics.uomFixes++;

    const origTaxType = String(pickField(row, ['TaxType', 'Tax Type', 'Tax Code']) || '').trim().toUpperCase();
    let taxType = origTaxType;
    const taxAmtRaw = pickField(row, ['Tax Amount', 'SST', 'TaxAmount']);
    let taxAmount = taxAmtRaw !== '' ? cleanMoney(taxAmtRaw) : 0;
    let taxNodeFixed = false;

    if (!['01', '02', '03', '04', '05', '06', 'E'].includes(taxType)) {
      taxType = taxAmount > 0 ? '01' : '06';
      diagnostics.myiL10TaxFixes++;
      taxNodeFixed = true;
    } else if ((taxType === '06' || taxType === 'E') && taxAmtRaw === '') {
      taxAmount = 0;
      diagnostics.myiL10TaxFixes++;
      taxNodeFixed = true;
    }

    let exemptionReason = pickField(row, ['TaxExemptionReason', 'Exemption Reason']);
    let exemptFixed = false;
    if (taxType === 'E' && !exemptionReason) {
      exemptionReason = 'Exempted under Malaysian Sales/Service Tax Exemption Order';
      diagnostics.exemptionReasonFixes++;
      exemptFixed = true;
    }

    const origBuyerTin = pickField(row, ['Buyer TIN', 'BuyerTIN', 'TIN']);
    let buyerTin = origBuyerTin;
    let buyerName = pickField(row, ['Buyer Name', 'Customer Name', 'Buyer Username']) || 'General Public';
    const origClassification = String(pickField(row, ['Classification', 'ClassificationCode', 'Class']) || '').trim();
    let classification = origClassification;

    const grandTotal = Number((lineSubtotal + taxAmount).toFixed(2));
    const isOver10k = grandTotal > 10000;
    let err236FixedOnRow = false;

    if (supplier.mode === 'CONSOLIDATED' && !isOver10k) {
      if (buyerTin !== 'EI00000000010' || classification !== '004') {
        diagnostics.err236Fixes++;
        err236FixedOnRow = true;
      }
      buyerTin = 'EI00000000010';
      buyerName = 'General Public';
      classification = '004';
    } else {
      if (isOver10k && supplier.mode === 'CONSOLIDATED') diagnostics.over10kSplitCount++;
      if (!classification || classification === '004') classification = '022';
      if (!buyerTin) buyerTin = 'EI00000000010';
    }

    const repairs = [];
    if (dateNorm.fixed) repairs.push({ field: 'issueDate' });
    if (uomNorm.fixed) repairs.push({ field: 'uom' });
    if (taxNodeFixed) repairs.push({ field: 'taxType' });
    if (exemptFixed) repairs.push({ field: 'exemptionReason' });
    if (err236FixedOnRow) {
      if (origClassification !== classification) repairs.push({ field: 'classification' });
      if (origBuyerTin !== buyerTin) repairs.push({ field: 'buyerTin' });
    }
    if (isOver10k && supplier.mode === 'CONSOLIDATED') repairs.push({ field: 'routing' });

    const normalizedItem = {
      rowIndex: idx + 1,
      docNo,
      issueDate: dateNorm.date,
      issueTime: '12:00:00Z',
      description: desc,
      qty,
      uom: uomNorm.uom,
      unitPrice,
      lineSubtotal,
      taxType,
      taxAmount,
      exemptionReason: exemptionReason || 'NA',
      grandTotal,
      classification,
      buyerName,
      buyerTin,
      buyerIdType: classification === '004' ? 'NA' : 'BRN',
      buyerIdValue: 'NA',
      over10kFlag: isOver10k,
      rawOriginal: {
        date: rawDate || '',
        uom: rawUom || '',
        taxType: origTaxType || '',
        buyerTin: origBuyerTin || '',
        classification: origClassification || '',
      },
      repairs,
    };

    if (supplier.mode === 'CONSOLIDATED' && !isOver10k) validConsolidatedRows.push(normalizedItem);
    else splitIndividualRows.push(normalizedItem);
  });

  const documents = [];
  if (supplier.mode === 'CONSOLIDATED' && validConsolidatedRows.length > 0) {
    for (let i = 0; i < validConsolidatedRows.length; i += 100) {
      const slice = validConsolidatedRows.slice(i, i + 100);
      const sub = Number(slice.reduce((s, r) => s + r.lineSubtotal, 0).toFixed(2));
      const tax = Number(slice.reduce((s, r) => s + r.taxAmount, 0).toFixed(2));
      documents.push({
        docNo: `CONS-${slice[0].issueDate.slice(0, 7).replace('-', '')}-${String(Math.floor(i / 100) + 1).padStart(2, '0')}`,
        docType: '01',
        issueDate: slice[slice.length - 1].issueDate,
        issueTime: '23:59:00Z',
        currency: 'MYR',
        supplier,
        buyer: { name: 'General Public', tin: 'EI00000000010', idType: 'NA', idValue: 'NA' },
        subtotal: sub,
        taxTotal: tax,
        grandTotal: Number((sub + tax).toFixed(2)),
        lines: slice,
      });
    }
  }

  for (const item of splitIndividualRows) {
    documents.push({
      docNo: item.docNo,
      docType: '01',
      issueDate: item.issueDate,
      issueTime: item.issueTime,
      currency: 'MYR',
      supplier,
      buyer: { name: item.buyerName, tin: item.buyerTin, idType: item.buyerIdType, idValue: item.buyerIdValue },
      subtotal: item.lineSubtotal,
      taxTotal: item.taxAmount,
      grandTotal: item.grandTotal,
      lines: [item],
    });
  }

  diagnostics.batchChunksCount = Math.max(1, Math.ceil(documents.length / 100));
  return {
    supplier,
    diagnostics,
    allNormalizedRows: [...validConsolidatedRows, ...splitIndividualRows],
    documents,
  };
}

export function buildSampleXmlSnippet(doc) {
  if (!doc) return '';
  const sampleLine = doc.lines?.[0];
  return [
    '<!-- Sample 1-Line PEPPOL UBL 2.1 Structure Preview (Full Batch Compiled via Cloud Engine) -->',
    '<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2">',
    `  <cbc:ID>${doc.docNo}</cbc:ID>`,
    `  <cbc:IssueDate>${doc.issueDate}</cbc:IssueDate>`,
    `  <cac:AccountingSupplierParty>...MSIC ${doc.supplier.msic} | TIN ${doc.supplier.tin}...</cac:AccountingSupplierParty>`,
    `  <cac:AccountingCustomerParty>...TIN ${doc.buyer.tin}...</cac:AccountingCustomerParty>`,
    sampleLine
      ? `  <cac:InvoiceLine><cbc:ID>1</cbc:ID><cbc:InvoicedQuantity unitCode="${sampleLine.uom}">${sampleLine.qty.toFixed(2)}</cbc:InvoicedQuantity><cbc:ItemClassificationCode>${sampleLine.classification}</cbc:ItemClassificationCode></cac:InvoiceLine>`
      : '',
    `  <!-- + ${Math.max(0, (doc.lines?.length || 1) - 1)} additional line(s) compiled on export -->`,
    '</Invoice>',
  ].join('\n');
}

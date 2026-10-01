const ExcelJS = require('exceljs');

// Column layout, in the same order as the original Excel register.
// `match` = keywords used to recognise the column header when importing.
const COLUMNS = [
  { key: 'serialNo', header: 'ક્રમ', match: ['ક્રમ', 'sr', 'serial'], width: 6 },
  { key: 'propertyId', header: 'આઈડી', match: ['આઈડી', 'id'], width: 10 },
  { key: 'location', header: 'ગામનું નામ તાલુકો જીલ્લો', match: ['ગામનું', 'village'], width: 22 },
  { key: 'surveyDetails', header: 'સર્વે/બ્લોક નંબરની વિગત', match: ['સર્વે', 'બ્લોક', 'survey'], width: 30 },
  { key: 'deed', header: 'દ.નં. તારીખ', match: ['દ.નં', 'દ. નં', 'તારીખ', 'deed'], width: 18 },
  { key: 'giverNameAddress', header: 'આપનારનું નામ અને સરનામું', match: ['આપનાર', 'giver'], width: 35 },
  { key: 'areaSqm', header: 'મિલકતનું ક્ષેત્રફળ (ચો.મી.)', match: ['મિલકત', 'area'], width: 14 },
  { key: 'in712', header: '૭/૧૨ નામ હા/ના', match: ['૭/૧૨', '7/12', '૭-૧૨'], width: 10 },
  { key: 'propertyCard', header: 'પ્રોપર્ટી કાર્ડ', match: ['પ્રોપર્ટી', 'property card'], width: 12 },
  { key: 'assessment', header: 'આકારણી', match: ['આકાર', 'assessment'], width: 12 },
  { key: 'villageForm2', header: 'ગામ નં. ૨', match: ['ગામ નં', 'form 2'], width: 10 },
  { key: 'purpose', header: 'વપરાશનો હેતુ', match: ['વપરાશ', 'હેતુ', 'purpose'], width: 22 },
  { key: 'isNA', header: 'એન.એ. થયેલ છે? હા/ના', match: ['થયેલ', 'is na'], width: 12 },
  { key: 'naAreaSqm', header: 'એન.એ. જમીનનું ક્ષેત્રફળ', match: ['જમીન', 'na area'], width: 14 },
  { key: 'remarks', header: 'રીમાર્કસ', match: ['રીમાર્ક', 'રિમાર્ક', 'remark'], width: 25 },
];

const GU_DIGITS = '૦૧૨૩૪૫૬૭૮૯';
const toAsciiDigits = (s) => String(s).replace(/[૦-૯]/g, (d) => GU_DIGITS.indexOf(d));
const norm = (s) => toAsciiDigits(s).toLowerCase().replace(/\s+/g, ' ').trim();

function cellText(cell) {
  const v = cell?.value;
  if (v == null) return '';
  if (v instanceof Date) return v;
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map((r) => r.text).join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return v.result instanceof Date ? v.result : String(v.result);
    return '';
  }
  return String(v).trim();
}

function parseYesNo(v) {
  const s = norm(v);
  if (!s) return '';
  if (['હા', 'yes', 'y', 'ha', 'haa'].includes(s)) return 'yes';
  if (['ના', 'no', 'n', 'na', 'naa'].includes(s)) return 'no';
  return '';
}

function parseNumber(v) {
  if (v === '' || v == null) return null;
  const m = toAsciiDigits(v).replace(/,/g, '').match(/-?\d+(\.\d+)?/);
  return m ? Number(m[0]) : null;
}

// "દ.નં. 1234 તા. 12/05/2020" → { deedNo: '1234', deedDate: Date }
function parseDeed(v) {
  if (v instanceof Date) return { deedNo: '', deedDate: v };
  const text = toAsciiDigits(v);
  const m = text.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  let deedDate = null;
  let rest = text;
  if (m) {
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    const d = new Date(Date.UTC(year, Number(m[2]) - 1, Number(m[1])));
    if (!isNaN(d)) {
      deedDate = d;
      rest = text.replace(m[0], '');
    }
  }
  const deedNo = rest.replace(/દ\.?\s*નં\.?|તા\.?|તારીખ|[:,]/g, ' ').replace(/\s+/g, ' ').trim();
  return { deedNo, deedDate };
}

// "ગામ, તાલુકો, જીલ્લો" → split on comma / newline / slash
function parseLocation(v) {
  const parts = String(v).split(/[,\n\/]+/).map((s) => s.trim()).filter(Boolean);
  if (parts.length >= 3) return { village: parts[0], taluka: parts[1], district: parts.slice(2).join(', ') };
  if (parts.length === 2) return { village: parts[0], taluka: parts[1], district: '' };
  return { village: parts[0] || '', taluka: '', district: '' };
}

function findHeaderRow(sheet) {
  for (let r = 1; r <= Math.min(sheet.rowCount, 15); r++) {
    const row = sheet.getRow(r);
    const map = {};
    row.eachCell((cell, col) => {
      const h = norm(cellText(cell));
      if (!h) return;
      const def = COLUMNS.find((c) => !Object.values(map).includes(c.key) && c.match.some((k) => h.includes(norm(k))));
      if (def) map[col] = def.key;
    });
    if (Object.keys(map).length >= 5) return { rowNumber: r, map };
  }
  return null;
}

async function parseWorkbook(buffer) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const sheet = wb.worksheets[0];
  if (!sheet) throw new Error('Excel file has no sheets');

  const header = findHeaderRow(sheet);
  if (!header) throw new Error('Could not find the header row (ક્રમ, આઈડી, ગામનું નામ ...)');

  const records = [];
  for (let r = header.rowNumber + 1; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    const raw = {};
    for (const [col, key] of Object.entries(header.map)) raw[key] = cellText(row.getCell(Number(col)));
    if (Object.values(raw).every((v) => v === '')) continue;

    const rec = {
      serialNo: parseNumber(raw.serialNo),
      propertyId: String(raw.propertyId || ''),
      ...parseLocation(raw.location || ''),
      surveyDetails: String(raw.surveyDetails || ''),
      ...parseDeed(raw.deed || ''),
      giverNameAddress: String(raw.giverNameAddress || ''),
      areaSqm: parseNumber(raw.areaSqm),
      in712: parseYesNo(raw.in712),
      propertyCard: String(raw.propertyCard || ''),
      assessment: String(raw.assessment || ''),
      villageForm2: String(raw.villageForm2 || ''),
      purpose: String(raw.purpose || ''),
      isNA: parseYesNo(raw.isNA),
      naAreaSqm: parseNumber(raw.naAreaSqm),
      remarks: String(raw.remarks || ''),
    };
    if (rec.serialNo == null) delete rec.serialNo;
    records.push(rec);
  }
  return records;
}

const YES_NO_LABEL = { yes: 'હા', no: 'ના', '': '' };

async function buildWorkbook(properties) {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('Properties', { views: [{ state: 'frozen', ySplit: 1 }] });
  sheet.columns = COLUMNS.map((c) => ({ header: c.header, key: c.key, width: c.width }));

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true };
  headerRow.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' };
  headerRow.height = 40;

  for (const p of properties) {
    const d = p.deedDate ? new Date(p.deedDate) : null;
    const dateStr = d
      ? `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`
      : '';
    sheet.addRow({
      serialNo: p.serialNo,
      propertyId: p.propertyId,
      location: [p.village, p.taluka, p.district].filter(Boolean).join(', '),
      surveyDetails: p.surveyDetails,
      deed: [p.deedNo, dateStr && `તા. ${dateStr}`].filter(Boolean).join(' '),
      giverNameAddress: p.giverNameAddress,
      areaSqm: p.areaSqm,
      in712: YES_NO_LABEL[p.in712 || ''],
      propertyCard: p.propertyCard,
      assessment: p.assessment,
      villageForm2: p.villageForm2,
      purpose: p.purpose,
      isNA: YES_NO_LABEL[p.isNA || ''],
      naAreaSqm: p.naAreaSqm,
      remarks: p.remarks,
    });
  }
  sheet.eachRow((row, n) => {
    if (n > 1) row.alignment = { wrapText: true, vertical: 'top' };
  });
  return wb.xlsx.writeBuffer();
}

module.exports = { parseWorkbook, buildWorkbook, toAsciiDigits };

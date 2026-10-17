/**
 * Minimal, dependency-free XLSX reader.
 *
 * XLSX is a ZIP archive of XML parts.  We unzip it with the vendored fflate
 * (MIT) and parse the handful of parts a character-card workbook needs:
 * workbook.xml, workbook.xml.rels, sharedStrings.xml, styles.xml and the
 * worksheet XML.  The reader intentionally returns plain values; character
 * extraction lives in `character.mjs`.
 *
 * Supported: .xlsx / .xlsm (not legacy binary .xls), shared strings, inline
 * strings, cached formula results, booleans, merged cells, hidden sheets are
 * still read, and common Excel date formats are converted to ISO strings.
 */

import { unzipSync } from './vendor/fflate.mjs';

const XML_ENTITIES = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  apos: "'",
};

/** Decode the five XML entities plus numeric character references. */
export function decodeXml(value) {
  return String(value).replace(/&(#x?[0-9a-fA-F]+|amp|lt|gt|quot|apos);/g, (match, body) => {
    if (body[0] === '#') {
      const hex = body[1] === 'x' || body[1] === 'X';
      const code = Number.parseInt(hex ? body.slice(2) : body.slice(1), hex ? 16 : 10);
      if (Number.isFinite(code) && code >= 0 && code <= 0x10ffff) {
        try {
          return String.fromCodePoint(code);
        } catch {
          return match;
        }
      }
      return match;
    }
    return XML_ENTITIES[body] ?? match;
  });
}

/** Read one XML attribute from a start tag. */
function attr(tag, name) {
  const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = new RegExp(`(?:^|\\s)${escaped}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(tag);
  return match === null ? undefined : decodeXml(match[2] ?? match[3] ?? '');
}

/** Return all text runs inside an XML fragment (shared/inline strings). */
function textRuns(fragment) {
  let out = '';
  const re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
  let match;
  while ((match = re.exec(fragment)) !== null) out += decodeXml(match[1]);
  return out;
}

/** Return the first `<v>` value inside a cell fragment. */
function cellValue(fragment) {
  const match = /<v(?:\s[^>]*)?>([\s\S]*?)<\/v>/.exec(fragment);
  return match === null ? undefined : decodeXml(match[1]);
}

/** Convert A1 style column letters to a zero-based column index. */
export function columnIndex(letters) {
  let value = 0;
  for (const ch of String(letters).toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    value = value * 26 + (code - 64);
  }
  return value - 1;
}

/** Convert a zero-based column index to A1 column letters. */
export function columnName(index) {
  let n = index + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

/** Split an A1 reference such as `$B$7` into row/column zero-based indexes. */
function parseRef(ref) {
  const match = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(String(ref).trim());
  if (match === null) return null;
  return { row: Number.parseInt(match[2], 10) - 1, col: columnIndex(match[1]) };
}

/** Parse a range such as `A1:C4` (or a single cell) into bounds. */
function parseRange(ref) {
  const parts = String(ref).split(':');
  const start = parseRef(parts[0] ?? '');
  if (start === null) return null;
  const end = parts.length > 1 ? parseRef(parts[1]) : start;
  if (end === null) return null;
  return {
    startRow: Math.min(start.row, end.row),
    endRow: Math.max(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endCol: Math.max(start.col, end.col),
  };
}

/** Resolve a relationship target against the workbook part directory. */
function resolvePart(base, target) {
  const raw = String(target).replace(/\\/g, '/');
  const joined = raw.startsWith('/') ? raw.slice(1) : `${base}/${raw}`;
  const parts = [];
  for (const segment of joined.split('/')) {
    if (segment === '' || segment === '.') continue;
    if (segment === '..') parts.pop();
    else parts.push(segment);
  }
  return parts.join('/');
}

/** Read text from a zip entry; returns '' when the entry is missing. */
function readPart(files, path) {
  const data = files[path];
  if (data === undefined) return '';
  try {
    return new TextDecoder('utf-8').decode(data);
  } catch {
    return '';
  }
}

/** Built-in Excel number-format ids that display a date or time. */
const BUILTIN_DATE_FORMATS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 30, 36, 45, 46, 47, 50, 57]);

/** Decide whether a cell style displays a numeric value as a date. */
function makeDateStyleLookup(stylesXml) {
  const custom = new Map();
  for (const match of stylesXml.matchAll(/<numFmt\b([^>]*?)\/?>/g)) {
    const id = Number.parseInt(attr(match[1], 'numFmtId') ?? '', 10);
    const code = attr(match[1], 'formatCode') ?? '';
    if (Number.isFinite(id)) custom.set(id, code);
  }
  const xfIds = [];
  const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml);
  if (cellXfs !== null) {
    for (const match of cellXfs[1].matchAll(/<xf\b([^>]*?)(?:\/>|>[\s\S]*?<\/xf>)/g)) {
      xfIds.push(Number.parseInt(attr(match[1], 'numFmtId') ?? '0', 10));
    }
  }
  return (styleIndex) => {
    if (!Number.isFinite(styleIndex) || styleIndex < 0 || styleIndex >= xfIds.length) return false;
    const id = xfIds[styleIndex];
    if (BUILTIN_DATE_FORMATS.has(id)) return true;
    const code = custom.get(id);
    if (!code) return false;
    const cleaned = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
    return /[ymdhs]/i.test(cleaned);
  };
}

/** Convert an Excel 1900-system serial number to an ISO date string. */
function excelSerialToIso(serial) {
  if (!Number.isFinite(serial)) return String(serial);
  const wholeDays = Math.floor(serial);
  const fraction = serial - wholeDays;
  const epoch = Date.UTC(1899, 11, 30);
  const dayAdjust = wholeDays < 60 ? 1 : 0;
  const ms = epoch + (wholeDays + dayAdjust) * 86400000 + Math.round(fraction * 86400000);
  const date = new Date(ms);
  if (!Number.isFinite(date.getTime())) return String(serial);
  const iso = date.toISOString();
  if (iso.endsWith('T00:00:00.000Z')) return iso.slice(0, 10);
  return iso.slice(0, 19).replace('T', ' ');
}

/** Coerce one numeric cell using its style. */
function numberCell(value, styleIndex, isDateStyle) {
  const number = Number(value);
  if (!Number.isFinite(number)) return value;
  if (isDateStyle(styleIndex)) return excelSerialToIso(number);
  return number;
}

/** Parse one worksheet XML into a dense-ish row array. */
function parseSheet(xml, sharedStrings, isDateStyle) {
  const rows = [];
  let maxCol = -1;
  let fallbackRow = 0;

  for (const rowMatch of xml.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const rowAttrs = rowMatch[1] ?? '';
    const rowNumber = Number.parseInt(attr(rowAttrs, 'r') ?? '', 10);
    const rowIndex = Number.isFinite(rowNumber) ? rowNumber - 1 : fallbackRow;
    fallbackRow = rowIndex + 1;
    const body = rowMatch[2] ?? '';
    const row = [];
    let fallbackCol = 0;

    for (const cellMatch of body.matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = cellMatch[1] ?? '';
      const ref = parseRef(attr(attrs, 'r') ?? '');
      const col = ref === null ? fallbackCol : ref.col;
      fallbackCol = col + 1;
      if (col < 0) continue;
      const fragment = cellMatch[2] ?? '';
      const type = attr(attrs, 't') ?? 'n';
      const styleIndex = Number.parseInt(attr(attrs, 's') ?? '0', 10);
      let value;

      if (type === 's') {
        const index = Number.parseInt(cellValue(fragment) ?? '', 10);
        value = Number.isFinite(index) ? sharedStrings[index] ?? '' : '';
      } else if (type === 'inlineStr') {
        value = textRuns(fragment);
      } else if (type === 'str') {
        value = cellValue(fragment) ?? textRuns(fragment);
      } else if (type === 'b') {
        const raw = cellValue(fragment);
        value = raw === '1' ? true : raw === '0' ? false : raw ?? '';
      } else if (type === 'e') {
        value = cellValue(fragment) ?? '';
      } else {
        const raw = cellValue(fragment);
        value = raw === undefined ? '' : numberCell(raw, styleIndex, isDateStyle);
      }

      row[col] = value;
      if (col > maxCol) maxCol = col;
    }

    if (row.length > 0) rows[rowIndex] = row;
  }

  // Fill merged cells with their top-left value.  Character sheets commonly
  // merge the value column across several rows.
  for (const match of xml.matchAll(/<mergeCell\b([^>]*?)\/?>/g)) {
    const range = parseRange(attr(match[1], 'ref') ?? '');
    if (range === null) continue;
    const source = rows[range.startRow]?.[range.startCol];
    if (source === undefined) continue;
    for (let r = range.startRow; r <= range.endRow; r++) {
      const row = (rows[r] ??= []);
      for (let c = range.startCol; c <= range.endCol; c++) {
        if (row[c] === undefined) row[c] = source;
      }
    }
    if (range.endCol > maxCol) maxCol = range.endCol;
  }

  const width = maxCol + 1;
  const dense = [];
  for (let r = 0; r < rows.length; r++) {
    const source = rows[r];
    if (source === undefined) {
      dense.push([]);
      continue;
    }
    const row = new Array(width).fill('');
    for (let c = 0; c < width; c++) row[c] = source[c] ?? '';
    dense.push(row);
  }
  return dense;
}

/** Trim fully empty trailing rows/columns without destroying interior gaps. */
function trimGrid(rows) {
  let lastRow = -1;
  let lastCol = -1;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] ?? [];
    for (let c = 0; c < row.length; c++) {
      const value = row[c];
      if (value !== '' && value !== undefined && value !== null) {
        if (r > lastRow) lastRow = r;
        if (c > lastCol) lastCol = c;
      }
    }
  }
  if (lastRow < 0) return [];
  return rows.slice(0, lastRow + 1).map((row) => {
    const out = new Array(lastCol + 1).fill('');
    for (let c = 0; c <= lastCol; c++) out[c] = row?.[c] ?? '';
    return out;
  });
}

/**
 * Parse an XLSX/XLSM buffer.
 *
 * @param {Uint8Array|ArrayBuffer|Buffer} input
 * @returns {{ sheets: Array<{ name: string, index: number, rows: unknown[][], hidden: boolean }> }}
 */
export function parseWorkbook(input) {
  let bytes;
  if (input instanceof Uint8Array) bytes = input;
  else if (input instanceof ArrayBuffer) bytes = new Uint8Array(input);
  else if (ArrayBuffer.isView(input)) bytes = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  else throw new Error('无法读取 Excel 文件：输入不是二进制数据');

  let files;
  try {
    files = unzipSync(bytes);
  } catch (error) {
    throw new Error(`无法解压 Excel 文件（仅支持 .xlsx / .xlsm）：${error instanceof Error ? error.message : String(error)}`);
  }

  const workbookXml = readPart(files, 'xl/workbook.xml');
  if (workbookXml === '') throw new Error('Excel 文件缺少 xl/workbook.xml，可能不是有效的 .xlsx 文件');

  const relsXml = readPart(files, 'xl/_rels/workbook.xml.rels');
  const relTargets = new Map();
  for (const match of relsXml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const id = attr(match[1], 'Id');
    const target = attr(match[1], 'Target');
    if (id && target) relTargets.set(id, target);
  }

  const sharedStrings = [];
  const sharedXml = readPart(files, 'xl/sharedStrings.xml');
  if (sharedXml !== '') {
    for (const match of sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>|<si\b[^>]*\/>/g)) {
      sharedStrings.push(match[1] === undefined ? '' : textRuns(match[1]));
    }
  }

  const isDateStyle = makeDateStyleLookup(readPart(files, 'xl/styles.xml'));

  const sheets = [];
  let index = 0;
  for (const match of workbookXml.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    const attrs = match[1] ?? '';
    const name = attr(attrs, 'name') ?? `Sheet${index + 1}`;
    const relId = attr(attrs, 'r:id') ?? attr(attrs, 'relationshipId');
    const hidden = attr(attrs, 'state') === 'hidden' || attr(attrs, 'state') === 'veryHidden';
    const target = relId === undefined ? undefined : relTargets.get(relId);
    let rows = [];
    if (target !== undefined) {
      const part = resolvePart('xl', target);
      const sheetXml = readPart(files, part);
      if (sheetXml !== '') rows = trimGrid(parseSheet(sheetXml, sharedStrings, isDateStyle));
    }
    sheets.push({ name, index, rows, hidden });
    index += 1;
  }

  if (sheets.length === 0) throw new Error('Excel 工作簿中没有工作表');
  return { sheets };
}

/** Stringify a cell without the "undefined"/"null" artefacts JSON would add. */
export function cellText(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return value.trim();
  if (typeof value === 'number') {
    if (Number.isInteger(value)) return String(value);
    return String(Number(value.toFixed(10)));
  }
  if (typeof value === 'boolean') return value ? '是' : '否';
  return String(value).trim();
}

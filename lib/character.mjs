/**
 * Turn a parsed workbook into character cards, and render a card as a
 * role-play system-prompt block.
 *
 * Two common Excel layouts are recognised:
 *   1. vertical key/value:  A 列是字段名，B 列是内容（可跨多行、可多角色分块）
 *   2. table:               第一行是表头，之后每行是一个角色
 * Horizontal key/value rows (`姓名 | 林晚 | 性别 | 女 | 年龄 | 24`) are also
 * accepted. Unknown fields are preserved under `extra`.
 */

import { createHash } from 'node:crypto';
import { cellText } from './xlsx.mjs';

/** Canonical fields, their Chinese labels, aliases, and prompt groups. */
export const FIELD_DEFS = [
  { key: 'gender', label: '性别', group: '基础', aliases: ['性别', '性別', 'gender', 'sex'] },
  { key: 'age', label: '年龄', group: '基础', aliases: ['年龄', '年纪', '岁数', 'age'] },
  { key: 'birthday', label: '生日', group: '基础', aliases: ['生日', '出生日期', '出生年月', '诞辰', 'birthday', 'birthdate', 'dob'] },
  { key: 'height', label: '身高', group: '基础', aliases: ['身高', 'height'] },
  { key: 'weight', label: '体重', group: '基础', aliases: ['体重', 'weight'] },
  { key: 'species', label: '种族', group: '基础', aliases: ['种族', '物种', '民族', 'species', 'race', 'ethnicity'] },
  { key: 'identity', label: '身份 / 职业', group: '基础', aliases: ['身份', '职业', '身份职业', '定位', '头衔', '职位', 'occupation', 'identity', 'role', 'title', 'job'] },
  { key: 'appearance', label: '外貌', group: '外貌', aliases: ['外貌', '外观', '长相', '容貌', '形象', '外表', '特征', 'appearance', 'looks', 'features'] },
  { key: 'personality', label: '性格', group: '性格', aliases: ['性格', '个性', '人格', '特质', '性情', 'personality', 'traits', 'character'] },
  { key: 'likes', label: '喜好', group: '性格', aliases: ['喜好', '喜欢', '爱好', '兴趣', 'likes', 'hobbies', 'interests', 'favorites'] },
  { key: 'dislikes', label: '讨厌', group: '性格', aliases: ['讨厌', '厌恶', '反感', '不喜欢', 'dislikes', 'hates'] },
  { key: 'fears', label: '恐惧', group: '性格', aliases: ['恐惧', '害怕', '畏惧', 'fears', 'fear'] },
  { key: 'background', label: '背景 / 经历', group: '背景', aliases: ['背景', '经历', '身世', '故事', '过往', '过往经历', 'background', 'backstory', 'history', 'bio'] },
  { key: 'motivation', label: '动机 / 目标', group: '背景', aliases: ['动机', '目标', '愿望', '追求', '信念', 'motivation', 'goal', 'desire', 'purpose'] },
  { key: 'scenario', label: '场景 / 世界观', group: '背景', aliases: ['场景', '设定', '世界观', '故事背景', 'scenario', 'setting', 'world'] },
  { key: 'skills', label: '技能 / 特长', group: '能力', aliases: ['技能', '能力', '特长', '擅长', '本领', 'skills', 'abilities', 'talents'] },
  { key: 'weaknesses', label: '弱点', group: '能力', aliases: ['弱点', '缺点', '缺陷', '短板', 'weaknesses', 'weakness', 'flaws'] },
  { key: 'speechStyle', label: '说话风格', group: '语言', aliases: ['说话风格', '语言风格', '语气', '口吻', '说话方式', '台词风格', 'speechstyle', 'speech', 'voice', 'tone', 'voicestyle'] },
  { key: 'catchphrases', label: '口头禅', group: '语言', aliases: ['口头禅', '常用语', '经典台词', '台词', '语录', 'catchphrase', 'catchphrases', 'quotes'] },
  { key: 'quote', label: '名言 / 座右铭', group: '语言', aliases: ['名言', '座右铭', '经典语录', 'quote', 'motto'] },
  { key: 'greeting', label: '开场白 / 问候', group: '语言', aliases: ['开场白', '问候语', '第一句话', 'greeting', 'opening', 'intro'] },
  { key: 'relationships', label: '人际关系', group: '关系', aliases: ['关系', '人际关系', '羁绊', '人物关系', 'relationships', 'relations', 'family'] },
  { key: 'secrets', label: '秘密', group: '秘密', aliases: ['秘密', '隐藏设定', '隐情', 'secrets', 'secret'] },
  { key: 'rules', label: '规则 / 禁忌', group: '规则', aliases: ['规则', '禁忌', '扮演要求', '注意事项', '行为准则', 'rules', 'taboos', 'constraints'] },
];

export const NAME_FIELD = {
  key: 'name',
  label: '姓名',
  aliases: ['姓名', '名字', '角色名', '角色姓名', '名称', '人物名', 'name', 'charactername', 'character', '角色', '人物'],
};

export const FIELD_KEYS = FIELD_DEFS.map((def) => def.key);

const ALL_DEFS = [NAME_FIELD, ...FIELD_DEFS];

const GROUP_ORDER = ['基础', '外貌', '性格', '背景', '能力', '语言', '关系', '秘密', '规则'];

/** Normalise a header/key for alias matching. */
export function normalizeLabel(value) {
  return String(value ?? '')
    .replace(/[（(【\[][^）)】\]]*[）)】\]]/g, '')
    .replace(/[\s\u3000]+/g, '')
    .replace(/[：:、,，。.;；/\\|<>《》"'“”‘’_\-—]+/g, '')
    .trim()
    .toLowerCase();
}

const ALIAS_INDEX = new Map();
for (const def of ALL_DEFS) {
  for (const alias of def.aliases) {
    const key = normalizeLabel(alias);
    if (key !== '' && !ALIAS_INDEX.has(key)) ALIAS_INDEX.set(key, def);
  }
}

/** Resolve a cell/header label to a field definition, if known. */
export function fieldByLabel(label) {
  return ALIAS_INDEX.get(normalizeLabel(label));
}

/** Whether a sheet name is a generic container rather than a character name. */
export function isGenericSheetName(name) {
  const key = normalizeLabel(name);
  return /^(sheet\d*|工作表\d*|worksheet\d*|worksheet|角色|角色卡|角色列表|人物|人物列表|群像|设定|说明|data|cards?|characters?|readme|info|list)$/.test(key);
}

/** Strip a little formatting that only makes a value noisier. */
function valueText(value) {
  return cellText(value)
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/** Create an empty card with all canonical fields present. */
function emptyCard(sheet, meta) {
  const fields = {};
  for (const key of FIELD_KEYS) fields[key] = '';
  return {
    id: '',
    name: '',
    fields,
    extra: [],
    source: {
      fileName: meta.fileName ?? 'uploaded.xlsx',
      sheet: sheet.name,
      sheetIndex: sheet.index,
      ...(meta.sourcePath ? { path: meta.sourcePath } : {}),
    },
    updatedAt: new Date().toISOString(),
  };
}

/** Append a value to a field, avoiding duplicated identical lines. */
function appendField(card, key, value) {
  const text = valueText(value);
  if (text === '') return;
  if (key === 'name') {
    if (card.name === '') card.name = text;
    return;
  }
  if (!Object.prototype.hasOwnProperty.call(card.fields, key)) return;
  const current = card.fields[key];
  if (current === '') card.fields[key] = text;
  else if (!current.includes(text)) card.fields[key] = `${current}\n${text}`;
}

/** Append a non-canonical key/value pair. */
function appendExtra(card, label, value) {
  const key = valueText(label).replace(/[：:]\s*$/, '');
  const text = valueText(value);
  if (key === '' || text === '') return;
  const existing = card.extra.find((item) => item.label === key);
  if (existing === undefined) card.extra.push({ label: key, value: text });
  else if (!existing.value.includes(text)) existing.value = `${existing.value}\n${text}`;
}

/** Number of non-empty fields, used for "does this block contain data" checks. */
function fieldCount(card) {
  let count = card.name === '' ? 0 : 1;
  for (const key of FIELD_KEYS) if (card.fields[key] !== '') count += 1;
  count += card.extra.length;
  return count;
}

/** Split a single cell of the form `字段：内容`. */
function splitInlinePair(text) {
  const match = /^([^：:]{1,24})[：:]\s*([\s\S]+)$/.exec(String(text).trim());
  if (match === null) return null;
  return { label: match[1].trim(), value: match[2].trim() };
}

/** Detect a horizontal key/value row such as 姓名 | 林晚 | 性别 | 女. */
function looksLikePairRow(row) {
  const cells = row.map(valueText);
  let pairs = 0;
  for (let i = 0; i < cells.length; i++) {
    if (cells[i] === '') continue;
    const def = fieldByLabel(cells[i]);
    if (def === undefined) continue;
    const next = cells[i + 1];
    if (next !== undefined && next !== '' && fieldByLabel(next) === undefined) {
      pairs += 1;
      i += 1;
    }
  }
  return pairs >= 2;
}

/** Apply one key/value pair to a card. */
function applyPair(card, label, value) {
  const def = fieldByLabel(label);
  if (def === undefined) {
    appendExtra(card, label, value);
    return;
  }
  appendField(card, def.key, value);
}

/** Detect a table header row and its column mapping. */
function detectTableHeader(rows) {
  const limit = Math.min(rows.length, 12);
  for (let r = 0; r < limit; r++) {
    const row = rows[r] ?? [];
    const cells = row.map(valueText);
    if (cells.every((cell) => cell === '')) continue;
    if (looksLikePairRow(row)) continue;

    const columns = [];
    let known = 0;
    let hasName = false;
    for (let c = 0; c < cells.length; c++) {
      const text = cells[c];
      if (text === '') continue;
      if (text.length > 30 || text.includes('\n')) {
        known = -999;
        break;
      }
      const def = fieldByLabel(text);
      if (def !== undefined) {
        columns.push({ col: c, key: def.key, label: def.label });
        known += 1;
        if (def.key === 'name') hasName = true;
      } else if (text.length <= 20) {
        columns.push({ col: c, key: 'extra', label: text });
      }
    }
    if (known < 2 || (!hasName && known < 3)) continue;

    // A header needs at least one following row with a non-empty cell.
    let dataRows = 0;
    for (let d = r + 1; d < rows.length; d++) {
      if ((rows[d] ?? []).some((cell) => valueText(cell) !== '')) dataRows += 1;
    }
    if (dataRows === 0) continue;
    return { rowIndex: r, columns };
  }
  return null;
}

/** Derive a name for a nameless card row. */
function deriveName(sheetName, ordinal) {
  return isGenericSheetName(sheetName) ? `角色${ordinal}` : sheetName;
}

/** Parse a table where every following row is one character. */
function parseTable(rows, header, sheet, meta) {
  const cards = [];
  for (let r = header.rowIndex + 1; r < rows.length; r++) {
    const row = rows[r] ?? [];
    if (!row.some((cell) => valueText(cell) !== '')) continue;
    const card = emptyCard(sheet, meta);
    for (const column of header.columns) {
      const value = valueText(row[column.col]);
      if (value === '') continue;
      if (column.key === 'extra') appendExtra(card, column.label, value);
      else appendField(card, column.key, value);
    }
    if (fieldCount(card) === 0) continue;
    if (card.name === '') card.name = deriveName(sheet.name, cards.length + 1);
    cards.push(card);
  }
  return cards;
}

/** Parse horizontal key/value rows (one or more rows per character). */
function parsePairRows(rows, sheet, meta) {
  const cards = [];
  let current = null;
  for (const row of rows) {
    if (!row.some((cell) => valueText(cell) !== '')) continue;
    if (!looksLikePairRow(row)) {
      // A plain one-cell title or note should not become a character.
      if (current !== null && fieldCount(current) > 0) {
        cards.push(current);
        current = null;
      }
      continue;
    }
    const cells = row.map(valueText);
    const rowHasName = cells.some((cell) => fieldByLabel(cell)?.key === 'name');
    if (current === null || (rowHasName && fieldCount(current) > 0 && current.name !== '')) {
      if (current !== null && fieldCount(current) > 0) cards.push(current);
      current = emptyCard(sheet, meta);
    }
    for (let c = 0; c < cells.length; c++) {
      const text = cells[c];
      if (text === '') continue;
      const def = fieldByLabel(text);
      if (def === undefined) continue;
      const next = cells[c + 1];
      if (next !== undefined && next !== '' && fieldByLabel(next) === undefined) {
        appendField(current, def.key, next);
        c += 1;
      }
    }
  }
  if (current !== null && fieldCount(current) > 0) cards.push(current);
  for (let i = 0; i < cards.length; i++) {
    if (cards[i].name === '') cards[i].name = deriveName(sheet.name, i + 1);
  }
  return cards;
}

/** Parse a vertical key/value sheet, including multiple blocks. */
function parseKeyValue(rows, sheet, meta) {
  const cards = [];
  let current = null;
  const ensure = () => {
    if (current === null) current = emptyCard(sheet, meta);
    return current;
  };
  const flush = () => {
    if (current !== null && fieldCount(current) > 0) cards.push(current);
    current = null;
  };

  for (const row of rows) {
    const cells = row.map(valueText);
    const nonEmpty = cells
      .map((text, index) => ({ text, index }))
      .filter((item) => item.text !== '');
    if (nonEmpty.length === 0) {
      flush();
      continue;
    }

    // One-cell `字段：内容` row.
    if (nonEmpty.length === 1) {
      const inline = splitInlinePair(nonEmpty[0].text);
      if (inline !== null) {
        const def = fieldByLabel(inline.label);
        if (def !== undefined) {
          if (def.key === 'name' && current !== null && fieldCount(current) > 0 && current.name !== '') flush();
          applyPair(ensure(), inline.label, inline.value);
          continue;
        }
      }
      // A lone title/note: keep it as extra only when we already have a card.
      if (current !== null) appendExtra(current, '备注', nonEmpty[0].text);
      continue;
    }

    // Horizontal pairs inside an otherwise vertical sheet.
    if (looksLikePairRow(row)) {
      const rowHasName = nonEmpty.some((item) => fieldByLabel(item.text)?.key === 'name');
      if (current !== null && rowHasName && fieldCount(current) > 0 && current.name !== '') flush();
      const card = ensure();
      for (const item of nonEmpty) {
        const def = fieldByLabel(item.text);
        if (def === undefined) continue;
        const next = cells[item.index + 1];
        if (next !== undefined && next !== '' && fieldByLabel(next) === undefined) {
          appendField(card, def.key, next);
        }
      }
      continue;
    }

    const first = nonEmpty[0];
    const def = fieldByLabel(first.text);
    if (def !== undefined) {
      const rest = nonEmpty.slice(1).map((item) => item.text);
      const value = rest.join('\n');
      if (def.key === 'name' && current !== null && fieldCount(current) > 0 && current.name !== '' && value !== current.name) {
        flush();
      }
      applyPair(ensure(), first.text, value);
      continue;
    }

    // Unknown key -> preserve it.
    if (current !== null) {
      appendExtra(current, first.text, nonEmpty.slice(1).map((item) => item.text).join('\n'));
    }
  }
  flush();
  return cards;
}

/** Last-resort freeform sheet: put the whole text into one card. */
function parseFreeform(rows, sheet, meta) {
  const lines = [];
  for (const row of rows) {
    const line = row.map(valueText).filter((text) => text !== '').join(' | ');
    if (line !== '') lines.push(line);
  }
  if (lines.length === 0) return [];
  const card = emptyCard(sheet, meta);
  card.name = isGenericSheetName(sheet.name) ? '未命名角色' : sheet.name;
  card.fields.background = lines.join('\n');
  return [card];
}

/** Extract every card from one worksheet. */
function extractFromSheet(sheet, meta) {
  const rows = (sheet.rows ?? []).map((row) => row.map(valueText));
  if (rows.length === 0) return [];

  const header = detectTableHeader(rows);
  if (header !== null) {
    const table = parseTable(rows, header, sheet, meta);
    if (table.length > 0) return table;
  }

  const pairs = parsePairRows(rows, sheet, meta);
  if (pairs.length > 0) return pairs;

  const keyValue = parseKeyValue(rows, sheet, meta);
  if (keyValue.length > 0) return keyValue;

  return parseFreeform(rows, sheet, meta);
}

/** Stable, content-derived card id. */
function makeCardId(sourceKey, sheetName, name, ordinal) {
  return `cc_${createHash('sha1')
    .update(`${sourceKey}\u0000${sheetName}\u0000${name}\u0000${ordinal}`)
    .digest('hex')
    .slice(0, 12)}`;
}

/**
 * Extract character cards from a parsed workbook.
 *
 * @param {{sheets: Array<{name:string,index:number,rows:unknown[][]}>}} workbook
 * @param {{fileName?:string, sourceKey?:string, sourcePath?:string, sheet?:string}} meta
 * @returns {Array<object>}
 */
export function extractCharacters(workbook, meta = {}) {
  const fileName = meta.fileName ?? 'uploaded.xlsx';
  const sourceKey = meta.sourceKey ?? fileName.toLowerCase();
  const sheetFilter = meta.sheet === undefined ? null : String(meta.sheet).trim().toLowerCase();
  const sheets = (workbook.sheets ?? []).filter((sheet) => {
    if (sheetFilter === null || sheetFilter === '') return true;
    return sheet.name.toLowerCase() === sheetFilter;
  });
  if (sheets.length === 0) {
    throw new Error(`找不到工作表「${meta.sheet}」。可用工作表：${(workbook.sheets ?? []).map((s) => s.name).join('、')}`);
  }

  const merged = new Map();
  for (const sheet of sheets) {
    const cards = extractFromSheet(sheet, { fileName, sourceKey, sourcePath: meta.sourcePath });
    for (const card of cards) {
      const key = card.name.trim().toLowerCase();
      const existing = merged.get(key);
      if (existing === undefined) {
        merged.set(key, card);
        continue;
      }
      // Same character across sheets: merge non-empty fields.
      for (const field of FIELD_KEYS) {
        if (card.fields[field] !== '' && existing.fields[field] === '') existing.fields[field] = card.fields[field];
      }
      for (const item of card.extra) {
        if (!existing.extra.some((entry) => entry.label === item.label && entry.value === item.value)) {
          existing.extra.push(item);
        }
      }
    }
  }

  const cards = [...merged.values()];
  if (cards.length === 0) {
    throw new Error('没有从 Excel 中识别到角色。请使用「A列字段名 / B列内容」或「第一行表头、每行一个角色」的格式。');
  }

  const usedIds = new Set();
  return cards.map((card, index) => {
    if (card.name.trim() === '') card.name = `未命名角色${index + 1}`;
    let id = makeCardId(sourceKey, card.source.sheet, card.name, index);
    let suffix = 2;
    while (usedIds.has(id)) {
      id = `${makeCardId(sourceKey, card.source.sheet, card.name, index)}_${suffix}`;
      suffix += 1;
    }
    usedIds.add(id);
    card.id = id;
    card.name = card.name.trim();
    return card;
  });
}

/** A short one-line description for lists. */
export function cardTagline(card) {
  const parts = [card.fields.gender, card.fields.age, card.fields.height, card.fields.identity]
    .map((value) => String(value ?? '').replace(/\s+/g, ' ').trim())
    .filter((value) => value !== '');
  return parts.slice(0, 4).join(' · ');
}

/** A compact public projection used by the UI and the list tool action. */
export function summarizeCard(card) {
  let count = card.extra.length;
  for (const key of FIELD_KEYS) if (card.fields[key] !== '') count += 1;
  return {
    id: card.id,
    name: card.name,
    tagline: cardTagline(card),
    source: {
      fileName: card.source.fileName,
      sheet: card.source.sheet,
    },
    updatedAt: card.updatedAt,
    fieldCount: count,
  };
}

/** Clip a prompt value without cutting in the middle of a surrogate pair. */
function clip(value, limit) {
  const text = String(value ?? '').trim();
  if (text.length <= limit) return text;
  return `${text.slice(0, Math.max(0, limit - 1))}…`;
}

/**
 * Render a character card as the system-prompt block that drives role-play.
 * Kept deterministic so prompt caching stays useful between turns.
 */
export function buildCharacterPrompt(card) {
  if (card === null || card === undefined) return '';
  const name = String(card.name ?? '').trim() || '未命名角色';
  const lines = [];
  lines.push(`【角色扮演模式 · 最高优先级角色指令】`);
  lines.push(`从现在起，你要扮演「${name}」，而不是以 AI 助手的身份回答。以下角色卡来自用户提供的设定，是你保持角色一致性的依据。`);
  lines.push('');
  lines.push(`# 角色：${name}`);

  for (const group of GROUP_ORDER) {
    const defs = FIELD_DEFS.filter((def) => def.group === group);
    const present = defs.filter((def) => String(card.fields[def.key] ?? '').trim() !== '');
    if (present.length === 0) continue;
    lines.push('');
    lines.push(`## ${group}`);
    for (const def of present) {
      lines.push(`- ${def.label}：${clip(card.fields[def.key], 1400)}`);
    }
  }

  if (card.extra.length > 0) {
    lines.push('');
    lines.push('## 其他设定');
    for (const item of card.extra) {
      lines.push(`- ${item.label}：${clip(item.value, 1000)}`);
    }
  }

  lines.push('');
  lines.push('## 扮演规则');
  lines.push(`1. 始终以「${name}」的第一人称「我」说话和行动；不要自称 AI、助手、模型或语言模型。`);
  lines.push('2. 严格遵守上面的性别、年龄、身高、身份、外貌、性格、经历、能力与语言风格；不要做出与档案冲突的行为。');
  lines.push('3. 对白要自然、口语化，符合角色的年龄、教养、职业和说话风格；可以带动作、神态和心理描写，但不要长篇旁白。');
  lines.push('4. 用户没有提供的细节，可以按角色性格自然补全；不要编造会推翻已有档案的重大设定。');
  lines.push('5. 不要主动提及“角色卡”“Excel”“设定文件”“系统提示词”或“根据设定”等出戏内容；除非用户明确要求。');
  lines.push('6. 如果用户要求你出戏或停止扮演，先用角色口吻回应，再另起一行以“（出戏）”开头简短说明。');
  lines.push('7. 在不违反平台安全规则的前提下，尽量让每一句回复都像这个角色会说的话。');
  if (String(card.fields.greeting ?? '').trim() !== '') {
    lines.push(`8. 如果这是你们第一次见面，可以用这句开场白起手：${clip(card.fields.greeting, 500)}`);
  }
  if (String(card.fields.rules ?? '').trim() !== '') {
    lines.push(`9. 额外扮演规则（优先遵守）：${clip(card.fields.rules, 1200)}`);
  }

  let prompt = lines.join('\n');
  const MAX_TOTAL = 14000;
  if (prompt.length > MAX_TOTAL) {
    prompt = `${prompt.slice(0, MAX_TOTAL - 40)}\n（角色卡内容过长，以上已截断）`;
  }
  return prompt;
}

/** Format a card for the tool's model-visible text. */
export function formatCardForModel(card) {
  const lines = [`角色：${card.name}`];
  if (cardTagline(card) !== '') lines.push(`简介：${cardTagline(card)}`);
  for (const def of FIELD_DEFS) {
    const value = String(card.fields[def.key] ?? '').trim();
    if (value !== '') lines.push(`${def.label}：${clip(value, 600)}`);
  }
  for (const item of card.extra) lines.push(`${item.label}：${clip(item.value, 400)}`);
  lines.push(`来源：${card.source.fileName} / ${card.source.sheet}`);
  return lines.join('\n');
}

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { zipSync } from '../lib/vendor/fflate.mjs';
import { parseWorkbook } from '../lib/xlsx.mjs';
import { extractCharacters, buildCharacterPrompt, summarizeCard } from '../lib/character.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const sample = path.join(here, '..', 'sample', '角色卡示例.xlsx');

const buffer = readFileSync(sample);
const workbook = parseWorkbook(buffer);
assert.equal(workbook.sheets.length, 2, 'sample should have two sheets');
assert.equal(workbook.sheets[0].name, '角色卡');
assert.ok(workbook.sheets[0].rows.length > 10);

const cards = extractCharacters(workbook, {
  fileName: '角色卡示例.xlsx',
  sourceKey: 'test',
  sourcePath: sample,
});
assert.equal(cards.length, 4, `expected 4 characters, got ${cards.length}: ${cards.map((c) => c.name).join(', ')}`);

const lin = cards.find((card) => card.name === '林晚');
assert.ok(lin, 'key/value sheet should produce 林晚');
assert.equal(lin.fields.gender, '女');
assert.equal(lin.fields.age, '24');
assert.equal(lin.fields.height, '168 cm');
assert.match(lin.fields.personality, /表面温和疏离/);
assert.match(lin.fields.background, /晚风/);
assert.equal(lin.fields.birthday, '2001-09-17', 'date style should convert to ISO');
assert.ok(lin.extra.some((item) => item.label === '备注'));

const shen = cards.find((card) => card.name === '沈砚');
assert.ok(shen, 'table sheet should produce 沈砚');
assert.equal(shen.fields.gender, '男');
assert.equal(shen.fields.age, '29');
assert.match(shen.fields.identity, /刑警/);

const summary = summarizeCard(lin);
assert.equal(summary.name, '林晚');
assert.match(summary.tagline, /女/);
assert.match(summary.tagline, /24/);

const prompt = buildCharacterPrompt(lin);
assert.match(prompt, /角色扮演模式/);
assert.match(prompt, /林晚/);
assert.match(prompt, /168 cm/);
assert.match(prompt, /表面温和疏离/);
assert.match(prompt, /不要主动提及/);

// A sheet-filtered import should only return the table characters.
const tableOnly = extractCharacters(workbook, {
  fileName: '角色卡示例.xlsx',
  sourceKey: 'test-table',
  sheet: '群像',
});
assert.equal(tableOnly.length, 3);
assert.deepEqual(tableOnly.map((card) => card.name), ['沈砚', '阿澄', '顾言']);

// Horizontal key/value rows.
const horizontal = extractCharacters(
  {
    sheets: [
      {
        name: '横排',
        index: 0,
        rows: [
          ['姓名', '周声', '性别', '男', '年龄', '31', '身高', '181cm', '性格', '懒散但敏锐'],
          [],
          ['姓名', '许棠', '性别', '女', '年龄', '22', '身高', '163cm', '性格', '安静、观察力强'],
        ],
      },
    ],
  },
  { fileName: 'horizontal.xlsx', sourceKey: 'horizontal' },
);
assert.equal(horizontal.length, 2);
assert.equal(horizontal[0].name, '周声');
assert.equal(horizontal[0].fields.age, '31');
assert.equal(horizontal[1].name, '许棠');

// Multi-block vertical key/value separated by a blank row.
const blocks = extractCharacters(
  {
    sheets: [
      {
        name: '角色卡',
        index: 0,
        rows: [
          ['姓名', '甲'],
          ['性别', '女'],
          ['年龄', '18'],
          [],
          ['姓名', '乙'],
          ['性别', '男'],
          ['年龄', '30'],
        ],
      },
    ],
  },
  { fileName: 'blocks.xlsx', sourceKey: 'blocks' },
);
assert.equal(blocks.length, 2);
assert.deepEqual(blocks.map((card) => card.name), ['甲', '乙']);

// Inline strings + merged cells + no sharedStrings part.
const encoder = new TextEncoder();
const minimalXlsx = zipSync({
  'xl/workbook.xml': encoder.encode(
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="角色卡" sheetId="1" r:id="rId1"/></sheets></workbook>',
  ),
  'xl/_rels/workbook.xml.rels': encoder.encode(
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="worksheets/sheet1.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"/></Relationships>',
  ),
  'xl/worksheets/sheet1.xml': encoder.encode(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>姓名</t></is></c><c r="B1" t="inlineStr"><is><t>林晚</t></is></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>性别</t></is></c><c r="B2" t="inlineStr"><is><t>女</t></is></c></row>' +
      '<row r="3"><c r="A3" t="inlineStr"><is><t>背景</t></is></c><c r="B3" t="inlineStr"><is><t>很长的一段背景</t></is></c></row>' +
      '<row r="4"><c r="A4" t="inlineStr"><is><t>性格</t></is></c></row>' +
      '</sheetData><mergeCells count="1"><mergeCell ref="B3:B4"/></mergeCells></worksheet>',
  ),
});
const inlineCards = extractCharacters(parseWorkbook(minimalXlsx), { fileName: 'inline.xlsx', sourceKey: 'inline' });
assert.equal(inlineCards.length, 1);
assert.equal(inlineCards[0].name, '林晚');
assert.equal(inlineCards[0].fields.gender, '女');
assert.equal(inlineCards[0].fields.background, '很长的一段背景');
assert.equal(inlineCards[0].fields.personality, '很长的一段背景');

console.log(`OK: ${cards.length} cards parsed; prompt ${prompt.length} chars`);
for (const card of cards) console.log(`- ${card.name}: ${summarizeCard(card).tagline}`);

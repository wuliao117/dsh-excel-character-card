/**
 * dsh-excel-character-card — Host half.
 *
 * Responsibilities:
 *   - parse .xlsx/.xlsm character cards (XLSX reader + character extractor)
 *   - persist the imported cards in $DSH_HOME/excel-character-card/state.json
 *   - register a model tool (`excel_character_card`) for import/list/activate/...
 *   - register a dynamic system-prompt section: once a card is active for a
 *     session, every model step receives the character sheet as a high-priority
 *     role-play instruction
 *   - expose a tiny same-origin JSON API used by the browser half's panel
 *
 * Zero third-party runtime imports (the XLSX unzip helper is vendored) so the
 * bundle can be installed from a local directory without an SDK link.
 */

import { mkdir, readFile, rename, stat, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { parseWorkbook } from './xlsx.mjs';
import {
  FIELD_KEYS,
  buildCharacterPrompt,
  extractCharacters,
  formatCardForModel,
  summarizeCard,
} from './character.mjs';

export const name = 'excel-character-card';
export const inject = ['tools', 'systemPrompt', 'webServer'];

const PACKAGE_ID = 'dsh-excel-character-card';
const API_PREFIX = '/excel-character-card/api';
const DSH_HOME = process.env.DSH_HOME || path.join(os.homedir(), '.dsh');
const DATA_DIR = path.join(DSH_HOME, 'excel-character-card');
const STATE_FILE = path.join(DATA_DIR, 'state.json');
const MAX_FILE_BYTES = 20 * 1024 * 1024;
const MAX_BODY_BYTES = 30 * 1024 * 1024;
const MAX_CARDS = 200;

/** Small logger that never lets a logging failure escape. */
function makeLogger(ctx) {
  let sink;
  try {
    sink = typeof ctx.logger === 'function' ? ctx.logger(PACKAGE_ID) : ctx.logger;
  } catch {
    sink = undefined;
  }
  const emit = (level, message) => {
    try {
      if (sink && typeof sink[level] === 'function') sink[level](`[${PACKAGE_ID}] ${message}`);
      else if (typeof console[level] === 'function') console[level](`[${PACKAGE_ID}] ${message}`);
    } catch {
      /* logging must never break the plugin */
    }
  };
  return {
    info: (message) => emit('info', message),
    warn: (message) => emit('warn', message),
    error: (message) => emit('error', message),
  };
}

function emptyState() {
  return {
    version: 1,
    cards: [],
    activeBySession: {},
    activeGlobal: null,
    updatedAt: new Date().toISOString(),
  };
}

function asString(value, fallback = '') {
  return typeof value === 'string' ? value : value === undefined || value === null ? fallback : String(value);
}

/** Validate and normalise one persisted card. */
function normalizeCard(raw) {
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return null;
  const fields = {};
  for (const key of FIELD_KEYS) fields[key] = asString(raw.fields?.[key]);
  const extra = Array.isArray(raw.extra)
    ? raw.extra
        .filter((item) => item && typeof item === 'object' && asString(item.label).trim() !== '')
        .map((item) => ({ label: asString(item.label).trim(), value: asString(item.value) }))
    : [];
  const id = asString(raw.id).trim();
  const cardName = asString(raw.name).trim();
  if (id === '' || cardName === '') return null;
  return {
    id,
    name: cardName,
    fields,
    extra,
    source: {
      fileName: asString(raw.source?.fileName, 'uploaded.xlsx'),
      sheet: asString(raw.source?.sheet, 'Sheet1'),
      sheetIndex: Number.isFinite(raw.source?.sheetIndex) ? raw.source.sheetIndex : 0,
      ...(asString(raw.source?.path) ? { path: asString(raw.source.path) } : {}),
    },
    createdAt: asString(raw.createdAt, new Date().toISOString()),
    updatedAt: asString(raw.updatedAt, new Date().toISOString()),
  };
}

function normalizeState(raw) {
  const state = emptyState();
  if (raw === null || typeof raw !== 'object' || Array.isArray(raw)) return state;
  if (Array.isArray(raw.cards)) {
    state.cards = raw.cards.map(normalizeCard).filter((card) => card !== null).slice(0, MAX_CARDS);
  }
  if (raw.activeBySession && typeof raw.activeBySession === 'object' && !Array.isArray(raw.activeBySession)) {
    for (const [sessionId, cardId] of Object.entries(raw.activeBySession)) {
      if (typeof sessionId !== 'string' || sessionId === '') continue;
      if (cardId === null) state.activeBySession[sessionId] = null;
      else if (typeof cardId === 'string' && cardId !== '') state.activeBySession[sessionId] = cardId;
    }
  }
  if (typeof raw.activeGlobal === 'string' && raw.activeGlobal !== '') state.activeGlobal = raw.activeGlobal;
  else state.activeGlobal = null;
  state.updatedAt = asString(raw.updatedAt, state.updatedAt);
  // Drop dangling active references.
  const ids = new Set(state.cards.map((card) => card.id));
  for (const [sessionId, cardId] of Object.entries(state.activeBySession)) {
    if (cardId !== null && !ids.has(cardId)) delete state.activeBySession[sessionId];
  }
  if (state.activeGlobal !== null && !ids.has(state.activeGlobal)) state.activeGlobal = null;
  return state;
}

/** Read a JSON request body with a hard byte cap. */
async function readBody(req, limit = MAX_BODY_BYTES) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += bytes.length;
    if (size > limit) throw Object.assign(new Error(`请求体超过 ${Math.round(limit / 1024 / 1024)} MB`), { status: 413 });
    chunks.push(bytes);
  }
  return Buffer.concat(chunks);
}

function sendJson(res, status, value) {
  const body = JSON.stringify(value);
  res.writeHead(status, {
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
    'content-length': Buffer.byteLength(body),
  });
  res.end(body);
}

function sendError(res, error) {
  const status = Number.isInteger(error?.status) ? error.status : 400;
  const message = error instanceof Error ? error.message : String(error);
  sendJson(res, status, { ok: false, error: message });
}

function assertXlsxPath(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext !== '.xlsx' && ext !== '.xlsm') {
    throw new Error('仅支持 .xlsx / .xlsm 文件。旧版 .xls 请先在 Excel 中“另存为” .xlsx。');
  }
}

function base64ToBuffer(dataBase64) {
  const text = String(dataBase64 ?? '').replace(/^data:[^;]+;base64,/, '').replace(/\s+/g, '');
  if (text === '') throw new Error('上传内容为空');
  if (text.length > Math.ceil((MAX_FILE_BYTES * 4) / 3) + 16) throw new Error('Excel 文件超过 20 MB');
  const buffer = Buffer.from(text, 'base64');
  if (buffer.length === 0) throw new Error('上传内容不是有效的 base64 数据');
  if (buffer.length > MAX_FILE_BYTES) throw new Error('Excel 文件超过 20 MB');
  return buffer;
}

/**
 * @param {import('@deepseek-ai/cordis').Context} ctx
 */
export function apply(ctx) {
  const log = makeLogger(ctx);

  let state = emptyState();
  let loadPromise = null;
  let writeChain = Promise.resolve();

  const load = () => {
    if (loadPromise !== null) return loadPromise;
    loadPromise = (async () => {
      try {
        const raw = JSON.parse(await readFile(STATE_FILE, 'utf8'));
        state = normalizeState(raw);
        log.info(`已加载 ${state.cards.length} 张角色卡`);
      } catch (error) {
        if (error?.code !== 'ENOENT') {
          log.warn(`角色卡状态无法读取，将使用空状态：${error instanceof Error ? error.message : String(error)}`);
          try {
            const backup = `${STATE_FILE}.bad-${Date.now()}`;
            await rename(STATE_FILE, backup);
            log.warn(`原状态文件已备份到 ${backup}`);
          } catch {
            /* ignore backup failure */
          }
        }
      }
    })();
    return loadPromise;
  };

  const persist = () => {
    const snapshot = JSON.stringify(state, null, 2);
    writeChain = writeChain
      .then(async () => {
        await mkdir(DATA_DIR, { recursive: true });
        const tmp = `${STATE_FILE}.tmp-${process.pid}`;
        await writeFile(tmp, snapshot, 'utf8');
        await rename(tmp, STATE_FILE);
      })
      .catch((error) => {
        log.error(`角色卡状态写入失败：${error instanceof Error ? error.message : String(error)}`);
      });
    return writeChain;
  };

  const cardById = (cardId) => state.cards.find((card) => card.id === cardId);

  const activeCardFor = (sessionId) => {
    if (typeof sessionId === 'string' && sessionId !== '') {
      if (Object.prototype.hasOwnProperty.call(state.activeBySession, sessionId)) {
        const cardId = state.activeBySession[sessionId];
        if (cardId === null) return null;
        return cardById(cardId) ?? null;
      }
    }
    if (state.activeGlobal === null) return null;
    return cardById(state.activeGlobal) ?? null;
  };

  const activeScopeFor = (sessionId) => {
    if (typeof sessionId === 'string' && sessionId !== '' && Object.prototype.hasOwnProperty.call(state.activeBySession, sessionId)) {
      return state.activeBySession[sessionId] === null ? 'none' : 'session';
    }
    return state.activeGlobal === null ? 'none' : 'global';
  };

  const publicState = (sessionId) => {
    const active = activeCardFor(sessionId);
    return {
      cards: state.cards.map(summarizeCard),
      activeCardId: active?.id ?? null,
      activeCard: active === null ? null : summarizeCard(active),
      activeScope: activeScopeFor(sessionId),
      total: state.cards.length,
      dataDir: DATA_DIR,
    };
  };

  const activate = async (sessionId, cardId) => {
    const card = cardById(cardId);
    if (card === undefined) throw new Error(`找不到角色卡「${cardId}」。请先 list 查看可用角色。`);
    if (typeof sessionId === 'string' && sessionId !== '') state.activeBySession[sessionId] = card.id;
    else state.activeGlobal = card.id;
    state.updatedAt = new Date().toISOString();
    await persist();
    return card;
  };

  const deactivate = async (sessionId) => {
    if (typeof sessionId === 'string' && sessionId !== '') state.activeBySession[sessionId] = null;
    else state.activeGlobal = null;
    state.updatedAt = new Date().toISOString();
    await persist();
  };

  const deleteCard = async (cardId) => {
    const index = state.cards.findIndex((card) => card.id === cardId);
    if (index < 0) throw new Error(`找不到角色卡「${cardId}」`);
    const [removed] = state.cards.splice(index, 1);
    for (const [sessionId, id] of Object.entries(state.activeBySession)) {
      if (id === cardId) delete state.activeBySession[sessionId];
    }
    if (state.activeGlobal === cardId) state.activeGlobal = null;
    state.updatedAt = new Date().toISOString();
    await persist();
    return removed;
  };

  /** Merge freshly parsed cards into state, replacing the same source file. */
  const mergeCards = (cards, sourceKey, sourcePath) => {
    const existing = new Map(state.cards.map((card) => [card.id, card]));
    const sameSource = state.cards.filter((card) => {
      const key = card.source?.path ? `path:${card.source.path.toLowerCase()}` : `upload:${card.source.fileName.toLowerCase()}`;
      return key === sourceKey;
    });
    const incomingIds = new Set(cards.map((card) => card.id));
    let added = 0;
    let updated = 0;
    const nextCards = state.cards.filter((card) => !sameSource.some((old) => old.id === card.id && !incomingIds.has(card.id)));

    const now = new Date().toISOString();
    for (const card of cards) {
      const old = existing.get(card.id);
      if (old !== undefined) {
        card.createdAt = old.createdAt ?? now;
        card.updatedAt = now;
        updated += 1;
      } else {
        card.createdAt = now;
        card.updatedAt = now;
        added += 1;
      }
      const index = nextCards.findIndex((entry) => entry.id === card.id);
      if (index >= 0) nextCards[index] = card;
      else nextCards.push(card);
    }

    if (nextCards.length > MAX_CARDS) throw new Error(`角色卡数量超过上限（${MAX_CARDS} 张），请先删除一些再导入。`);
    state.cards = nextCards;
    state.updatedAt = now;

    const ids = new Set(state.cards.map((card) => card.id));
    for (const [sessionId, id] of Object.entries(state.activeBySession)) {
      if (id !== null && !ids.has(id)) delete state.activeBySession[sessionId];
    }
    if (state.activeGlobal !== null && !ids.has(state.activeGlobal)) state.activeGlobal = null;

    return {
      added,
      updated,
      removed: sameSource.filter((card) => !incomingIds.has(card.id)).length,
      cards: state.cards.map(summarizeCard),
      sourceKey,
      sourcePath: sourcePath ?? null,
    };
  };

  const importBuffer = async (buffer, meta, options = {}) => {
    const workbook = parseWorkbook(buffer);
    const cards = extractCharacters(workbook, meta);
    const report = mergeCards(cards, meta.sourceKey, meta.sourcePath);
    let activeCard = null;
    if (options.activate === true) {
      activeCard = await activate(options.sessionId, cards[0].id);
    } else {
      await persist();
    }
    return { cards, report, activeCard, workbook };
  };

  const importPath = async (filePath, options = {}) => {
    const absolute = path.resolve(filePath);
    assertXlsxPath(absolute);
    let info;
    try {
      info = await stat(absolute);
    } catch {
      throw new Error(`找不到文件：${absolute}`);
    }
    if (!info.isFile()) throw new Error(`不是文件：${absolute}`);
    if (info.size > MAX_FILE_BYTES) throw new Error('Excel 文件超过 20 MB');
    const buffer = await readFile(absolute);
    const sourceKey = `path:${absolute.toLowerCase()}`;
    return importBuffer(buffer, {
      fileName: path.basename(absolute),
      sourceKey,
      sourcePath: absolute,
      sheet: options.sheet,
    }, options);
  };

  const importUpload = async (fileName, dataBase64, options = {}) => {
    const safeName = path.basename(String(fileName ?? '').trim()) || 'uploaded.xlsx';
    assertXlsxPath(safeName);
    const buffer = base64ToBuffer(dataBase64);
    const sourceKey = `upload:${safeName.toLowerCase()}`;
    return importBuffer(buffer, { fileName: safeName, sourceKey, sheet: options.sheet }, options);
  };

  /** Human/model-facing message helpers. */
  const listMessage = (sessionId) => {
    if (state.cards.length === 0) return '还没有导入任何 Excel 角色卡。';
    const active = activeCardFor(sessionId);
    return `共有 ${state.cards.length} 张角色卡${active ? `，当前启用「${active.name}」` : ''}：`;
  };

  const cardMessage = (card) => {
    const summary = summarizeCard(card);
    return [`已选中角色「${summary.name}」。`, summary.tagline ? `简介：${summary.tagline}` : '', `来源：${summary.source.fileName} / ${summary.source.sheet}`]
      .filter((line) => line !== '')
      .join('\n');
  };

  // ---------------------------------------------------------------------
  // Model tool
  // ---------------------------------------------------------------------
  const tool = {
    name: 'excel_character_card',
    description:
      '读取 Excel 角色卡并启用角色扮演。支持 .xlsx / .xlsm：可以是「A列字段名、B列内容」的键值表，或「第一行表头、之后每行一个角色」的角色表。' +
      '常用流程：action=import 且 path=用户给的 Excel 路径、activate=true，立即在当前会话启用；' +
      '之后系统提示词会自动注入角色档案，你的所有回复都必须保持该角色的性别、年龄、身高、性格、身份、说话风格等设定。' +
      '也支持 list（列出角色卡）、preview（只解析不保存）、get（查看角色档案）、activate、deactivate、delete。',
    parameters: {
      type: 'object',
      properties: {
        action: {
          type: 'string',
          enum: ['list', 'import', 'preview', 'get', 'activate', 'deactivate', 'delete'],
          description: '要执行的操作。',
        },
        path: {
          type: 'string',
          description: 'action=import 或 preview 时的 .xlsx/.xlsm 文件路径（绝对路径或相对当前工作目录）。',
        },
        sheet: {
          type: 'string',
          description: '可选：只读取指定名称的工作表。',
        },
        activate: {
          type: 'boolean',
          description: 'action=import 时设为 true，导入后立即在当前会话启用第一个角色。',
        },
        cardId: {
          type: 'string',
          description: 'action=get / activate / delete 时的角色卡 id（先用 list 获取）。',
        },
      },
      required: ['action'],
      additionalProperties: false,
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          ok: { type: 'boolean' },
          action: { type: 'string' },
          message: { type: 'string' },
          cards: { type: 'array', items: { type: 'object', additionalProperties: true } },
          card: { type: 'object', additionalProperties: true },
          activeCardId: { type: 'string' },
          total: { type: 'number' },
          report: { type: 'object', additionalProperties: true },
          prompt: { type: 'string' },
          samples: { type: 'array', items: { type: 'string' } },
        },
      },
      render: (_args, value) => {
        const lines = [value.message];
        if (Array.isArray(value.cards) && value.cards.length > 0) {
          for (const card of value.cards) {
            lines.push(
              `- ${card.name}${card.tagline ? `（${card.tagline}）` : ''}  [${card.id}]${value.activeCardId === card.id ? '  ← 当前启用' : ''}`,
            );
          }
        }
        if (Array.isArray(value.samples)) lines.push(...value.samples);
        if (value.card && value.card.fields) {
          lines.push('');
          lines.push(formatCardForModel(value.card));
        }
        if (typeof value.prompt === 'string' && value.prompt !== '') {
          lines.push('');
          lines.push('已注入的系统提示词：');
          lines.push(value.prompt);
        }
        return [{ type: 'text', text: lines.filter((line) => line !== undefined && line !== null).join('\n') }];
      },
    },
    async execute(args, exec) {
      await load();
      const action = String(args?.action ?? '').trim();
      const sessionId = exec?.agent?.id === undefined ? undefined : String(exec.agent.id);

      if (action === 'list') {
        const active = activeCardFor(sessionId);
        return {
          ok: true,
          action,
          message: listMessage(sessionId),
          cards: state.cards.map(summarizeCard),
          total: state.cards.length,
          ...(active ? { activeCardId: active.id } : {}),
        };
      }

      if (action === 'import' || action === 'preview') {
        const filePath = String(args?.path ?? '').trim();
        if (filePath === '') throw new Error(`action=${action} 需要提供 path（.xlsx/.xlsm 文件路径）`);
        if (action === 'preview') {
          const absolute = path.resolve(filePath);
          assertXlsxPath(absolute);
          const info = await stat(absolute).catch(() => null);
          if (info === null || !info.isFile()) throw new Error(`找不到文件：${absolute}`);
          if (info.size > MAX_FILE_BYTES) throw new Error('Excel 文件超过 20 MB');
          const workbook = parseWorkbook(await readFile(absolute));
          const cards = extractCharacters(workbook, {
            fileName: path.basename(absolute),
            sourceKey: `preview:${absolute.toLowerCase()}`,
            sourcePath: absolute,
            sheet: args?.sheet,
          });
          return {
            ok: true,
            action,
            message: `预览成功：识别到 ${cards.length} 个角色（尚未保存）。`,
            cards: cards.map(summarizeCard),
            samples: cards.slice(0, 3).map(formatCardForModel),
            total: cards.length,
          };
        }
        const result = await importPath(filePath, {
          activate: args?.activate === true,
          sessionId,
          sheet: args?.sheet,
        });
        const active = activeCardFor(sessionId);
        const lines = [`导入完成：新增 ${result.report.added}，更新 ${result.report.updated}，移除 ${result.report.removed}。`];
        if (result.activeCard !== null) lines.push(`已在当前会话启用「${result.activeCard.name}」。之后的回复会保持该角色的设定。`);
        else lines.push('角色卡已保存。调用 action=activate 或使用输入框右侧的「角色卡」面板启用。');
        return {
          ok: true,
          action,
          message: lines.join('\n'),
          cards: state.cards.map(summarizeCard),
          total: state.cards.length,
          report: result.report,
          ...(active ? { activeCardId: active.id } : {}),
        };
      }

      const cardId = String(args?.cardId ?? '').trim();
      if (action === 'get') {
        if (cardId === '') throw new Error('action=get 需要提供 cardId');
        const card = cardById(cardId);
        if (card === undefined) throw new Error(`找不到角色卡「${cardId}」。`);
        return {
          ok: true,
          action,
          message: cardMessage(card),
          card,
          prompt: buildCharacterPrompt(card),
          ...(activeCardFor(sessionId)?.id === card.id ? { activeCardId: card.id } : {}),
        };
      }

      if (action === 'activate') {
        if (cardId === '') throw new Error('action=activate 需要提供 cardId');
        const card = await activate(sessionId, cardId);
        return {
          ok: true,
          action,
          message: `已启用角色「${card.name}」。${sessionId ? '当前会话' : '全局默认'}之后的回复会以该角色的身份进行。`,
          card: summarizeCard(card),
          activeCardId: card.id,
        };
      }

      if (action === 'deactivate') {
        await deactivate(sessionId);
        return {
          ok: true,
          action,
          message: `已停止角色扮演（${sessionId ? '当前会话' : '全局默认'}）。之后的回复恢复为普通助手。`,
        };
      }

      if (action === 'delete') {
        if (cardId === '') throw new Error('action=delete 需要提供 cardId');
        const removed = await deleteCard(cardId);
        return {
          ok: true,
          action,
          message: `已删除角色卡「${removed.name}」。`,
          cards: state.cards.map(summarizeCard),
          total: state.cards.length,
          ...(activeCardFor(sessionId) ? { activeCardId: activeCardFor(sessionId).id } : {}),
        };
      }

      throw new Error(`未知 action「${action}」。可用：list / import / preview / get / activate / deactivate / delete。`);
    },
  };

  ctx.effect(() => ctx.tools.register(tool), 'excel-character-card: tool');

  // ---------------------------------------------------------------------
  // Dynamic role-play system prompt
  // ---------------------------------------------------------------------
  let promptOrder = 10100;
  try {
    const suffixOrder = ctx.systemPrompt.getSectionOrder('DEPLOYMENT_PERSONA_SUFFIX');
    if (Number.isFinite(suffixOrder)) promptOrder = Math.max(1, suffixOrder - 1);
  } catch {
    /* older/newer prompt registry: keep the numeric fallback */
  }

  ctx.effect(
    () =>
      ctx.systemPrompt.section({
        name: 'excel-character-card:roleplay',
        order: promptOrder,
        interpolate: false,
        text: (context) => {
          try {
            const sessionId = context?.agent?.id === undefined ? undefined : String(context.agent.id);
            const card = activeCardFor(sessionId);
            return card === null ? '' : buildCharacterPrompt(card);
          } catch {
            return '';
          }
        },
      }),
    'excel-character-card: role-play prompt',
  );

  // ---------------------------------------------------------------------
  // Browser JSON API (same origin; the panel uses it for uploads/activation)
  // ---------------------------------------------------------------------
  ctx.effect(
    () =>
      ctx.webServer.register({
        kind: 'prefix',
        path: API_PREFIX,
        handler: async (req, res) => {
          try {
            const url = new URL(req.url ?? '/', 'http://localhost');
            const route = url.pathname.startsWith(API_PREFIX)
              ? url.pathname.slice(API_PREFIX.length) || '/'
              : '/';
            await load();

            if (req.method === 'GET' && route === '/state') {
              const sessionId = url.searchParams.get('sessionId') ?? undefined;
              return sendJson(res, 200, { ok: true, result: publicState(sessionId) });
            }

            if (req.method === 'GET' && route === '/card') {
              const cardId = url.searchParams.get('id') ?? '';
              const card = cardById(cardId);
              if (card === undefined) return sendJson(res, 404, { ok: false, error: `找不到角色卡「${cardId}」` });
              return sendJson(res, 200, { ok: true, result: { card, prompt: buildCharacterPrompt(card) } });
            }

            if (req.method === 'GET' && route === '/health') {
              return sendJson(res, 200, { ok: true, result: { cards: state.cards.length, dataDir: DATA_DIR } });
            }

            if (req.method !== 'POST') {
              return sendJson(res, 405, { ok: false, error: 'Method Not Allowed' });
            }

            // Requiring JSON forces a CORS preflight for cross-origin callers
            // (this route does not answer preflights), which keeps a random
            // local web page from silently importing cards with a form POST.
            const contentType = String(req.headers['content-type'] ?? '').toLowerCase();
            if (!contentType.includes('application/json')) {
              return sendJson(res, 415, { ok: false, error: 'Content-Type 必须是 application/json' });
            }

            const raw = await readBody(req);
            let body = {};
            if (raw.length > 0) {
              try {
                body = JSON.parse(raw.toString('utf8'));
              } catch {
                return sendJson(res, 400, { ok: false, error: '请求体不是合法 JSON' });
              }
            }
            if (body === null || typeof body !== 'object' || Array.isArray(body)) {
              return sendJson(res, 400, { ok: false, error: '请求体必须是 JSON 对象' });
            }

            const sessionId = typeof body.sessionId === 'string' && body.sessionId.trim() !== '' ? body.sessionId.trim() : undefined;

            if (route === '/import') {
              if (typeof body.fileName !== 'string' || typeof body.dataBase64 !== 'string') {
                return sendJson(res, 400, { ok: false, error: '缺少 fileName 或 dataBase64' });
              }
              const result = await importUpload(body.fileName, body.dataBase64, {
                activate: body.activate === true,
                sessionId,
                sheet: typeof body.sheet === 'string' ? body.sheet : undefined,
              });
              return sendJson(res, 200, {
                ok: true,
                result: {
                  ...publicState(sessionId),
                  report: result.report,
                  imported: result.cards.map(summarizeCard),
                },
              });
            }

            if (route === '/activate') {
              const cardId = typeof body.cardId === 'string' ? body.cardId.trim() : '';
              if (cardId === '') return sendJson(res, 400, { ok: false, error: '缺少 cardId' });
              const card = await activate(sessionId, cardId);
              return sendJson(res, 200, {
                ok: true,
                result: { ...publicState(sessionId), activated: summarizeCard(card) },
              });
            }

            if (route === '/deactivate') {
              await deactivate(sessionId);
              return sendJson(res, 200, { ok: true, result: publicState(sessionId) });
            }

            if (route === '/delete') {
              const cardId = typeof body.cardId === 'string' ? body.cardId.trim() : '';
              if (cardId === '') return sendJson(res, 400, { ok: false, error: '缺少 cardId' });
              await deleteCard(cardId);
              return sendJson(res, 200, { ok: true, result: publicState(sessionId) });
            }

            return sendJson(res, 404, { ok: false, error: 'Not Found' });
          } catch (error) {
            return sendError(res, error);
          }
        },
      }),
    'excel-character-card: HTTP API',
  );

  // Load eagerly so the prompt and list tool are correct before the first call.
  void load();
}

export default { name, inject, apply };

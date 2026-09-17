/**
 * <tableEdit> 状态块 → shujuku 表格 的「状态桥」源码骨架（通用版，可直接抄）
 *
 * 说明文档：同目录 README.md
 *
 * 用法：
 *   1) 把 SCHEMA 换成你自己卡模板里的物理表名与 DDL 列名（示例全是占位名 demo_* / example_*）；
 *   2) 整段贴进酒馆助手（JS-Slash-Runner）脚本，脚本名建议「结构化状态同步桥」；
 *   3) 世界书里配一条 constant 协议条目（表名、列约束、where 定位键写在里面）；
 *   4) 走状态块的表把 updateConfig.updateFrequency 设为 0，再回包。
 *
 * 骨架与真实实现结构一致，但错误提示与边界做了简化：真机上建议把每个 throw 换成
 * 带表名/列名的明确文案，并用 toast 提示（成功 / 排队 / 失败各一种）。
 */
(function installTableEditBridge(){
  'use strict';
  const BUTTON_NAME = '同步最后一条状态';
  const TIME_PATTERN = '^\\d{4}-\\d{2}-\\d{2} \\d{2}:\\d{2}$';
  const COUNTER_PATTERN = '^项A=\\d+,项B=\\d+$';

  // ① 表/列白名单：校验与写入都读它，必须与世界书协议条目逐条对应
  const SCHEMA = {
    demo_state: {                                  // 单行状态表：只有 row_id=1 一行，只更新不新增
      display: '示例状态表', singleton: true, key: 'row_id',
      insertColumns: { row_id: { display: 'row_id', type: 'int', default: 1 } },
      columns: {
        current_location: { display: '当前位置', type: 'text', required: true },
        example_time: { display: '示例时间', type: 'text', required: true, pattern: TIME_PATTERN, patternHint: 'YYYY-MM-DD HH:MM' },
        example_flag: { display: '示例标记', type: 'text', enum: ['是', '否'], default: '否' },
        example_counter: { display: '示例计数', type: 'text', required: true, pattern: COUNTER_PATTERN, patternHint: '项A=N,项B=N' }
      }
    },
    demo_npc_state: {                              // 按名字定位的关系/数值表：首次互动可新增
      display: '示例NPC状态表', singleton: false, key: 'npc_name',
      columns: {
        npc_name: { display: 'NPC姓名', type: 'text', required: true },
        example_score: { display: '示例数值', type: 'int', min: -100, max: 100, default: 0 },
        example_tag: { display: '示例标签', type: 'text', maxLength: 10, default: '普通' },
        updated_at: { display: '更新时间', type: 'text', default: '' }
      }
    },
    demo_npc_profile: {                            // 列多的档案表：新增行缺列自动补占位（避免空白列被填表 AI 跳过）
      display: '示例NPC档案表', singleton: false, key: 'name',
      insertFillMissing: '未知',
      columns: {
        name: { display: '姓名', type: 'text', required: true },
        appearance: { display: '外貌', type: 'text', required: true },
        brief_intro: { display: '人物简介', type: 'text' },
        alias: { display: '别称', type: 'text' }
      }
    }
  };

  // ② 跨楼层共享状态：楼层是独立 iframe，这些必须挂在顶层窗口上
  function hostWindow() { try { return (window.top && window.top !== window) ? window.top : window; } catch (e) { return window; } }
  const done = hostWindow().__tableEditDone = hostWindow().__tableEditDone || new Set();          // 去重：同一状态块只写一次
  const config = hostWindow().__tableEditBridgeConfig = hostWindow().__tableEditBridgeConfig || { strictColumns: false };
  const fillState = hostWindow().__tableEditFillState = hostWindow().__tableEditFillState || {
    running: false, queued: false, inflight: false, hooked: false, activityHooked: false,
    timer: null, startedAt: 0, lastActivityAt: 0, triggerAt: 0, attempt: 0,
    delays: [30000, 45000, 60000],   // 试写退避节奏
    maxMs: 10 * 60 * 1000,           // 超过这么久改为慢等（不放弃）
    staleMs: 3 * 60 * 1000,          // 「填表进行中」标记安静这么久就视为过期
    quietMs: 20000,                  // 距最近一次表格写入超过这么久才考虑写
    minWaitMs: 45000                 // 收到回复后至少等这么久（给「填表开始」通知留时间）
  };
  const instanceToken = {};
  hostWindow().__tableEditActiveInstance = instanceToken;                                        // 只让最新实例处理事件
  const isActiveInstance = () => { try { return hostWindow().__tableEditActiveInstance === instanceToken; } catch (e) { return false; } };

  // ③ 找 API：不同宿主里 AutoCardUpdaterAPI 可能挂在 window / parent / top
  function roots() {
    const list = [];
    const add = (w) => { if (w && !list.includes(w)) list.push(w); };
    add(window);
    let w = window;
    for (let i = 0; i < 6 && w; i += 1) { try { if (!w.parent || w.parent === w) break; add(w.parent); w = w.parent; } catch (e) { break; } }
    return list;
  }
  function apiInstances(method) {
    const list = [];
    const push = (A) => { try { if (A && (!method || typeof A[method] === 'function') && !list.includes(A)) list.push(A); } catch (e) {} };
    try { roots().forEach((r) => push(r.AutoCardUpdaterAPI)); } catch (e) {}
    try { push(window.AutoCardUpdaterAPI); } catch (e) {}
    try { if (window.parent && window.parent !== window) push(window.parent.AutoCardUpdaterAPI); } catch (e) {}
    try { if (window.top && window.top !== window) push(window.top.AutoCardUpdaterAPI); } catch (e) {}
    return list;
  }
  const api = () => apiInstances('updateCell')[0] || apiInstances('insertRow')[0] || null;

  // ④ 取状态块：取「最后一对 <tableEdit> … </tableEdit>」，从后往前找第一个能解析成 JSON 的
  function decodeEscaped(text) {
    return String(text || '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, '&');
  }
  function normalizeBody(raw) {
    return String(raw || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```$/i, '').trim();
  }
  function extractBlock(rawText) {
    const source = String(rawText || '');
    let failure = null;
    for (const candidate of [source, decodeEscaped(source)]) {
      const opens = [];
      const openRe = /<tableEdit\b[^>]*>/gi;
      let open = null;
      while ((open = openRe.exec(candidate)) !== null) opens.push(open.index + open[0].length);
      for (let i = opens.length - 1; i >= 0; i -= 1) {
        const rest = candidate.slice(opens[i]);
        const close = rest.match(/<\/tableEdit\s*>/i);
        if (!close) continue;
        const body = normalizeBody(rest.slice(0, close.index));
        try { return { ok: true, payload: JSON.parse(body) }; }
        catch (error) { if (!failure) failure = { error: String(error.message || error), body: body.slice(0, 160), blocks: opens.length }; }
      }
    }
    return failure ? { ok: false, ...failure } : null;
  }

  // ⑤ 校验与规划：整块先过一遍，任何一项不合法就抛错（不会写半条）
  function specOf(tableKey, name, allowInsertOnly) {
    const spec = SCHEMA[tableKey];
    const raw = String(name == null ? '' : name).trim();
    if (spec.columns[raw]) return { ddl: raw, spec: spec.columns[raw] };
    if (allowInsertOnly && spec.insertColumns && spec.insertColumns[raw]) return { ddl: raw, spec: spec.insertColumns[raw] };
    for (const ddl in spec.columns) { if (spec.columns[ddl].display === raw) return { ddl, spec: spec.columns[ddl] }; }
    return null;
  }
  function resolveTable(name) {
    const raw = String(name == null ? '' : name).trim();
    if (SCHEMA[raw]) return raw;
    for (const key in SCHEMA) { if (SCHEMA[key].display === raw || ('sheet_' + key) === raw) return key; }
    return null;
  }
  function coerce(display, spec, raw, notes) {
    if (spec.type === 'int') {
      const value = Number(String(raw).trim());
      if (!Number.isInteger(value)) throw new Error(display + '.' + spec.display + ' 必须是整数');
      if (spec.min != null && value < spec.min) throw new Error(display + '.' + spec.display + ' 不能小于 ' + spec.min);
      if (spec.max != null && value > spec.max) throw new Error(display + '.' + spec.display + ' 不能大于 ' + spec.max);
      return value;
    }
    let text = String(raw == null ? '' : raw).trim();
    if (spec.enum && text && !spec.enum.includes(text)) throw new Error(display + '.' + spec.display + ' 只能写 ' + spec.enum.join('／'));
    if (spec.pattern && text && !new RegExp(spec.pattern).test(text)) throw new Error(display + '.' + spec.display + ' 格式必须是 ' + (spec.patternHint || spec.pattern));
    if (spec.maxLength && text.length > spec.maxLength) { notes.push(spec.display + ' 超过 ' + spec.maxLength + ' 字，已截断'); text = text.slice(0, spec.maxLength); }
    return text;
  }
  function buildPlan(payload) {
    const plan = new Map();
    const notes = [];
    const push = (tableKey, keyValue, columns) => {
      const signature = tableKey + '|' + keyValue;
      const current = plan.get(signature);
      if (!current) {
        plan.set(signature, { tableKey, display: SCHEMA[tableKey].display, keyValue, columns: new Map(columns.map((column) => [column.ddl, column])) });
        return;
      }
      columns.forEach((column) => current.columns.set(column.ddl, column));
      notes.push(SCHEMA[tableKey].display + ' 的 ' + keyValue + ' 出现多次写入，已合并为一次');
    };
    const collect = (tableKey, source, allowInsertOnly) => {
      const spec = SCHEMA[tableKey];
      const columns = [];
      for (const name of Object.keys(source || {})) {
        const found = specOf(tableKey, name, allowInsertOnly);
        if (!found) {
          if (config.strictColumns) throw new Error(spec.display + ' 不允许写入列：' + name);
          notes.push(spec.display + ' 没有可写列「' + name + '」，已忽略该列');
          continue;
        }
        columns.push({ ddl: found.ddl, display: found.spec.display, value: coerce(spec.display, found.spec, source[name], notes) });
      }
      return columns;
    };
    const keyOf = (spec, source) => {
      if (spec.singleton) return 'singleton';
      const raw = source[spec.key] != null ? source[spec.key] : source[spec.columns[spec.key].display];
      const value = String(raw == null ? '' : raw).trim();
      if (!value) throw new Error(spec.display + ' 必须提供定位键 ' + spec.key);
      return value;
    };
    for (const item of payload.updates || []) {
      const tableKey = resolveTable(item.table);
      if (!tableKey) throw new Error('表不允许或表名不符合 DDL：' + item.table);
      const spec = SCHEMA[tableKey];
      const where = item.where || {};
      for (const name of Object.keys(where)) {
        const allowed = spec.singleton ? ['row_id'] : [spec.key, 'row_id'];
        const found = specOf(tableKey, name, true);
        if (!found || !allowed.includes(found.ddl)) throw new Error(spec.display + ' 的 where 只允许 ' + allowed.join(' 或 '));
      }
      push(tableKey, keyOf(spec, where), collect(tableKey, item.set || {}, false));
    }
    for (const item of payload.inserts || []) {
      const tableKey = resolveTable(item.table);
      if (!tableKey) throw new Error('表不允许或表名不符合 DDL：' + item.table);
      const spec = SCHEMA[tableKey];
      const row = item.row || {};
      const columns = collect(tableKey, row, true);
      if (spec.insertFillMissing) {
        const provided = new Set(columns.map((column) => column.ddl));
        const missing = Object.keys(spec.columns).filter((ddl) => !provided.has(ddl));
        missing.forEach((ddl) => {
          const columnSpec = spec.columns[ddl];
          columns.push({ ddl, display: columnSpec.display, value: columnSpec.enum ? columnSpec.enum[0] : spec.insertFillMissing });
        });
        if (missing.length) notes.push(spec.display + ' 缺少 ' + missing.map((ddl) => spec.columns[ddl].display).join('、') + '，已补「' + spec.insertFillMissing + '」');
      }
      push(tableKey, keyOf(spec, row), columns);
    }
    return { entries: [...plan.values()].map((entry) => ({ ...entry, columns: [...entry.columns.values()] })), notes };
  }

  // ⑥ 写入：读表 → 定位行（表头同时认显示名与 DDL 列名）→ updateRow / updateCell / insertRow → 刷新
  function readRows(A, tableKey) {
    try {
      const all = A.exportTableAsJson() || {};
      const display = SCHEMA[tableKey].display;
      const byKey = all['sheet_' + tableKey];
      if (byKey && Array.isArray(byKey.content)) return byKey.content;
      if (Array.isArray(byKey)) return byKey;
      for (const key in all) { if (!key.startsWith('sheet_')) continue; const sheet = all[key]; if (sheet && sheet.name === display && Array.isArray(sheet.content)) return sheet.content; }
      if (Array.isArray(all[display])) return all[display];
      if (all[display] && Array.isArray(all[display].content)) return all[display].content;
      return null;
    } catch (e) { return null; }
  }
  function headerIndexOf(header, ...names) {
    if (!Array.isArray(header)) return -1;
    for (const name of names) { if (!name) continue; const index = header.indexOf(name); if (index >= 0) return index; }
    return -1;
  }
  function findRowIndex(tableKey, keyValue, rows, header) {
    const spec = SCHEMA[tableKey];
    const names = spec.singleton ? ['row_id'] : [spec.columns[spec.key].display, spec.key];
    const index = headerIndexOf(header, ...names);
    if (index < 0) return spec.singleton && rows.length > 1 ? 1 : -1;
    for (let row = 1; row < rows.length; row += 1) {
      const cell = rows[row] ? rows[row][index] : null;
      const want = spec.singleton ? 1 : keyValue;
      if (cell != null && String(cell) === String(want)) return row;
    }
    return -1;
  }
  function lockStateOf(A, tableKey) {
    try { return typeof A.getTableLockState === 'function' ? A.getTableLockState('sheet_' + tableKey) : null; } catch (e) { return null; }
  }
  function writableColumns(lock, rowIndex, header, entry, notes) {
    if (!lock) return entry.columns;
    if (rowIndex >= 0 && Array.isArray(lock.rows) && lock.rows.some((row) => Number(row) === rowIndex)) {
      notes.push(entry.display + ' 第 ' + rowIndex + ' 行已被锁定，已跳过');
      return [];
    }
    return entry.columns.filter((column) => {
      const index = headerIndexOf(header, column.display, column.ddl);
      if (Array.isArray(lock.cols) && lock.cols.some((col) => Number(col) === index)) { notes.push(entry.display + ' 列「' + column.display + '」已被锁定，已跳过'); return false; }
      const hit = (lock.cells || []).some((cell) => Array.isArray(cell) ? (Number(cell[0]) === rowIndex && Number(cell[1]) === index) : String(cell) === rowIndex + ':' + index);
      if (hit) { notes.push(entry.display + ' 单元格 ' + rowIndex + ':' + index + ' 已被锁定，已跳过'); return false; }
      return true;
    });
  }
  async function writeEntry(A, entry, rowIndex, header) {
    const data = {};
    entry.columns.forEach((column) => { data[column.display] = column.value; });
    if (rowIndex >= 0) {
      if (typeof A.updateRow === 'function') {
        const result = await A.updateRow({ tableName: entry.display, rowIndex, data, skipNotify: true });
        if (result === false || result == null) throw new Error(entry.display + ' 更新第 ' + rowIndex + ' 行失败');
        return;
      }
      for (const column of entry.columns) {
        const index = headerIndexOf(header, column.display, column.ddl);
        if (index < 0) throw new Error(entry.display + ' 缺少列 ' + column.display);
        await A.updateCell({ tableName: entry.display, rowIndex, colIdentifier: index, value: column.value, skipNotify: true });
      }
      return;
    }
    const inserted = await A.insertRow({ tableName: entry.display, data, skipNotify: true });
    if (inserted === -1 || inserted === false || inserted == null) throw new Error(entry.display + ' 插入新行失败');
  }

  // ⑦ 与自动填表抢窗口的闸门：没有「填表结束」回调，只能靠只读信号 + 试写推断
  const setIntervalFn = typeof setInterval === 'function' ? setInterval : null;
  const clearIntervalFn = typeof clearInterval === 'function' ? clearInterval : null;
  const setTimeoutFn = typeof setTimeout === 'function' ? setTimeout : null;
  const clearTimeoutFn = typeof clearTimeout === 'function' ? clearTimeout : null;
  function fillUiActive() {
    const docs = [];
    try { docs.push(window.document); } catch (e) {}
    try { if (window.parent && window.parent !== window) docs.push(window.parent.document); } catch (e) {}
    try { if (window.top && window.top !== window) docs.push(window.top.document); } catch (e) {}
    return docs.some((doc) => { try { return !!doc.querySelector('[id^="acu-stop-"]'); } catch (e) { return false; } });
  }
  function fillBlocksWrite(allowRunning) {
    if (fillUiActive()) return true;
    if (allowRunning || !fillState.running) return false;
    const quietSince = Math.max(fillState.lastActivityAt || 0, fillState.startedAt || 0);
    return Date.now() - quietSince < fillState.staleMs;
  }
  function canWriteNow(allowRunning, skipTimeGates) {
    if (fillBlocksWrite(allowRunning)) return false;
    if (skipTimeGates) return true;
    const now = Date.now();
    if (now - (fillState.triggerAt || 0) < fillState.minWaitMs) return false;
    if (now - (fillState.lastActivityAt || 0) < fillState.quietMs) return false;
    return true;
  }
  function clearRetryTimer() { if (fillState.timer && clearTimeoutFn) clearTimeoutFn(fillState.timer); fillState.timer = null; }
  function scheduleRetry() {
    if (fillState.timer || !setTimeoutFn) return;
    const base = fillState.delays[Math.min(fillState.attempt, fillState.delays.length - 1)];
    const slow = fillState.startedAt && Date.now() - fillState.startedAt > fillState.maxMs;
    fillState.timer = setTimeoutFn(() => {
      fillState.timer = null;
      fillState.attempt += 1;
      if (!canWriteNow(true)) { scheduleRetry(); return; }
      runGuarded({ probe: true });
    }, slow ? Math.max(base, 300000) : base);
  }
  function onFillStart() {
    fillState.running = true;
    fillState.startedAt = fillState.startedAt || Date.now();
    fillState.lastActivityAt = Date.now();
    fillState.attempt = 0;
    clearRetryTimer();
    if (fillState.queued) scheduleRetry();
  }
  function touchActivity() { fillState.lastActivityAt = Date.now(); }
  function ensureHooks() {
    apiInstances('registerTableFillStartCallback').forEach((A) => {
      A.__fillStartDispatch = A.__fillStartDispatch || function () { try { onFillStart(); } catch (e) {} };
      if (!A.__fillStartRegistered) { try { A.registerTableFillStartCallback(A.__fillStartDispatch); A.__fillStartRegistered = true; fillState.hooked = true; } catch (e) {} }
    });
    apiInstances('registerTableUpdateCallback').forEach((A) => {
      A.__activityDispatch = A.__activityDispatch || function () { try { touchActivity(); } catch (e) {} };
      if (!A.__activityRegistered) { try { A.registerTableUpdateCallback(A.__activityDispatch); A.__activityRegistered = true; fillState.activityHooked = true; } catch (e) {} }
    });
  }
  function watchHooks() {
    const attempt = () => { try { return !!(ensureHooks(), fillState.hooked && fillState.activityHooked); } catch (e) { return false; } };
    if (attempt() || !setIntervalFn) return;
    let waited = 0;
    const timer = setIntervalFn(() => { waited += 2000; if (attempt() || waited >= 60000) { if (clearIntervalFn) clearIntervalFn(timer); } }, 2000);
  }

  // ⑧ 读最后一条 AI 回复并抽出状态块（兼容 ChatMessage.message / 原生 mes / swipes）
  function findGlobal(name) { for (const root of roots()) { try { if (root && root[name] != null) return root[name]; } catch (e) {} } return undefined; }
  async function pickCandidate() {
    const helper = findGlobal('TavernHelper');
    let list = [];
    if (helper && typeof helper.getChatMessages === 'function') {
      let last = null;
      try { last = helper.getLastMessageId && helper.getLastMessageId(); } catch (e) {}
      try { list = await helper.getChatMessages(typeof last === 'number' ? '0-' + last : -1, { include_swipes: true }) || []; } catch (e) { list = []; }
    }
    if (!list.length) {
      const st = findGlobal('SillyTavern');
      const context = (st && st.getContext && st.getContext()) || {};
      list = context.chat || [];
    }
    for (let index = list.length - 1; index >= 0; index -= 1) {
      const message = list[index] || {};
      if (message.is_user) continue;
      const text = message.message != null ? message.message
        : message.mes != null ? message.mes
          : Array.isArray(message.swipes) ? message.swipes[Math.max(0, Number(message.swipe_id || 0))] : (message.content || '');
      const block = extractBlock(text);
      if (block) return { messageId: message.message_id != null ? message.message_id : (message.id != null ? message.id : index), block };
    }
    return null;
  }
  function hash(text) { let h = 2166136261; for (const ch of String(text)) h = Math.imul(h ^ ch.charCodeAt(0), 16777619); return (h >>> 0).toString(16); }

  // ⑨ 一次同步：闸门 → 取块 → 规划 → 写 → 刷新
  async function apply(options) {
    if (!(options && options.force) && fillBlocksWrite(!!(options && options.probe))) return { queued: true };
    const A = api();
    if (!A) throw new Error('AutoCardUpdaterAPI 不可用');
    const found = await pickCandidate();
    if (!found) return { empty: true };
    if (!found.block.ok) throw new Error('tableEdit 状态块 JSON 解析失败：' + found.block.error + '；内容片段：' + found.block.body);
    const key = found.messageId + ':' + hash(JSON.stringify(found.block.payload));
    if (done.has(key)) return { duplicate: true };
    const plan = buildPlan(found.block.payload);
    let count = 0;
    for (const entry of plan.entries) {
      const rows = readRows(A, entry.tableKey);
      if (!Array.isArray(rows)) throw new Error(entry.display + ' 读不到当前表格内容，已跳过写入以避免重复插入');
      const header = Array.isArray(rows[0]) ? rows[0] : [];
      const rowIndex = findRowIndex(entry.tableKey, entry.keyValue, rows, header);
      const columns = writableColumns(lockStateOf(A, entry.tableKey), rowIndex, header, entry, plan.notes);
      if (!columns.length) continue;
      await writeEntry(A, { ...entry, columns }, rowIndex, header);
      count += 1;
    }
    if (typeof A.refreshDataAndWorldbook === 'function') await A.refreshDataAndWorldbook();
    done.add(key);
    return { ok: true, count, notes: plan.notes };
  }

  // ⑩ 唯一入口：probe 只用于探测填表是否结束；force 只用于用户手动越权
  async function runGuarded(options) {
    const probe = !!(options && options.probe);
    const manual = !!(options && options.manual);
    const force = !!(options && options.force);
    if (!probe && !manual) fillState.triggerAt = Date.now();
    watchHooks();
    if (fillState.inflight) { fillState.queued = true; scheduleRetry(); return { queued: true }; }
    if (!force && !canWriteNow(probe, manual)) { fillState.queued = true; scheduleRetry(); return { queued: true }; }
    fillState.inflight = true;
    try {
      const result = await apply({ probe, force });
      if (result && result.ok) {
        fillState.queued = false; fillState.running = false; fillState.startedAt = 0;
        fillState.triggerAt = 0; fillState.attempt = 0; clearRetryTimer();
      }
      return result;
    } catch (error) {
      const text = String((error && error.message) || error);
      if (/AI 填表正在进行中|TableUpdateCommit|precondition/.test(text)) {
        fillState.queued = true; fillState.running = true;
        fillState.startedAt = fillState.startedAt || Date.now();
        scheduleRetry();
        return { queued: true };
      }
      throw error;
    } finally {
      fillState.inflight = false;
    }
  }

  // ⑪ 入口注册：手动按钮（含越权确认）+ 事件钩子（只让最新实例干活）
  function reg() {
    const on = findGlobal('eventOn');
    const getButtonEvent = findGlobal('getButtonEvent');
    const appendButtons = findGlobal('appendInexistentScriptButtons');
    const events = findGlobal('tavern_events');
    if (appendButtons) appendButtons([{ name: BUTTON_NAME, visible: true }]);
    if (on && getButtonEvent) {
      on(getButtonEvent(BUTTON_NAME), async () => {
        const result = await runGuarded({ manual: true });
        if (!(result && result.queued)) return;
        const approved = typeof window.confirm === 'function'
          && window.confirm('状态桥判断「AI 填表仍在进行」。\n\n确认填表已结束就点「确定」：越权立刻同步（真在填表会被拒绝，不会写坏数据）。');
        if (approved) await runGuarded({ manual: true, force: true });
      });
    }
    if (on && events) {
      const hook = (event, delay) => {
        if (!event) return;
        on(event, () => {
          if (!isActiveInstance()) return;
          if (setTimeoutFn) setTimeoutFn(() => { if (isActiveInstance()) runGuarded(); }, delay);
        });
      };
      hook(events.MESSAGE_RECEIVED, 300);
      hook(events.MESSAGE_UPDATED, 300);
      hook(events.MESSAGE_SWIPED, 300);
      hook(events.GENERATION_ENDED, 300);
      hook(events.MESSAGE_RENDERED, 500);
    }
    watchHooks();
  }

  window.TableEditBridge = {
    apply,
    runGuarded,
    extractBlock,
    schema: SCHEMA,
    fillState,
    config,
    forceSync: () => runGuarded({ manual: true, force: true }),
    setStrictColumns: (value) => { config.strictColumns = value !== false; },
    setProbePolicy: (policy) => { ['quietMs', 'minWaitMs', 'staleMs', 'maxMs'].forEach((key) => { if (policy && Number.isFinite(policy[key])) fillState[key] = policy[key]; }); }
  };
  reg();
})();

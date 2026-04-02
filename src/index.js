const DEFAULT_TIMEZONE = 'Europe/Kyiv';
const DEFAULT_CALENDAR_ID = 'techno.perspektiva@gmail.com';
const DEFAULT_MODEL = 'gemini-2.5-flash';
const MAX_TELEGRAM_MESSAGE = 3800;

const SHEETS = {
  TASKS: 'Задачі',
  BUDGET: 'Бюджет',
  SHOPPING: 'Покупки',
  LOG: 'Журнал',
  NOTES: 'Нотатки'
};

const TASK_HEADERS = ['task_id','created_at','updated_at','chat_id','user_id','username','title','details','due_at','due_date','status','calendar_event_id','source_text'];
const BUDGET_HEADERS = ['entry_id','created_at','updated_at','chat_id','user_id','username','type','amount','currency','category','note','source_text'];
const SHOPPING_HEADERS = ['item_id','created_at','updated_at','chat_id','user_id','username','list_type','item_name','details','quantity','status','source_text'];
const NOTE_HEADERS = ['note_id','created_at','chat_id','user_id','username','note_text','source'];
const LOG_HEADERS = ['created_at','chat_id','user_id','username','input_type','original_text','normalized_text','action','result_summary','calendar_event_id','status'];

const MENU = {
  CREATE_TASK: '➕ Створити задачу',
  LIST_TASKS: '📋 Список задач',
  COMPLETE_TODAY: '✅ Виконати сьогодні',
  TODAY: '📅 Сьогодні',
  BUDGET: '💰 Бюджет',
  PRODUCTS: '🛒 Купити продукти',
  NOTE: '📝 Нотатка',
  BUY: '📦 Що купити',
  SEARCH: '🔎 Знайти',
  NEARBY: '📍 Місця поруч',
  WEEK_REPORT: '📊 Звіт за тиждень'
};

const STATES = {
  WAIT_TASK: 'wait_task',
  WAIT_NOTE: 'wait_note',
  WAIT_BUDGET: 'wait_budget',
  WAIT_PRODUCTS: 'wait_products',
  WAIT_BUY: 'wait_buy',
  WAIT_SEARCH: 'wait_search',
  WAIT_NEARBY_LOCATION: 'wait_nearby_location',
  WAIT_NEARBY_QUERY: 'wait_nearby_query',
  WAIT_COMPLETE_TODAY: 'wait_complete_today'
};

const TASK_STATUS = { NEW:'Нова', DONE:'Виконано', MOVED:'Перенесено', CANCELED:'Скасовано' };
const SHOPPING_STATUS = { OPEN:'Купити', DONE:'Куплено' };

const TEXT = {
  START: '<b>Привіт. Я твій особистий асистент.</b>\n\nПрацюю стабільно в 4 речах:\n• задачі\n• календар\n• бюджет\n• покупки\n\nТакож можу зберігати нотатки і робити простий пошук.\n\nГолосові поки працюють в обережному режимі: короткі голосові до 30 секунд.',
  HELP: '<b>Приклади:</b>\n\n• Подзвонити клієнту завтра о 15:00\n• Тренування сьогодні о 19:00\n• Витрата 250 грн таксі\n• Молоко 2 шт\n• Купити кабель USB-C\n• Занотуй: подзвонити майстру\n• Знайди новини про ...\n• Кав\'ярня біля Позняків',
  RESET: 'Контекст очищено.',
  SEARCH_FAIL: 'Пошук тимчасово не відповідає. Спробуй коротший запит.',
  VOICE_FAIL: 'Не вдалося обробити голосове. Зараз стабільніше працює текст. Спробуй коротке голосове до 30 секунд або напиши текстом.',
  CALENDAR_FAIL: 'Не зміг записати або прочитати Google Calendar. Перевір, чи календар точно пошарений на service account.',
  SHEETS_FAIL: 'Не зміг записати в Google Sheets. Перевір доступ таблиці для service account.',
  GENERIC_FAIL: 'Сталася помилка. Спробуй ще раз.',
};

export default {
  async fetch(request, env, ctx) {
    try {
      const url = new URL(request.url);
      if (request.method === 'GET' && url.pathname === '/') return jsonResponse({ok:true, service:'personal-assistant-stable', status:'running'});
      if (request.method === 'GET' && url.pathname === '/setup') return await handleSetup(request, env);
      if (request.method === 'POST' && url.pathname === '/webhook') return await handleWebhook(request, env, ctx);
      return new Response('Not Found', {status:404});
    } catch (error) {
      return jsonResponse({ok:false, error:String(error?.message || error)}, 500);
    }
  }
};

async function handleSetup(request, env) {
  validateEnv(env);
  const url = new URL(request.url);
  if (url.searchParams.get('key') !== env.SETUP_KEY) return jsonResponse({ok:false, error:'Unauthorized'}, 401);
  const webhookUrl = `${url.protocol}//${url.host}/webhook`;
  const webhookResult = await telegramApi(env, 'setWebhook', {url:webhookUrl, secret_token:env.TELEGRAM_WEBHOOK_SECRET, allowed_updates:['message']});
  const commandsResult = await telegramApi(env, 'setMyCommands', {commands:[
    {command:'start', description:'Запустити асистента'},
    {command:'today', description:'План на сьогодні'},
    {command:'week', description:'Звіт за тиждень'},
    {command:'help', description:'Підказка'},
    {command:'reset', description:'Очистити контекст'}
  ]});
  return jsonResponse({ok:true, webhookUrl, webhookResult, commandsResult});
}

async function handleWebhook(request, env, ctx) {
  validateEnv(env);
  if (env.TELEGRAM_WEBHOOK_SECRET && request.headers.get('X-Telegram-Bot-Api-Secret-Token') !== env.TELEGRAM_WEBHOOK_SECRET) {
    return jsonResponse({ok:false, error:'Forbidden'}, 403);
  }
  const update = await request.json();
  const updateId = update.update_id;
  if (typeof updateId === 'number') {
    const dup = await env.STATE_KV.get(`update:${updateId}`);
    if (dup === '1') return jsonResponse({ok:true, duplicate:true});
    await env.STATE_KV.put(`update:${updateId}`, '1', {expirationTtl: 1800});
  }
  ctx.waitUntil(processUpdate(update, env));
  return jsonResponse({ok:true});
}

async function processUpdate(update, env) {
  try {
    if (update.message) await handleMessage(update.message, env);
  } catch (error) {
    console.error('processUpdate', error?.stack || error?.message || String(error));
  }
}

async function handleMessage(message, env) {
  const chatId = message.chat?.id;
  const from = message.from || {};
  if (!chatId) return;
  const rawText = (message.text || message.caption || '').trim();
  const state = await getState(env, chatId);

  if (rawText === '/start') {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.START, {reply_markup: mainMenuKeyboard()});
    return;
  }
  if (rawText === '/help') {
    await sendMessage(env, chatId, TEXT.HELP, {reply_markup: mainMenuKeyboard()});
    return;
  }
  if (rawText === '/reset') {
    await clearState(env, chatId);
    await sendMessage(env, chatId, TEXT.RESET, {reply_markup: mainMenuKeyboard()});
    return;
  }
  if (rawText === '/today' || sameBtn(rawText, MENU.TODAY)) {
    await sendTodayOverview(env, chatId);
    return;
  }
  if (rawText === '/week' || sameBtn(rawText, MENU.WEEK_REPORT)) {
    await sendWeeklyReport(env, chatId);
    return;
  }

  if (sameBtn(rawText, MENU.CREATE_TASK)) {
    await saveState(env, chatId, {step:STATES.WAIT_TASK});
    await sendMessage(env, chatId, 'Напиши задачу одним повідомленням.\n\nНаприклад: <i>Тренування сьогодні о 19:00</i>', {reply_markup:mainMenuKeyboard()});
    return;
  }
  if (sameBtn(rawText, MENU.LIST_TASKS)) return await sendTaskList(env, chatId);
  if (sameBtn(rawText, MENU.COMPLETE_TODAY)) return await prepareCompleteToday(env, chatId);
  if (sameBtn(rawText, MENU.BUDGET)) {
    await saveState(env, chatId, {step:STATES.WAIT_BUDGET});
    await sendMessage(env, chatId, 'Напиши дохід або витрату.\n\nНаприклад: <i>Витрата 250 грн таксі</i>', {reply_markup:mainMenuKeyboard()});
    return;
  }
  if (sameBtn(rawText, MENU.PRODUCTS)) {
    await saveState(env, chatId, {step:STATES.WAIT_PRODUCTS});
    await sendShoppingSummaryAndPrompt(env, chatId, 'products');
    return;
  }
  if (sameBtn(rawText, MENU.BUY)) {
    await saveState(env, chatId, {step:STATES.WAIT_BUY});
    await sendShoppingSummaryAndPrompt(env, chatId, 'buy');
    return;
  }
  if (sameBtn(rawText, MENU.NOTE)) {
    await saveState(env, chatId, {step:STATES.WAIT_NOTE});
    await sendMessage(env, chatId, 'Напиши нотатку одним повідомленням.', {reply_markup:mainMenuKeyboard()});
    return;
  }
  if (sameBtn(rawText, MENU.SEARCH)) {
    await saveState(env, chatId, {step:STATES.WAIT_SEARCH});
    await sendMessage(env, chatId, 'Напиши, що саме знайти.', {reply_markup:mainMenuKeyboard()});
    return;
  }
  if (sameBtn(rawText, MENU.NEARBY)) {
    await saveState(env, chatId, {step:STATES.WAIT_NEARBY_LOCATION});
    await sendMessage(env, chatId, 'Надішли геолокацію або напиши район / місто.', {reply_markup:nearbyLocationKeyboard()});
    return;
  }

  if (message.location && state?.step === STATES.WAIT_NEARBY_LOCATION) {
    await saveState(env, chatId, {step:STATES.WAIT_NEARBY_QUERY, location:{latitude:message.location.latitude, longitude:message.location.longitude}});
    await sendMessage(env, chatId, 'Напиши, що саме знайти поруч. Наприклад: <i>кав\'ярня</i>', {reply_markup:mainMenuKeyboard()});
    return;
  }

  if (message.voice || message.audio) {
    try {
      await sendMessage(env, chatId, 'Отримав голосове. Спробую коротко розшифрувати…', {reply_markup:mainMenuKeyboard()});
      const transcript = await transcribeTelegramAudio(env, message.voice || message.audio);
      if (!transcript?.trim()) throw new Error('empty transcript');
      await handleTextInput(env, chatId, from, transcript, 'voice');
    } catch (error) {
      console.error('voice', error?.stack || error?.message || String(error));
      await sendMessage(env, chatId, TEXT.VOICE_FAIL, {reply_markup:mainMenuKeyboard()});
    }
    return;
  }

  if (!rawText) {
    await sendMessage(env, chatId, 'Напиши текстом, що потрібно.', {reply_markup:mainMenuKeyboard()});
    return;
  }

  if (state?.step === STATES.WAIT_NEARBY_LOCATION) {
    await saveState(env, chatId, {step:STATES.WAIT_NEARBY_QUERY, areaText: rawText});
    await sendMessage(env, chatId, 'Тепер напиши, що саме знайти поруч.', {reply_markup:mainMenuKeyboard()});
    return;
  }
  if (state?.step === STATES.WAIT_NEARBY_QUERY) {
    await clearState(env, chatId);
    await runNearbyPlaces(env, chatId, rawText, state);
    return;
  }
  if (state?.step === STATES.WAIT_COMPLETE_TODAY) {
    await completeTaskFromSelection(env, chatId, rawText);
    await clearState(env, chatId);
    return;
  }

  await handleTextInput(env, chatId, from, rawText, 'text');
}

async function handleTextInput(env, chatId, from, rawText, inputType='text') {
  const state = await getState(env, chatId);
  try {
    if (state?.step === STATES.WAIT_NOTE) {
      await saveNote(env, chatId, from, rawText, inputType);
      await clearState(env, chatId);
      return await replyAndLog(env, chatId, from, inputType, rawText, 'add_note', 'Нотатку збережено.', 'ok');
    }
    if (state?.step === STATES.WAIT_BUDGET) {
      await addBudgetFromText(env, chatId, from, rawText, inputType);
      await clearState(env, chatId);
      return;
    }
    if (state?.step === STATES.WAIT_PRODUCTS) {
      await addShoppingItem(env, chatId, from, rawText, 'products', inputType);
      await clearState(env, chatId);
      return;
    }
    if (state?.step === STATES.WAIT_BUY) {
      await addShoppingItem(env, chatId, from, rawText, 'buy', inputType);
      await clearState(env, chatId);
      return;
    }
    if (state?.step === STATES.WAIT_SEARCH) {
      await clearState(env, chatId);
      return await runSearch(env, chatId, from, rawText, inputType);
    }
    if (state?.step === STATES.WAIT_TASK) {
      await clearState(env, chatId);
      return await createTaskFromText(env, chatId, from, rawText, inputType);
    }

    if (/^(занотуй|нотатка)[:\s]/i.test(rawText)) return await saveNoteFlow(env, chatId, from, rawText.replace(/^(занотуй|нотатка)[:\s]*/i, ''), inputType);
    if (/^(витрата|дохід)\b/i.test(rawText)) return await addBudgetFromText(env, chatId, from, rawText, inputType);
    if (/^(купи|купити|додай у продукти|додай в продукти|продукти)\b/i.test(rawText)) {
      const isProducts = /(продукти|молоко|хліб|сир|яйц|яблук|банан|м\'яс|овоч|фрукт|круп|вода)/i.test(rawText);
      return await addShoppingItem(env, chatId, from, rawText.replace(/^(додай у продукти|додай в продукти|продукти|купи|купити)\s*/i, ''), isProducts ? 'products' : 'buy', inputType);
    }
    if (/^(знайди|пошук|хто такий|що таке)\b/i.test(rawText)) return await runSearch(env, chatId, from, rawText, inputType);
    if (/\b(біля|поруч|near|район|позняк|осокорк|центр|київ)\b/i.test(rawText)) {
      return await runNearbyPlaces(env, chatId, rawText, {areaText: extractArea(rawText)});
    }

    return await createTaskFromText(env, chatId, from, rawText, inputType);
  } catch (error) {
    console.error('handleTextInput', error?.stack || error?.message || String(error));
    const msg = String(error?.message || error);
    const userMsg = msg.includes('CALENDAR_') ? TEXT.CALENDAR_FAIL : msg.includes('SHEETS_') ? TEXT.SHEETS_FAIL : msg.includes('SEARCH_') ? TEXT.SEARCH_FAIL : TEXT.GENERIC_FAIL;
    await sendMessage(env, chatId, userMsg, {reply_markup:mainMenuKeyboard()});
    await logInteraction(env, {chatId, user:from, inputType, originalText:rawText, normalizedText:rawText, action:'error', resultSummary:truncate(msg,200), status:'error'});
  }
}

async function saveNoteFlow(env, chatId, from, text, inputType) {
  await saveNote(env, chatId, from, text, inputType);
  await replyAndLog(env, chatId, from, inputType, text, 'add_note', 'Готово, занотував.', 'ok');
}

async function createTaskFromText(env, chatId, from, rawText, inputType='text') {
  const parsed = parseTaskInput(rawText, env.TIMEZONE || DEFAULT_TIMEZONE);
  const now = nowLocalString(env);
  let calendarEventId = '';
  if (parsed.due_at || parsed.due_date) {
    try {
      const event = await calendarInsertEvent(env, buildCalendarEvent(parsed.title, parsed.details, parsed.due_at, parsed.due_date, from, env));
      calendarEventId = event.id || '';
    } catch (error) {
      console.error('calendar insert', error?.message || String(error));
      throw new Error(`CALENDAR_WRITE_FAILED:${error?.message || error}`);
    }
  }
  try {
    await ensureWorksheet(env, SHEETS.TASKS, TASK_HEADERS);
    await sheetsAppend(env, `${SHEETS.TASKS}!A:M`, [[generateId('TASK'), now, now, String(chatId), String(from?.id || ''), usernameLabel(from), parsed.title, parsed.details, parsed.due_at, parsed.due_date, TASK_STATUS.NEW, calendarEventId, rawText]]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED:${error?.message || error}`);
  }
  const parts = ['Готово, задачу збережено.', '', `<b>${escapeHtml(parsed.title)}</b>`];
  if (parsed.due_at) parts.push(`🕒 ${escapeHtml(formatDateTime(parsed.due_at, env.TIMEZONE || DEFAULT_TIMEZONE))}`);
  else if (parsed.due_date) parts.push(`📅 ${escapeHtml(parsed.due_date)}`);
  if (calendarEventId) parts.push('📌 Також додано в Google Calendar.');
  await replyAndLog(env, chatId, from, inputType, rawText, 'create_task', parts.join('\n'), 'ok', calendarEventId);
}

function parseTaskInput(text, timezone) {
  const clean = text.trim();
  let remaining = clean;
  const now = new Date();
  const dateObj = extractDateFromText(remaining, timezone, now);
  remaining = dateObj.text;
  const timeObj = extractTimeFromText(remaining);
  remaining = timeObj.text;
  let title = remaining
    .replace(/^(створи задачу|додай задачу|задача|нагадай( мені)?|нагадування)[:\s]*/i, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!title) title = clean;

  let due_at = '';
  let due_date = '';
  if (dateObj.date) {
    if (timeObj.time) due_at = buildIso(dateObj.date, timeObj.time, timezone);
    else due_date = dateObj.date;
  } else if (timeObj.time) {
    const today = localDateKey(now, timezone);
    due_at = buildIso(today, timeObj.time, timezone);
  }
  return {title, details:'', due_at, due_date};
}

function extractDateFromText(text, timezone, now) {
  let t = text;
  const today = localDateKey(now, timezone);
  const tomorrow = addDaysToIsoDate(today, 1);
  if (/сьогодні/i.test(t)) return {date: today, text: t.replace(/сьогодні/ig, '').trim()};
  if (/завтра/i.test(t)) return {date: tomorrow, text: t.replace(/завтра/ig, '').trim()};
  const dmY = t.match(/(\d{1,2})[.\/](\d{1,2})(?:[.\/](\d{4}))?/);
  if (dmY) {
    const day = dmY[1].padStart(2,'0');
    const month = dmY[2].padStart(2,'0');
    const year = dmY[3] || String(now.getFullYear());
    return {date: `${year}-${month}-${day}`, text: t.replace(dmY[0], '').trim()};
  }
  const ymd = t.match(/(20\d{2})-(\d{2})-(\d{2})/);
  if (ymd) return {date: ymd[0], text: t.replace(ymd[0], '').trim()};
  return {date:'', text:t};
}

function extractTimeFromText(text) {
  const m = text.match(/(?:о|на)?\s*(\d{1,2})[:.](\d{2})/i);
  if (!m) return {time:'', text};
  const hh = String(Math.min(23, Number(m[1]))).padStart(2,'0');
  const mm = String(Math.min(59, Number(m[2]))).padStart(2,'0');
  return {time:`${hh}:${mm}:00`, text:text.replace(m[0], '').trim()};
}

function buildIso(date, time, timezone) {
  const offset = timezone === 'Europe/Kyiv' ? '+03:00' : 'Z';
  return `${date}T${time}${offset}`;
}

async function addBudgetFromText(env, chatId, from, rawText, inputType='text') {
  const parsed = parseBudgetInput(rawText);
  const now = nowLocalString(env);
  try {
    await ensureWorksheet(env, SHEETS.BUDGET, BUDGET_HEADERS);
    await sheetsAppend(env, `${SHEETS.BUDGET}!A:L`, [[generateId('BUD'), now, now, String(chatId), String(from?.id || ''), usernameLabel(from), parsed.type, parsed.amount, parsed.currency, parsed.category, parsed.note, rawText]]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED:${error?.message || error}`);
  }
  await replyAndLog(env, chatId, from, inputType, rawText, 'add_budget', `Готово, записав у бюджет.\n\n<b>${escapeHtml(parsed.type)}</b>: ${escapeHtml(String(parsed.amount))} ${escapeHtml(parsed.currency)}\n${escapeHtml(parsed.note)}`, 'ok');
}

function parseBudgetInput(text) {
  const amountMatch = text.match(/(\d+(?:[.,]\d+)?)/);
  const amount = Number((amountMatch?.[1] || '0').replace(',', '.'));
  const isIncome = /дохід|отримав|зароб/i.test(text);
  const currency = /\$|usd/i.test(text) ? 'USD' : /eur|€/i.test(text) ? 'EUR' : 'грн';
  const note = text.replace(/^(витрата|дохід)\s*/i, '').trim();
  return {type: isIncome ? 'Дохід' : 'Витрата', amount, currency, category: 'Інше', note};
}

async function addShoppingItem(env, chatId, from, rawText, listType='buy', inputType='text') {
  const parsed = parseShoppingInput(rawText);
  const now = nowLocalString(env);
  try {
    await ensureWorksheet(env, SHEETS.SHOPPING, SHOPPING_HEADERS);
    await sheetsAppend(env, `${SHEETS.SHOPPING}!A:L`, [[generateId('BUY'), now, now, String(chatId), String(from?.id || ''), usernameLabel(from), listType === 'products' ? 'Продукти' : 'Покупки', parsed.name, parsed.details, parsed.quantity, SHOPPING_STATUS.OPEN, rawText]]);
  } catch (error) {
    throw new Error(`SHEETS_WRITE_FAILED:${error?.message || error}`);
  }
  await replyAndLog(env, chatId, from, inputType, rawText, 'add_shopping', `Готово, додав у список.\n\n<b>${escapeHtml(parsed.name)}</b>${parsed.quantity ? `\nКількість: ${escapeHtml(parsed.quantity)}` : ''}${parsed.details ? `\n${escapeHtml(parsed.details)}` : ''}`, 'ok');
}

function parseShoppingInput(text) {
  const qty = text.match(/(\d+\s*(?:шт|кг|л|уп|пач|г)?)/i)?.[1] || '';
  let name = text.replace(/^(додай у продукти|додай в продукти|продукти|купи|купити)\s*/i, '').trim();
  if (qty) name = name.replace(qty, '').trim();
  return {name: name || text.trim(), quantity: qty.trim(), details:''};
}

async function saveNote(env, chatId, user, text, source) {
  await ensureWorksheet(env, SHEETS.NOTES, NOTE_HEADERS);
  await sheetsAppend(env, `${SHEETS.NOTES}!A:G`, [[generateId('NOTE'), nowLocalString(env), String(chatId), String(user?.id || ''), usernameLabel(user), text.trim(), source]]);
}

async function sendTaskList(env, chatId) {
  const tasks = await getTasks(env, chatId);
  const active = tasks.filter(t => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED);
  const lines = ['<b>Список задач</b>', ''];
  if (!active.length) lines.push('Поки активних задач немає.');
  for (const task of active.slice(0,20)) lines.push(formatTaskLine(task, env));
  await sendLongMessage(env, chatId, lines.join('\n'), {reply_markup:mainMenuKeyboard()});
}

async function prepareCompleteToday(env, chatId) {
  const tasks = (await getTasks(env, chatId)).filter(t => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED && isTaskToday(t, env));
  if (!tasks.length) return await sendMessage(env, chatId, 'На сьогодні активних задач немає.', {reply_markup:mainMenuKeyboard()});
  const selectionMap = {};
  const lines = ['<b>Що виконати сьогодні:</b>'];
  tasks.slice(0,15).forEach((task, idx) => { const n = String(idx+1); selectionMap[n] = task.task_id; lines.push(`\n${n}. ${formatTaskLine(task, env)}`); });
  lines.push('\nВідправ номер задачі, яку треба позначити як виконану.');
  await saveState(env, chatId, {step:STATES.WAIT_COMPLETE_TODAY, selectionMap});
  await sendLongMessage(env, chatId, lines.join('\n'), {reply_markup:mainMenuKeyboard()});
}

async function completeTaskFromSelection(env, chatId, rawText) {
  const state = await getState(env, chatId);
  const taskId = state?.selectionMap?.[rawText.trim()];
  if (!taskId) return await sendMessage(env, chatId, 'Не знайшов такої задачі.', {reply_markup:mainMenuKeyboard()});
  const ok = await updateTaskStatus(env, taskId, TASK_STATUS.DONE);
  await sendMessage(env, chatId, ok ? 'Позначив задачу як виконану.' : 'Не зміг оновити статус.', {reply_markup:mainMenuKeyboard()});
}

async function sendTodayOverview(env, chatId) {
  const tasks = (await getTasks(env, chatId)).filter(t => t.status !== TASK_STATUS.DONE && t.status !== TASK_STATUS.CANCELED && isTaskToday(t, env));
  let events = [];
  try {
    const range = resolveListRange('today', env);
    events = await listCalendarEvents(env, range.timeMin, range.timeMax, 15);
  } catch (error) {
    console.error('calendar list', error?.message || String(error));
  }
  const lines = ['<b>Сьогодні</b>'];
  if (!tasks.length && !events.length) lines.push('\nПоки нічого не заплановано.');
  if (tasks.length) {
    lines.push('\n<b>Задачі:</b>');
    for (const t of tasks) lines.push(formatTaskLine(t, env));
  }
  if (events.length) {
    lines.push('\n<b>Календар:</b>');
    for (const e of events) {
      const when = e.start?.dateTime ? formatDateTime(e.start.dateTime, env.TIMEZONE || DEFAULT_TIMEZONE) : `${e.start?.date || '-'} (весь день)`;
      lines.push(`• <b>${escapeHtml(e.summary || 'Без назви')}</b> — ${escapeHtml(when)}`);
    }
  }
  await sendLongMessage(env, chatId, lines.join('\n'), {reply_markup:mainMenuKeyboard()});
}

async function sendWeeklyReport(env, chatId) {
  const tasks = await getTasks(env, chatId);
  const budget = await getBudgetEntries(env, chatId);
  const shopping = await getShoppingItems(env, chatId);
  const weekAgo = Date.now() - 7*24*3600*1000;
  const created = tasks.filter(t => parseLocalDateGuess(t.created_at) >= weekAgo).length;
  const done = tasks.filter(t => t.status === TASK_STATUS.DONE && parseLocalDateGuess(t.updated_at || t.created_at) >= weekAgo).length;
  const expenses = budget.filter(b => b.type === 'Витрата' && parseLocalDateGuess(b.created_at) >= weekAgo).reduce((s,b)=>s+Number(b.amount||0),0);
  const income = budget.filter(b => b.type === 'Дохід' && parseLocalDateGuess(b.created_at) >= weekAgo).reduce((s,b)=>s+Number(b.amount||0),0);
  const openProducts = shopping.filter(i => i.status !== SHOPPING_STATUS.DONE && i.list_type === 'Продукти').length;
  const openBuy = shopping.filter(i => i.status !== SHOPPING_STATUS.DONE && i.list_type === 'Покупки').length;
  const text = `<b>Звіт за тиждень</b>\n\n• Створено задач: <b>${created}</b>\n• Виконано: <b>${done}</b>\n\n• Витрати: <b>${formatMoney(expenses)}</b>\n• Доходи: <b>${formatMoney(income)}</b>\n\n• Продукти: <b>${openProducts}</b>\n• Покупки: <b>${openBuy}</b>`;
  await sendMessage(env, chatId, text, {reply_markup:mainMenuKeyboard()});
}

async function sendShoppingSummaryAndPrompt(env, chatId, kind) {
  const items = (await getShoppingItems(env, chatId)).filter(i => i.status !== SHOPPING_STATUS.DONE && i.list_type === (kind === 'products' ? 'Продукти' : 'Покупки')).slice(0, 10);
  const title = kind === 'products' ? 'Продукти' : 'Покупки';
  const lines = [`<b>${title}</b>`];
  if (!items.length) lines.push('\nПоки список порожній.');
  else items.forEach(i => lines.push(`• ${escapeHtml(i.item_name || '-')}${i.quantity ? ` — ${escapeHtml(i.quantity)}` : ''}`));
  lines.push(`\nНапиши, що додати.`);
  await sendLongMessage(env, chatId, lines.join('\n'), {reply_markup:mainMenuKeyboard()});
}

async function runSearch(env, chatId, from, query, inputType='text') {
  await sendMessage(env, chatId, 'Шукаю…', {reply_markup:mainMenuKeyboard()});
  try {
    const answer = await answerWithGoogleSearch(env, query);
    await replyAndLog(env, chatId, from, inputType, query, 'search', answer, 'ok');
  } catch (error) {
    console.error('search', error?.message || String(error));
    throw new Error(`SEARCH_FAILED:${error?.message || error}`);
  }
}

async function runNearbyPlaces(env, chatId, query, state={}) {
  const area = buildAreaLabel(state);
  const searchQuery = `${query} ${area}`.trim();
  const url = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(searchQuery)}`;
  const text = `<b>Місця поруч</b>\n\nЩо шукати: <b>${escapeHtml(query)}</b>${area ? `\nДе: <b>${escapeHtml(area)}</b>` : ''}\n\nGoogle Maps:\n${escapeHtml(url)}`;
  await sendMessage(env, chatId, text, {reply_markup:mainMenuKeyboard()});
}

function buildAreaLabel(state) {
  if (state?.areaText) return state.areaText;
  if (state?.location?.latitude && state?.location?.longitude) return `${state.location.latitude},${state.location.longitude}`;
  return '';
}

function extractArea(text) {
  const m = text.match(/(?:біля|поруч|в районі)\s+(.+)$/i);
  return m?.[1] || '';
}

async function transcribeTelegramAudio(env, audio) {
  const maxBytes = Number(env.MAX_VOICE_BYTES || 1800000);
  const fileInfo = await telegramApi(env, 'getFile', {file_id: audio.file_id});
  const path = fileInfo.result?.file_path;
  const fileSize = Number(fileInfo.result?.file_size || audio.file_size || 0);
  if (!path || fileSize > maxBytes) throw new Error('voice not available');
  const tgUrl = `https://api.telegram.org/file/bot${env.BOT_TOKEN}/${path}`;
  const response = await fetchWithTimeout(tgUrl, {}, Number(env.TELEGRAM_FILE_TIMEOUT_MS || 12000));
  if (!response.ok) throw new Error(`telegram audio ${response.status}`);
  const bytes = new Uint8Array(await response.arrayBuffer());
  const base64 = arrayBufferToBase64(bytes);
  return await geminiText(env, {
    contents:[{role:'user', parts:[
      {text:'Розшифруй це коротке голосове повідомлення українською. Поверни тільки текст без пояснень.'},
      {inline_data:{mime_type: detectTelegramAudioMime(audio, path), data: base64}}
    ]}],
    useSearch:false,
    timeoutMs: Number(env.GEMINI_TIMEOUT_MS || 18000)
  });
}

async function answerWithGoogleSearch(env, text) {
  return await geminiText(env, {
    contents:[{role:'user', parts:[{text:`Дай коротку практичну відповідь українською на запит користувача. Якщо доречно, використай Google Search. Запит: ${text}`}]}],
    useSearch:true,
    timeoutMs: Number(env.GEMINI_TIMEOUT_MS || 18000)
  });
}

function mainMenuKeyboard() {
  return {keyboard:[
    [{text:MENU.CREATE_TASK}, {text:MENU.LIST_TASKS}],
    [{text:MENU.COMPLETE_TODAY}, {text:MENU.TODAY}],
    [{text:MENU.BUDGET}, {text:MENU.PRODUCTS}],
    [{text:MENU.NOTE}, {text:MENU.BUY}],
    [{text:MENU.SEARCH}, {text:MENU.NEARBY}],
    [{text:MENU.WEEK_REPORT}]
  ], resize_keyboard:true};
}
function nearbyLocationKeyboard() {
  return {keyboard:[[{text:'Надіслати геолокацію', request_location:true}], [{text:MENU.NEARBY}]], resize_keyboard:true, one_time_keyboard:true};
}
function sameBtn(a,b){ return normalize(a) === normalize(b); }
function normalize(s=''){ return s.toLowerCase().replace(/\s+/g,' ').trim(); }

async function sendLongMessage(env, chatId, text, extra={}) { for (const c of splitMessage(text, MAX_TELEGRAM_MESSAGE)) await sendMessage(env, chatId, c, extra); }
function splitMessage(text, maxLen){ if(text.length<=maxLen) return [text]; const out=[]; let cur=''; for(const line of text.split('\n')){ if((cur+'\n'+line).length>maxLen && cur){ out.push(cur); cur=line; } else cur = cur ? `${cur}\n${line}` : line; } if(cur) out.push(cur); return out; }
async function sendMessage(env, chatId, text, extra={}) { return await telegramApi(env, 'sendMessage', {chat_id:chatId, text, parse_mode:'HTML', disable_web_page_preview:false, ...extra}); }
async function telegramApi(env, method, payload) {
  const response = await fetch(`https://api.telegram.org/bot${env.BOT_TOKEN}/${method}`, {method:'POST', headers:{'content-type':'application/json'}, body:JSON.stringify(payload)});
  const result = await response.json();
  if (!response.ok || !result.ok) throw new Error(`Telegram API ${method}: ${JSON.stringify(result)}`);
  return result;
}

async function getTasks(env, chatId){ await ensureWorksheet(env, SHEETS.TASKS, TASK_HEADERS); return mapRows(await sheetsGetValues(env, `${SHEETS.TASKS}!A:M`)).filter(r=>String(r.chat_id||'')===String(chatId)); }
async function getBudgetEntries(env, chatId){ await ensureWorksheet(env, SHEETS.BUDGET, BUDGET_HEADERS); return mapRows(await sheetsGetValues(env, `${SHEETS.BUDGET}!A:L`)).filter(r=>String(r.chat_id||'')===String(chatId)); }
async function getShoppingItems(env, chatId){ await ensureWorksheet(env, SHEETS.SHOPPING, SHOPPING_HEADERS); return mapRows(await sheetsGetValues(env, `${SHEETS.SHOPPING}!A:L`)).filter(r=>String(r.chat_id||'')===String(chatId)); }
async function updateTaskStatus(env, taskId, newStatus) {
  const values = await sheetsGetValues(env, `${SHEETS.TASKS}!A:M`);
  if (values.length <= 1) return false;
  const headers = values[0]; const idIndex = headers.indexOf('task_id'); const statusIndex = headers.indexOf('status'); const updatedIndex = headers.indexOf('updated_at');
  for (let i=1; i<values.length; i++) if ((values[i][idIndex]||'') === taskId) { const row = i+1; await sheetsUpdateValues(env, `${SHEETS.TASKS}!${indexToColumn(statusIndex)}${row}`, [[newStatus]]); await sheetsUpdateValues(env, `${SHEETS.TASKS}!${indexToColumn(updatedIndex)}${row}`, [[nowLocalString(env)]]); return true; }
  return false;
}
function mapRows(values){ if(!values.length) return []; const headers=values[0]; return values.slice(1).map(row=>{ const o={}; headers.forEach((h,i)=>o[h]=row[i]??''); return o; }); }
function formatTaskLine(task, env){ const when = task.due_at ? formatDateTime(task.due_at, env.TIMEZONE||DEFAULT_TIMEZONE) : (task.due_date || 'без дати'); return `• <b>${escapeHtml(task.title || 'Без назви')}</b> — ${escapeHtml(when)} <i>(${escapeHtml(task.status || TASK_STATUS.NEW)})</i>`; }
function isTaskToday(task, env){ const today=localDateKey(new Date(), env.TIMEZONE||DEFAULT_TIMEZONE); if(task.due_date) return task.due_date===today; if(task.due_at) return localDateKey(new Date(task.due_at), env.TIMEZONE||DEFAULT_TIMEZONE)===today; return false; }
function resolveListRange(range, env){ const tz=env.TIMEZONE||DEFAULT_TIMEZONE; const now=new Date(); const timeMin=startOfDayInTimezone(now, tz, 0); const timeMax=startOfDayInTimezone(now, tz, range==='week' ? 7 : 1); return {timeMin, timeMax}; }

async function replyAndLog(env, chatId, from, inputType, originalText, action, replyText, status='ok', calendarEventId='') {
  await logInteraction(env, {chatId, user:from, inputType, originalText, normalizedText:originalText, action, resultSummary:truncate(stripHtml(replyText),220), calendarEventId, status});
  await sendLongMessage(env, chatId, replyText, {reply_markup:mainMenuKeyboard()});
}
async function logInteraction(env, entry) {
  try {
    await ensureWorksheet(env, SHEETS.LOG, LOG_HEADERS);
    await sheetsAppend(env, `${SHEETS.LOG}!A:K`, [[nowLocalString(env), String(entry.chatId||''), String(entry.user?.id||''), usernameLabel(entry.user), entry.inputType||'', entry.originalText||'', entry.normalizedText||'', entry.action||'', entry.resultSummary||'', entry.calendarEventId||'', entry.status||'']]);
  } catch (error) { console.error('log', error?.message || String(error)); }
}

async function getState(env, chatId){ const raw = await env.STATE_KV.get(`state:${chatId}`); return raw ? JSON.parse(raw) : null; }
async function saveState(env, chatId, state){ await env.STATE_KV.put(`state:${chatId}`, JSON.stringify(state), {expirationTtl: 86400}); }
async function clearState(env, chatId){ await env.STATE_KV.delete(`state:${chatId}`); }

async function ensureWorksheet(env, sheetTitle, headers) {
  const metadata = await sheetsGetMetadata(env);
  let sheet = metadata.sheets?.find(s => s.properties?.title === sheetTitle);
  if (!sheet) {
    await sheetsBatchUpdate(env, {requests:[{addSheet:{properties:{title:sheetTitle}}}]});
    sheet = true;
  }
  const firstRow = await sheetsGetValues(env, `${sheetTitle}!A1:Z1`);
  const existing = firstRow[0] || [];
  if (!existing.length) {
    await sheetsUpdateValues(env, `${sheetTitle}!A1:${indexToColumn(headers.length-1)}1`, [headers]);
    return;
  }
  let changed = false;
  const merged = [...existing];
  for (const header of headers) if (!merged.includes(header)) { merged.push(header); changed = true; }
  if (changed) await sheetsUpdateValues(env, `${sheetTitle}!A1:${indexToColumn(merged.length-1)}1`, [merged]);
}

async function sheetsGetMetadata(env){ return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}`); }
async function sheetsBatchUpdate(env, body){ return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}:batchUpdate`, {method:'POST', body}); }
async function sheetsGetValues(env, range){ const data = await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}`); return data.values || []; }
async function sheetsUpdateValues(env, range, values){ return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=USER_ENTERED`, {method:'PUT', body:{range, majorDimension:'ROWS', values}}); }
async function sheetsAppend(env, range, values){ return await googleApi(env, `https://sheets.googleapis.com/v4/spreadsheets/${env.GOOGLE_SHEET_ID}/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED`, {method:'POST', body:{range, majorDimension:'ROWS', values}}); }
async function calendarInsertEvent(env, event){ return await googleApi(env, `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(env.CALENDAR_ID || DEFAULT_CALENDAR_ID)}/events`, {method:'POST', body:event}); }
async function listCalendarEvents(env, timeMin, timeMax, maxResults=10){ const qs = new URLSearchParams({timeMin, timeMax, singleEvents:'true', orderBy:'startTime', maxResults:String(maxResults)}); const data = await googleApi(env, `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(env.CALENDAR_ID || DEFAULT_CALENDAR_ID)}/events?${qs.toString()}`); return data.items || []; }
let cachedGoogleToken = null;
async function googleApi(env, url, options={}) {
  const token = await getGoogleAccessToken(env);
  const response = await fetch(url, {method:options.method || 'GET', headers:{authorization:`Bearer ${token}`, 'content-type':'application/json'}, body:options.body ? JSON.stringify(options.body) : undefined});
  const text = await response.text();
  let data = {}; try { data = text ? JSON.parse(text) : {}; } catch { data = {raw:text}; }
  if (!response.ok) throw new Error(`Google API ${response.status}: ${JSON.stringify(data)}`);
  return data;
}
async function getGoogleAccessToken(env) {
  if (cachedGoogleToken && cachedGoogleToken.expiresAt > Date.now() + 60000) return cachedGoogleToken.token;
  const sa = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now()/1000);
  const header = {alg:'RS256', typ:'JWT'};
  const claim = {iss:sa.client_email, scope:'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/calendar', aud:sa.token_uri, exp:now+3600, iat:now};
  const unsigned = `${base64UrlEncode(JSON.stringify(header))}.${base64UrlEncode(JSON.stringify(claim))}`;
  const signature = await signJwt(unsigned, sa.private_key);
  const response = await fetch(sa.token_uri, {method:'POST', headers:{'content-type':'application/x-www-form-urlencoded'}, body:new URLSearchParams({grant_type:'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion:`${unsigned}.${signature}`})});
  const tokenData = await response.json();
  if (!response.ok) throw new Error(`Google OAuth: ${JSON.stringify(tokenData)}`);
  cachedGoogleToken = {token:tokenData.access_token, expiresAt:Date.now() + (tokenData.expires_in || 3600)*1000};
  return cachedGoogleToken.token;
}

async function geminiText(env, {contents, useSearch=false, timeoutMs=18000}) {
  const body = {contents, generationConfig:{temperature:0.2, responseMimeType:'text/plain'}};
  if (useSearch) body.tools = [{google_search:{}}];
  const response = await fetchWithTimeout(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(env.GEMINI_MODEL || DEFAULT_MODEL)}:generateContent`, {method:'POST', headers:{'content-type':'application/json', 'x-goog-api-key':env.GEMINI_API_KEY}, body:JSON.stringify(body)}, timeoutMs);
  const data = await response.json();
  if (!response.ok) throw new Error(`Gemini: ${JSON.stringify(data)}`);
  const text = (data.candidates?.[0]?.content?.parts || []).map(p => p.text || '').join('').trim();
  if (!text) throw new Error(`Gemini empty: ${JSON.stringify(data)}`);
  return text;
}

async function fetchWithTimeout(url, options={}, timeoutMs=15000) { const controller = new AbortController(); const timer = setTimeout(()=>controller.abort('timeout'), timeoutMs); try { return await fetch(url, {...options, signal:controller.signal}); } finally { clearTimeout(timer); } }
function detectTelegramAudioMime(audio, path=''){ const raw=String(audio?.mime_type||'').toLowerCase(); const lower=String(path||'').toLowerCase(); if(raw) return raw; if(lower.endsWith('.oga')||lower.endsWith('.ogg')) return 'audio/ogg'; if(lower.endsWith('.mp3')) return 'audio/mpeg'; if(lower.endsWith('.m4a')) return 'audio/mp4'; return 'audio/ogg'; }
async function signJwt(unsignedToken, pemPrivateKey) { const keyData = pemToArrayBuffer(pemPrivateKey); const cryptoKey = await crypto.subtle.importKey('pkcs8', keyData, {name:'RSASSA-PKCS1-v1_5', hash:'SHA-256'}, false, ['sign']); const signature = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', cryptoKey, new TextEncoder().encode(unsignedToken)); return arrayBufferToBase64Url(signature); }
function pemToArrayBuffer(pem){ const base64 = pem.replace('-----BEGIN PRIVATE KEY-----','').replace('-----END PRIVATE KEY-----','').replace(/\s+/g,''); const binary = atob(base64); const bytes = new Uint8Array(binary.length); for(let i=0;i<binary.length;i++) bytes[i] = binary.charCodeAt(i); return bytes.buffer; }
function base64UrlEncode(value){ return arrayBufferToBase64Url(new TextEncoder().encode(value)); }
function arrayBufferToBase64Url(buffer){ const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer); let binary=''; for(const b of bytes) binary += String.fromCharCode(b); return btoa(binary).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,''); }
function arrayBufferToBase64(buffer){ const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer); let binary=''; for(const b of bytes) binary += String.fromCharCode(b); return btoa(binary); }
function indexToColumn(index){ let col=''; let num=index+1; while(num>0){ const rem=(num-1)%26; col=String.fromCharCode(65+rem)+col; num=Math.floor((num-1)/26);} return col; }
function generateId(prefix){ return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2,7).toUpperCase()}`; }
function usernameLabel(user){ return user?.username ? `@${user.username}` : ((user?.first_name || '') + (user?.last_name ? ` ${user.last_name}` : '')).trim(); }
function nowLocalString(env){ return new Intl.DateTimeFormat('uk-UA', {timeZone:env.TIMEZONE||DEFAULT_TIMEZONE, dateStyle:'short', timeStyle:'medium'}).format(new Date()); }
function formatDateTime(value, timezone){ return new Intl.DateTimeFormat('uk-UA', {timeZone:timezone, dateStyle:'short', timeStyle:'short'}).format(new Date(value)); }
function formatMoney(value){ return `${Number(value||0).toFixed(2)} грн`; }
function localDateKey(date, timezone){ const parts = new Intl.DateTimeFormat('en-CA', {timeZone:timezone, year:'numeric', month:'2-digit', day:'2-digit'}).formatToParts(date); const y=parts.find(p=>p.type==='year')?.value; const m=parts.find(p=>p.type==='month')?.value; const d=parts.find(p=>p.type==='day')?.value; return `${y}-${m}-${d}`; }
function startOfDayInTimezone(date, timezone, addDays=0){ const day = localDateKey(new Date(date.getTime() + addDays*86400000), timezone); const offset = timezone === 'Europe/Kyiv' ? '+03:00' : 'Z'; return `${day}T00:00:00${offset}`; }
function addDaysToIsoDate(date, days){ const d = new Date(`${date}T00:00:00Z`); d.setUTCDate(d.getUTCDate()+days); return d.toISOString().slice(0,10); }
function buildCalendarEvent(title, details, dueAt, dueDate, user, env){ const tz=env.TIMEZONE||DEFAULT_TIMEZONE; const desc=[details, user?.username ? `Telegram: @${user.username}` : '', user?.id ? `Telegram ID: ${user.id}` : ''].filter(Boolean).join('\n'); const event={summary:title, description:desc, reminders:{useDefault:false, overrides:[{method:'popup', minutes:30}]}}; if(dueAt){ event.start={dateTime:dueAt, timeZone:tz}; event.end={dateTime:addHourKeepingOffset(dueAt,1), timeZone:tz}; } else { const date=dueDate || localDateKey(new Date(), tz); event.start={date}; event.end={date:addDaysToIsoDate(date,1)}; } return event; }
function addHourKeepingOffset(iso, hours){ const d = new Date(iso); const next = new Date(d.getTime()+hours*3600000); const offset = iso.match(/([+-]\d{2}:\d{2}|Z)$/)?.[1] || 'Z'; if (offset === 'Z') return next.toISOString().replace('.000Z','Z'); const ms = Date.parse(next.toISOString()); const sign = offset[0] === '-' ? -1 : 1; const hh = Number(offset.slice(1,3)); const mm = Number(offset.slice(4,6)); const shifted = new Date(ms + sign * (hh*60 + mm) * 60000).toISOString().replace('Z',''); return shifted + offset; }
function parseLocalDateGuess(value){ if(!value) return 0; const normalized = value.replace(/\u00A0/g,' '); const m = normalized.match(/(\d{2})[.](\d{2})[.](\d{2,4}),?\s*(\d{2}):(\d{2})(?::(\d{2}))?/); if(m){ const year = m[3].length===2 ? `20${m[3]}` : m[3]; return new Date(`${year}-${m[2]}-${m[1]}T${m[4]}:${m[5]}:${m[6]||'00'}+03:00`).getTime(); } const ts = Date.parse(value); return Number.isNaN(ts) ? 0 : ts; }
function stripHtml(text){ return String(text||'').replace(/<[^>]*>/g,''); }
function truncate(text,n){ const s=String(text||''); return s.length>n ? `${s.slice(0,n-1)}…` : s; }
function escapeHtml(value){ return String(value ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;'); }
function jsonResponse(data,status=200){ return new Response(JSON.stringify(data,null,2), {status, headers:{'content-type':'application/json; charset=utf-8'}}); }
function validateEnv(env){ const required = ['BOT_TOKEN','GOOGLE_SHEET_ID','GOOGLE_SERVICE_ACCOUNT_JSON','STATE_KV','SETUP_KEY','TELEGRAM_WEBHOOK_SECRET']; for (const k of required) if (!env[k]) throw new Error(`Missing env ${k}`); }

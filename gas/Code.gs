const SHEETS = {
  CONFIG: 'config',
  PARTICIPANTS: 'participants',
  EVENTS: 'events',
  ENTRIES: 'entries',
  PREDICTIONS: 'predictions',
  RESULTS: 'results',
};

const CONFIG_KEYS = {
  DEADLINE_ISO: 'predictionDeadlineIso',
  REVEAL_PREDICTIONS: 'revealPredictions',
  PUBLISH_SCOREBOARD: 'publishScoreboard',
  PUBLISH_CROWD_FORECAST: 'publishCrowdForecast',
  TIMEZONE: 'timezone',
  APP_TITLE: 'appTitle',
};

const DEFAULT_CONFIG = {
  [CONFIG_KEYS.DEADLINE_ISO]: '',
  [CONFIG_KEYS.REVEAL_PREDICTIONS]: 'false',
  [CONFIG_KEYS.PUBLISH_SCOREBOARD]: 'false',
  [CONFIG_KEYS.PUBLISH_CROWD_FORECAST]: 'false',
  [CONFIG_KEYS.TIMEZONE]: 'Asia/Tokyo',
  [CONFIG_KEYS.APP_TITLE]: '日本選手権 予想アプリ',
};

const HEADERS = {
  [SHEETS.CONFIG]: ['key', 'value'],
  [SHEETS.PARTICIPANTS]: ['participantId', 'displayName', 'accessKey', 'isActive', 'createdAt'],
  [SHEETS.EVENTS]: ['eventId', 'sortOrder', 'gender', 'eventName'],
  [SHEETS.ENTRIES]: ['entryId', 'eventId', 'seedOrder', 'athleteName', 'team', 'entryTime'],
  [SHEETS.PREDICTIONS]: ['participantId', 'eventId', 'pick1EntryId', 'pick2EntryId', 'pick3EntryId', 'pick4EntryId', 'updatedAt'],
  [SHEETS.RESULTS]: ['eventId', 'firstEntryId', 'secondEntryId', 'thirdEntryId', 'fourthEntryId', 'updatedAt'],
};

const PREDICTION_FIELDS = ['pick1EntryId', 'pick2EntryId', 'pick3EntryId', 'pick4EntryId'];
const RESULT_FIELDS = ['firstEntryId', 'secondEntryId', 'thirdEntryId', 'fourthEntryId'];
const FOUR_PLACE_EVENT_NAMES = { '100m 自由形': true, '200m 自由形': true };

// Standalone Apps Script として使う場合のみ設定してください（通常は空でOK）。
const STANDALONE_SPREADSHEET_ID = '';

function doGet(e) {
  const page = e && e.parameter ? String(e.parameter.page || '').toLowerCase() : '';
  if (page === 'bridge') {
    logApiDebug_('doGet.bridge', {
      hasParams: !!(e && e.parameter),
      ua: e && e.parameter ? String(e.parameter._ua || '') : '',
    });
    return HtmlService.createTemplateFromFile('Bridge')
      .evaluate()
      .setTitle('RankMaker Bridge')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.isAdminPage = page === 'admin';
  let title = DEFAULT_CONFIG[CONFIG_KEYS.APP_TITLE];
  try {
    title = getAppConfig_()[CONFIG_KEYS.APP_TITLE] || title;
  } catch (err) {
    // 初期セットアップ前 / スプレッドシート未接続でも画面は開けるようにする。
  }
  return template
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupApp() {
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  return { ok: true };
}

function getParticipantOptions() {
  logApiDebug_('getParticipantOptions');
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  return {
    participants: getParticipants_().map(function (p) {
      return { participantId: p.participantId, displayName: p.displayName };
    }),
  };
}

function participantLogin(participantId) {
  logApiDebug_('participantLogin', { participantId: participantId });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const participant = getParticipantById_(participantId);
  if (!participant) throw new Error('参加者を選択してください。');
  return buildParticipantDashboard_(participant);
}

function saveParticipantPredictions(participantId, predictionRows) {
  logApiDebug_('saveParticipantPredictions', {
    participantId: participantId,
    rows: Array.isArray(predictionRows) ? predictionRows.length : 0,
  });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const participant = getParticipantById_(participantId);
  if (!participant) throw new Error('参加者を選択してください。');
  const config = getAppConfig_();
  assertWithinDeadline_(config);

  const master = getMasterData_();
  const validEventMap = toMap_(master.events, 'eventId');
  const validEntryMap = toMap_(master.entries, 'entryId');

  const eventsToDelete = [];
  const sanitized = (predictionRows || []).map(function (row) {
    if (!row || !row.eventId) throw new Error('eventId が不足しています。');
    if (!validEventMap[row.eventId]) throw new Error('不正な種目です: ' + row.eventId);
    const placementCount = getPlacementCountForEvent_(validEventMap[row.eventId]);
    const picks = PREDICTION_FIELDS.slice(0, placementCount).map(function (field) { return row[field]; });
    const filtered = picks.filter(Boolean);
    if (filtered.length === 0) {
      eventsToDelete.push(String(row.eventId));
      return null;
    }
    if (filtered.length !== placementCount) {
      throw new Error('入力する種目は1〜' + placementCount + '位を全て選択してください。');
    }
    if (new Set(filtered).size !== placementCount) {
      const counts = {};
      filtered.forEach(function (entryId) {
        counts[entryId] = (counts[entryId] || 0) + 1;
      });
      const duplicateNames = Object.keys(counts)
        .filter(function (entryId) { return counts[entryId] > 1; })
        .map(function (entryId) {
          const entry = validEntryMap[entryId];
          return entry ? entry.athleteName : entryId;
        });
      const eventName = validEventMap[row.eventId] ? validEventMap[row.eventId].eventName : row.eventId;
      throw new Error('「' + eventName + '」で重複選択: ' + duplicateNames.join('、'));
    }
    filtered.forEach(function (entryId) {
      if (!validEntryMap[entryId]) throw new Error('不正な選手です: ' + entryId);
      if (validEntryMap[entryId].eventId !== row.eventId) {
        throw new Error('種目外の選手が選択されています。');
      }
    });
    return {
      participantId: participant.participantId,
      eventId: row.eventId,
      pick1EntryId: filtered[0],
      pick2EntryId: filtered[1],
      pick3EntryId: filtered[2],
      pick4EntryId: filtered[3] || '',
      updatedAt: new Date().toISOString(),
    };
  }).filter(function (r) { return !!r; });

  upsertPredictions_(participant.participantId, sanitized, eventsToDelete);
  return buildParticipantDashboard_(participant);
}

function getVisiblePredictions(participantId) {
  logApiDebug_('getVisiblePredictions', { participantId: participantId });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  if (!getParticipantById_(participantId)) throw new Error('参加者を選択してください。');
  const config = getAppConfig_();
  if (!isTrue_(config[CONFIG_KEYS.REVEAL_PREDICTIONS])) {
    throw new Error('管理人が公開するまで他の人の予想は閲覧できません。');
  }
  return { submissions: buildVisiblePredictionList_() };
}

function getParticipantScoreboard(participantId) {
  logApiDebug_('getParticipantScoreboard', { participantId: participantId });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  if (!getParticipantById_(participantId)) throw new Error('参加者を選択してください。');
  const config = getAppConfig_();
  if (!isTrue_(config[CONFIG_KEYS.PUBLISH_SCOREBOARD])) {
    throw new Error('管理人が公開するまで結果比較は閲覧できません。');
  }
  return { scoreboard: buildScoreboard_() };
}

function getParticipantCrowdForecast(participantId) {
  logApiDebug_('getParticipantCrowdForecast', { participantId: participantId });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  if (!getParticipantById_(participantId)) throw new Error('参加者を選択してください。');
  const config = getAppConfig_();
  if (!isTrue_(config[CONFIG_KEYS.PUBLISH_CROWD_FORECAST])) {
    throw new Error('管理人が公開するまで投票集計予想は閲覧できません。');
  }
  return { crowdForecast: buildCrowdForecast_() };
}

function adminGetDashboard() {
  logApiDebug_('adminGetDashboard');
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  return buildAdminDashboard_();
}

function adminUpdateConfig(patch) {
  logApiDebug_('adminUpdateConfig', { keys: patch ? Object.keys(patch) : [] });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const safePatch = {};
  if (patch && Object.prototype.hasOwnProperty.call(patch, CONFIG_KEYS.DEADLINE_ISO)) {
    safePatch[CONFIG_KEYS.DEADLINE_ISO] = String(patch[CONFIG_KEYS.DEADLINE_ISO] || '').trim();
  }
  if (patch && Object.prototype.hasOwnProperty.call(patch, CONFIG_KEYS.REVEAL_PREDICTIONS)) {
    safePatch[CONFIG_KEYS.REVEAL_PREDICTIONS] = String(!!patch[CONFIG_KEYS.REVEAL_PREDICTIONS]);
  }
  if (patch && Object.prototype.hasOwnProperty.call(patch, CONFIG_KEYS.PUBLISH_SCOREBOARD)) {
    safePatch[CONFIG_KEYS.PUBLISH_SCOREBOARD] = String(!!patch[CONFIG_KEYS.PUBLISH_SCOREBOARD]);
  }
  if (patch && Object.prototype.hasOwnProperty.call(patch, CONFIG_KEYS.PUBLISH_CROWD_FORECAST)) {
    safePatch[CONFIG_KEYS.PUBLISH_CROWD_FORECAST] = String(!!patch[CONFIG_KEYS.PUBLISH_CROWD_FORECAST]);
  }
  if (patch && Object.prototype.hasOwnProperty.call(patch, CONFIG_KEYS.APP_TITLE)) {
    safePatch[CONFIG_KEYS.APP_TITLE] = String(patch[CONFIG_KEYS.APP_TITLE] || DEFAULT_CONFIG[CONFIG_KEYS.APP_TITLE]).trim();
  }
  setConfigValues_(safePatch);
  return buildAdminDashboard_();
}

function adminSaveParticipants(participants) {
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  replaceSheetRows_(SHEETS.PARTICIPANTS, HEADERS[SHEETS.PARTICIPANTS], (participants || []).map(function (p, i) {
    if (!p.displayName) {
      throw new Error('participants の各行に displayName が必要です。 index=' + i);
    }
    const participantId = p.participantId ? String(p.participantId).trim() : normalizeParticipantName_(p.displayName);
    return [
      participantId,
      String(p.displayName).trim(),
      '',
      String(p.isActive !== false),
      p.createdAt ? String(p.createdAt) : new Date().toISOString(),
    ];
  }));
  return buildAdminDashboard_();
}

function adminAddParticipant(displayName) {
  logApiDebug_('adminAddParticipant', { displayName: displayName });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const name = String(displayName || '').trim();
  if (!name) throw new Error('表示名を入力してください。');
  const participantId = normalizeParticipantName_(name);
  if (!participantId) throw new Error('表示名を入力してください。');
  if (getParticipantById_(participantId) || getParticipants_().some(function (p) { return p.displayName === name; })) {
    throw new Error('同名の参加者が既に存在します。');
  }
  getSheet_(SHEETS.PARTICIPANTS).appendRow([participantId, name, '', 'true', new Date().toISOString()]);
  return buildAdminDashboard_();
}

function adminDeleteParticipant(participantId) {
  logApiDebug_('adminDeleteParticipant', { participantId: participantId });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const id = String(participantId || '').trim();
  if (!id) throw new Error('participantId が必要です。');

  const participantRows = readSheetObjects_(SHEETS.PARTICIPANTS).filter(function (r) {
    return String(r.participantId || '').trim() !== id;
  }).map(function (r) {
    return [
      String(r.participantId || '').trim(),
      String(r.displayName || '').trim(),
      String(r.accessKey || ''),
      String(r.isActive == null ? 'true' : r.isActive),
      String(r.createdAt || ''),
    ];
  });

  const predictionRows = readSheetObjects_(SHEETS.PREDICTIONS).filter(function (r) {
    return String(r.participantId || '').trim() !== id;
  }).map(function (r) {
    return [
      String(r.participantId || ''),
      String(r.eventId || ''),
      String(r.pick1EntryId || ''),
      String(r.pick2EntryId || ''),
      String(r.pick3EntryId || ''),
      String(r.pick4EntryId || ''),
      String(r.updatedAt || ''),
    ];
  });

  replaceSheetRows_(SHEETS.PARTICIPANTS, HEADERS[SHEETS.PARTICIPANTS], participantRows);
  replaceSheetRows_(SHEETS.PREDICTIONS, HEADERS[SHEETS.PREDICTIONS], predictionRows);
  return buildAdminDashboard_();
}

function adminImportMasterData(masterData) {
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  importMasterData_(masterData);
  return buildAdminDashboard_();
}

function adminSaveResults(resultRows) {
  logApiDebug_('adminSaveResults', { rows: Array.isArray(resultRows) ? resultRows.length : 0 });
  ensureSheets_();
  ensureOfficialMasterDataLoaded_();
  const master = getMasterData_();
  const eventMap = toMap_(master.events, 'eventId');
  const entryMap = toMap_(master.entries, 'entryId');

  const rows = (resultRows || []).map(function (r) {
    if (!r || !r.eventId) throw new Error('結果に eventId が必要です。');
    if (!eventMap[r.eventId]) throw new Error('不正な種目です: ' + r.eventId);
    const placementCount = getPlacementCountForEvent_(eventMap[r.eventId]);
    const picks = RESULT_FIELDS.slice(0, placementCount).map(function (field) { return r[field]; }).filter(Boolean);
    if (picks.length && picks.length !== placementCount) throw new Error('結果は1〜' + placementCount + '位を全て入力してください。');
    if (picks.length === placementCount && new Set(picks).size !== placementCount) throw new Error('結果で同一選手を重複できません。');
    picks.forEach(function (entryId) {
      if (!entryMap[entryId]) throw new Error('不正な選手です: ' + entryId);
      if (entryMap[entryId].eventId !== r.eventId) throw new Error('種目外の選手が結果に含まれています。');
    });
    return [
      r.eventId,
      r.firstEntryId || '',
      r.secondEntryId || '',
      r.thirdEntryId || '',
      r.fourthEntryId || '',
      new Date().toISOString(),
    ];
  });

  replaceSheetRows_(SHEETS.RESULTS, HEADERS[SHEETS.RESULTS], rows);
  return buildAdminDashboard_();
}

function logApiDebug_(name, detail) {
  try {
    console.log('[RankMaker]', name, JSON.stringify(detail || {}));
  } catch (err) {
    try {
      Logger.log('[RankMaker] ' + name + ' ' + JSON.stringify(detail || {}));
    } catch (ignored) {}
  }
}

function buildParticipantDashboard_(participant) {
  const config = getAppConfig_();
  return {
    participant: participant,
    config: publicConfig_(config),
    masterData: getMasterData_(),
    myPredictions: getPredictionsByParticipant_(participant.participantId),
    revealEnabled: isTrue_(config[CONFIG_KEYS.REVEAL_PREDICTIONS]),
    scoreboardEnabled: isTrue_(config[CONFIG_KEYS.PUBLISH_SCOREBOARD]),
    crowdForecastEnabled: isTrue_(config[CONFIG_KEYS.PUBLISH_CROWD_FORECAST]),
  };
}

function buildAdminDashboard_() {
  const config = getAppConfig_();
  return {
    config: config,
    participants: getParticipants_(),
    masterData: getMasterData_(),
    predictions: getAllPredictions_(),
    results: getResults_(),
    scoreboard: buildScoreboard_(),
    crowdForecast: buildCrowdForecast_(),
    visiblePredictions: buildVisiblePredictionList_(),
  };
}

function ensureOfficialMasterDataLoaded_() {
  const eventsSheet = getSheet_(SHEETS.EVENTS);
  const entriesSheet = getSheet_(SHEETS.ENTRIES);
  if (eventsSheet.getLastRow() > 1 && entriesSheet.getLastRow() > 1) return;
  importMasterData_(OFFICIAL_MASTER_DATA_2025);
}

function importMasterData_(masterData) {
  if (!masterData || !Array.isArray(masterData.events) || !Array.isArray(masterData.entries)) {
    throw new Error('masterData は events / entries 配列を含む必要があります。');
  }

  const events = masterData.events.map(function (e, i) {
    if (!e.eventId || !e.eventName) {
      throw new Error('events[' + i + '] に eventId / eventName が必要です。');
    }
    return [
      String(e.eventId).trim(),
      Number(e.sortOrder || i + 1),
      String(e.gender || '').trim(),
      String(e.eventName).trim(),
    ];
  });

  const eventIdSet = {};
  events.forEach(function (row) { eventIdSet[row[0]] = true; });

  const entries = masterData.entries.map(function (r, i) {
    if (!r.entryId || !r.eventId || !r.athleteName) {
      throw new Error('entries[' + i + '] に entryId / eventId / athleteName が必要です。');
    }
    if (!eventIdSet[String(r.eventId).trim()]) {
      throw new Error('entries[' + i + '] の eventId が events に存在しません。');
    }
    return [
      String(r.entryId).trim(),
      String(r.eventId).trim(),
      Number(r.seedOrder || i + 1),
      String(r.athleteName).trim(),
      String(r.team || '').trim(),
      String(r.entryTime || '').trim(),
    ];
  });

  replaceSheetRows_(SHEETS.EVENTS, HEADERS[SHEETS.EVENTS], events);
  replaceSheetRows_(SHEETS.ENTRIES, HEADERS[SHEETS.ENTRIES], entries);
}

function publicConfig_(config) {
  return {
    [CONFIG_KEYS.DEADLINE_ISO]: config[CONFIG_KEYS.DEADLINE_ISO] || '',
    [CONFIG_KEYS.REVEAL_PREDICTIONS]: config[CONFIG_KEYS.REVEAL_PREDICTIONS] || 'false',
    [CONFIG_KEYS.PUBLISH_SCOREBOARD]: config[CONFIG_KEYS.PUBLISH_SCOREBOARD] || 'false',
    [CONFIG_KEYS.PUBLISH_CROWD_FORECAST]: config[CONFIG_KEYS.PUBLISH_CROWD_FORECAST] || 'false',
    [CONFIG_KEYS.TIMEZONE]: config[CONFIG_KEYS.TIMEZONE] || DEFAULT_CONFIG[CONFIG_KEYS.TIMEZONE],
    [CONFIG_KEYS.APP_TITLE]: config[CONFIG_KEYS.APP_TITLE] || DEFAULT_CONFIG[CONFIG_KEYS.APP_TITLE],
  };
}

function assertWithinDeadline_(config) {
  const iso = (config[CONFIG_KEYS.DEADLINE_ISO] || '').trim();
  if (!iso) return;
  const deadline = new Date(iso);
  if (isNaN(deadline.getTime())) throw new Error('管理設定の締切日時が不正です。');
  if (new Date().getTime() > deadline.getTime()) {
    throw new Error('締切を過ぎているため編集できません。');
  }
}

function ensureSheets_() {
  const ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(function (name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const header = HEADERS[name];
    const currentHeader = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, header.length).getValues()[0] : [];
    const mismatch = header.some(function (h, i) { return currentHeader[i] !== h; });
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    } else if (mismatch) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);
    }
  });
  const config = getAppConfig_();
  const patch = {};
  Object.keys(DEFAULT_CONFIG).forEach(function (k) {
    if (!Object.prototype.hasOwnProperty.call(config, k)) patch[k] = DEFAULT_CONFIG[k];
  });
  if (Object.keys(patch).length) setConfigValues_(patch);
}

function getAppConfig_() {
  const rows = readSheetObjects_(SHEETS.CONFIG);
  const map = {};
  rows.forEach(function (row) {
    if (row.key) map[String(row.key)] = String(row.value == null ? '' : row.value);
  });
  return map;
}

function setConfigValues_(patch) {
  const sheet = getSheet_(SHEETS.CONFIG);
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues() : [];
  const rowIndexByKey = {};
  rows.forEach(function (row, idx) {
    rowIndexByKey[String(row[0])] = idx + 2;
  });
  Object.keys(patch || {}).forEach(function (key) {
    const value = patch[key];
    const rowIndex = rowIndexByKey[key];
    if (rowIndex) {
      sheet.getRange(rowIndex, 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  });
}

function getParticipants_() {
  return readSheetObjects_(SHEETS.PARTICIPANTS)
    .filter(function (r) { return String(r.isActive) !== 'false'; })
    .map(function (r) {
      const displayName = String(r.displayName || '').trim();
      const participantId = String(r.participantId || '').trim() || normalizeParticipantName_(displayName);
      if (!displayName) return null;
      return {
        participantId: participantId,
        displayName: displayName,
        isActive: String(r.isActive) !== 'false',
        createdAt: String(r.createdAt || ''),
      };
    })
    .filter(function (r) { return !!r; });
}

function getParticipantById_(participantId) {
  const id = String(participantId || '').trim();
  if (!id) return null;
  return getParticipants_().find(function (p) { return p.participantId === id; }) || null;
}

function getMasterData_() {
  const events = readSheetObjects_(SHEETS.EVENTS)
    .map(function (r) {
      return {
        eventId: String(r.eventId),
        sortOrder: Number(r.sortOrder || 0),
        gender: String(r.gender || ''),
        eventName: String(r.eventName || ''),
      };
    })
    .sort(function (a, b) { return a.sortOrder - b.sortOrder; });

  const entries = readSheetObjects_(SHEETS.ENTRIES)
    .map(function (r) {
      return {
        entryId: String(r.entryId),
        eventId: String(r.eventId),
        seedOrder: Number(r.seedOrder || 0),
        athleteName: String(r.athleteName || ''),
        team: String(r.team || ''),
        entryTime: String(r.entryTime || ''),
      };
    })
    .sort(function (a, b) {
      if (a.eventId !== b.eventId) return a.eventId < b.eventId ? -1 : 1;
      return a.seedOrder - b.seedOrder;
    });

  return { events: events, entries: entries };
}

function getPredictionsByParticipant_(participantId) {
  const rows = readSheetObjects_(SHEETS.PREDICTIONS).filter(function (r) {
    return String(r.participantId) === String(participantId);
  });
  const map = {};
  rows.forEach(function (r) {
    map[String(r.eventId)] = {
      eventId: String(r.eventId),
      pick1EntryId: String(r.pick1EntryId || ''),
      pick2EntryId: String(r.pick2EntryId || ''),
      pick3EntryId: String(r.pick3EntryId || ''),
      pick4EntryId: String(r.pick4EntryId || ''),
      updatedAt: String(r.updatedAt || ''),
    };
  });
  return map;
}

function getAllPredictions_() {
  return readSheetObjects_(SHEETS.PREDICTIONS).map(function (r) {
    return {
      participantId: String(r.participantId),
      eventId: String(r.eventId),
      pick1EntryId: String(r.pick1EntryId || ''),
      pick2EntryId: String(r.pick2EntryId || ''),
      pick3EntryId: String(r.pick3EntryId || ''),
      pick4EntryId: String(r.pick4EntryId || ''),
      updatedAt: String(r.updatedAt || ''),
    };
  });
}

function upsertPredictions_(participantId, newRows, deleteEventIds) {
  const sheet = getSheet_(SHEETS.PREDICTIONS);
  const all = readSheetObjects_(SHEETS.PREDICTIONS);
  const byKey = {};
  const deleteMap = {};
  (deleteEventIds || []).forEach(function (eventId) {
    deleteMap[String(participantId) + '::' + String(eventId)] = true;
  });
  all.forEach(function (row) {
    const key = String(row.participantId) + '::' + String(row.eventId);
    if (deleteMap[key]) return;
    byKey[key] = row;
  });
  newRows.forEach(function (row) {
    byKey[String(participantId) + '::' + String(row.eventId)] = row;
  });

  const merged = Object.keys(byKey).map(function (key) { return byKey[key]; });
  merged.sort(function (a, b) {
    if (a.participantId !== b.participantId) return a.participantId < b.participantId ? -1 : 1;
    return a.eventId < b.eventId ? -1 : 1;
  });

  replaceSheetRows_(SHEETS.PREDICTIONS, HEADERS[SHEETS.PREDICTIONS], merged.map(function (r) {
    return [r.participantId, r.eventId, r.pick1EntryId, r.pick2EntryId, r.pick3EntryId, r.pick4EntryId || '', r.updatedAt];
  }));
}

function getResults_() {
  const map = {};
  readSheetObjects_(SHEETS.RESULTS).forEach(function (r) {
    map[String(r.eventId)] = {
      eventId: String(r.eventId),
      firstEntryId: String(r.firstEntryId || ''),
      secondEntryId: String(r.secondEntryId || ''),
      thirdEntryId: String(r.thirdEntryId || ''),
      fourthEntryId: String(r.fourthEntryId || ''),
      updatedAt: String(r.updatedAt || ''),
    };
  });
  return map;
}

function buildVisiblePredictionList_() {
  const participants = getParticipants_();
  const participantMap = toMap_(participants, 'participantId');
  const predictions = getAllPredictions_();
  return predictions.map(function (row) {
    return {
      participantId: row.participantId,
      participantName: participantMap[row.participantId] ? participantMap[row.participantId].displayName : row.participantId,
      eventId: row.eventId,
      pick1EntryId: row.pick1EntryId,
      pick2EntryId: row.pick2EntryId,
      pick3EntryId: row.pick3EntryId,
      pick4EntryId: row.pick4EntryId || '',
      updatedAt: row.updatedAt,
    };
  });
}

function buildScoreboard_() {
  const participants = getParticipants_();
  const predictions = getAllPredictions_();
  const results = getResults_();
  const master = getMasterData_();
  const entryMap = toMap_(master.entries, 'entryId');

  const predByUser = {};
  predictions.forEach(function (p) {
    if (!predByUser[p.participantId]) predByUser[p.participantId] = {};
    predByUser[p.participantId][p.eventId] = [p.pick1EntryId, p.pick2EntryId, p.pick3EntryId, p.pick4EntryId || ''];
  });

  const resultByEvent = {};
  Object.keys(results).forEach(function (eventId) {
    const r = results[eventId];
    const event = (master.events || []).find(function (evt) { return evt.eventId === eventId; });
    const placementCount = getPlacementCountForEvent_(event);
    const picks = [r.firstEntryId, r.secondEntryId, r.thirdEntryId, r.fourthEntryId || ''].slice(0, placementCount);
    if (picks.every(Boolean)) {
      resultByEvent[eventId] = picks;
    }
  });

  const rows = participants.map(function (p) {
    let totalScore = 0;
    let exactHits = 0;
    let top3Hits = 0;
    const details = [];
    Object.keys(resultByEvent).forEach(function (eventId) {
      const pred = predByUser[p.participantId] && predByUser[p.participantId][eventId];
      const res = resultByEvent[eventId];
      if (!pred) return;
      const scored = scorePrediction_(pred, res);
      totalScore += scored.score;
      exactHits += scored.exactHits;
      top3Hits += scored.top3Hits;
      details.push({
        eventId: eventId,
        score: scored.score,
        exactHits: scored.exactHits,
        top3Hits: scored.top3Hits,
      });
    });
    return {
      participantId: p.participantId,
      participantName: p.displayName,
      totalScore: totalScore,
      exactHits: exactHits,
      top3Hits: top3Hits,
      details: details,
    };
  });

  rows.sort(function (a, b) {
    if (b.totalScore !== a.totalScore) return b.totalScore - a.totalScore;
    if (b.exactHits !== a.exactHits) return b.exactHits - a.exactHits;
    if (b.top3Hits !== a.top3Hits) return b.top3Hits - a.top3Hits;
    return a.participantName < b.participantName ? -1 : 1;
  });

  const eventSummaries = Object.keys(resultByEvent).map(function (eventId) {
    const result = resultByEvent[eventId].map(function (entryId) { return entryDisplay_(entryMap[entryId]); });
    return { eventId: eventId, resultTop3: result };
  });

  return {
    scoringRule: 'exact=3pt, top3/4(順不同)=1pt',
    ranking: rows,
    completedEvents: Object.keys(resultByEvent).length,
    eventSummaries: eventSummaries,
  };
}

function buildCrowdForecast_() {
  const master = getMasterData_();
  const events = master.events || [];
  const entries = master.entries || [];
  const predictions = getAllPredictions_();
  const entryMap = toMap_(entries, 'entryId');
  const entrySeedMap = {};
  entries.forEach(function (ent) {
    entrySeedMap[String(ent.entryId)] = Number(ent.seedOrder || 9999);
  });

  const predByEvent = {};
  predictions.forEach(function (p) {
    if (!predByEvent[p.eventId]) predByEvent[p.eventId] = [];
    predByEvent[p.eventId].push(p);
  });

  const eventForecasts = events.map(function (evt) {
    const placementCount = getPlacementCountForEvent_(evt);
    const eventPreds = predByEvent[evt.eventId] || [];
    const voteMap = {};

    eventPreds.forEach(function (p) {
      const picks = [p.pick1EntryId, p.pick2EntryId, p.pick3EntryId, p.pick4EntryId || ''].slice(0, placementCount);
      picks.forEach(function (entryId, idx) {
        if (!entryId) return;
        if (!voteMap[entryId]) {
          voteMap[entryId] = {
            entryId: entryId,
            totalPoints: 0,
            rankVotes: [0, 0, 0, 0],
            appearances: 0,
          };
        }
        voteMap[entryId].totalPoints += (placementCount - idx);
        voteMap[entryId].rankVotes[idx] += 1;
        voteMap[entryId].appearances += 1;
      });
    });

    const ranking = Object.keys(voteMap).map(function (entryId) {
      return voteMap[entryId];
    }).sort(function (a, b) {
      if (b.totalPoints !== a.totalPoints) return b.totalPoints - a.totalPoints;
      for (let i = 0; i < 4; i += 1) {
        if ((b.rankVotes[i] || 0) !== (a.rankVotes[i] || 0)) {
          return (b.rankVotes[i] || 0) - (a.rankVotes[i] || 0);
        }
      }
      const seedA = entrySeedMap[a.entryId] == null ? 9999 : entrySeedMap[a.entryId];
      const seedB = entrySeedMap[b.entryId] == null ? 9999 : entrySeedMap[b.entryId];
      if (seedA !== seedB) return seedA - seedB;
      return a.entryId < b.entryId ? -1 : 1;
    }).slice(0, placementCount).map(function (row, idx) {
      const entry = entryMap[row.entryId];
      return {
        predictedRank: idx + 1,
        entryId: row.entryId,
        athleteName: entry ? entry.athleteName : row.entryId,
        team: entry ? entry.team : '',
        totalPoints: row.totalPoints,
        appearances: row.appearances,
        rankVotes: row.rankVotes.slice(0, placementCount),
      };
    });

    return {
      eventId: evt.eventId,
      gender: evt.gender,
      eventName: evt.eventName,
      placementCount: placementCount,
      submissionCount: eventPreds.length,
      ranking: ranking,
    };
  });

  return {
    rule: '各種目で1位=高得点 ... （3位種目は3-2-1点、4位種目は4-3-2-1点）',
    events: eventForecasts,
  };
}

function scorePrediction_(pred, res) {
  let score = 0;
  let exactHits = 0;
  let top3Hits = 0;
  const n = Math.min(pred.length, res.length);
  for (let i = 0; i < n; i += 1) {
    if (!pred[i] || !res[i]) continue;
    if (pred[i] === res[i]) {
      score += 3;
      exactHits += 1;
      top3Hits += 1;
    } else if (res.indexOf(pred[i]) !== -1) {
      score += 1;
      top3Hits += 1;
    }
  }
  return { score: score, exactHits: exactHits, top3Hits: top3Hits };
}

function getPlacementCountForEvent_(event) {
  if (!event) return 3;
  return FOUR_PLACE_EVENT_NAMES[String(event.eventName || '').trim()] ? 4 : 3;
}

function entryDisplay_(entry) {
  if (!entry) return '';
  return [entry.athleteName, entry.team ? '(' + entry.team + ')' : ''].join(' ').trim();
}

function replaceSheetRows_(sheetName, header, rows) {
  const sheet = getSheet_(sheetName);
  sheet.clearContents();
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
}

function readSheetObjects_(sheetName) {
  const sheet = getSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  return values.slice(1).filter(function (row) {
    return row.some(function (cell) { return String(cell) !== ''; });
  }).map(function (row) {
    const obj = {};
    header.forEach(function (key, idx) {
      obj[String(key)] = row[idx];
    });
    return obj;
  });
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('シートが見つかりません: ' + name);
  return sheet;
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  let spreadsheetId = String(STANDALONE_SPREADSHEET_ID || '').trim();
  if (!spreadsheetId) {
    try {
      spreadsheetId = String(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '').trim();
    } catch (err) {
      spreadsheetId = '';
    }
  }
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  throw new Error(
    'スプレッドシートに紐づけた Apps Script で実行してください。' +
    '単体プロジェクトの場合は Script Properties に SPREADSHEET_ID を設定してください。'
  );
}

function toMap_(rows, keyField) {
  const map = {};
  (rows || []).forEach(function (row) {
    map[String(row[keyField])] = row;
  });
  return map;
}

function normalizeParticipantName_(name) {
  return String(name || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function isTrue_(value) {
  return String(value).toLowerCase() === 'true';
}

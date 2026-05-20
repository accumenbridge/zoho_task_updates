  /***************
 * Daily Work Tracker Ã¢â€ â€ Zoho Projects Sync
 * Replace Code.gs with this file.
 * Keep your real SPREADSHEET_ID and FORM_ID in CFG.
 ***************/
const CFG = {
  // Google Sheet ID from your Daily Work Tracker URL:
  // https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  SPREADSHEET_ID: '1Q-J9ZY5WmyG5Gak-r1kRTvgcEXVHYGRCTS8MG70CCLQ',

  // Google Form ID from your form editor URL:
  // https://docs.google.com/forms/d/FORM_ID/edit
  FORM_ID: '1JNSxjplWeSPViWOg2xcBwWAGYpzdBKwvN71UGvC_B9o  ',

  // Zoho DC examples: zoho.in, zoho.com, zoho.eu, zoho.com.au
  DEFAULT_DC: 'zoho.in',

  PORTALS_SHEET: 'ZohoPortals',
  PROJECTS_SHEET: 'ZohoProjects',
  TASKS_SHEET: 'ZohoTasks',
  UPDATES_SHEET: 'TaskUpdates',
  TEAM_SHEET: 'Team'
};

function getTrackerSpreadsheet_() {
  const id = String(CFG.SPREADSHEET_ID || '').trim();
  if (!id || id === 'PASTE_DAILY_WORK_TRACKER_SPREADSHEET_ID_HERE') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
    throw new Error('Set CFG.SPREADSHEET_ID from the Daily Work Tracker Google Sheet URL.');
  }
  return SpreadsheetApp.openById(id);
}

function notify_(message) {
  Logger.log(message);
  try {
    const ss = getTrackerSpreadsheet_();
    ss.toast(message, 'Zoho Sync', 8);
  } catch (err) {
    Logger.log('UI notification skipped: ' + err.message);
  }
}

function withScriptLock_(name, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Another Zoho Sync operation is already running. Try again in a minute: ' + name);
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function getTargetForm_() {
  const ss = getTrackerSpreadsheet_();

  const manualId = String(CFG.FORM_ID || '').trim();
  const isBadManualId =
    !manualId ||
    manualId.includes('PASTE_') ||
    manualId === 'FORM_ID';

  if (!isBadManualId) {
    try {
      return FormApp.openById(manualId);
    } catch (err) {
      Logger.log('Manual FORM_ID failed. Trying linked form URL. Error: ' + err.message);
    }
  }

  let formUrl = '';

  const responseSheet = ss.getSheetByName('Responses');

  if (responseSheet && typeof responseSheet.getFormUrl === 'function') {
    formUrl = responseSheet.getFormUrl();
  }

  if (!formUrl && typeof ss.getFormUrl === 'function') {
    formUrl = ss.getFormUrl();
  }

  if (!formUrl) {
    throw new Error(
      'Cannot find linked Google Form. Open the Form editor, copy the ID from /forms/d/FORM_ID/edit, put it in CFG.FORM_ID, and make sure this Google account has edit access to the Form.'
    );
  }

  return FormApp.openByUrl(formUrl);
}

/***************
 * ZOHO SETUP
 ***************/
function setupZohoProperties() {
  PropertiesService.getScriptProperties().setProperties({
    ZOHO_DC: CFG.DEFAULT_DC,

    // Paste real values only when generating refresh token.
    // After exchangeGrantCodeForRefreshToken() succeeds, replace these with placeholders again.
    ZOHO_CLIENT_ID: 'PASTE_ZOHO_CLIENT_ID_HERE',
    ZOHO_CLIENT_SECRET: 'PASTE_ZOHO_CLIENT_SECRET_HERE',
    ZOHO_GRANT_CODE: 'PASTE_ZOHO_GRANT_CODE_HERE'
  }, true);

  notify_('Zoho properties saved. Now run exchangeGrantCodeForRefreshToken().');
}

function exchangeGrantCodeForRefreshToken() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const url = `https://accounts.${props.ZOHO_DC}/oauth/v2/token`;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      grant_type: 'authorization_code',
      client_id: props.ZOHO_CLIENT_ID,
      client_secret: props.ZOHO_CLIENT_SECRET,
      code: props.ZOHO_GRANT_CODE
    },
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  const json = JSON.parse(text);

  if (!json.refresh_token) {
    throw new Error('Refresh token not returned. Response: ' + text);
  }

  PropertiesService.getScriptProperties().setProperty('ZOHO_REFRESH_TOKEN', json.refresh_token);
  notify_('Refresh token saved successfully. Now replace visible Zoho secrets with placeholders and do not run setupZohoProperties() again unless rotating credentials.');
}

/***************
 * MENU
 ***************/
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Zoho Sync')
      .addItem('1. Setup Zoho Properties', 'setupZohoProperties')
      .addItem('2. Exchange Grant Code', 'exchangeGrantCodeForRefreshToken')
      .addItem('3. List Zoho Portals', 'listZohoPortals')
      .addItem('4. List Zoho Projects', 'listZohoProjects')
      .addItem('5. Sync Zoho Tasks + Update Form', 'syncZohoTasksAndUpdateForm')
      .addItem('6. Install Triggers', 'installTriggers')
      .addItem('7. Retry Latest Quota-Failed Update', 'retryLatestFailedTaskUpdateOnce')
      .addToUi();
  } catch (err) {
    Logger.log('Zoho Sync menu was not created: ' + err.message);
  }
}

/***************
 * ZOHO API HELPERS
 ***************/
function getZohoAccessToken_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('ZOHO_ACCESS_TOKEN');
  if (cached) return cached;

  const props = PropertiesService.getScriptProperties().getProperties();
  const url = `https://accounts.${props.ZOHO_DC}/oauth/v2/token`;

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      grant_type: 'refresh_token',
      client_id: props.ZOHO_CLIENT_ID,
      client_secret: props.ZOHO_CLIENT_SECRET,
      refresh_token: props.ZOHO_REFRESH_TOKEN
    },
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  const json = JSON.parse(text);

  if (!json.access_token) {
    throw new Error('Access token not returned. Response: ' + text);
  }

  cache.put('ZOHO_ACCESS_TOKEN', json.access_token, 3300);
  return json.access_token;
}

function zohoApi_(method, path, params) {
  const props = PropertiesService.getScriptProperties().getProperties();
  const token = getZohoAccessToken_();
  const base = `https://projectsapi.${props.ZOHO_DC}/restapi`;

  let url = base + path;
  const options = {
    method: method.toLowerCase(),
    headers: {
      Authorization: `Zoho-oauthtoken ${token}`
    },
    muteHttpExceptions: true
  };

  if (method.toUpperCase() === 'GET' && params && Object.keys(params).length) {
    url += '?' + toQuery_(params);
  } else if (params && Object.keys(params).length) {
    options.payload = params;
    options.contentType = 'application/x-www-form-urlencoded';
  }

  const maxAttempts = 4;
  let lastCode = 0;
  let lastText = '';

  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    let response;
    let code;
    let text;

    try {
      response = UrlFetchApp.fetch(url, options);
      code = response.getResponseCode();
      text = response.getContentText();
    } catch (err) {
      code = 0;
      text = err.message || String(err);
    }

    if (code >= 200 && code < 300) {
      return text ? JSON.parse(text) : {};
    }

    lastCode = code;
    lastText = text;

    if (!isTransientZohoQuotaError_(code, text) || attempt === maxAttempts - 1) {
      break;
    }

    Utilities.sleep(zohoRetryDelayMs_(attempt));
  }

  throw new Error(`Zoho API error ${lastCode || 'fetch'}: ${truncate_(lastText, 600)}`);
}

function isTransientZohoQuotaError_(code, text) {
  const body = String(text || '').toLowerCase();
  return code === 429 ||
    code === 408 ||
    code === 503 ||
    code === 504 ||
    (code >= 500 && code < 600) ||
    body.indexOf('bandwidth quota') !== -1 ||
    body.indexOf('quota exceeded') !== -1 ||
    body.indexOf('rate limit') !== -1 ||
    body.indexOf('too many requests') !== -1 ||
    body.indexOf('try reducing the rate') !== -1 ||
    body.indexOf('temporarily') !== -1;
}

function zohoRetryDelayMs_(attempt) {
  const base = [2000, 5000, 10000, 20000][attempt] || 20000;
  return base + Math.floor(Math.random() * 1000);
}

function truncate_(value, maxLen) {
  const text = String(value || '');
  return text.length > maxLen ? text.slice(0, maxLen) + '...' : text;
}

function toQuery_(obj) {
  return Object.keys(obj)
    .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(obj[k]))
    .join('&');
}

/***************
 * SHEET HELPERS
 ***************/
function getOrCreateSheet_(name, headers) {
  const ss = getTrackerSpreadsheet_();
  let sh = ss.getSheetByName(name);

  if (!sh) sh = ss.insertSheet(name);

  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }

  return sh;
}

function clearAndSet_(sheet, headers, rows) {
  sheet.clear();

  // Important: keep Zoho IDs as text. Long numeric IDs break if Google Sheets auto-converts them.
  sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setNumberFormat('@');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
}

function getTeamOwnerMap_() {
  const sh = getTrackerSpreadsheet_().getSheetByName(CFG.TEAM_SHEET);
  if (!sh) throw new Error('Team sheet not found.');

  const values = sh.getDataRange().getValues();
  const map = {};
  const allowedOwners = new Set();

  for (let i = 1; i < values.length; i++) {
    const formName = String(values[i][0] || '').trim();
    const zohoOwner = String(values[i][1] || values[i][0] || '').trim();

    if (formName) {
      map[formName] = zohoOwner;
      allowedOwners.add(zohoOwner);
    }
  }

  return { map, allowedOwners };
}

/***************
 * PORTALS + PROJECTS
 ***************/
function listZohoPortals() {
  const json = zohoApi_('GET', '/portals/', {});
  const portals = json.portals || [];

  const rows = portals.map((p, idx) => [
    p.name || p.portal_name || '',
    String(p.id_string || p.id || ''),
    idx === 0 ? 'Yes' : ''
  ]);

  const sh = getOrCreateSheet_(CFG.PORTALS_SHEET);
  clearAndSet_(sh, ['Portal Name', 'Portal ID', 'Use?'], rows);

  notify_('Zoho portals listed. Keep Yes for the portal you want to use.');
}

function getSelectedPortalId_() {
  const sh = getTrackerSpreadsheet_().getSheetByName(CFG.PORTALS_SHEET);
  if (!sh) throw new Error('Run listZohoPortals() first.');

  const values = sh.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][2] || '').trim().toLowerCase() === 'yes') {
      return String(values[i][1]).trim();
    }
  }

  throw new Error('No portal selected. In ZohoPortals sheet, put Yes in Use?.');
}

function listZohoProjects() {
  const portalId = getSelectedPortalId_();

  const json = zohoApi_('GET', `/portal/${portalId}/projects/`, {});
  const projects = json.projects || [];

  const rows = projects.map(p => [
    p.name || '',
    String(p.id_string || p.id || ''),
    ''
  ]);

  const sh = getOrCreateSheet_(CFG.PROJECTS_SHEET);
  clearAndSet_(sh, ['Project Name', 'Project ID', 'Sync?'], rows);

  notify_('Zoho projects listed. Put Yes in Sync? for the projects you want in the form.');
}

/***************
 * TASK SYNC Ã¢â€ â€™ FORM DROPDOWNS
 ***************/
function syncZohoTasksAndUpdateForm() {
  return withScriptLock_('syncZohoTasksAndUpdateForm', syncZohoTasksAndUpdateFormLocked_);
}

function syncZohoTasksAndUpdateFormLocked_() {
  const ss = getTrackerSpreadsheet_();
  const projectsSheet = ss.getSheetByName(CFG.PROJECTS_SHEET);
  if (!projectsSheet) throw new Error('Run listZohoProjects() first.');

  const portalId = getSelectedPortalId_();
  const { allowedOwners } = getTeamOwnerMap_();
  const projectRows = projectsSheet.getDataRange().getValues();
  const lastFormByTaskId = getLatestSuccessfulFormPercentByTaskId_();
  const rows = [];

  for (let i = 1; i < projectRows.length; i++) {
    const projectName = String(projectRows[i][0] || '').trim();
    const projectId = String(projectRows[i][1] || '').trim();
    const sync = String(projectRows[i][2] || '').trim().toLowerCase();

    if (sync !== 'yes') continue;

    let index = 1;
    const range = 200;

    while (true) {
      const json = zohoApi_(
        'GET',
        `/portal/${portalId}/projects/${projectId}/tasks/`,
        { index, range }
      );

      const tasks = json.tasks || [];
      if (!tasks.length) break;

      tasks.forEach(task => {
        const taskInfo = makeZohoTaskSyncInfo_(
          portalId,
          projectId,
          projectName,
          task,
          'Task',
          null,
          lastFormByTaskId
        );

        if (!taskInfo) return;

        if (isVisibleZohoTaskForOwners_(taskInfo, allowedOwners)) {
          rows.push(taskInfo.row);
        }

        if (hasZohoSubtasks_(task)) {
          appendZohoSubtaskRows_(
            rows,
            portalId,
            projectId,
            projectName,
            taskInfo,
            allowedOwners,
            lastFormByTaskId,
            1
          );
        }
      });

      if (tasks.length < range) break;
      index += tasks.length;
      if (index > 2000) break;
    }
  }

  const sh = getOrCreateSheet_(CFG.TASKS_SHEET);
  clearAndSet_(sh, [
    'Task Choice',
    'Portal ID',
    'Project ID',
    'Project Name',
    'Task ID',
    'Task Key',
    'Task Name',
    'Owner',
    'Current %',
    'Status',
    'Task URL',
    'Last Synced',
    'Task Type',
    'Parent Task ID',
    'Parent Task Key',
    'Parent Task Name',
    'Previous Form %',
    'Previous Form Time'
  ], rows);

  updateZohoTaskDropdown_();
  notify_(`Synced ${rows.length} open assigned Zoho tasks and updated the Google Form dropdown.`);
}

function makeZohoTaskSyncInfo_(portalId, projectId, projectName, task, taskType, parentInfo, lastFormByTaskId) {
  const taskId = String(task.id_string || task.id || '').trim();
  const taskName = String(task.name || task.title || '').trim();
  const taskKey = String(task.key || '').trim();
  const percent = String(extractZohoTaskPercent_(task) || '').trim();
  const statusName = getZohoTaskStatusName_(task);
  const completed = isZohoTaskCompleted_(task, percent);

  if (!taskId || !taskName) return null;

  const ownerNames = extractZohoOwnerNames_(task);
  const ownerText = ownerNames.join(', ');
  const displayKey = taskKey || taskId;
  const parentKey = parentInfo ? (parentInfo.taskKey || parentInfo.taskId) : '';
  const lastForm = lastFormByTaskId[taskId] || {};
  const choice = buildTaskChoice_(projectName, displayKey, taskName, percent, lastForm.percent, parentKey);

  const taskUrl = task.link && task.link.self && task.link.self.url
    ? task.link.self.url
    : '';

  return {
    taskId,
    taskKey,
    taskName,
    ownerNames,
    ownerText,
    completed,
    row: [
      String(choice),
      String(portalId),
      String(projectId),
      String(projectName),
      String(taskId),
      String(taskKey),
      String(taskName),
      String(ownerText),
      String(percent),
      String(statusName),
      String(taskUrl),
      new Date(),
      String(taskType || 'Task'),
      parentInfo ? String(parentInfo.taskId || '') : '',
      parentInfo ? String(parentInfo.taskKey || '') : '',
      parentInfo ? String(parentInfo.taskName || '') : '',
      String(lastForm.percent || ''),
      lastForm.timestamp || ''
    ]
  };
}

function appendZohoSubtaskRows_(rows, portalId, projectId, projectName, parentInfo, allowedOwners, lastFormByTaskId, depth) {
  if (!parentInfo || !parentInfo.taskId || depth > 4) return;

  let index = 1;
  const range = 200;

  while (true) {
    const json = zohoApi_(
      'GET',
      `/portal/${portalId}/projects/${projectId}/tasks/${parentInfo.taskId}/subtasks/`,
      { index, range }
    );

    const subtasks = json.tasks || json.subtasks || [];
    if (!subtasks.length) break;

    subtasks.forEach(subtask => {
      const subtaskInfo = makeZohoTaskSyncInfo_(
        portalId,
        projectId,
        projectName,
        subtask,
        'Subtask',
        parentInfo,
        lastFormByTaskId
      );

      if (!subtaskInfo) return;

      if (isVisibleZohoTaskForOwners_(subtaskInfo, allowedOwners)) {
        rows.push(subtaskInfo.row);
      }

      if (hasZohoSubtasks_(subtask)) {
        appendZohoSubtaskRows_(
          rows,
          portalId,
          projectId,
          projectName,
          subtaskInfo,
          allowedOwners,
          lastFormByTaskId,
          depth + 1
        );
      }
    });

    if (subtasks.length < range) break;
    index += subtasks.length;
    if (index > 2000) break;
  }
}

function isVisibleZohoTaskForOwners_(taskInfo, allowedOwners) {
  if (!taskInfo || taskInfo.completed || !taskInfo.ownerNames.length) return false;
  return taskInfo.ownerNames.some(o => allowedOwners.has(o));
}

function hasZohoSubtasks_(task) {
  const value = task.subtasks || task.subtask_count || task.subtasks_count || task.subtaskCount || task.child_count;
  if (value === true) return true;
  if (Array.isArray(value)) return value.length > 0;
  if (typeof value === 'number') return value > 0;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || (Number.isFinite(Number(text)) && Number(text) > 0);
}

function isZohoTaskCompleted_(task, percent) {
  const completedText = String(task.completed || task.closed || task.is_completed || '').trim().toLowerCase();
  const statusText = getZohoTaskStatusName_(task).toLowerCase();
  return completedText === 'true' ||
    statusText === 'closed' ||
    statusText === 'completed' ||
    Number(percent) >= 100;
}

function getZohoTaskStatusName_(task) {
  if (task.status && task.status.name) return String(task.status.name || '').trim();
  if (task.status && typeof task.status === 'string') return String(task.status || '').trim();
  if (task.status_name) return String(task.status_name || '').trim();
  return '';
}

function buildTaskChoice_(projectName, taskKey, taskName, currentPercent, previousFormPercent, parentKey) {
  const keyPath = parentKey ? `${parentKey} > ${taskKey}` : taskKey;
  const previousText = formatChoicePercent_(previousFormPercent, '-');
  return [
    projectName,
    keyPath,
    taskName,
    `Previous: ${previousText === '-' ? previousText : previousText + '%'}`
  ].join(' | ');
}

function formatChoicePercent_(value, emptyText) {
  const text = String(value || '').replace(/%/g, '').trim();
  if (!text) return emptyText === undefined ? '0' : emptyText;
  const n = Number(text);
  return Number.isFinite(n) ? String(Math.round(n)) : text;
}

function getLatestSuccessfulFormPercentByTaskId_() {
  const sh = getTrackerSpreadsheet_().getSheetByName(CFG.UPDATES_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};

  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const taskIdCol = headers.indexOf('Task ID');
  const percentCol = headers.indexOf('% Completion');
  const statusCol = headers.indexOf('Zoho Sync Status');
  const timestampCol = headers.indexOf('Timestamp');

  if (taskIdCol < 0 || percentCol < 0 || statusCol < 0) return {};

  const out = {};
  for (let i = 1; i < values.length; i++) {
    const status = String(values[i][statusCol] || '').trim();
    const taskId = String(values[i][taskIdCol] || '').trim();
    const percent = String(values[i][percentCol] || '').trim();
    if (status !== 'SUCCESS' || !taskId || !percent) continue;
    out[taskId] = {
      percent,
      timestamp: timestampCol >= 0 ? values[i][timestampCol] : ''
    };
  }
  return out;
}

function extractZohoOwnerNames_(task) {
  const names = [];

  function addName(value) {
    if (!value) return;

    if (typeof value === 'string') {
      names.push(value.trim());
      return;
    }

    if (Array.isArray(value)) {
      value.forEach(addName);
      return;
    }

    if (typeof value === 'object') {
      if (value.name) names.push(String(value.name).trim());
      if (value.full_name) names.push(String(value.full_name).trim());
      if (value.display_name) names.push(String(value.display_name).trim());
      if (value.zp_user_name) names.push(String(value.zp_user_name).trim());
      if (value.email) names.push(String(value.email).trim());
      if (value.email_id) names.push(String(value.email_id).trim());
    }
  }

  if (task.details && task.details.owners) addName(task.details.owners);
  if (task.owners) addName(task.owners);
  if (task.owner) addName(task.owner);
  if (task.assignee) addName(task.assignee);
  if (task.assignees) addName(task.assignees);

  return [...new Set(names.filter(Boolean))];
}

function updateZohoTaskDropdown_() {
  updatePersonTaskDropdowns_();
}

function updatePersonTaskDropdowns_() {
  const form = getTargetForm_();
  const ss = getTrackerSpreadsheet_();
  const taskSheet = ss.getSheetByName(CFG.TASKS_SHEET);

  if (!taskSheet) {
    throw new Error('ZohoTasks sheet not found. Run Zoho task sync first.');
  }

  const { map } = getTeamOwnerMap_();
  const taskValues = taskSheet.getDataRange().getValues();
  const listItems = form
    .getItems(FormApp.ItemType.LIST)
    .map(item => item.asListItem());

  Object.keys(map).forEach(formName => {
    const zohoOwnerName = String(map[formName] || '').trim();
    const questionTitle = `Zoho Task - ${formName}`;

    const taskItem = listItems.find(item => String(item.getTitle() || '').trim() === questionTitle);

    if (!taskItem) {
      Logger.log(`Missing Google Form dropdown: ${questionTitle}`);
      return;
    }

    const choices = [];

    for (let i = 1; i < taskValues.length; i++) {
      const taskChoice = String(taskValues[i][0] || '').trim();
      const ownerText = String(taskValues[i][7] || '').trim();

      if (!taskChoice || !ownerText) continue;

      const owners = ownerText.split(',').map(x => x.trim()).filter(Boolean);
      if (owners.includes(zohoOwnerName)) choices.push(taskChoice);
    }

    taskItem.setChoiceValues(
      choices.length ? choices : [`No open assigned Zoho tasks for ${formName}`]
    );

    Logger.log(`${questionTitle}: ${choices.length} tasks`);
  });

  applyPercentValidation_(form);
  notify_('Updated person-wise Zoho task dropdowns in the Google Form.');
}

function applyPercentValidation_(form) {
  const validation = FormApp.createTextValidation()
    .setHelpText('% Completion must be one of: 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100.')
    .requireTextMatchesPattern('^(0|10|20|30|40|50|60|70|80|90|100)$')
    .build();

  form.getItems(FormApp.ItemType.TEXT).forEach(item => {
    if (String(item.getTitle() || '').trim() === '% Completion') {
      item.asTextItem().setValidation(validation);
    }
  });
}

/***************
 * FORM SUBMIT Ã¢â€ â€™ ZOHO UPDATE + TASKUPDATES LOG
 ***************/
function onTaskUpdateFormSubmit(e) {
  return withScriptLock_('onTaskUpdateFormSubmit', () => onTaskUpdateFormSubmitLocked_(e));
}

function onTaskUpdateFormSubmitLocked_(e) {
  let update = null;
  let meta = null;
  let status = 'SUCCESS';
  const errors = [];

  try {
    update = parseTaskUpdateFromSubmitEvent_(e);
    meta = getTaskMetaByChoice_(update.taskChoice);
    validateOwner_(update.name, meta.owner);

    let percentUpdated = false;
    try {
      updateZohoTaskPercent_(meta, update.percent);
      percentUpdated = true;
    } catch (err) {
      errors.push('Percent update failed: ' + (err.message || String(err)));
    }

    if (percentUpdated) {
      try {
        addZohoTaskComment_(meta, update);
      } catch (err) {
        errors.push('Comment add failed: ' + (err.message || String(err)));
      }
    }
  } catch (err) {
    errors.push(err.message || String(err));
  }

  if (!update) {
    update = buildParseErrorUpdate_(e, errors.join(' | '));
  }

  if (errors.length) status = 'ERROR';
  appendTaskUpdate_(update, meta, status, errors.join(' | '));
}

function retryLatestFailedTaskUpdateOnce() {
  return withScriptLock_('retryLatestFailedTaskUpdateOnce', retryLatestFailedTaskUpdateOnceLocked_);
}

function retryLatestFailedTaskUpdateOnceLocked_() {
  const sh = getTrackerSpreadsheet_().getSheetByName(CFG.UPDATES_SHEET);
  if (!sh) throw new Error('TaskUpdates sheet not found.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('No TaskUpdates rows found.');

  const headers = values[0].map(h => String(h || '').trim());
  const statusCol = headers.indexOf('Zoho Sync Status');
  const errorCol = headers.indexOf('Zoho Error');
  if (statusCol < 0 || errorCol < 0) {
    throw new Error('TaskUpdates sheet is missing Zoho Sync Status or Zoho Error columns.');
  }

  for (let r = values.length - 1; r >= 1; r--) {
    const status = String(values[r][statusCol] || '').trim();
    const error = String(values[r][errorCol] || '').trim();
    if (status !== 'ERROR') continue;
    if (!isRetryableTaskUpdateError_(error)) {
      throw new Error('Latest ERROR row is not a transient quota/rate failure: ' + truncate_(error, 240));
    }

    const update = buildUpdateFromTaskUpdateRow_(headers, values[r]);
    update.proof = update.proof
      ? update.proof + '\nRetry of TaskUpdates row ' + (r + 1)
      : 'Retry of TaskUpdates row ' + (r + 1);

    const meta = getTaskMetaByChoice_(update.taskChoice);
    validateOwner_(update.name, meta.owner);

    let statusOut = 'SUCCESS';
    const errors = [];
    try {
      updateZohoTaskPercent_(meta, update.percent);
    } catch (err) {
      errors.push('Percent update failed: ' + (err.message || String(err)));
    }

    const oldCommentFailed = error.indexOf('Comment add failed:') !== -1;
    if (!errors.length && oldCommentFailed) {
      try {
        addZohoTaskComment_(meta, update);
      } catch (err) {
        errors.push('Comment add failed: ' + (err.message || String(err)));
      }
    }

    if (errors.length) statusOut = 'ERROR';
    appendTaskUpdate_(update, meta, statusOut, errors.join(' | '));

    const message = statusOut === 'SUCCESS'
      ? 'Retried TaskUpdates row ' + (r + 1) + ' successfully.'
      : 'Retried TaskUpdates row ' + (r + 1) + ' but it still failed.';
    notify_(message);
    return;
  }

  throw new Error('No ERROR rows found to retry.');
}

function isRetryableTaskUpdateError_(error) {
  const text = String(error || '').toLowerCase();
  return text.indexOf('bandwidth quota') !== -1 ||
    text.indexOf('quota exceeded') !== -1 ||
    text.indexOf('rate limit') !== -1 ||
    text.indexOf('too many requests') !== -1 ||
    text.indexOf('try reducing the rate') !== -1 ||
    text.indexOf('temporarily') !== -1 ||
    text.indexOf('zoho api error 429') !== -1 ||
    text.indexOf('zoho api error 503') !== -1 ||
    text.indexOf('zoho api error 504') !== -1;
}

function buildUpdateFromTaskUpdateRow_(headers, row) {
  const get = name => {
    const idx = headers.indexOf(name);
    return idx >= 0 ? row[idx] : '';
  };
  const workDateRaw = get('Work Date');
  const percentRaw = get('% Completion');
  return {
    timestamp: new Date(),
    workDateRaw: workDateRaw,
    workDate: parseDateOnly_(workDateRaw),
    name: cellToText_(get('Name')),
    taskChoice: cellToText_(get('Zoho Task')),
    percent: parsePercent_(percentRaw),
    workDone: cellToText_(get('Work Done Today')),
    blockers: cellToText_(get('Blockers')),
    nextSteps: cellToText_(get('Next Steps')),
    proof: cellToText_(get('Proof Link / Notes'))
  };
}

function parseTaskUpdateFromSubmitEvent_(e) {
  const rows = getSubmittedHeaderValueRows_(e);
  const nv = e && e.namedValues ? e.namedValues : {};

  let workDateRaw = getNamedOrRowValue_(nv, rows, 'Work Date');
  let name = getNamedOrRowValue_(nv, rows, 'Name');

  // Fallback: if Work Date question is renamed/missing, use timestamp date instead of losing the update.
  if (!workDateRaw) {
    workDateRaw = getNamedOrRowValue_(nv, rows, 'Timestamp') || new Date();
  }

  if (!name) throw new Error('Name is missing. Check the Google Form question title is exactly "Name".');

  const taskChoice =
    getNamedOrRowValue_(nv, rows, `Zoho Task - ${name}`) ||
    getFirstNonEmptyHeaderPrefix_(rows, 'Zoho Task -') ||
    getFirstNonEmptyNamedPrefix_(nv, 'Zoho Task -') ||
    getNamedOrRowValue_(nv, rows, 'Zoho Task');

  if (!taskChoice) throw new Error(`No Zoho task selected for ${name}.`);
  if (String(taskChoice).startsWith('No open assigned Zoho tasks')) {
    throw new Error(`No valid Zoho task selected for ${name}.`);
  }

  const percentRaw = getLastNonEmptyExact_(rows, '% Completion') || getNamedOrRowValue_(nv, rows, '% Completion');
  const workDone = getLastNonEmptyExact_(rows, 'Work Done Today') || getLastNonEmptyExact_(rows, 'Work done today');
  const blockers = getLastNonEmptyExact_(rows, 'Blockers');
  const nextSteps = getLastNonEmptyExact_(rows, 'Next Steps');
  const proof = getLastNonEmptyExact_(rows, 'Proof Link / Notes');

  if (!percentRaw) throw new Error('% Completion is missing.');
  if (!workDone) throw new Error('Work Done Today is missing.');
  if (!blockers) throw new Error('Blockers is missing. Use N.A if no blocker.');
  if (!nextSteps) throw new Error('Next Steps is missing.');

  return {
    timestamp: new Date(),
    workDateRaw: workDateRaw,
    workDate: parseDateOnly_(workDateRaw),
    name: name,
    taskChoice: taskChoice,
    percent: parsePercent_(percentRaw),
    workDone: workDone,
    blockers: blockers,
    nextSteps: nextSteps,
    proof: proof
  };
}

function getSubmittedHeaderValueRows_(e) {
  if (e && e.range) {
    const responseSheet = e.range.getSheet();
    const lastColumn = responseSheet.getLastColumn();
    const headers = responseSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const values = e.values && e.values.length
      ? e.values
      : responseSheet.getRange(e.range.getRow(), 1, 1, lastColumn).getValues()[0];

    const rows = [];
    for (let i = 0; i < lastColumn; i++) {
      rows.push({
        header: String(headers[i] || '').trim(),
        value: cellToText_(values[i])
      });
    }
    return rows;
  }

  if (e && e.namedValues) {
    return Object.keys(e.namedValues).map(k => ({
      header: String(k || '').trim(),
      value: getNamedValue_(e.namedValues, k)
    }));
  }

    throw new Error('Form submit event is missing. Submit the Google Form. To retry a quota-failed TaskUpdates row, run retryLatestFailedTaskUpdateOnce().');
}

function cellToText_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'M/d/yyyy');
  }
  return String(value || '').trim();
}

function getNamedValue_(namedValues, key) {
  if (!namedValues || !namedValues[key] || !namedValues[key].length) return '';
  return cellToText_(namedValues[key][0]);
}

function getNamedOrRowValue_(namedValues, rows, title) {
  return getNamedValue_(namedValues, title) || getFirstNonEmptyExact_(rows, title);
}

function normalizeHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function getFirstNonEmptyExact_(rows, title) {
  const target = normalizeHeader_(title);
  for (let i = 0; i < rows.length; i++) {
    if (normalizeHeader_(rows[i].header) === target && rows[i].value !== '') return rows[i].value;
  }
  return '';
}

function getLastNonEmptyExact_(rows, title) {
  const target = normalizeHeader_(title);
  for (let i = rows.length - 1; i >= 0; i--) {
    if (normalizeHeader_(rows[i].header) === target && rows[i].value !== '') return rows[i].value;
  }
  return '';
}

function getFirstNonEmptyHeaderPrefix_(rows, prefix) {
  const targetPrefix = normalizeHeader_(prefix);
  for (let i = 0; i < rows.length; i++) {
    if (normalizeHeader_(rows[i].header).startsWith(targetPrefix) && rows[i].value !== '') return rows[i].value;
  }
  return '';
}

function getFirstNonEmptyNamedPrefix_(namedValues, prefix) {
  const targetPrefix = normalizeHeader_(prefix);
  const keys = Object.keys(namedValues || {});
  for (let i = 0; i < keys.length; i++) {
    if (normalizeHeader_(keys[i]).startsWith(targetPrefix)) {
      const value = getNamedValue_(namedValues, keys[i]);
      if (value) return value;
    }
  }
  return '';
}

function buildParseErrorUpdate_(e, error) {
  let name = 'UNKNOWN';
  let workDate = new Date();

  try {
    const rows = getSubmittedHeaderValueRows_(e);
    const nv = e && e.namedValues ? e.namedValues : {};
    name = getNamedOrRowValue_(nv, rows, 'Name') || 'UNKNOWN';
    const rawDate = getNamedOrRowValue_(nv, rows, 'Work Date') || getNamedOrRowValue_(nv, rows, 'Timestamp');
    if (rawDate) workDate = parseDateOnly_(rawDate);
  } catch (ignored) {}

  return {
    timestamp: new Date(),
    workDateRaw: '',
    workDate: workDate,
    name: name,
    taskChoice: 'PARSE ERROR',
    percent: '',
    workDone: '',
    blockers: '',
    nextSteps: '',
    proof: ''
  };
}

function buildNamedValuesFromRow_(sheet, row) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  const out = {};

  for (let i = 0; i < lastColumn; i++) {
    const h = String(headers[i] || '').trim();
    const v = cellToText_(values[i]);
    if (!h || !v) continue;
    if (!out[h]) out[h] = [];
    out[h].push(v);
  }
  return out;
}

/***************
 * VALIDATION + ZOHO WRITE
 ***************/
function parsePercent_(value) {
  const cleaned = String(value || '').replace(/%/g, '').trim();
  const n = Number(cleaned);
  if (!Number.isFinite(n) || n < 0 || n > 100) {
    throw new Error('% Completion must be one of 0, 10, 20, ... 100.');
  }

  const rounded = Math.round(n);
  if (Math.abs(n - rounded) > 0.0001 || rounded % 10 !== 0) {
    throw new Error('% Completion must be one of 0, 10, 20, ... 100. Zoho Projects accepts task completion only in multiples of 10.');
  }

  return rounded;
}
function parseDateOnly_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const d = new Date(value);
  if (isNaN(d.getTime())) throw new Error('Invalid Work Date: ' + value);
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function getTaskMetaByChoice_(choice) {
  const sh = getTrackerSpreadsheet_().getSheetByName(CFG.TASKS_SHEET);
  if (!sh) throw new Error('ZohoTasks sheet not found. Run task sync first.');

  const values = sh.getDataRange().getValues();
  const normalizedChoice = normalizeTaskChoice_(choice);
  const parsed = parseTaskChoice_(choice);
  for (let i = 1; i < values.length; i++) {
    if (normalizeTaskChoice_(values[i][0]) === normalizedChoice) {
      return taskMetaFromRow_(values[i]);
    }
  }

  if (parsed.taskKey || parsed.projectName || parsed.taskName) {
    const matches = [];

    for (let i = 1; i < values.length; i++) {
      const rowProject = normalizeLookupText_(values[i][3]);
      const rowTaskId = normalizeLookupText_(values[i][4]);
      const rowTaskKey = normalizeLookupText_(values[i][5]);
      const rowTaskName = normalizeLookupText_(values[i][6]);

      const idMatches = parsed.taskId && rowTaskId === normalizeLookupText_(parsed.taskId);
      const keyMatches = parsed.taskKey && rowTaskKey === normalizeLookupText_(parsed.taskKey);
      const projectMatches = !parsed.projectName || rowProject === normalizeLookupText_(parsed.projectName);
      const nameMatches = !parsed.taskName || rowTaskName === normalizeLookupText_(parsed.taskName);

      if ((idMatches || keyMatches) && projectMatches && nameMatches) {
        matches.push(values[i]);
      }
    }

    if (!matches.length && (parsed.taskId || parsed.taskKey)) {
      for (let i = 1; i < values.length; i++) {
        const rowTaskId = normalizeLookupText_(values[i][4]);
        const rowTaskKey = normalizeLookupText_(values[i][5]);
        const idMatches = parsed.taskId && rowTaskId === normalizeLookupText_(parsed.taskId);
        const keyMatches = parsed.taskKey && rowTaskKey === normalizeLookupText_(parsed.taskKey);
        if (idMatches || keyMatches) {
          matches.push(values[i]);
        }
      }
    }

    if (matches.length === 1) {
      return taskMetaFromRow_(matches[0]);
    }

    if (matches.length > 1) {
      throw new Error(`Selected Zoho Task matched multiple ZohoTasks rows for key "${parsed.taskKey}". Re-run Zoho Sync, then choose the task again from the refreshed form.`);
    }
  }

  throw new Error('Selected Zoho Task was not found in ZohoTasks sheet. Run Zoho Sync Ã¢â€ â€™ 5. Sync Zoho Tasks + Update Form, then submit again.');
}

function taskMetaFromRow_(row) {
  const meta = {
    choice: String(row[0] || '').trim(),
    portalId: String(row[1] || '').trim(),
    projectId: String(row[2] || '').trim(),
    projectName: String(row[3] || '').trim(),
    taskId: String(row[4] || '').trim(),
    taskKey: String(row[5] || '').trim(),
    taskName: String(row[6] || '').trim(),
    owner: String(row[7] || '').trim(),
    currentPercent: String(row[8] || '').trim(),
    taskUrl: String(row[10] || '').trim(),
    taskType: String(row[12] || 'Task').trim(),
    parentTaskId: String(row[13] || '').trim(),
    parentTaskKey: String(row[14] || '').trim(),
    parentTaskName: String(row[15] || '').trim()
  };
  validateZohoIds_(meta);
  return meta;
}

function normalizeTaskChoice_(value) {
  return String(value || '')
    .replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig, '')
    .replace(/\s*,\s*,\s*/g, ', ')
    .replace(/\s*,\s*\|/g, ' |')
    .replace(/\s*\|\s*/g, ' | ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeLookupText_(value) {
  return normalizeTaskChoice_(value).toLowerCase();
}

function parseTaskChoice_(choice) {
  const parts = stripTaskChoiceMetaParts_(
    normalizeTaskChoice_(choice).split(' | ').map(s => s.trim()).filter(Boolean)
  );

  if (parts.length >= 3 && looksLikeTaskKeyPath_(parts[1])) {
    const taskPath = parts[1];
    const taskKeyOrId = lastTaskKeyFromPath_(taskPath);
    return {
      projectName: parts[0],
      taskKey: /^\d+$/.test(taskKeyOrId) ? '' : taskKeyOrId,
      taskId: /^\d+$/.test(taskKeyOrId) ? taskKeyOrId : '',
      taskName: parts.slice(2).join(' | ')
    };
  }

  if (parts.length < 4) return {};

  return {
    owner: parts[0],
    projectName: parts[1],
    taskKey: /^\d+$/.test(parts[2]) ? '' : parts[2],
    taskId: /^\d+$/.test(parts[2]) ? parts[2] : '',
    taskName: parts.slice(3).join(' | ')
  };
}

function stripTaskChoiceMetaParts_(parts) {
  const cleaned = parts.slice();
  while (cleaned.length) {
    const last = String(cleaned[cleaned.length - 1] || '').trim().toLowerCase();
    if (/^(current|previous form|last form|previous|zoho)\s*:/.test(last)) {
      cleaned.pop();
      continue;
    }
    break;
  }
  return cleaned;
}

function looksLikeTaskKeyPath_(value) {
  const text = String(value || '').trim();
  return text.indexOf('>') !== -1 ||
    /^[A-Z]{1,8}\d*-T\d+/i.test(text) ||
    /^\d{8,}$/.test(text);
}

function lastTaskKeyFromPath_(value) {
  const parts = String(value || '').split('>').map(s => s.trim()).filter(Boolean);
  return parts.length ? parts[parts.length - 1] : String(value || '').trim();
}

function validateOwner_(formName, taskOwnerText) {
  const { map } = getTeamOwnerMap_();
  const expectedZohoOwner = map[formName];
  if (!expectedZohoOwner) throw new Error(`Name "${formName}" is not found in Team sheet.`);

  const taskOwners = String(taskOwnerText || '').split(',').map(s => s.trim()).filter(Boolean);
  if (!taskOwners.includes(expectedZohoOwner)) {
    throw new Error(`Task is assigned to "${taskOwnerText}", but form was submitted by "${formName}".`);
  }
}

function validateZohoIds_(meta) {
  ['portalId', 'projectId', 'taskId'].forEach(key => {
    const value = String(meta[key] || '').trim();
    if (!/^\d+$/.test(value)) {
      throw new Error(`Invalid Zoho ${key}: ${value}. Re-run Zoho Sync so IDs are stored as text.`);
    }
  });
}

function updateZohoTaskPercent_(meta, percent) {
  validateZohoIds_(meta);

  const p = String(parsePercent_(percent));
  const path = `/portal/${meta.portalId}/projects/${meta.projectId}/tasks/${meta.taskId}/`;

  const result = zohoApi_('POST', path, { percent_complete: p });

  for (let i = 0; i < 3; i++) {
    Utilities.sleep(1000);
    const current = getZohoTaskPercentFromProject_(meta);
    if (String(current) === p) return result;
  }

  throw new Error('Zoho accepted the percent update request, but the task still does not show ' + p + '%. Recheck task percent settings or API response.');
}

function getZohoTaskPercentFromProject_(meta) {
  try {
    const json = zohoApi_('GET', `/portal/${meta.portalId}/projects/${meta.projectId}/tasks/${meta.taskId}/`, {});
    const task = firstZohoTaskFromResponse_(json);
    if (task) {
      const percent = extractZohoTaskPercent_(task);
      if (percent !== '') return percent;
    }
  } catch (err) {
    Logger.log('Direct Zoho task readback skipped: ' + (err.message || String(err)));
  }

  if (meta.parentTaskId) {
    const subtaskPercent = getZohoSubtaskPercentFromParent_(meta);
    if (subtaskPercent !== '') return subtaskPercent;
  }

  const range = 100;
  let index = 1;

  for (let page = 0; page < 25; page++) {
    const json = zohoApi_('GET', `/portal/${meta.portalId}/projects/${meta.projectId}/tasks/`, {
      range: range,
      index: index
    });
    const tasks = json.tasks || [];

    for (let i = 0; i < tasks.length; i++) {
      const taskId = String(tasks[i].id_string || tasks[i].id || '').trim();
      if (taskId === String(meta.taskId)) {
        return extractZohoTaskPercent_(tasks[i]);
      }
    }

    if (tasks.length < range) break;
    index += tasks.length;
  }

  throw new Error('Could not re-read Zoho task ' + meta.taskId + ' to verify percent update.');
}

function getZohoSubtaskPercentFromParent_(meta) {
  const range = 100;
  let index = 1;

  for (let page = 0; page < 25; page++) {
    const json = zohoApi_('GET', `/portal/${meta.portalId}/projects/${meta.projectId}/tasks/${meta.parentTaskId}/subtasks/`, {
      range: range,
      index: index
    });
    const tasks = json.tasks || json.subtasks || [];

    for (let i = 0; i < tasks.length; i++) {
      const taskId = String(tasks[i].id_string || tasks[i].id || '').trim();
      if (taskId === String(meta.taskId)) {
        return extractZohoTaskPercent_(tasks[i]);
      }
    }

    if (tasks.length < range) break;
    index += tasks.length;
  }

  return '';
}

function firstZohoTaskFromResponse_(json) {
  if (!json) return null;
  if (json.task) return Array.isArray(json.task) ? json.task[0] : json.task;
  if (json.tasks && json.tasks.length) return json.tasks[0];
  if (json.subtasks && json.subtasks.length) return json.subtasks[0];
  return null;
}

function extractZohoTaskPercent_(task) {
  const keys = ['percent_complete', 'percentcomplete', 'pcomplete'];
  for (let i = 0; i < keys.length; i++) {
    const value = task[keys[i]];
    if (value !== undefined && value !== null && String(value).trim() !== '') {
      return String(Math.round(Number(value))).trim();
    }
  }
  return '';
}
function addZohoTaskComment_(meta, update) {
  validateZohoIds_(meta);

  const dateText = Utilities.formatDate(update.workDate, Session.getScriptTimeZone(), 'dd MMM yyyy');
  const content =
    `<b>Daily Task Update - ${html_(dateText)}</b><br>` +
    `<b>Updated by:</b> ${html_(update.name)}<br>` +
    `<b>% Completion:</b> ${html_(update.percent)}%<br><br>` +
    `<b>Work Done Today:</b><br>${html_(update.workDone)}<br><br>` +
    `<b>Blockers:</b><br>${html_(update.blockers || 'None')}<br><br>` +
    `<b>Next Steps:</b><br>${html_(update.nextSteps)}<br><br>` +
    `<b>Proof / Notes:</b><br>${html_(update.proof || '-')}`;

  zohoApi_('POST', `/portal/${meta.portalId}/projects/${meta.projectId}/tasks/${meta.taskId}/comments/`, { content: content });
}

function html_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}

function appendTaskUpdate_(update, meta, status, error) {
  const sh = getOrCreateSheet_(CFG.UPDATES_SHEET, [
    'Timestamp',
    'Work Date',
    'Name',
    'Zoho Task',
    '% Completion',
    'Work Done Today',
    'Blockers',
    'Next Steps',
    'Proof Link / Notes',
    'Portal ID',
    'Project ID',
    'Task ID',
    'Zoho Sync Status',
    'Zoho Sync Time',
    'Zoho Error'
  ]);

  // Keep Zoho IDs as text in the log too.
  sh.getRange('J:L').setNumberFormat('@');

  sh.appendRow([
    update.timestamp,
    update.workDate,
    update.name,
    update.taskChoice,
    update.percent,
    update.workDone,
    update.blockers,
    update.nextSteps,
    update.proof,
    meta ? String(meta.portalId) : '',
    meta ? String(meta.projectId) : '',
    meta ? String(meta.taskId) : '',
    status,
    new Date(),
    error
  ]);

  sh.getRange('B:B').setNumberFormat('dd-mmm-yyyy');
}

/***************
 * TRIGGERS
 ***************/
function installTriggers() {
  const ss = getTrackerSpreadsheet_();

  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === 'onTaskUpdateFormSubmit' || fn === 'syncZohoTasksAndUpdateForm') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('onTaskUpdateFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  ScriptApp.newTrigger('syncZohoTasksAndUpdateForm')
    .timeBased()
    .everyHours(1)
    .create();

  notify_('Triggers installed: form submit sync + hourly Zoho task refresh.');
}

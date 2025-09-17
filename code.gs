const SHEET_ID = '1O8Vnw24dWCRukdjtZC3XNBW2Je7IkEwKD7qz2a-Djj0';
const SHEET_NAMES = {
  USERS: 'Users',
  TASKS: 'Tasks',
  SUBTASKS: 'Subtasks',
  ACTIVITY: 'ActivityLog',
  MOODS: 'Moods',
  ATTACHMENTS: 'Attachments'
};

const BRANDING = {
  name: 'Aura Flow V2',
  logo: '🌐',
  primary: '#4C1D95',
  accent: '#7C3AED',
  dark: '#1F1B2E'
};

const ROLE_LEVEL = {
  'Admin': 4,
  'Sub-Admin': 3,
  'Manager': 2,
  'Intern': 1
};

const STATUS_LIST = ['Planned', 'In Progress', 'Review', 'Blocked', 'Done'];
const STATUS_COLORS = {
  'Planned': '#CBD5F5',
  'In Progress': '#D97706',
  'Review': '#60A5FA',
  'Blocked': '#F87171',
  'Done': '#34D399'
};
const PRIORITY_COLORS = {
  'Critical': '#EF4444',
  'High': '#F97316',
  'Medium': '#FBBF24',
  'Low': '#60A5FA',
  'Backlog': '#9CA3AF'
};

/**
 * Entry point that serves the bundled single-page app.
 */
function doGet() {
  initializeWorkspace();
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle(`${BRANDING.logo} ${BRANDING.name}`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Ensure that all configured sheets exist with the required headers.
 */
function initializeWorkspace() {
  const db = SpreadsheetApp.openById(SHEET_ID);
  const sheetConfigs = {
    [SHEET_NAMES.USERS]: ['Email', 'Password', 'Role', 'ManagerEmail', 'IsActive', 'NotificationEmail', 'CreatedAt'],
    [SHEET_NAMES.TASKS]: ['TaskID', 'Name', 'Category', 'Priority', 'Status', 'DurationMins', 'Labels', 'Notes', 'ResourcesCSV', 'Assigner', 'Assignee', 'Timestamp', 'DueAt', 'UpdatedAt', 'ParentTaskID'],
    [SHEET_NAMES.SUBTASKS]: ['SubtaskID', 'TaskID', 'Name', 'DurationMins', 'Status', 'CreatedAt'],
    [SHEET_NAMES.ACTIVITY]: ['LogID', 'ActorEmail', 'Action', 'TargetType', 'TargetID', 'MetaJSON', 'At'],
    [SHEET_NAMES.MOODS]: ['MoodID', 'ActorEmail', 'Mood', 'Notes', 'At'],
    [SHEET_NAMES.ATTACHMENTS]: ['AttachmentID', 'TaskID', 'FileName', 'DriveId', 'Url', 'AddedBy', 'At']
  };

  Object.keys(sheetConfigs).forEach(name => {
    let sheet = db.getSheetByName(name);
    if (!sheet) {
      sheet = db.insertSheet(name);
    }
    const headers = sheetConfigs[name];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
      if (!headers.every((h, idx) => (existingHeaders[idx] || '').toString().trim() === h)) {
        sheet.clear();
        sheet.appendRow(headers);
      }
    }
  });
}

function getDbSpreadsheet() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  return getDbSpreadsheet().getSheetByName(name);
}

function toMap(list, keyField) {
  return list.reduce((acc, item) => {
    acc[item[keyField]] = item;
    return acc;
  }, {});
}

function readSheetRecords(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length <= 1) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    const record = {};
    headers.forEach((header, idx) => {
      record[header] = row[idx];
    });
    return record;
  });
}

function writeRow(sheet, rowValues) {
  sheet.appendRow(rowValues);
}

function updateRowById(sheet, idField, id, updates) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const index = headers.indexOf(idField);
  if (index === -1) {
    throw new Error(`Field ${idField} not found in ${sheet.getName()}`);
  }
  for (let r = 1; r < values.length; r++) {
    if (values[r][index] === id) {
      const row = values[r];
      Object.keys(updates).forEach(key => {
        const cIndex = headers.indexOf(key);
        if (cIndex !== -1) {
          row[cIndex] = updates[key];
        }
      });
      sheet.getRange(r + 1, 1, 1, headers.length).setValues([row]);
      return;
    }
  }
  throw new Error(`Record ${id} not found in ${sheet.getName()}`);
}

function parseBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') return value.toLowerCase() === 'true';
  return !!value;
}

function normalizeEmail(email) {
  return (email || '').trim().toLowerCase();
}

function getActiveUsers() {
  const sheet = getSheet(SHEET_NAMES.USERS);
  return readSheetRecords(sheet)
    .filter(u => parseBoolean(u.IsActive))
    .map(u => ({
      Email: (u.Email || '').trim(),
      NormalizedEmail: normalizeEmail(u.Email),
      Password: u.Password,
      Role: u.Role || 'Intern',
      ManagerEmail: (u.ManagerEmail || '').trim(),
      NotificationEmail: u.NotificationEmail || '',
      CreatedAt: u.CreatedAt
    }));
}

function getUserByEmail(email) {
  const target = normalizeEmail(email);
  const users = getActiveUsers();
  return users.find(u => u.NormalizedEmail === target);
}

function assertUser(auth) {
  if (!auth || !auth.email) {
    throw new Error('Missing authentication payload.');
  }
  const user = getUserByEmail(auth.email);
  if (!user) {
    throw new Error('User record not found.');
  }
  if (auth.role && user.Role !== auth.role) {
    throw new Error('Role mismatch for authenticated user.');
  }
  return user;
}

function loginUser(email, password) {
  initializeWorkspace();
  const users = getActiveUsers();
  const lookup = users.find(u => normalizeEmail(u.Email) === normalizeEmail(email));
  if (!lookup) {
    throw new Error('Account not found or inactive.');
  }
  if (lookup.Password !== password) {
    throw new Error('Invalid credentials.');
  }
  const bootstrap = buildBootstrapPayload(lookup);
  return {
    success: true,
    user: sanitizeUser(lookup),
    bootstrap
  };
}

function sanitizeUser(user) {
  return {
    email: user.Email,
    role: user.Role,
    managerEmail: user.ManagerEmail,
    notificationEmail: user.NotificationEmail
  };
}

function buildBootstrapPayload(user) {
  const tasks = getTasksForUser(user);
  const subtasks = getSubtasksForTasks(tasks.map(t => t.TaskID));
  const attachments = getAttachmentsForTasks(tasks.map(t => t.TaskID));
  const activity = getRecentActivity(user);
  const metrics = buildDashboardMetrics(user, tasks, subtasks);
  const roles = Object.keys(ROLE_LEVEL);
  const users = getUsersVisibleTo(user);

  return {
    tasks,
    subtasks,
    attachments,
    activity,
    metrics,
    statuses: STATUS_LIST,
    statusColors: STATUS_COLORS,
    priorityColors: PRIORITY_COLORS,
    roles,
    branding: BRANDING,
    users
  };
}

function getUsersVisibleTo(user) {
  const users = getActiveUsers();
  const roleLevel = ROLE_LEVEL[user.Role] || 0;
  if (roleLevel >= ROLE_LEVEL['Admin']) {
    return dedupeUsers(users.map(sanitizeUser));
  }
  const myEmail = normalizeEmail(user.Email);
  const team = users.filter(u => normalizeEmail(u.ManagerEmail) === myEmail || normalizeEmail(u.Email) === myEmail);
  if (roleLevel >= ROLE_LEVEL['Sub-Admin']) {
    const combined = team.concat(users.filter(u => normalizeEmail(u.ManagerEmail) === normalizeEmail(user.ManagerEmail))).map(sanitizeUser);
    return dedupeUsers(combined);
  }
  return dedupeUsers(team.map(sanitizeUser));
}

function dedupeUsers(list) {
  const seen = {};
  return list.filter(user => {
    const key = normalizeEmail(user.email);
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function getTasksForUser(user) {
  const sheet = getSheet(SHEET_NAMES.TASKS);
  const allTasks = readSheetRecords(sheet).map(formatTaskRecord);
  const users = getActiveUsers();
  const userMap = toMap(users, 'Email');
  return allTasks.filter(task => canUserSeeTask(user, task, userMap));
}

function formatTaskRecord(record) {
  return {
    TaskID: record.TaskID || Utilities.getUuid(),
    Name: record.Name || '',
    Category: record.Category || '',
    Priority: record.Priority || 'Medium',
    Status: record.Status || 'Planned',
    DurationMins: Number(record.DurationMins) || 0,
    Labels: (record.Labels || '').toString().split(',').map(s => s.trim()).filter(Boolean),
    Notes: record.Notes || '',
    Resources: (record.ResourcesCSV || '').toString().split(',').map(s => s.trim()).filter(Boolean),
    Assigner: record.Assigner || '',
    Assignee: record.Assignee || '',
    Timestamp: record.Timestamp || '',
    DueAt: record.DueAt || '',
    UpdatedAt: record.UpdatedAt || record.Timestamp || '',
    ParentTaskID: record.ParentTaskID || '',
    Color: STATUS_COLORS[record.Status] || '#A855F7'
  };
}

function getSubtasksForTasks(taskIds) {
  if (!taskIds || taskIds.length === 0) return [];
  const sheet = getSheet(SHEET_NAMES.SUBTASKS);
  return readSheetRecords(sheet)
    .filter(sub => taskIds.indexOf(sub.TaskID) !== -1)
    .map(sub => ({
      SubtaskID: sub.SubtaskID,
      TaskID: sub.TaskID,
      Name: sub.Name,
      DurationMins: Number(sub.DurationMins) || 0,
      Status: sub.Status || 'Planned',
      CreatedAt: sub.CreatedAt || ''
    }));
}

function getAttachmentsForTasks(taskIds) {
  if (!taskIds || taskIds.length === 0) return [];
  const sheet = getSheet(SHEET_NAMES.ATTACHMENTS);
  return readSheetRecords(sheet)
    .filter(att => taskIds.indexOf(att.TaskID) !== -1)
    .map(att => ({
      AttachmentID: att.AttachmentID,
      TaskID: att.TaskID,
      FileName: att.FileName,
      DriveId: att.DriveId,
      Url: att.Url,
      AddedBy: att.AddedBy,
      At: att.At
    }));
}

function getRecentActivity(user) {
  const sheet = getSheet(SHEET_NAMES.ACTIVITY);
  const records = readSheetRecords(sheet);
  const limit = 40;
  const myEmail = normalizeEmail(user.Email || user.email);
  const visible = records.filter(rec => {
    const actor = normalizeEmail(rec.ActorEmail);
    return actor === myEmail || ROLE_LEVEL[user.Role] >= ROLE_LEVEL['Manager'];
  }).slice(-limit);
  return visible.map(rec => ({
    LogID: rec.LogID,
    ActorEmail: rec.ActorEmail,
    Action: rec.Action,
    TargetType: rec.TargetType,
    TargetID: rec.TargetID,
    MetaJSON: rec.MetaJSON,
    At: rec.At
  })).reverse();
}

function canUserSeeTask(user, task, userMap) {
  const level = ROLE_LEVEL[user.Role] || 0;
  const me = normalizeEmail(user.Email || user.email);
  const assignee = normalizeEmail(task.Assignee);
  const assigner = normalizeEmail(task.Assigner);

  if (level >= ROLE_LEVEL['Admin']) {
    return true;
  }

  if (assignee === me || assigner === me) {
    return true;
  }

  if (level >= ROLE_LEVEL['Sub-Admin']) {
    return assignee && normalizeEmail(userMap[task.Assignee] && userMap[task.Assignee].ManagerEmail) === me;
  }

  if (level >= ROLE_LEVEL['Manager']) {
    const userRecord = userMap[task.Assignee];
    return userRecord && normalizeEmail(userRecord.ManagerEmail) === me;
  }

  return false;
}

function ensureAssignmentAllowed(assigner, assigneeEmail) {
  const assignee = getUserByEmail(assigneeEmail);
  if (!assignee) {
    throw new Error('Assignee not found or inactive.');
  }
  const assignerLevel = ROLE_LEVEL[assigner.Role] || 0;
  const assigneeLevel = ROLE_LEVEL[assignee.Role] || 0;

  if (assignerLevel >= ROLE_LEVEL['Admin']) {
    return assignee;
  }
  if (assignerLevel >= ROLE_LEVEL['Sub-Admin']) {
    const isManaged = normalizeEmail(assignee.ManagerEmail) === normalizeEmail(assigner.Email);
    if (!isManaged) {
      throw new Error('Sub-Admins can only assign to their managed team.');
    }
    return assignee;
  }
  if (assignerLevel >= ROLE_LEVEL['Manager']) {
    if (assigneeLevel >= ROLE_LEVEL['Manager']) {
      throw new Error('Managers can only assign to interns.');
    }
    const isReport = normalizeEmail(assignee.ManagerEmail) === normalizeEmail(assigner.Email);
    if (!isReport) {
      throw new Error('Managers can only assign to their direct reports.');
    }
    return assignee;
  }

  if (assignerLevel >= ROLE_LEVEL['Intern']) {
    if (normalizeEmail(assigner.Email) !== normalizeEmail(assignee.Email)) {
      throw new Error('Interns can only create tasks for themselves.');
    }
    return assignee;
  }
  throw new Error('Role is not authorized for assignments.');
}

function createTask(auth, data) {
  const assigner = assertUser(auth);
  const assignee = ensureAssignmentAllowed(assigner, data.Assignee || assigner.Email);
  const now = new Date();
  const taskId = data.TaskID || Utilities.getUuid();
  const sheet = getSheet(SHEET_NAMES.TASKS);
  const payload = [
    taskId,
    data.Name || 'Untitled Task',
    data.Category || '',
    data.Priority || 'Medium',
    data.Status || 'Planned',
    Number(data.DurationMins) || 0,
    (data.Labels || []).join(', '),
    data.Notes || '',
    (data.Resources || []).join(', '),
    assigner.Email,
    assignee.Email,
    now,
    data.DueAt || '',
    now,
    data.ParentTaskID || ''
  ];
  writeRow(sheet, payload);

  logActivity(assigner.Email, 'create', 'Task', taskId, {
    name: data.Name,
    assignee: assignee.Email,
    priority: data.Priority
  });

  sendTaskCreationEmail(payload);

  return {
    task: formatTaskRecord({
      TaskID: taskId,
      Name: payload[1],
      Category: payload[2],
      Priority: payload[3],
      Status: payload[4],
      DurationMins: payload[5],
      Labels: payload[6],
      Notes: payload[7],
      ResourcesCSV: payload[8],
      Assigner: payload[9],
      Assignee: payload[10],
      Timestamp: payload[11],
      DueAt: payload[12],
      UpdatedAt: payload[13],
      ParentTaskID: payload[14]
    })
  };
}

function updateTask(auth, updates) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.TASKS);
  const now = new Date();
  const allowed = ['Name', 'Category', 'Priority', 'Status', 'DurationMins', 'Labels', 'Notes', 'ResourcesCSV', 'Assigner', 'Assignee', 'DueAt', 'ParentTaskID'];
  const updatePayload = {};
  if (updates.Assignee) {
    ensureAssignmentAllowed(user, updates.Assignee);
  }
  Object.keys(updates).forEach(key => {
    if (allowed.indexOf(key) !== -1) {
      if (key === 'Labels' && Array.isArray(updates[key])) {
        updatePayload[key] = updates[key].join(', ');
        return;
      }
      if (key === 'ResourcesCSV' && Array.isArray(updates[key])) {
        updatePayload[key] = updates[key].join(', ');
        return;
      }
      updatePayload[key] = updates[key];
    }
  });
  updatePayload['UpdatedAt'] = now;
  updateRowById(sheet, 'TaskID', updates.TaskID, updatePayload);

  logActivity(user.Email, 'update', 'Task', updates.TaskID, updatePayload);

  return {
    taskId: updates.TaskID,
    updatedAt: now
  };
}

function updateTaskStatus(auth, payload) {
  const user = assertUser(auth);
  const now = new Date();
  updateTask(auth, {
    TaskID: payload.TaskID,
    Status: payload.Status,
    UpdatedAt: now
  });
  logActivity(user.Email, 'status_change', 'Task', payload.TaskID, {
    status: payload.Status
  });
  return { success: true, updatedAt: now };
}

function saveSubtask(auth, subtask) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.SUBTASKS);
  const now = new Date();
  if (!subtask.TaskID) {
    throw new Error('Task reference required for subtask.');
  }
  if (subtask.SubtaskID) {
    updateRowById(sheet, 'SubtaskID', subtask.SubtaskID, {
      Name: subtask.Name,
      DurationMins: subtask.DurationMins,
      Status: subtask.Status
    });
    logActivity(user.Email, 'update', 'Subtask', subtask.SubtaskID, subtask);
    return { success: true, subtaskId: subtask.SubtaskID };
  }
  const id = Utilities.getUuid();
  writeRow(sheet, [id, subtask.TaskID, subtask.Name || '', Number(subtask.DurationMins) || 0, subtask.Status || 'Planned', now]);
  logActivity(user.Email, 'create', 'Subtask', id, subtask);
  return { success: true, subtaskId: id };
}

function recordAttachment(auth, attachment) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.ATTACHMENTS);
  const now = new Date();
  const id = Utilities.getUuid();
  writeRow(sheet, [id, attachment.TaskID, attachment.FileName, attachment.DriveId, attachment.Url, user.Email, now]);
  logActivity(user.Email, 'attach', 'Task', attachment.TaskID, attachment);
  return { success: true, attachmentId: id };
}

function buildDashboardMetrics(user, tasks, subtasks) {
  const totalTasks = tasks.length;
  const completed = tasks.filter(t => t.Status === 'Done').length;
  const inProgress = tasks.filter(t => t.Status === 'In Progress').length;
  const blocked = tasks.filter(t => t.Status === 'Blocked').length;
  const totalMinutes = tasks.reduce((acc, t) => acc + (Number(t.DurationMins) || 0), 0);
  const subtaskMinutes = subtasks.reduce((acc, st) => acc + (Number(st.DurationMins) || 0), 0);

  const byStatus = STATUS_LIST.map(status => ({
    status,
    color: STATUS_COLORS[status],
    count: tasks.filter(t => t.Status === status).length
  }));

  const byPriority = Object.keys(PRIORITY_COLORS).map(priority => ({
    priority,
    color: PRIORITY_COLORS[priority],
    count: tasks.filter(t => (t.Priority || 'Medium') === priority).length
  }));

  return {
    totalTasks,
    completed,
    inProgress,
    blocked,
    totalMinutes,
    subtaskMinutes,
    byStatus,
    byPriority,
    user: sanitizeUser(user)
  };
}

function refreshBootstrap(auth) {
  const user = assertUser(auth);
  return buildBootstrapPayload(user);
}

function logActivity(actorEmail, action, targetType, targetId, meta) {
  const sheet = getSheet(SHEET_NAMES.ACTIVITY);
  const now = new Date();
  const payload = [
    Utilities.getUuid(),
    actorEmail,
    action,
    targetType,
    targetId,
    JSON.stringify(meta || {}),
    now
  ];
  writeRow(sheet, payload);
}

function bulkImport(auth, payload) {
  const user = assertUser(auth);
  const results = {
    success: true,
    imported: [],
    skipped: [],
    errors: []
  };
  const rows = payload.rows || [];
  rows.forEach((row, idx) => {
    try {
      const task = {
        Name: row.Name || row.TaskName || `Imported Task ${idx + 1}`,
        Category: row.Category || '',
        Priority: row.Priority || 'Medium',
        Status: row.Status || 'Planned',
        DurationMins: Number(row.DurationMins || row.Duration || 0),
        Labels: (row.Labels || '').split(',').map(s => s.trim()).filter(Boolean),
        Notes: row.Notes || '',
        Resources: (row.Resources || row.Links || '').split(',').map(s => s.trim()).filter(Boolean),
        Assignee: row.Assignee || user.Email,
        DueAt: row.DueAt || row.Due || ''
      };
      const output = createTask({ email: user.Email, role: user.Role }, task);
      results.imported.push(output.task.TaskID);
    } catch (err) {
      results.skipped.push(idx + 1);
      results.errors.push(err.message);
    }
  });
  return results;
}

function exportTasks(auth, options) {
  const user = assertUser(auth);
  const tasks = getTasksForUser(user);
  const subtasks = getSubtasksForTasks(tasks.map(t => t.TaskID));
  if (options && options.format === 'pdf') {
    return exportTasksPdf(user, tasks, subtasks, options);
  }
  return exportTasksCsv(user, tasks, subtasks, options);
}

function exportTasksCsv(user, tasks, subtasks, options) {
  const headers = ['TaskID', 'Name', 'Category', 'Priority', 'Status', 'DurationMins', 'Labels', 'Assigner', 'Assignee', 'Timestamp', 'DueAt'];
  const rows = tasks.map(task => headers.map(header => {
    if (header === 'Labels') {
      return task.Labels.join('; ');
    }
    return task[header];
  }));
  const csv = [
    `${BRANDING.logo} ${BRANDING.name} Export`,
    headers.join(','),
    ...rows.map(r => r.map(value => typeof value === 'string' && value.indexOf(',') !== -1 ? `"${value}"` : value).join(','))
  ].join('\n');
  const blob = Utilities.newBlob(csv, 'text/csv', `aura-flow-tasks.csv`);
  return {
    mimeType: 'text/csv',
    fileName: 'aura-flow-tasks.csv',
    data: Utilities.base64Encode(blob.getBytes())
  };
}

function exportTasksPdf(user, tasks, subtasks, options) {
  const template = HtmlService.createTemplateFromFile('index');
  template.isExport = true;
  template.exportData = { tasks, subtasks, user: sanitizeUser(user), branding: BRANDING };
  const html = template.evaluate().getContent();
  const pdf = Utilities.newBlob(html, 'text/html', 'report.html').getAs('application/pdf');
  return {
    mimeType: 'application/pdf',
    fileName: 'aura-flow-report.pdf',
    data: Utilities.base64Encode(pdf.getBytes())
  };
}

function recordFocusSession(auth, payload) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.MOODS);
  const now = new Date();
  writeRow(sheet, [Utilities.getUuid(), user.Email, payload.Mood || 'Focused', payload.Notes || `Task ${payload.TaskID}`, now]);
  logActivity(user.Email, 'focus_session', 'Task', payload.TaskID, payload);
  return { success: true };
}

function sendTaskCreationEmail(taskRow) {
  const [taskId, name, , priority, status, , , notes, resourcesCSV, assigner, assignee, timestamp, dueAt] = taskRow;
  const subject = `${BRANDING.logo} ${BRANDING.name} - Task Assigned: ${name}`;
  const htmlBody = `
    <div style="font-family:Arial;padding:16px;color:#111">
      <h2 style="color:${BRANDING.primary}">${BRANDING.logo} ${BRANDING.name}</h2>
      <p>A new task has been created for <strong>${assignee}</strong>.</p>
      <table style="border-collapse:collapse;width:100%;max-width:520px">
        <tr><td style="padding:6px;border:1px solid #ddd">Task</td><td style="padding:6px;border:1px solid #ddd">${name}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ddd">Priority</td><td style="padding:6px;border:1px solid #ddd">${priority}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ddd">Status</td><td style="padding:6px;border:1px solid #ddd">${status}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ddd">Due</td><td style="padding:6px;border:1px solid #ddd">${dueAt || 'Not set'}</td></tr>
      </table>
      <p>${notes || ''}</p>
      <p>Resources: ${(resourcesCSV || '').split(',').map(s => `<a href="${s.trim()}">${s.trim()}</a>`).join(', ') || 'N/A'}</p>
      <p style="font-size:12px;color:#666">Task ID: ${taskId} &bull; Created by ${assigner} on ${timestamp}</p>
    </div>
  `;
  try {
    MailApp.sendEmail({ to: assignee, subject, htmlBody });
  } catch (err) {
    // ignore if MailApp is unavailable (e.g., in test environment)
  }
}

function sendDailySummaries() {
  initializeWorkspace();
  const users = getActiveUsers();
  const tasksSheet = getSheet(SHEET_NAMES.TASKS);
  const tasks = readSheetRecords(tasksSheet).map(formatTaskRecord);
  const now = new Date();
  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const summaryDate = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  users.forEach(user => {
    const myTasks = tasks.filter(task => normalizeEmail(task.Assignee) === normalizeEmail(user.Email));
    if (!myTasks.length) {
      return;
    }
    const started = myTasks.filter(task => task.Timestamp && task.Timestamp.toString().indexOf(summaryDate) !== -1);
    const completed = myTasks.filter(task => task.Status === 'Done' && task.UpdatedAt && task.UpdatedAt.toString().indexOf(summaryDate) !== -1);
    const totalMinutes = myTasks.reduce((acc, t) => acc + (Number(t.DurationMins) || 0), 0);

    const htmlBody = `
      <div style="font-family:Arial;padding:18px;background:#f9f9ff">
        <h2 style="color:${BRANDING.primary}">${BRANDING.logo} ${BRANDING.name}</h2>
        <p>Daily summary for <strong>${user.Email}</strong> (${summaryDate})</p>
        <ul>
          <li><strong>${started.length}</strong> tasks started</li>
          <li><strong>${completed.length}</strong> tasks completed</li>
          <li><strong>${totalMinutes}</strong> minutes logged</li>
        </ul>
        <h3 style="margin-top:16px;color:${BRANDING.accent}">Highlights</h3>
        <ol>
          ${completed.slice(0, 5).map(task => `<li>${task.Name} &mdash; ${task.DurationMins} mins</li>`).join('')}
        </ol>
        <p style="font-size:12px;color:#666">Generated automatically by ${BRANDING.logo} ${BRANDING.name}.</p>
      </div>
    `;
    const recipients = [user.Email];
    if (user.NotificationEmail) {
      recipients.push(user.NotificationEmail);
    }
    if (ROLE_LEVEL[user.Role] >= ROLE_LEVEL['Admin']) {
      recipients.push(user.Email);
    }
    try {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: `${BRANDING.logo} ${BRANDING.name} Daily Summary (${summaryDate})`,
        htmlBody
      });
    } catch (err) {
      // ignore silently in environments without MailApp permissions
    }
  });
}

function registerDailySummaryTrigger(hour, minute) {
  ScriptApp.newTrigger('sendDailySummaries')
    .timeBased()
    .atHour(hour || 17)
    .nearMinute(minute || 0)
    .everyDays(1)
    .create();
}

function lazyLoadData(auth, payload) {
  const user = assertUser(auth);
  switch (payload.section) {
    case 'dashboard':
      const tasks = getTasksForUser(user);
      const subtasks = getSubtasksForTasks(tasks.map(t => t.TaskID));
      return {
        metrics: buildDashboardMetrics(user, tasks, subtasks)
      };
    case 'kanban':
      return {
        tasks: getTasksForUser(user)
      };
    case 'activity':
      return {
        activity: getRecentActivity(user)
      };
    default:
      return { message: 'No lazy data for section.' };
  }
}

function setUserNotificationEmail(auth, email) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.USERS);
  updateRowById(sheet, 'Email', user.Email, { NotificationEmail: email });
  return { success: true };
}

function upsertUser(auth, payload) {
  const actor = assertUser(auth);
  if (ROLE_LEVEL[actor.Role] < ROLE_LEVEL['Admin']) {
    throw new Error('Only admins may manage users.');
  }
  const sheet = getSheet(SHEET_NAMES.USERS);
  const existing = getUserByEmail(payload.Email);
  if (existing) {
    updateRowById(sheet, 'Email', existing.Email, {
      Password: payload.Password || existing.Password,
      Role: payload.Role || existing.Role,
      ManagerEmail: payload.ManagerEmail || existing.ManagerEmail,
      IsActive: payload.IsActive !== undefined ? payload.IsActive : existing.IsActive,
      NotificationEmail: payload.NotificationEmail || existing.NotificationEmail
    });
    logActivity(actor.Email, 'update', 'User', existing.Email, payload);
    return { success: true, updated: true };
  }
  writeRow(sheet, [payload.Email, payload.Password || '', payload.Role || 'Intern', payload.ManagerEmail || '', true, payload.NotificationEmail || '', new Date()]);
  logActivity(actor.Email, 'create', 'User', payload.Email, payload);
  return { success: true, created: true };
}

function archiveTask(auth, payload) {
  const user = assertUser(auth);
  updateTask(auth, { TaskID: payload.TaskID, Status: 'Done' });
  logActivity(user.Email, 'archive', 'Task', payload.TaskID, {});
  return { success: true };
}

function removeAttachment(auth, payload) {
  const user = assertUser(auth);
  const sheet = getSheet(SHEET_NAMES.ATTACHMENTS);
  const range = sheet.getDataRange();
  const values = range.getValues();
  for (let r = 1; r < values.length; r++) {
    if (values[r][0] === payload.AttachmentID) {
      sheet.deleteRow(r + 1);
      logActivity(user.Email, 'detach', 'Attachment', payload.AttachmentID, payload);
      return { success: true };
    }
  }
  throw new Error('Attachment not found.');
}

function toggleTaskTimer(auth, payload) {
  const user = assertUser(auth);
  const taskSheet = getSheet(SHEET_NAMES.TASKS);
  const taskId = payload.TaskID;
  const range = taskSheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const idIndex = headers.indexOf('TaskID');
  const durationIndex = headers.indexOf('DurationMins');
  for (let r = 1; r < values.length; r++) {
    if (values[r][idIndex] === taskId) {
      const current = Number(values[r][durationIndex]) || 0;
      const increment = Number(payload.Minutes || 0);
      values[r][durationIndex] = current + increment;
      values[r][headers.indexOf('UpdatedAt')] = new Date();
      taskSheet.getRange(r + 1, 1, 1, headers.length).setValues([values[r]]);
      logActivity(user.Email, 'timer', 'Task', taskId, { increment });
      return { success: true, duration: current + increment };
    }
  }
  throw new Error('Task not found for timer.');
}

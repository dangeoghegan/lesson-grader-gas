/*
 * Lesson Grader web integration. Use this ONE file instead of the old
 * WebAppIntegration.gs + WebAppWorkflow.gs; do not keep either old copy.
 * Keep the existing Code.gs and WebApp.html. Exactly one doGet must exist.
 * This does not make existing spreadsheet-bound grading APIs web-safe.
 */
var LessonGraderWeb = (function () {
  var ROOT_NAME = 'Graded Assessments';
  var STAGES = ['Stage 4', 'Stage 5', 'Stage 6'];
  var ID_RE = /^[A-Za-z0-9_-]{20,200}$/;
  var SUB_HEADERS = ['SubmissionRecordID', 'StudentUserID', 'Status'];
  var CONFIG_HEADERS = [
    'ConfigID', 'CourseID', 'CourseName', 'CourseSection', 'CourseWorkID',
    'AssignmentTitle', 'MaxPoints', 'SavedAt', 'SavedBy', 'Active',
    'TaskName', 'DueDate', 'AutoImported'
  ];
  var NON_TASK_SHEETS = {
    Submissions: true, ClassroomConfig: true, Setup: true, SubmissionFiles: true,
    CriteriaConfig: true, CriterionAssessments: true, AIAssessments: true,
    AssessmentHistory: true, ErrorLog: true, ClassLists: true,
    RubricProfiles: true, Rubric: true, Criteria: true, Summary: true,
    Sheet1: true
  };

  function fail_(message) { return { success: false, message: message }; }
  function validId_(id) { return typeof id === 'string' && ID_RE.test(id); }
  function ownRoots_() { return DriveApp.getRootFolder().getFoldersByName(ROOT_NAME); }
  function header_(sheet, names) {
    if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < names.length) return false;
    var actual = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
    return names.every(function (name) { return actual.indexOf(name) !== -1; });
  }
  function validWorkbook_(ss) {
    return header_(ss.getSheetByName(Config.SHEET_SUBMISSIONS), SUB_HEADERS) &&
      header_(ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG), CONFIG_HEADERS);
  }
  function walk_(callback) {
    var roots = ownRoots_();
    while (roots.hasNext()) {
      var root = roots.next();
      for (var i = 0; i < STAGES.length; i++) {
        var stages = root.getFoldersByName(STAGES[i]);
        while (stages.hasNext()) {
          var courses = stages.next().getFolders();
          while (courses.hasNext()) {
            var categories = courses.next().getFolders();
            while (categories.hasNext()) {
              var files = categories.next().getFilesByType(MimeType.GOOGLE_SHEETS);
              while (files.hasNext()) callback(files.next());
            }
          }
        }
      }
    }
  }
  function member_(id) {
    var found = false;
    walk_(function (file) { if (file.getId() === id) found = true; });
    return found;
  }
  function open_(id) {
    if (!validId_(id)) throw new Error('Invalid workbook ID.');
    var file = DriveApp.getFileById(id);
    if (file.isTrashed() || file.getMimeType() !== MimeType.GOOGLE_SHEETS || !member_(id)) {
      throw new Error('Workbook is not in the Graded Assessments library.');
    }
    var ss = SpreadsheetApp.openById(id);
    if (!validWorkbook_(ss)) throw new Error('Workbook is not initialised for Lesson Grader.');
    return ss;
  }
  function taskNames_(ss) {
    return ss.getSheets().filter(function (sheet) {
      return !NON_TASK_SHEETS[sheet.getName()] && !sheet.isSheetHidden();
    }).map(function (sheet) { return sheet.getName(); }).sort();
  }
  function internalCategory_() {
    // Legacy backend still requires categoryName. Never derive an unknown
    // value from user input or silently create a new category.
    var explicit = PropertiesService.getScriptProperties().getProperty('LESSON_GRADER_WEB_CATEGORY');
    if (explicit && String(explicit).trim()) return String(explicit).trim();
    var categories = {};
    var roots = ownRoots_();
    while (roots.hasNext()) {
      var root = roots.next();
      for (var i = 0; i < STAGES.length; i++) {
        var stages = root.getFoldersByName(STAGES[i]);
        while (stages.hasNext()) {
          var courses = stages.next().getFolders();
          while (courses.hasNext()) {
            var folder = courses.next().getFolders();
            while (folder.hasNext()) categories[folder.next().getName()] = true;
          }
        }
      }
    }
    var names = Object.keys(categories);
    if (names.length === 1) return names[0];
    return null;
  }
  function list_() {
    try {
      var seen = {}, result = [];
      walk_(function (file) {
        var id = file.getId();
        if (seen[id] || file.isTrashed()) return;
        seen[id] = true;
        try {
          if (validWorkbook_(SpreadsheetApp.openById(id))) {
            result.push({ id: id, name: file.getName() });
          }
        } catch (err) { console.error('Skipped inaccessible workbook: ' + err); }
      });
      result.sort(function (a, b) {
        var x = a.name.toLowerCase(), y = b.name.toLowerCase();
        return x < y ? -1 : x > y ? 1 : (a.id < b.id ? -1 : a.id > b.id ? 1 : 0);
      });
      return { success: true, data: result };
    } catch (err) {
      console.error('Web workbook list: ' + err);
      return fail_('Unable to list grading workbooks.');
    }
  }
  function metadata_(id) {
    try {
      var ss = open_(id), sheet = ss.getSheetByName(Config.SHEET_SUBMISSIONS);
      var last = sheet.getLastRow(), count = 0;
      if (last > 1) {
        var values = sheet.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < values.length; i++) if (String(values[i][0] || '').trim()) count++;
      }
      return { success: true, data: { id: id, name: ss.getName(), submissionCount: count } };
    } catch (err) {
      console.error('Web workbook metadata: ' + err);
      return fail_('Workbook is invalid, uninitialised, or inaccessible.');
    }
  }
  function rubrics_() {
    try {
      var data = [], seen = {}, roots = ownRoots_();
      while (roots.hasNext()) {
        var shared = roots.next().getFoldersByName('Shared Rubrics');
        while (shared.hasNext()) {
          var files = shared.next().getFilesByType(MimeType.GOOGLE_SHEETS);
          while (files.hasNext()) {
            var f = files.next();
            if (!seen[f.getId()] && !f.isTrashed()) {
              seen[f.getId()] = true;
              data.push({ id: f.getId(), name: f.getName() });
            }
          }
        }
      }
      data.sort(function (a, b) { return a.name.toLowerCase().localeCompare(b.name.toLowerCase()); });
      return { success: true, data: data };
    } catch (err) {
      console.error('Web rubric list: ' + err);
      return fail_('Unable to list shared rubrics.');
    }
  }
  function tasks_(id) {
    try { return { success: true, data: taskNames_(open_(id)) }; }
    catch (err) { console.error('Web task list: ' + err); return fail_('Unable to list tasks for this workbook.'); }
  }

  function ensureRequiredSheet_(ss, sheetName, requiredHeaders) {
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, requiredHeaders.length)
      .setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var lastColumn = sheet.getLastColumn();
  var actual = lastColumn
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
        .map(function (value) { return String(value).trim(); })
    : [];

  // Add missing headers without overwriting existing rows or columns.
  var missing = requiredHeaders.filter(function (name) {
    return actual.indexOf(name) === -1;
  });

  if (missing.length) {
    var firstNewColumn = actual.length + 1;
    var requiredLastColumn = firstNewColumn + missing.length - 1;

    if (sheet.getMaxColumns() < requiredLastColumn) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        requiredLastColumn - sheet.getMaxColumns()
      );
    }

    sheet.getRange(1, firstNewColumn, 1, missing.length)
      .setValues([missing]);
  }

  return sheet;
}

function initialiseWorkbook_(workbookId) {
  var ss = SpreadsheetApp.openById(workbookId);

  ensureRequiredSheet_(
    ss,
    Config.SHEET_SUBMISSIONS,
    SUB_HEADERS
  );

  ensureRequiredSheet_(
    ss,
    Config.SHEET_CLASSROOM_CONFIG,
    CONFIG_HEADERS
  );

  SpreadsheetApp.flush();

  if (!validWorkbook_(ss)) {
    throw new Error(
      'Workbook initialization failed: required sheets or headers are missing.'
    );
  }

  return ss;
}

  function create_(courseCode, taskName, year, rubricId, rubricName, document) {
    var code = String(courseCode || '').trim(), task = String(taskName || '').trim();
    var yr = String(year == null || String(year).trim() === ''
  ? new Date().getFullYear()
  : year).trim();
var category = internalCategory_();
    if (!category) return fail_('Internal category is not configured. Ask an administrator to set the LESSON_GRADER_WEB_CATEGORY script property to the existing category name.');
    if (!deriveStageFromCourseCode(code) || !task || /^\s*$/.test(task) || !/^20\d{2}$/.test(yr)) {
      return fail_('Enter a valid course code, task name and four-digit year.');
    }
    if (rubricId && rubricId !== '__UPLOAD__') {
      var rubrics = rubrics_();
      if (!rubrics.success || !rubrics.data.some(function (r) { return r.id === rubricId; })) return fail_('Select an existing shared rubric.');
    }
    if (rubricId === '__UPLOAD__' && (!String(rubricName || '').trim() || !document)) {
      return fail_('Upload a rubric document and enter its name.');
    }
    try {
      return apiSubmitGradingWorkbookSetup(code, category, task, yr,
        rubricId || '', rubricName || '', document || null);
    } catch (err) {
      console.error('Web create workbook: ' + err);
      return fail_('Workbook setup failed. Check Apps Script executions.');
    }
  }
  function link_(id, taskName, courseId, workId) {
    if (!validId_(id) || !/^\d{1,40}$/.test(String(courseId || '')) || !/^\d{1,40}$/.test(String(workId || ''))) {
      return fail_('Invalid workbook or Classroom selection.');
    }
    var lock = LockService.getScriptLock();
    try {
      var ss = open_(id);
      var task = String(taskName || '').trim();
      if (taskNames_(ss).indexOf(task) === -1) return fail_('Choose a task tab in the selected workbook.');
      var course = Classroom.Courses.get(String(courseId));
      var cw = Classroom.Courses.CourseWork.get(String(courseId), String(workId));
      if (!cw || !cw.dueDate || !cw.dueDate.year || !cw.dueDate.month || !cw.dueDate.day) {
        return fail_('A due date is required for scheduled import.');
      }
      if (cw.courseId && String(cw.courseId) !== String(courseId)) return fail_('Assignment does not belong to the selected course.');
      var time = cw.dueTime || {};
      var due = new Date(Date.UTC(cw.dueDate.year, cw.dueDate.month - 1, cw.dueDate.day,
        time.hours === undefined ? 23 : time.hours,
        time.minutes === undefined ? 59 : time.minutes));
      if (isNaN(due.getTime())) return fail_('Invalid Classroom due date.');
      var sheet = ss.getSheetByName(Config.SHEET_CLASSROOM_CONFIG);
      if (!header_(sheet, CONFIG_HEADERS)) return fail_('Workbook ClassroomConfig schema is incomplete.');
      lock.waitLock(10000);
      var values = sheet.getDataRange().getValues();
      var headers = values[0].map(function (v) { return String(v).trim(); });
      var columns = {};
      CONFIG_HEADERS.forEach(function (h) { columns[h] = headers.indexOf(h); });
      for (var r = 1; r < values.length; r++) {
        if (String(values[r][columns.TaskName]) === task && String(values[r][columns.CourseID]) === String(courseId) &&
            String(values[r][columns.CourseWorkID]) === String(workId)) return fail_('Assignment is already linked to this task.');
      }
      var row = new Array(headers.length).fill('');
      row[columns.ConfigID] = Utilities.getUuid();
      row[columns.CourseID] = String(courseId);
      row[columns.CourseName] = course.name || '';
      row[columns.CourseSection] = course.section || '';
      row[columns.CourseWorkID] = String(workId);
      row[columns.AssignmentTitle] = cw.title || '';
      row[columns.MaxPoints] = cw.maxPoints == null ? 0 : cw.maxPoints;
      row[columns.SavedAt] = new Date().toISOString();
      row[columns.SavedBy] = Session.getEffectiveUser().getEmail() || '';
      row[columns.Active] = true;
      row[columns.TaskName] = task;
      row[columns.DueDate] = due.toISOString();
      row[columns.AutoImported] = false;
      sheet.appendRow(row);
      return { success: true, message: 'Linked “' + (cw.title || 'Untitled assignment') + '” to “' + task + '”. Hourly auto-import trigger must be configured separately.' };
    } catch (err) {
      console.error('Web Classroom link: ' + err);
      return fail_('Unable to link assignment. Check Classroom access and workbook edit permission.');
    } finally {
      try { lock.releaseLock(); } catch (ignore) {}
    }
  }
  return { list: list_, metadata: metadata_, rubrics: rubrics_, tasks: tasks_, create: create_, link: link_ };
})();

function apiWebListCandidateWorkbooks() { return LessonGraderWeb.list(); }
function apiWebGetWorkbookMetadata(spreadsheetId) { return LessonGraderWeb.metadata(spreadsheetId); }
function apiWebListSharedRubrics() { return LessonGraderWeb.rubrics(); }
function apiWebListWorkbookTasks(spreadsheetId) { return LessonGraderWeb.tasks(spreadsheetId); }
function apiWebCreateWorkbook(courseCode, taskName, year, rubricSelection, rubricName, rubricDocument) {
  return LessonGraderWeb.create(courseCode, taskName, year, rubricSelection, rubricName, rubricDocument);
}
function apiWebLinkClassroomAssignment(spreadsheetId, taskName, courseId, courseWorkId) {
  return LessonGraderWeb.link(spreadsheetId, taskName, courseId, courseWorkId);
}

/* Remove this definition if an existing Code.gs already defines doGet. */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('WebApp').setTitle('Lesson Grader');
}

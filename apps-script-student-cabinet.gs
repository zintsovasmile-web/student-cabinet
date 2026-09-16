const SPREADSHEET_ID = '1OwfkLNY33gw1vWAIdQNyekyEVkeAr3G5hfHMBYFkT9M';

const SHEETS = {
  students: 'УЧЕНИКИ',
  lessons: 'ЗАНЯТИЯ',
  personalMaterials: 'ЛИЧНЫЕ МАТЕРИАЛЫ',
  submissions: 'ДЗ СДАЧИ',
  questions: 'ВОПРОСЫ',
};

function doGet(e) {
  const params = (e && e.parameter) || {};

  try {
    if (params.action === 'submitHomework') {
      return respond_(params, submitHomework_(params));
    }

    if (params.action === 'askQuestion') {
      return respond_(params, askQuestion_(params));
    }

    return respond_(params, { ok: false, error: 'Введите код' });
  } catch (err) {
    return respond_(params, {
      ok: false,
      error: err && err.message ? err.message : String(err),
    });
  }
}

function submitHomework_(params) {
  const code = String(params.code || '').trim();
  const lessonId = String(params.lessonId || '').trim();
  if (!code) throw new Error('Не передан код ученика');
  if (!lessonId) throw new Error('Не передан lessonId');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const student = findStudentByCode_(ss, code);
    const lesson = findLesson_(ss, lessonId);
    const workUrl = findHomeworkUrl_(ss, student.id, lessonId);

    upsertSubmission_(ss, {
      timestamp: new Date(),
      lessonId,
      courseId: lesson.group || student.group || '',
      studentId: student.id,
      studentName: student.name,
      status: 'сдано',
      workUrl,
      source: 'сайт',
    });

    return {
      ok: true,
      lessonId,
      studentId: student.id,
      studentName: student.name,
      status: 'сдано',
    };
  } finally {
    lock.releaseLock();
  }
}

function askQuestion_(params) {
  const code = String(params.code || '').trim();
  const lessonId = String(params.lessonId || '').trim();
  const question = String(params.question || '').trim();
  if (!code) throw new Error('Не передан код ученика');
  if (!lessonId) throw new Error('Не передан lessonId');
  if (!question) throw new Error('Вопрос пустой');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const student = findStudentByCode_(ss, code);
  const lesson = findLesson_(ss, lessonId);

  ss.getSheetByName(SHEETS.questions).appendRow([
    new Date(),
    lessonId,
    lesson.group || student.group || '',
    student.id,
    student.name,
    question,
    '',
    'нет',
    '',
    'сайт',
  ]);

  return { ok: true };
}

function findStudentByCode_(ss, code) {
  const sheet = ss.getSheetByName(SHEETS.students);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const studentId = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const group = String(row[2] || '').trim();
    const secretCode = String(row[3] || '').trim();
    const active = String(row[4] || '').trim().toLowerCase();

    if (secretCode === code && active === 'да') {
      return { id: studentId, name, group };
    }
  }

  throw new Error('Код ученика не найден');
}

function findLesson_(ss, lessonId) {
  const sheet = ss.getSheetByName(SHEETS.lessons);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i += 1) {
    const rowLessonId = String(rows[i][0] || '').trim();
    if (rowLessonId === lessonId) {
      return {
        id: rowLessonId,
        group: String(rows[i][1] || '').trim(),
        number: rows[i][2],
        date: rows[i][3],
        topic: String(rows[i][4] || '').trim(),
      };
    }
  }

  return { id: lessonId, group: '' };
}

function findHomeworkUrl_(ss, studentId, lessonId) {
  const sheet = ss.getSheetByName(SHEETS.personalMaterials);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i += 1) {
    const rowLessonId = String(rows[i][0] || '').trim();
    const rowStudentId = String(rows[i][1] || '').trim();
    const type = String(rows[i][3] || '').trim();

    if (rowLessonId === lessonId && rowStudentId === studentId && type === 'ДЗ') {
      return String(rows[i][5] || '').trim();
    }
  }

  return '';
}

function upsertSubmission_(ss, submission) {
  const sheet = ss.getSheetByName(SHEETS.submissions);
  const rows = sheet.getDataRange().getValues();
  let targetRow = 0;

  for (let i = 1; i < rows.length; i += 1) {
    const rowLessonId = String(rows[i][1] || '').trim();
    const rowStudentId = String(rows[i][3] || '').trim();

    if (rowLessonId === submission.lessonId && rowStudentId === submission.studentId) {
      targetRow = i + 1;
      break;
    }
  }

  const leftValues = [[
    submission.timestamp,
    submission.lessonId,
    submission.courseId,
    submission.studentId,
    submission.studentName,
    submission.status,
    submission.workUrl,
  ]];

  if (targetRow) {
    sheet.getRange(targetRow, 1, 1, 7).setValues(leftValues);
    sheet.getRange(targetRow, 10).setValue(submission.source);
    return;
  }

  sheet.appendRow([
    submission.timestamp,
    submission.lessonId,
    submission.courseId,
    submission.studentId,
    submission.studentName,
    submission.status,
    submission.workUrl,
    'нет',
    '',
    submission.source,
  ]);
}

function respond_(params, payload) {
  const json = JSON.stringify(payload);
  const callback = String(params.callback || '').trim();

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

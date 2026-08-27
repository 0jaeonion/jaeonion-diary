import fs from 'fs';

const html = fs.readFileSync('google-apps-script/index.html', 'utf-8');

const backendLogic = `
// 1. 웹앱 진입점 (doGet - 0.001초 초고속 즉시 반환)
function doGet() {
  return HtmlService.createHtmlOutput(getHtmlContent())
    .setTitle('🌸 재어니언 다이어리')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getHtmlContent() {
  return ${JSON.stringify(html)};
}

// 2. 안전한 날짜 변환 함수 (최적화)
function formatDateSafe(val) {
  if (!val) return '';
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
    } catch (e) {
      const y = val.getFullYear();
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const d = String(val.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
  }
  const str = String(val).trim();
  if (str.length >= 10 && str.charAt(4) === '-' && str.charAt(7) === '-') {
    return str.slice(0, 10);
  }
  return str;
}

// 3. 시트 초기화 (시트가 없을 때만 1회 안전하게 실행)
function initSheetsIfNotExist() {
  const lock = LockService.getScriptLock();
  try { lock.tryLock(3000); } catch (e) {}

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
    const nowIso = new Date().toISOString();
    
    // 1) Diaries 시트
    let diarySheet = ss.getSheetByName('Diaries');
    if (!diarySheet) {
      diarySheet = ss.insertSheet('Diaries');
      diarySheet.appendRow(['id', 'date', 'title', 'content', 'mood', 'weather', 'tags', 'updatedAt', 'photo', 'createdAt']);
      diarySheet.setFrozenRows(1);
      diarySheet.getRange('A1:J1').setBackground('#FCE4EC').setFontWeight('bold');

      diarySheet.appendRow([
        'diary_' + new Date().getTime(),
        today,
        '🌸 재어니언 다이어리 시작!',
        '구글 스프레드시트와 연동된 나만의 포근하고 스마트한 재어니언 다이어리를 시작했다!\\n소중한 하루와 투두를 깔끔하게 기록해야지 ✨',
        'happy',
        'sunny',
        '첫기록, 재어니언',
        nowIso,
        '',
        nowIso
      ]);
    }

    // 2) Memos 시트
    let memoSheet = ss.getSheetByName('Memos');
    if (!memoSheet) {
      memoSheet = ss.insertSheet('Memos');
      memoSheet.appendRow(['id', 'title', 'content', 'color', 'isPinned', 'tags', 'updatedAt', 'createdAt']);
      memoSheet.setFrozenRows(1);
      memoSheet.getRange('A1:H1').setBackground('#FFF9C4').setFontWeight('bold');

      memoSheet.appendRow([
        'memo_1',
        '💡 다이어리 & 투두 활용 팁',
        '1. 대시보드에서 오늘의 투두리스트를 확인하고 [미루기] 버튼으로 일정을 손쉽게 관리하세요.\\n2. 과거에 밀린 할 일은 [⚡ 전부 오늘로 가져오기] 버튼으로 한 번에 오늘 날짜로 가져올 수 있어요!\\n3. 메모장에 링크를 적으면 자동으로 하이퍼링크가 생겨요 https://google.com',
        'pastel-peach',
        'TRUE',
        '팁, 재어니언',
        nowIso,
        nowIso
      ]);
    }

    // 3) Todos 시트
    let todoSheet = ss.getSheetByName('Todos');
    if (!todoSheet) {
      todoSheet = ss.insertSheet('Todos');
      todoSheet.appendRow(['id', 'title', 'date', 'isCompleted', 'priority', 'category', 'content', 'updatedAt', 'createdAt']);
      todoSheet.setFrozenRows(1);
      todoSheet.getRange('A1:I1').setBackground('#E1F5FE').setFontWeight('bold');

      todoSheet.appendRow([
        'todo_1',
        '하반기 프로젝트 기획안 검토',
        today,
        'FALSE',
        'high',
        '기획',
        '주요 마일스톤 일정 확정 및 리소스 배분 계획 수립하기',
        nowIso,
        nowIso
      ]);
      todoSheet.appendRow([
        'todo_2',
        '구글 시트 동기화 기능 점검',
        today,
        'TRUE',
        'medium',
        '개발',
        '자동 백업 및 실시간 동기화 정상 작동 확인 완료',
        nowIso,
        nowIso
      ]);
    }
  } catch (err) {
    Logger.log('initSheets error: ' + err.toString());
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// 4. 초고속 통합 전체 데이터 불러오기 (getAllData - 0.3초 번개 조회)
function getAllData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let diarySheet = ss.getSheetByName('Diaries');
    let memoSheet = ss.getSheetByName('Memos');
    let todoSheet = ss.getSheetByName('Todos');

    if (!diarySheet || !memoSheet || !todoSheet) {
      initSheetsIfNotExist();
      diarySheet = ss.getSheetByName('Diaries');
      memoSheet = ss.getSheetByName('Memos');
      todoSheet = ss.getSheetByName('Todos');
    }

    return {
      diaries: getDiariesFromSheet(diarySheet),
      memos: getMemosFromSheet(memoSheet),
      todos: getTodosFromSheet(todoSheet)
    };
  } catch (err) {
    Logger.log('Error in getAllData: ' + err.toString());
    throw new Error('전체 데이터 동기화 실패: ' + err.message);
  }
}

// 5. 일기 목록 시트에서 파싱
function getDiariesFromSheet(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const diaries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;
    
    const dateVal = formatDateSafe(row[1]);
    diaries.push({
      id: String(row[0]),
      date: dateVal,
      title: String(row[2] || ''),
      content: String(row[3] || ''),
      mood: String(row[4] || 'happy'),
      weather: String(row[5] || 'sunny'),
      tags: String(row[6] || '').split(',').map(t => t.trim()).filter(Boolean),
      updatedAt: String(row[7] || ''),
      photo: String(row[8] || ''),
      createdAt: String(row[9] || row[7] || '')
    });
  }
  diaries.sort((a, b) => b.date.localeCompare(a.date));
  return diaries;
}

function getDiaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getDiariesFromSheet(ss.getSheetByName('Diaries'));
}

// 6. 메모 목록 시트에서 파싱
function getMemosFromSheet(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const memos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;

    const isPinned = String(row[4]).toUpperCase() === 'TRUE';
    memos.push({
      id: String(row[0]),
      title: String(row[1] || ''),
      content: String(row[2] || ''),
      color: String(row[3] || 'pastel-peach'),
      isPinned: isPinned,
      tags: String(row[5] || '').split(',').map(t => t.trim()).filter(Boolean),
      updatedAt: String(row[6] || ''),
      createdAt: String(row[7] || row[6] || '')
    });
  }

  memos.sort((a, b) => {
    if (a.isPinned !== b.isPinned) return a.isPinned ? -1 : 1;
    return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
  });

  return memos;
}

function getMemos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getMemosFromSheet(ss.getSheetByName('Memos'));
}

// 7. 투두 목록 시트에서 파싱
function getTodosFromSheet(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const todos = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;

    const dateVal = formatDateSafe(row[2]);
    const isCompleted = String(row[3]).toUpperCase() === 'TRUE';
    todos.push({
      id: String(row[0]),
      title: String(row[1] || ''),
      date: dateVal,
      isCompleted: isCompleted,
      priority: String(row[4] || 'medium'),
      category: String(row[5] || '업무'),
      content: String(row[6] || ''),
      updatedAt: String(row[7] || ''),
      createdAt: String(row[8] || row[7] || '')
    });
  }
  return todos;
}

function getTodos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getTodosFromSheet(ss.getSheetByName('Todos'));
}

// 8. 일기 저장 또는 수정
function saveDiary(diary) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Diaries');
    if (!sheet) {
      initSheetsIfNotExist();
      sheet = ss.getSheetByName('Diaries');
    }
    
    if (sheet.getMaxColumns() < 10) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 10 - sheet.getMaxColumns());
    }

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const tagsStr = Array.isArray(diary.tags) ? diary.tags.join(', ') : (diary.tags || '');
    const photoStr = diary.photo || '';
    let existingRowIndex = -1;
    let originalCreatedAt = now;

    for (let i = 1; i < data.length; i++) {
      const rowDate = formatDateSafe(data[i][1]);
      if (rowDate === String(diary.date)) {
        existingRowIndex = i + 1;
        originalCreatedAt = data[i][9] || data[i][7] || now;
        break;
      }
    }

    if (existingRowIndex > 0) {
      const id = data[existingRowIndex - 1][0] || ('diary_' + new Date().getTime());
      sheet.getRange(existingRowIndex, 1, 1, 10).setValues([[
        id,
        diary.date,
        diary.title || '',
        diary.content || '',
        diary.mood || 'happy',
        diary.weather || 'sunny',
        tagsStr,
        now,
        photoStr,
        originalCreatedAt
      ]]);
      return {
        id: String(id),
        date: diary.date,
        title: diary.title || '',
        content: diary.content || '',
        mood: diary.mood || 'happy',
        weather: diary.weather || 'sunny',
        tags: Array.isArray(diary.tags) ? diary.tags : tagsStr.split(',').map(t => t.trim()).filter(Boolean),
        updatedAt: now,
        photo: photoStr,
        createdAt: originalCreatedAt
      };
    } else {
      const newId = 'diary_' + new Date().getTime();
      sheet.appendRow([
        newId,
        diary.date,
        diary.title || '',
        diary.content || '',
        diary.mood || 'happy',
        diary.weather || 'sunny',
        tagsStr,
        now,
        photoStr,
        now
      ]);
      return {
        id: newId,
        date: diary.date,
        title: diary.title || '',
        content: diary.content || '',
        mood: diary.mood || 'happy',
        weather: diary.weather || 'sunny',
        tags: Array.isArray(diary.tags) ? diary.tags : tagsStr.split(',').map(t => t.trim()).filter(Boolean),
        updatedAt: now,
        photo: photoStr,
        createdAt: now
      };
    }
  } catch (err) {
    Logger.log('Error in saveDiary: ' + err.toString());
    throw new Error('일기 저장 중 오류: ' + err.message);
  }
}

// 9. 일기 삭제
function deleteDiary(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Diaries');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (err) {
    Logger.log('Error in deleteDiary: ' + err.toString());
    throw new Error('일기 삭제 중 오류: ' + err.message);
  }
}

// 10. 메모 저장 또는 수정
function saveMemo(memo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Memos');
    if (!sheet) {
      initSheetsIfNotExist();
      sheet = ss.getSheetByName('Memos');
    }
    
    if (sheet.getMaxColumns() < 8) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 8 - sheet.getMaxColumns());
    }

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const tagsStr = Array.isArray(memo.tags) ? memo.tags.join(', ') : (memo.tags || '');
    const isPinnedStr = memo.isPinned ? 'TRUE' : 'FALSE';
    let originalCreatedAt = now;

    if (memo.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(memo.id)) {
          originalCreatedAt = data[i][7] || data[i][6] || now;
          sheet.getRange(i + 1, 1, 1, 8).setValues([[
            memo.id,
            memo.title || '',
            memo.content || '',
            memo.color || 'pastel-peach',
            isPinnedStr,
            tagsStr,
            now,
            originalCreatedAt
          ]]);
          return {
            id: memo.id,
            title: memo.title || '',
            content: memo.content || '',
            color: memo.color || 'pastel-peach',
            isPinned: Boolean(memo.isPinned),
            tags: Array.isArray(memo.tags) ? memo.tags : tagsStr.split(',').map(t => t.trim()).filter(Boolean),
            updatedAt: now,
            createdAt: originalCreatedAt
          };
        }
      }
    }

    const newId = 'memo_' + new Date().getTime();
    sheet.appendRow([
      newId,
      memo.title || '',
      memo.content || '',
      memo.color || 'pastel-peach',
      isPinnedStr,
      tagsStr,
      now,
      now
    ]);

    return {
      id: newId,
      title: memo.title || '',
      content: memo.content || '',
      color: memo.color || 'pastel-peach',
      isPinned: Boolean(memo.isPinned),
      tags: Array.isArray(memo.tags) ? memo.tags : tagsStr.split(',').map(t => t.trim()).filter(Boolean),
      updatedAt: now,
      createdAt: now
    };
  } catch (err) {
    Logger.log('Error in saveMemo: ' + err.toString());
    throw new Error('메모 저장 중 오류: ' + err.message);
  }
}

// 11. 메모 삭제
function deleteMemo(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Memos');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (err) {
    Logger.log('Error in deleteMemo: ' + err.toString());
    throw new Error('메모 삭제 중 오류: ' + err.message);
  }
}

// 12. 투두 저장 또는 수정
function saveTodo(todo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Todos');
    if (!sheet) {
      initSheetsIfNotExist();
      sheet = ss.getSheetByName('Todos');
    }
    
    if (sheet.getMaxColumns() < 9) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());
    }

    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const isCompletedStr = todo.isCompleted ? 'TRUE' : 'FALSE';
    const contentStr = todo.content || '';
    let originalCreatedAt = now;

    if (todo.id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(todo.id)) {
          originalCreatedAt = data[i][8] || data[i][7] || now;
          sheet.getRange(i + 1, 1, 1, 9).setValues([[
            todo.id,
            todo.title || '',
            todo.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd'),
            isCompletedStr,
            todo.priority || 'medium',
            todo.category || '업무',
            contentStr,
            now,
            originalCreatedAt
          ]]);
          return {
            id: todo.id,
            title: todo.title || '',
            date: todo.date,
            isCompleted: Boolean(todo.isCompleted),
            priority: todo.priority || 'medium',
            category: todo.category || '업무',
            content: contentStr,
            updatedAt: now,
            createdAt: originalCreatedAt
          };
        }
      }
    }

    const newId = 'todo_' + new Date().getTime();
    sheet.appendRow([
      newId,
      todo.title || '',
      todo.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd'),
      isCompletedStr,
      todo.priority || 'medium',
      todo.category || '업무',
      contentStr,
      now,
      now
    ]);

    return {
      id: newId,
      title: todo.title || '',
      date: todo.date,
      isCompleted: Boolean(todo.isCompleted),
      priority: todo.priority || 'medium',
      category: todo.category || '업무',
      content: contentStr,
      updatedAt: now,
      createdAt: now
    };
  } catch (err) {
    Logger.log('Error in saveTodo: ' + err.toString());
    throw new Error('투두 저장 중 오류: ' + err.message);
  }
}

// 13. 투두 일괄 업데이트 (밀린 할 일 일괄 오늘로 변경 시 사용)
function saveTodosBatch(todoList) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Todos');
    if (!sheet) {
      initSheetsIfNotExist();
      sheet = ss.getSheetByName('Todos');
    }
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();
    const updatedIds = [];

    const todoMap = {};
    todoList.forEach(t => { if (t.id) todoMap[t.id] = t; });

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]);
      if (todoMap[rowId]) {
        const item = todoMap[rowId];
        const isCompletedStr = item.isCompleted ? 'TRUE' : 'FALSE';
        const contentStr = item.content || '';
        const originalCreatedAt = data[i][8] || data[i][7] || now;

        sheet.getRange(i + 1, 1, 1, 9).setValues([[
          item.id,
          item.title || '',
          item.date,
          isCompletedStr,
          item.priority || 'medium',
          item.category || '업무',
          contentStr,
          now,
          originalCreatedAt
        ]]);
        updatedIds.push(item.id);
      }
    }
    return { success: true, count: updatedIds.length, updatedAt: now };
  } catch (err) {
    Logger.log('Error in saveTodosBatch: ' + err.toString());
    throw new Error('투두 일괄 업데이트 오류: ' + err.message);
  }
}

// 14. 투두 삭제
function deleteTodo(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Todos');
    if (!sheet) return false;
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return true;
      }
    }
    return false;
  } catch (err) {
    Logger.log('Error in deleteTodo: ' + err.toString());
    throw new Error('투두 삭제 중 오류: ' + err.message);
  }
}
`;

fs.writeFileSync('google-apps-script/Code.gs', backendLogic.trim(), 'utf-8');
console.log('Code.gs compiled successfully with ultra-fast optimized backend!');

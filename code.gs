const SS_ID = ""; // 비워두면 현재 스프레드시트 사용

function getActiveSheet_() {
  if (SS_ID) {
    return SpreadsheetApp.openById(SS_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === 'Teachers') {
      sheet.appendRow(['className', 'password']);
    } else if (sheetName === 'Articles') {
      sheet.appendRow(['id', 'timestamp', 'targetTeacher', 'name', 'mood', 'shortMsg', 'activityTitle', 'w6', 'draftText']);
    }
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter ? e.parameter.action : '';
  const ss = getActiveSheet_();
  
  try {
    if (action === 'getData') {
      const sheet = getOrCreateSheet_(ss, 'Articles');
      const rows = sheet.getDataRange().getValues();
      const articles = [];
      for (let i = 1; i < rows.length; i++) {
        if (!rows[i][0]) continue;
        let w6Obj = {};
        try { w6Obj = JSON.parse(rows[i][7] || '{}'); } catch(err) {}
        
        articles.push({
          id: rows[i][0],
          timestamp: rows[i][1],
          targetTeacher: String(rows[i][2]).trim(),
          name: rows[i][3],
          mood: rows[i][4],
          shortMsg: rows[i][5],
          activityTitle: rows[i][6],
          w6: w6Obj,
          draftText: rows[i][8]
        });
      }
      return createJsonResponse({ success: true, data: articles });
      
    } else if (action === 'clearData') {
      const teacher = String(e.parameter.teacher || '').trim();
      const sheet = getOrCreateSheet_(ss, 'Articles');
      const rows = sheet.getDataRange().getValues();
      
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][2]).trim() === teacher) {
          sheet.deleteRow(i + 1);
        }
      }
      return createJsonResponse({ success: true, message: `[${teacher}] 반의 데이터가 초기화되었습니다.` });
    }
    
    return createJsonResponse({ success: false, message: '잘못된 요청입니다.' });
  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseErr) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    const action = data.action;
    const ss = getActiveSheet_();
    
    if (action === 'registerTeacher') {
      const className = String(data.className || '').trim();
      const password = String(data.password || '').trim();
      
      if (!className || !password) {
        return createJsonResponse({ success: false, message: '반 이름과 비밀번호를 입력해주세요.' });
      }

      const sheet = getOrCreateSheet_(ss, 'Teachers');
      const rows = sheet.getDataRange().getValues();
      
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === className) {
          return createJsonResponse({ success: false, message: `이미 등록된 반 이름입니다: [${className}]` });
        }
      }
      
      sheet.appendRow([className, password]);
      return createJsonResponse({ success: true, message: '반이 성공적으로 생성되었습니다.' });
      
    } else if (action === 'loginTeacher') {
      const className = String(data.className || '').trim();
      const password = String(data.password || '').trim();
      
      const sheet = getOrCreateSheet_(ss, 'Teachers');
      const rows = sheet.getDataRange().getValues();
      
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === className && String(rows[i][1]).trim() === password) {
          return createJsonResponse({ success: true });
        }
      }
      return createJsonResponse({ success: false, message: '비밀번호가 일치하지 않습니다.' });
      
    } else if (action === 'saveArticle') {
      const sheet = getOrCreateSheet_(ss, 'Articles');
      const id = 'art_' + new Date().getTime() + Math.random().toString(36).substring(2, 7);
      const timestamp = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
      
      sheet.appendRow([
        id,
        timestamp,
        String(data.targetTeacher || '').trim(),
        data.name || '',
        data.mood || '',
        data.shortMsg || '',
        data.activityTitle || '',
        JSON.stringify(data.w6 || {}),
        data.draftText || ''
      ]);
      
      return createJsonResponse({ success: true, message: '🎉 기사가 성공적으로 제출되었습니다!' });
    }
    
    return createJsonResponse({ success: false, message: '지원하지 않는 액션입니다.' });
  } catch (err) {
    return createJsonResponse({ success: false, message: err.toString() });
  }
}

function createJsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
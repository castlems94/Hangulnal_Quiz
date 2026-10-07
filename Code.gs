/**
 * 세종대왕의 한글날 퀴즈 - 구글 시트 기록용 Apps Script
 *
 * [시트 구조]
 *   A1 = 이름 / B1~P1 = 문제1~문제15 / Q1 = 맞춘 개수
 *   기록은 2행부터 한 줄씩 쌓입니다.
 *
 * [설정 방법]
 * 1. 위 구조로 만든 구글 시트에서 [확장 프로그램] > [Apps Script] 를 열고, 이 코드를 전부 붙여넣는다.
 * 2. [배포] > [새 배포] > 유형 '웹 앱'
 *      - 다음 사용자 인증으로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 * 3. 처음 배포할 때 나오는 권한 승인을 허용한다.
 * 4. 발급된 '웹 앱 URL'(…/exec)을 index.html 맨 위의 SHEET_URL 에 붙여넣는다.
 *
 * ※ 코드를 수정한 뒤에는 [배포 관리] > 수정(연필) > 버전 '새 버전' 으로 다시 배포해야 반영됩니다.
 */

const SHEET_NAME = '';   // 기록할 탭 이름. 비워 두면 이 스프레드시트의 첫 번째 탭에 기록합니다.
const TOTAL = 15;        // 문제 수

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SHEET_NAME) return ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  return ss.getSheets()[0];
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const data = JSON.parse(e.postData.contents);
    const sh = getSheet_();

    // 1행(머리글)이 비어 있을 때만 만들어 줍니다. 이미 있으면 건드리지 않습니다.
    if (String(sh.getRange(1, 1).getValue()).trim() === '') {
      const header = ['이름'];
      for (let i = 1; i <= TOTAL; i++) header.push('문제' + i);
      header.push('맞춘 개수');
      sh.getRange(1, 1, 1, header.length).setValues([header]);
    }

    // 이름이 수식으로 해석되지 않도록 처리
    let name = String(data.name || '').trim().slice(0, 20);
    if (/^[=+\-@]/.test(name)) name = "'" + name;

    const results = (data.results || []).slice(0, TOTAL).map(r => (r === 'O' ? 'O' : 'X'));
    while (results.length < TOTAL) results.push('');
    const score = results.filter(r => r === 'O').length;

    // A열에서 마지막으로 이름이 적힌 행을 찾아 그 다음 행(최소 2행)에 기록
    const colA = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
    let lastNameRow = 1;
    for (let i = colA.length - 1; i >= 1; i--) {
      if (String(colA[i][0]).trim() !== '') { lastNameRow = i + 1; break; }
    }
    const row = Math.max(2, lastNameRow + 1);

    sh.getRange(row, 1, 1, TOTAL + 2).setValues([[name].concat(results, [score])]);

    // 보기 좋게 정렬 + O/X 색 표시
    sh.getRange(row, 2, 1, TOTAL + 1).setHorizontalAlignment('center');
    for (let i = 0; i < TOTAL; i++) {
      sh.getRange(row, 2 + i)
        .setFontColor(results[i] === 'O' ? '#2e8b57' : '#c0392b')
        .setFontWeight('bold');
    }

    return json({ ok: true, row: row });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 주소를 브라우저에서 열어 동작 확인용
function doGet() {
  return json({ ok: true, message: '한글날 퀴즈 기록 서버가 작동 중입니다.' });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

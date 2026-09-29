/**
 * 프리미엘네일 VIP 회원권 계약서 저장용 Apps Script
 *
 * 계약서 페이지(index.html)에서 제출한 내용을 받아
 *  1) '계약내역' 시트에 한 줄 추가하고
 *  2) 서명이 들어간 계약서 PDF와 서명 이미지를 드라이브 폴더에 저장합니다.
 *
 * 붙여넣는 곳: 스프레드시트 > 확장 프로그램 > Apps Script
 * 배포: 배포 > 새 배포 > 유형 '웹 앱' > 실행 계정 '나' > 액세스 권한 '모든 사용자'
 */

const SHEET_ID = '1HYnOtU4Y9hXjP0z16Jy6b-vUbwAnHuNRCKDmsKTvONo';
const SHEET_NAME = '계약내역';
const FOLDER_NAME = '프리미엘네일 계약서';
const TOKEN = 'premiel-vip-2026';   // index.html 의 API.token 과 같아야 합니다
const ID_COL = 18;                  // R열: 계약ID
const PDF_COL = 19;                 // S열: 계약서 PDF

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== TOKEN) return json_({ ok: false, error: 'unauthorized' });
    const r = body.record || {};
    if (!r.id || !r.name || !r.phone || !r.sigStaff || !r.sigMember) {
      return json_({ ok: false, error: 'invalid' });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

      // 같은 계약서를 두 번 보내도 한 번만 저장
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(2, ID_COL, last - 1, 1).getValues().map(function (v) { return v[0]; });
        const idx = ids.indexOf(r.id);
        if (idx >= 0) {
          return json_({ ok: true, id: r.id, duplicate: true, pdfUrl: sh.getRange(idx + 2, PDF_COL).getValue() });
        }
      }

      const folder = folder_();
      const base = '프리미엘네일_계약서_' + r.name + '_' + String(r.signDate || '').replace(/-/g, '') + '_' + r.id;
      const staffUrl = saveImage_(folder, r.sigStaff, base + '_담당자확인.png');
      const memberUrl = saveImage_(folder, r.sigMember, base + '_회원서명.png');
      let pdfUrl = '';
      try {
        const pdf = HtmlService.createHtmlOutput(body.html || '').getBlob().getAs('application/pdf').setName(base + '.pdf');
        pdfUrl = folder.createFile(pdf).getUrl();
      } catch (err) {
        pdfUrl = 'PDF 생성 실패: ' + err;
      }

      const bonus = [];
      if (r.bonusVisit) bonus.push('3회 방문 내 등록 +1만원');
      if (r.bonusRenew) bonus.push('연장 +1만원');

      sh.appendRow([
        Utilities.formatDate(new Date(r.createdAt || Date.now()), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'),
        r.signDate, r.name, String(r.birth || ''), r.phone, r.kind, r.planLabel,
        r.method === 'cash' ? '현금·이체' : '카드',
        Number(r.price) || 0, Number(r.credit) || 0, bonus.join(', '),
        r.startDate, r.endDate,
        r.agreeTerms ? '동의' : '미동의', r.agreePrivacy ? '동의' : '미동의', r.agreeMarketing ? '동의' : '미동의',
        r.termsVersion || '', r.id,
        pdfUrl, staffUrl, memberUrl
      ]);
      return json_({ ok: true, id: r.id, pdfUrl: pdfUrl });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return ContentService.createTextOutput('프리미엘네일 계약서 저장 서버가 동작 중입니다.');
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function saveImage_(folder, dataUrl, name) {
  const m = /^data:image\/png;base64,(.+)$/.exec(dataUrl || '');
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[1]), 'image/png', name);
  return folder.createFile(blob).getUrl();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 처음 한 번 실행해서 드라이브·시트 권한을 허용해 주세요. (편집기 상단에서 'authorize' 선택 후 실행) */
function authorize() {
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).getRange('A1').getValue();
  folder_();
}

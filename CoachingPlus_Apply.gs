const SHEET_NAME  = 'raw';
const ADMIN_EMAIL = 'discplus2010@naver.com';

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);

        // honeypot 차단
        if (data._honey) {
            return ContentService
                .createTextOutput(JSON.stringify({ result: 'blocked' }))
                .setMimeType(ContentService.MimeType.JSON);
        }

        const email = String(data.email || '').trim();
        const tel   = String(data.tel || '').replace(/\D/g, '');

        const emailOk = /^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)+$/.test(email);
        const telOk = tel.length >= 8;

        if (!emailOk || !telOk) {
            return ContentService
                .createTextOutput(JSON.stringify({ result: 'invalid_contact' }))
                .setMimeType(ContentService.MimeType.JSON);
        }

        const ss    = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName(SHEET_NAME)
                     || ss.insertSheet(SHEET_NAME);

        // 헤더가 없으면 생성
        if (sheet.getLastRow() === 0) {
            sheet.appendRow([
                '고유번호', '신청일시',
                '성함', '소속/직함', '이메일', '전화번호',
                '유형', '선택항목', '코칭상세', '메모', '상태'
            ]);
            sheet.setFrozenRows(1);
        }

        // 고유번호 생성: CP-YYYYMMDD-NNN
        const today   = Utilities.formatDate(
                            new Date(), 'Asia/Seoul', 'yyyyMMdd');
        const lastRow = sheet.getLastRow();
        const seq     = String(lastRow).padStart(3, '0');
        const uid     = 'CP-' + today + '-' + seq;

        // 신청일시
        const timestamp = Utilities.formatDate(
                            new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

        // 유형 한글 변환
        const typeMap = {
            program:  '정규 프로그램',
            lecture:  '특강',
            coaching: '개인 코칭·상담'
        };
        const typeKr = typeMap[data.type] || data.type;

        // 시트에 기록
        sheet.appendRow([
            uid, timestamp,
            data.name    || '',
            data.org     || '',
            email,
            tel,
            typeKr,
            data.selection      || '',
            data.coachingDetail || '',
            data.memo           || '',
            '미확인'   // 상태 초기값 (수동 관리)
        ]);

        // 관리자 이메일 발송
        const subject = `[코칭플러스 신청] ${uid} — ${data.name} (${typeKr})`;
        const body = [
            `■ 신청번호: ${uid}`,
            `■ 신청일시: ${timestamp}`,
            ``,
            `■ 성함: ${data.name}`,
            `■ 소속: ${data.org || '—'}`,
            `■ 이메일: ${email}`,
            `■ 전화번호: ${tel}`,
            ``,
            `■ 유형: ${typeKr}`,
            `■ 선택항목: ${data.selection || '—'}`,
            data.coachingDetail
                ? `■ 코칭 상세: ${data.coachingDetail}`
                : '',
            ``,
            `■ 추가 메모: ${data.memo || '—'}`,
            ``,
            `─────────────────────`,
            `스프레드시트에서 확인: ${ss.getUrl()}`
        ].filter(Boolean).join('\n');

        MailApp.sendEmail(ADMIN_EMAIL, subject, body);

        return ContentService
            .createTextOutput(JSON.stringify({ result: 'ok', uid }))
            .setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService
            .createTextOutput(JSON.stringify({ result: 'error', msg: err.message }))
            .setMimeType(ContentService.MimeType.JSON);
    }
}

/**
 * ============================================================================
 * 🌸 포항 향수 공방 「윤슬향방」 사장님 비서 - Google Apps Script (Code.gs)
 * ============================================================================
 * [보안 안내]
 * Gemini API 키는 코드에 직접 적지 않고, Apps Script 좌측 메뉴의
 * [프로젝트 설정(⚙️) > 스크립트 속성]에 아래 이름으로 안전하게 저장해 사용합니다.
 *   - 속성(Property): GEMINI_API_KEY
 *   - 값(Value): 발급받은 Gemini API 키
 * (이렇게 하면 이 폴더를 그대로 GitHub에 올려도 API 키가 절대 유출되지 않습니다!)
 */

// 1. 접속을 허용할 사장님과 스태프의 구글 이메일 주소를 적어주세요.
//    (여기에 적힌 구글 계정으로 로그인한 사람만 분석 및 장부 저장이 가능합니다)
const ALLOWED_EMAILS = [
  'owner@gmail.com',   // 사장님 구글 이메일로 변경하세요
  'staff@gmail.com'    // 스태프 구글 이메일로 변경하세요 (더 추가하려면 쉼표 찍고 아래에 추가)
];

// 2. 구글 시트 하단의 탭 이름 (기본값: '윤슬향방 매출 장부')
//    해당 이름의 시트 탭이 없으면 자동으로 만들어 줍니다.
const SHEET_NAME = '윤슬향방 매출 장부';

// (선택) Google Cloud에서 발급받은 웹 클라이언트 ID를 넣으면 보안을 한층 더 강화합니다.
// 스크립트 속성(GOOGLE_CLIENT_ID)에 넣거나 아래에 적어두셔도 됩니다.
const GOOGLE_CLIENT_ID = '';


/**
 * ============================================================================
 * 🛠️ 처음 한 번만 실행하는 초기 설정 함수
 * Apps Script 상단 메뉴에서 [초기설정_및_장부만들기]를 선택하고 [▷ 실행]을 눌러주세요.
 * ============================================================================
 */
function 초기설정_및_장부만들기() {
  const sheet = getOrCreateLedgerSheet_();
  Logger.log('✅ 「' + sheet.getName() + '」 시트 준비가 완료되었습니다!');

  const apiKey = getGeminiApiKey_();
  if (apiKey) {
    Logger.log('✅ 스크립트 속성(GEMINI_API_KEY)이 정상적으로 등록되어 있습니다.');
  } else {
    Logger.log('⚠️ 아직 스크립트 속성에 GEMINI_API_KEY가 없습니다. 좌측 [프로젝트 설정(⚙️) > 스크립트 속성]에서 추가해 주세요.');
  }
}


/**
 * 웹 앱 상태 확인용 (브라우저에서 웹 앱 주소로 바로 접속했을 때 안내 문구 표시)
 */
function doGet(e) {
  const output = {
    ok: true,
    service: '윤슬향방 사장님 비서 웹 앱 창구',
    message: '정상 작동 중입니다. index.html 화면에서 연결해 사용해 주세요.'
  };
  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * index.html 화면에서 보내는 모든 요청(로그인 확인, AI 음성 정리, 시트 저장) 처리
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: '요청 데이터가 비어 있습니다.' });
    }

    const body = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const idToken = body.idToken || '';

    // 1단계: 구글 로그인 토큰 검증 및 사장님/스태프 이메일 승인 여부 확인
    const auth = verifyGoogleUser_(idToken);
    if (!auth.ok) {
      return jsonResponse_({
        ok: false,
        code: 'UNAUTHORIZED',
        email: auth.email || '',
        error: auth.error
      });
    }

    // 2단계: 요청 종류(action)에 따라 처리
    if (action === 'auth_check') {
      // 로그인 직후 승인된 계정인지 확인 + 최근 저장 내역 5건 함께 반환
      const recent = getRecentRows_(5);
      return jsonResponse_({
        ok: true,
        user: {
          email: auth.email,
          name: auth.name,
          picture: auth.picture
        },
        recentRows: recent
      });
    }

    if (action === 'parse') {
      // 말한 문장을 날짜·구분·인원·금액·결제수단·유입경로·비고 7칸으로 정리
      const text = (body.text || '').trim();
      if (!text) {
        return jsonResponse_({ ok: false, error: '정리할 내용을 말씀해 주세요.' });
      }
      const parsedResult = parseSalesText_(text);
      return jsonResponse_({
        ok: true,
        data: parsedResult.data,
        aiUsed: parsedResult.aiUsed,
        notice: parsedResult.notice || ''
      });
    }

    if (action === 'save') {
      // 확인을 마친 7개 항목을 구글 시트에 한 줄로 저장
      const item = body.item || {};
      const savedRow = appendRowToSheet_(item, auth.email);
      const recent = getRecentRows_(5);
      return jsonResponse_({
        ok: true,
        message: '「윤슬향방 매출 장부」에 한 줄 저장되었습니다!',
        savedRow: savedRow,
        recentRows: recent
      });
    }

    return jsonResponse_({ ok: false, error: '알 수 없는 요청입니다: ' + action });

  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: '서버 처리 중 오류가 발생했습니다: ' + (err && err.message ? err.message : String(err))
    });
  }
}


/**
 * ============================================================================
 * 🔐 구글 로그인 검증 및 지정된 계정(사장님·스태프) 확인
 * ============================================================================
 */
function verifyGoogleUser_(idToken) {
  if (!idToken) {
    return {
      ok: false,
      error: '구글 로그인이 필요합니다. 화면에서 [Google 계정으로 로그인]을 먼저 진행해 주세요.'
    };
  }

  // Google 공식 tokeninfo 엔드포인트로 ID 토큰 진위 여부 검증
  const verifyUrl = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const res = UrlFetchApp.fetch(verifyUrl, { muteHttpExceptions: true });

  if (res.getResponseCode() !== 200) {
    return {
      ok: false,
      error: '구글 로그인 세션이 만료되었거나 올바르지 않습니다. 다시 로그인해 주세요.'
    };
  }

  const payload = JSON.parse(res.getContentText());
  const email = String(payload.email || '').trim().toLowerCase();
  const emailVerified = String(payload.email_verified) === 'true';

  if (!email || !emailVerified) {
    return {
      ok: false,
      error: '이메일 인증이 완료되지 않은 구글 계정입니다.'
    };
  }

  // GOOGLE_CLIENT_ID를 설정해 둔 경우(스크립트 속성 또는 상단 변수) 클라이언트 ID 일치 여부도 확인
  const expectedClientId = (
    PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID') ||
    GOOGLE_CLIENT_ID ||
    ''
  ).trim();
  if (expectedClientId !== '') {
    if (payload.aud !== expectedClientId) {
      return {
        ok: false,
        email: email,
        error: '지정된 앱(Client ID)에서 발급된 로그인 토큰이 아닙니다.'
      };
    }
  }

  // 허용된 이메일 목록 가져오기 (Code.gs 상단 ALLOWED_EMAILS + 시트 '승인계정' 탭이 있으면 합산)
  const allowedList = getAllowedEmails_();

  // 아직 기본 예시 이메일('owner@gmail.com')만 있고 실제 이메일을 안 넣으신 경우 친절하게 안내
  const isDefaultOnly = allowedList.every(function(e) {
    return e === 'owner@gmail.com' || e === 'staff@gmail.com';
  });

  if (!allowedList.includes(email)) {
    if (isDefaultOnly) {
      return {
        ok: false,
        email: email,
        error: '아직 Code.gs의 ALLOWED_EMAILS에 사장님 이메일이 등록되지 않았습니다. Code.gs 상단 ALLOWED_EMAILS에 [\'' + email + '\']을 넣고 새 배포를 해 주세요!'
      };
    }
    return {
      ok: false,
      email: email,
      error: '승인되지 않은 구글 계정(' + email + ')입니다. 사장님 또는 지정된 스태프 계정으로만 접속할 수 있습니다.'
    };
  }

  return {
    ok: true,
    email: email,
    name: payload.name || email.split('@')[0],
    picture: payload.picture || ''
  };
}


/**
 * 허용된 이메일 목록 불러오기 (Code.gs 상단 배열 + 시트 내 '승인계정' 탭 지원)
 */
function getAllowedEmails_() {
  const list = (ALLOWED_EMAILS || [])
    .map(function(e) { return String(e || '').trim().toLowerCase(); })
    .filter(Boolean);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      const authSheet = ss.getSheetByName('승인계정');
      if (authSheet && authSheet.getLastRow() >= 1) {
        const values = authSheet.getRange(1, 1, authSheet.getLastRow(), 1).getValues();
        values.forEach(function(row) {
          const cell = String(row[0] || '').trim().toLowerCase();
          if (cell && cell.indexOf('@') !== -1 && list.indexOf(cell) === -1) {
            list.push(cell);
          }
        });
      }
    }
  } catch (ignore) {}

  return list;
}


/**
 * 스크립트 속성(Script Properties)에서 GEMINI_API_KEY 불러오기
 */
function getGeminiApiKey_() {
  const propKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return String(propKey || '').trim();
}


/**
 * ============================================================================
 * 🤖 Gemini AI로 음성/텍스트 내용을 7개 항목으로 정리 (실패 시 똑똑한 규칙 방식 자동 작동)
 * ============================================================================
 */
function parseSalesText_(text) {
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  const apiKey = getGeminiApiKey_();

  // 스크립트 속성에 API 키가 아직 입력되지 않은 경우 규칙 기반 분석기로 바로 처리
  if (!apiKey) {
    return {
      aiUsed: false,
      notice: '스크립트 속성(GEMINI_API_KEY) 등록 전이어서 기본 똑똑한 규칙 모드로 정리했습니다.',
      data: parseByRules_(text, todayStr)
    };
  }

  const prompt = [
    '당신은 포항 향수 공방 "윤슬향방"의 매출 장부 정리 비서입니다.',
    '사장님이나 스태프가 수업/판매 후 말한 문장을 아래 7개 항목의 JSON으로 정확히 정리하세요.',
    '',
    '[기준 정보]',
    '- 오늘 날짜(한국 시간): ' + todayStr,
    '- "오늘" 또는 날짜 언급이 없으면 "' + todayStr + '"로 설정하세요. "어제", "그제"나 특정 날짜(예: 10월 25일)가 있으면 YYYY-MM-DD 형식으로 계산하세요.',
    '',
    '[7개 항목 규칙]',
    '1. date (문자열): YYYY-MM-DD 형식의 날짜',
    '2. category (문자열): 반드시 다음 4가지 중 가장 알맞은 하나로 분류하세요.',
    '   - "원데이 클래스" (예: 원데이, 일일체험, 향수 만들기, 커플 수업, 키즈 수업 등)',
    '   - "취미·자격증반" (예: 취미반, 자격증반, 조향사 과정, 전문가반, 정규반 등)',
    '   - "완제품·답례품 판매" (예: 향수 판매, 디퓨저, 룸스프레이, 답례품, 완제품, 선물세트 등)',
    '   - "기타" (위 3가지에 해당하지 않는 경우)',
    '3. people (정수): 인원수(숫자만). 언급이 없으면 1, "커플/두분"이면 2.',
    '4. amount (정수): 총 결제 금액(원 단위 숫자만). 예: "7만원" -> 70000, "3만5천원" -> 35000. 만약 "1인 35000원 2명"처럼 단가만 말했으면 곱해서 총액으로 계산.',
    '5. payment (문자열): 결제수단. 다음 중 하나로 표준화 권장: "카드", "계좌이체", "현금", "네이버페이", "카카오페이", "포항사랑상품권", "기타". 언급 없으면 "카드".',
    '6. channel (문자열): 유입경로. 다음 중 하나로 표준화 권장: "인스타그램", "네이버 예약", "네이버 검색/블로그", "지인 추천", "재방문", "워크인(지나가다)", "당근", "기타". (예: "인스타 보고 옴" -> "인스타그램"). 언급 없으면 "기타".',
    '7. memo (문자열): 위 6개 항목 외에 남아있는 특이사항, 향 취향, 고객 메모, 예약자명 등이 있으면 간단히 적고, 없으면 빈 문자열("")로 두세요.',
    '',
    '[사장님/스태프 입력 문장]',
    '"' + text + '"'
  ].join('\n');

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + encodeURIComponent(apiKey);
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const result = JSON.parse(response.getContentText());
      const rawText = result.candidates[0].content.parts[0].text;
      const parsed = JSON.parse(rawText);
      return {
        aiUsed: true,
        data: normalizeParsedItem_(parsed, todayStr)
      };
    } else {
      // 키 오류 또는 할당량 초과 시 규칙 기반으로 안전하게 대체
      return {
        aiUsed: false,
        notice: 'Gemini 응답 지연으로 기본 규칙 모드로 정리했습니다.',
        data: parseByRules_(text, todayStr)
      };
    }
  } catch (err) {
    return {
      aiUsed: false,
      notice: 'AI 연결 대신 기본 규칙 모드로 정리했습니다.',
      data: parseByRules_(text, todayStr)
    };
  }
}


/**
 * AI가 반환한 값을 안전하게 다듬는 보조 함수
 */
function normalizeParsedItem_(item, todayStr) {
  const validCategories = ['원데이 클래스', '취미·자격증반', '완제품·답례품 판매', '기타'];
  let category = String(item.category || '원데이 클래스').trim();
  if (validCategories.indexOf(category) === -1) {
    if (category.indexOf('원데이') !== -1) category = '원데이 클래스';
    else if (category.indexOf('취미') !== -1 || category.indexOf('자격증') !== -1) category = '취미·자격증반';
    else if (category.indexOf('완제품') !== -1 || category.indexOf('답례품') !== -1 || category.indexOf('판매') !== -1) category = '완제품·답례품 판매';
    else category = '기타';
  }

  const people = Math.max(1, parseInt(item.people, 10) || 1);
  const amount = Math.max(0, parseInt(String(item.amount || '0').replace(/[^0-9]/g, ''), 10) || 0);

  return {
    date: /^\d{4}-\d{2}-\d{2}$/.test(String(item.date || '')) ? String(item.date) : todayStr,
    category: category,
    people: people,
    amount: amount,
    payment: String(item.payment || '카드').trim(),
    channel: String(item.channel || '기타').trim(),
    memo: String(item.memo || '').trim()
  };
}


/**
 * API 키가 없거나 일시적 오류일 때도 한국어 문장을 정확히 분석하는 보조 규칙 분석기
 */
function parseByRules_(text, todayStr) {
  let date = todayStr;
  const now = new Date();
  if (text.indexOf('어제') !== -1) {
    const d = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    date = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  } else if (text.indexOf('그제') !== -1 || text.indexOf('그저께') !== -1) {
    const d = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
    date = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }

  // 구분 추출
  let category = '원데이 클래스';
  if (/취미|자격증|전문가|정규|조향사/.test(text)) {
    category = '취미·자격증반';
  } else if (/완제품|답례품|판매|디퓨저|룸스프레이|선물|구매/.test(text)) {
    category = '완제품·답례품 판매';
  } else if (/원데이|클래스|수업|체험/.test(text)) {
    category = '원데이 클래스';
  } else {
    category = '기타';
  }

  // 인원 추출
  let people = 1;
  const peopleMatch = text.match(/(\d+)\s*(명|분|인)/);
  if (peopleMatch) {
    people = parseInt(peopleMatch[1], 10);
  } else if (/커플|두\s*명|두\s*분/.test(text)) {
    people = 2;
  } else if (/세\s*명|세\s*분/.test(text)) {
    people = 3;
  } else if (/네\s*명|네\s*분/.test(text)) {
    people = 4;
  }

  // 금액 추출 (예: 7만원, 7만5천원, 10만 5천원, 70,000원, 70000)
  let amount = 0;
  const manCheonMatch = text.match(/(\d+)\s*만\s*(\d+)\s*천\s*원?/);
  const manDetailMatch = text.match(/(\d+)\s*만\s*(\d{3,4})\s*원?/);
  const manMatch = text.match(/(\d+(?:\.\d+)?)\s*만\s*원?/);
  const cheonMatch = text.match(/(\d+)\s*천\s*원?/);
  const wonMatch = text.match(/(\d{1,3}(?:,\d{3})+|\d{4,8})\s*원?/);

  if (manCheonMatch) {
    amount = parseInt(manCheonMatch[1], 10) * 10000 + parseInt(manCheonMatch[2], 10) * 1000;
  } else if (manDetailMatch) {
    amount = parseInt(manDetailMatch[1], 10) * 10000 + parseInt(manDetailMatch[2], 10);
  } else if (manMatch) {
    amount = Math.round(parseFloat(manMatch[1]) * 10000);
  } else if (cheonMatch) {
    amount = parseInt(cheonMatch[1], 10) * 1000;
  } else if (wonMatch) {
    amount = parseInt(wonMatch[1].replace(/,/g, ''), 10);
  }

  // 결제수단 추출
  let payment = '카드';
  if (/포항사랑|상품권|지역화폐/.test(text)) payment = '포항사랑상품권';
  else if (/네이버\s*페이/.test(text)) payment = '네이버페이';
  else if (/카카오\s*페이/.test(text)) payment = '카카오페이';
  else if (/계좌|이체|입금|무통장/.test(text)) payment = '계좌이체';
  else if (/현금/.test(text)) payment = '현금';
  else if (/카드/.test(text)) payment = '카드';

  // 유입경로 추출
  let channel = '기타';
  if (/인스타|인스타그램|릴스|DM|디엠/.test(text)) channel = '인스타그램';
  else if (/네이버\s*예약|예약/.test(text)) channel = '네이버 예약';
  else if (/블로그|네이버|검색/.test(text)) channel = '네이버 검색/블로그';
  else if (/지인|소개|추천|친구/.test(text)) channel = '지인 추천';
  else if (/재방문|단골|또\s*오/.test(text)) channel = '재방문';
  else if (/지나가다|워크인|길가다|간판/.test(text)) channel = '워크인(지나가다)';
  else if (/당근/.test(text)) channel = '당근';

  return {
    date: date,
    category: category,
    people: people,
    amount: amount,
    payment: payment,
    channel: channel,
    memo: ''
  };
}


/**
 * ============================================================================
 * 📊 구글 시트 「윤슬향방 매출 장부」에 한 줄 저장 및 최근 내역 조회
 * ============================================================================
 */
function getOrCreateLedgerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('이 스크립트가 구글 시트 안의 [확장 프로그램 > Apps Script]에 들어있는지 확인해 주세요.');
  }

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // 만약 첫 번째 시트가 비어 있으면 이름을 바꾸고, 아니면 새로 생성
    const firstSheet = ss.getSheets()[0];
    if (firstSheet && firstSheet.getLastRow() === 0) {
      firstSheet.setName(SHEET_NAME);
      sheet = firstSheet;
    } else {
      sheet = ss.insertSheet(SHEET_NAME);
    }
  }

  // 첫 줄(헤더 7칸)이 비어 있으면 예쁘게 자동 생성
  if (sheet.getLastRow() === 0) {
    const headers = ['날짜', '구분', '인원', '금액', '결제수단', '유입경로', '비고'];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#F5EFE6');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }

  return sheet;
}


function appendRowToSheet_(item, userEmail) {
  const sheet = getOrCreateLedgerSheet_();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  const rowData = [
    String(item.date || todayStr).trim(),
    String(item.category || '원데이 클래스').trim(),
    Number(item.people) || 1,
    Number(String(item.amount || '0').replace(/[^0-9]/g, '')) || 0,
    String(item.payment || '카드').trim(),
    String(item.channel || '기타').trim(),
    String(item.memo || '').trim()
  ];

  sheet.appendRow(rowData);

  // 새로 추가된 줄의 인원(C열), 금액(D열) 서식 깔끔하게 지정
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setHorizontalAlignment('center');
  sheet.getRange(lastRow, 2).setHorizontalAlignment('center');
  sheet.getRange(lastRow, 3).setNumberFormat('#,##0"명"').setHorizontalAlignment('right');
  sheet.getRange(lastRow, 4).setNumberFormat('#,##0"원"').setHorizontalAlignment('right');
  sheet.getRange(lastRow, 5).setHorizontalAlignment('center');
  sheet.getRange(lastRow, 6).setHorizontalAlignment('center');

  return {
    date: rowData[0],
    category: rowData[1],
    people: rowData[2],
    amount: rowData[3],
    payment: rowData[4],
    channel: rowData[5],
    memo: rowData[6],
    savedBy: userEmail || ''
  };
}


function getRecentRows_(limit) {
  try {
    const sheet = getOrCreateLedgerSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const count = Math.min(limit || 5, lastRow - 1);
    const startRow = lastRow - count + 1;
    const values = sheet.getRange(startRow, 1, count, 7).getValues();
    const timezone = 'Asia/Seoul';

    return values.reverse().map(function(r) {
      let dateVal = r[0];
      if (Object.prototype.toString.call(dateVal) === '[object Date]') {
        dateVal = Utilities.formatDate(dateVal, timezone, 'yyyy-MM-dd');
      } else {
        dateVal = String(dateVal || '');
      }
      return {
        date: dateVal,
        category: String(r[1] || ''),
        people: Number(r[2]) || 1,
        amount: Number(r[3]) || 0,
        payment: String(r[4] || ''),
        channel: String(r[5] || ''),
        memo: String(r[6] || '')
      };
    });
  } catch (e) {
    return [];
  }
}


function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

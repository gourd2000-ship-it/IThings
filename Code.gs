/**
 * ==============================================================================
 * 학교 기자재 관리 시스템 - Google Apps Script (QR 코드 연동형)
 * ==============================================================================
 * 
 * [시트 요구 조건]
 * 본 스크립트 실행 전 최초 1회 setupSheets() 함수를 실행하면 아래 시트들이 자동 생성됩니다.
 * 
 * 1. [마스터] 시트: 기자재 정보 저장 및 관리
 *    - 칼럼 구성: 연번 | 관리번호 | 설치장소 | 취급자 | 종류 | 제조사 | 모델명 | 도입일자 | 비고 | QR링크 | QR이미지 | 최종수정일
 *    - ※ 보안 주의: '비밀번호' 관련 정보는 취급하지 않으며 헤더에 포함하지 않습니다.
 * 
 * 2. [설정] 시트: 드롭다운 옵션 및 URL 환경변수 설정
 *    - 웹앱URL: 웹앱 배포 시 생성되는 주소 입력
 *    - 설치장소목록: 쉼표(,)로 구분된 리스트 (예: 교무실,행정실,1-1,1-2,1-3,과학실,컴퓨터실)
 *    - 취급자목록: 쉼표(,)로 구분된 리스트 (예: 홍길동,김철수,이영희,박민수)
 *    - 관리자이름: 모든 라벨에 공통으로 인쇄될 관리책임자 이름 (예: 교사 이영희)
 * 
 * 3. [기기관리라벨] 시트: 출력용 QR 바코드 라벨 자동 생성 영역
 * 
 * ==============================================================================
 */

/**
 * 웹 브라우저 최초 접속 시 실행되는 진입 함수 (HTTP GET)
 * @param {Object} e - 웹앱에 전달된 이벤트 파라미터 객체
 * @return {HtmlOutput} 렌더링될 HTML 화면 데이터
 */
function doGet(e) {
  // 1. URL 쿼리 파라미터에서 기기 식별을 위한 'id' (관리번호) 추출
  var id = e.parameter.id;
  
  if (id === "DEBUG_TEST") {
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var settingsSheet = ss.getSheetByName("설정");
      var action = e.parameter.action;
      if (action === "lock" || action === "unlock") {
        if (settingsSheet) {
          var lastRow = settingsSheet.getLastRow();
          var found = false;
          if (lastRow >= 2) {
            var values = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
            for (var i = 0; i < values.length; i++) {
              if (values[i][0].toString().trim() === "모든입력잠금") {
                settingsSheet.getRange(i + 2, 2).setValue(action === "lock" ? "Y" : "N");
                found = true;
                break;
              }
            }
          }
          if (!found) {
            settingsSheet.appendRow(["모든입력잠금", action === "lock" ? "Y" : "N"]);
          }
          SpreadsheetApp.flush();
        }
      }
      var sheet = ss.getSheetByName("마스터");
      var ids = [];
      if (sheet) {
        var lastRow = sheet.getLastRow();
        if (lastRow >= 2) {
          var headerMapping = getHeaderMapping_(sheet);
          var idCol = headerMapping["관리번호"];
          if (idCol) {
            ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues().map(function(r) { return r[0]; });
          }
        }
      }
      var settings = getSettings_();
      var debugObj = {
        ids: ids,
        settings: settings,
        sheets: ss.getSheets().map(function(s) { return s.getName(); })
      };
      return HtmlService.createHtmlOutput("DEBUG_DATA:" + JSON.stringify(debugObj))
        .setTitle("Debug Info")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (e) {
      return HtmlService.createHtmlOutput("DEBUG_ERROR:" + e.toString() + "\n" + e.stack)
        .setTitle("Debug Error")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  }
  
  // 2. id 값이 전달되지 않은 경우 에러 화면 반환
  if (!id) {
    return HtmlService.createHtmlOutput(createErrorHtml_("관리번호(id)가 전달되지 않았습니다."))
      .setTitle("오류 - 기자재 관리 시스템")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  try {
    // 3. 마스터 시트에서 관리번호로 기기 상세 정보 조회
    var device = getDeviceById_(id);
    
    // 4. 해당 관리번호에 해당하는 기기가 없는 경우 에러 화면 반환
    if (!device) {
      return HtmlService.createHtmlOutput(createErrorHtml_("해당 관리번호의 기기를 찾을 수 없습니다. (관리번호: " + id + ")"))
        .setTitle("조회 실패 - 기자재 관리 시스템")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // 5. 드롭다운 목록 조회를 위해 설정값 호출
    var settings = getSettings_();
    
    // 6. 정상적인 정보 수정 HTML 폼 화면 생성 및 반환
    var htmlContent = createHtml_(device, settings);
    return HtmlService.createHtmlOutput(htmlContent)
      .setTitle("기자재 정보 확인 및 수정 - " + device["관리번호"])
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no');
      
  } catch (err) {
    // 예외 발생 시 디버깅을 위한 에러 페이지 반환
    return HtmlService.createHtmlOutput(createErrorHtml_("시스템 로드 중 오류가 발생했습니다: " + err.message))
      .setTitle("오류 - 기자재 관리 시스템")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

/**
 * 마스터 시트에서 동적으로 헤더를 읽어 칼럼명과 열 번호(1-indexed)를 매핑하는 헬퍼 함수
 * @param {Sheet} sheet - 구글 스프레드시트의 시트 객체
 * @return {Object} 헤더명과 열 번호가 매핑된 객체 { "칼럼명": 열번호 }
 */
function getHeaderMapping_(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) {
    return {};
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var mapping = {};
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i].toString().trim();
    if (header) {
      mapping[header] = i + 1;
    }
  }
  return mapping;
}

/**
 * [설정] 시트의 환경변수 목록을 읽어오는 함수
 * @return {Object} 웹앱 URL 및 드롭다운용 배열이 담긴 설정 객체
 */
function getSettings_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("설정");
  
  var defaultSettings = {
    webAppUrl: "",
    locations: [],
    managers: [],
    deviceTypes: ["PC", "노트북", "프린터", "3D프린터", "TV", "모니터", "태블릿", "기타"],
    pcSensitiveEditAllowed: false,
    basicInfoEditAllowed: false,
    allEditLocked: false,
    schoolName: "비아초등학교",
    managerName: "교사 이영희"
  };
  
  if (!sheet) {
    return defaultSettings;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return defaultSettings;
  }
  
  // 설정 시트의 키-값 데이터 일괄 리드
  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    var key = values[i][0].toString().trim();
    var val = values[i][1].toString().trim();
    
    if (key === "웹앱URL") {
      defaultSettings.webAppUrl = val;
    } else if (key === "설치장소목록") {
      defaultSettings.locations = val.split(",").map(function(x) { return x.trim(); }).filter(Boolean);
    } else if (key === "취급자목록") {
      defaultSettings.managers = val.split(",").map(function(x) { return x.trim(); }).filter(Boolean);
    } else if (key === "PC민감정보수정허용") {
      defaultSettings.pcSensitiveEditAllowed = (val.toUpperCase() === "Y" || val.toUpperCase() === "TRUE");
    } else if (key === "기본정보수정허용") {
      defaultSettings.basicInfoEditAllowed = (val.toUpperCase() === "Y" || val.toUpperCase() === "TRUE");
    } else if (key === "기관명") {
      defaultSettings.schoolName = val;
    } else if (key === "관리자이름") {
      defaultSettings.managerName = val;
    } else if (key === "모든입력잠금") {
      defaultSettings.allEditLocked = (val.toUpperCase() === "Y" || val.toUpperCase() === "TRUE");
    }
  }
  
  // 마스터 시트에서 실시간 취급자 및 설치장소 고유 데이터 수집하여 추천 풀에 통합
  var masterSheet = ss.getSheetByName("마스터");
  if (masterSheet) {
    var masterLastRow = masterSheet.getLastRow();
    if (masterLastRow >= 2) {
      var headerMapping = getHeaderMapping_(masterSheet);
      
      // 1. 설치장소 실시간 수집 및 병합
      var locCol = headerMapping["설치장소"];
      if (locCol) {
        var masterLocs = masterSheet.getRange(2, locCol, masterLastRow - 1, 1).getValues();
        for (var k = 0; k < masterLocs.length; k++) {
          var val = masterLocs[k][0].toString().trim();
          if (val && defaultSettings.locations.indexOf(val) === -1) {
            defaultSettings.locations.push(val);
          }
        }
      }
      
      // 2. 취급자 실시간 수집 및 병합
      var mgrCol = headerMapping["취급자"];
      if (mgrCol) {
        var masterMgrs = masterSheet.getRange(2, mgrCol, masterLastRow - 1, 1).getValues();
        for (var k = 0; k < masterMgrs.length; k++) {
          var val = masterMgrs[k][0].toString().trim();
          if (val && defaultSettings.managers.indexOf(val) === -1) {
            defaultSettings.managers.push(val);
          }
        }
      }
      
      // 3. 기기 종류 실시간 수집 및 병합
      var typeCol = headerMapping["종류"];
      if (typeCol) {
        var masterTypes = masterSheet.getRange(2, typeCol, masterLastRow - 1, 1).getValues();
        for (var k = 0; k < masterTypes.length; k++) {
          var val = masterTypes[k][0].toString().trim();
          if (val && defaultSettings.deviceTypes.indexOf(val) === -1) {
            defaultSettings.deviceTypes.push(val);
          }
        }
      }
    }
  }
  
  // 추천 편의를 위해 가나다순 오름차순 정렬
  defaultSettings.locations.sort();
  defaultSettings.managers.sort();
  defaultSettings.deviceTypes.sort();
  
  return defaultSettings;
}

/**
 * [마스터] 시트에서 관리번호를 기반으로 기기 정보를 조회하는 함수
 * @param {string} id - 조회하고자 하는 기기의 고유 관리번호
 * @return {Object|null} 기기 정보 데이터 및 해당 행(Row) 번호 객체
 */
function getDeviceById_(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("마스터");
  if (!sheet) {
    throw new Error("마스터 시트가 존재하지 않습니다. setupSheets()를 실행해 주세요.");
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  
  var headerMapping = getHeaderMapping_(sheet);
  var idCol = headerMapping["관리번호"];
  if (!idCol) {
    throw new Error("마스터 시트에 '관리번호' 헤더가 존재하지 않습니다.");
  }
  
  // ID 열 전체 데이터 조회 (2행부터 마지막 행까지)
  var idValues = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  var targetRow = -1;
  
  for (var i = 0; i < idValues.length; i++) {
    if (idValues[i][0].toString().trim() === id.toString().trim()) {
      targetRow = i + 2; // 인덱스는 0부터 시작하고, 헤더를 제외한 2행부터 시작하므로 +2
      break;
    }
  }
  
  if (targetRow === -1) {
    return null;
  }
  
  // 조회한 타겟 행 전체 데이터 읽기
  var rowValues = sheet.getRange(targetRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  var device = {
    rowNum: targetRow
  };
  
  // 헤더 정보를 기반으로 데이터를 동적 객체 맵핑
  var settings = getSettings_();
  var typeCol = headerMapping["종류"];
  var deviceType = typeCol ? rowValues[typeCol - 1].toString().trim() : "";
  var isPc = (deviceType.toUpperCase() === "PC");
  var allowSensitive = isPc && settings.pcSensitiveEditAllowed;
  
  for (var header in headerMapping) {
    var colIdx = headerMapping[header];
    
    // [보안] 기기가 PC이고 설정에서 허용한 경우가 아니라면 IP 및 비밀번호 관련 필드는 조회에서 제외
    if (header === "IP" || header.indexOf("비밀번호") !== -1) {
      if (!allowSensitive) {
        continue;
      }
    }
    
    var val = rowValues[colIdx - 1];
    
    // 날짜형 데이터가 JSON 직렬화 중 깨지지 않도록 문자열 포맷 변환
    if (val instanceof Date) {
      device[header] = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    } else {
      device[header] = val;
    }
  }
  
  return device;
}

/**
 * 웹 폼에서 제출한 내용으로 마스터 시트 데이터를 갱신하는 함수
 * @param {string} id - 기기 관리번호
 * @param {Object} formData - 웹 폼으로부터 전달받은 정보 객체
 * @return {Object} 처리 성공 상태 및 메시지 객체
 */
function updateDevice(id, formData) {
  // [보안] 모든 입력 잠금이 활성화되어 있는 경우, 저장 요청 즉시 차단
  var settings = getSettings_();
  if (settings.allEditLocked) {
    return {
      success: false,
      message: "저장 실패: 현재 시스템이 입력 잠금 상태이므로 수정 내용을 저장할 수 없습니다."
    };
  }

  // 동시 쓰기 작업 시 데이터 유실 방지를 위한 Lock 서비스 획득
  var lock = LockService.getScriptLock();
  try {
    // 10초간 잠금 대기
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("동시 수정을 처리하지 못했습니다. 잠시 후 다시 제출해주세요.");
  }
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마스터");
    if (!sheet) {
      throw new Error("마스터 시트를 찾을 수 없습니다.");
    }
    
    // 1. 수정 대상 기기 탐색
    var device = getDeviceById_(id);
    if (!device) {
      throw new Error("수정하려는 기기가 마스터 시트에 없습니다.");
    }
    
    var rowNum = device.rowNum;
    var headerMapping = getHeaderMapping_(sheet);
    var settings = getSettings_();
    
    // PC 기기 및 민감정보 수정 권한 여부 확인
    var isPc = (device["종류"] && device["종류"].toString().trim().toUpperCase() === "PC");
    var allowSensitive = isPc && settings.pcSensitiveEditAllowed;
    var allowBasicInfo = settings.basicInfoEditAllowed;
    
    // 2. 수정 데이터 정의
    var updateFields = {
      "설치장소": formData.location,
      "취급자": formData.manager,
      "비고": formData.notes,
      "최종수정일": new Date() // 현재 날짜시간 자동 기록
    };
    
    // PC이고 관리자 허가 시 IP 및 비밀번호 수정 데이터 포함
    if (allowSensitive) {
      if (formData.ip !== undefined) updateFields["IP"] = formData.ip;
      if (formData.password1st !== undefined) updateFields["비밀번호(1차)"] = formData.password1st;
      if (formData.password2nd !== undefined) updateFields["비밀번호(2차)"] = formData.password2nd;
    }
    
    // 기본 정보 수정 권한 허용 시 기본 정보도 포함
    if (allowBasicInfo) {
      if (formData.deviceType !== undefined && headerMapping["종류"]) updateFields["종류"] = formData.deviceType;
      if (formData.manufacturer !== undefined && headerMapping["제조사"]) updateFields["제조사"] = formData.manufacturer;
      if (formData.modelName !== undefined && headerMapping["모델명"]) updateFields["모델명"] = formData.modelName;
      if (formData.introDate !== undefined && headerMapping["도입일자"]) updateFields["도입일자"] = formData.introDate;
    }
    
    // 3. 실제 마스터 시트에 값 쓰기
    for (var header in updateFields) {
      // [보안] 권한이 없는 경우 IP 및 비밀번호 수정 차단
      if (header === "IP" || header.indexOf("비밀번호") !== -1) {
        if (!allowSensitive) {
          continue;
        }
      }
      
      // [보안] 권한이 없는 경우 기본 정보 수정 차단
      if (header === "종류" || header === "제조사" || header === "모델명" || header === "도입일자") {
        if (!allowBasicInfo) {
          continue;
        }
      }
      
      var colIdx = headerMapping[header];
      if (colIdx) {
        sheet.getRange(rowNum, colIdx).setValue(updateFields[header]);
      }
    }
    
    // 4. [선택적 확장] 수정로그 시트 기록 시도
    logToHistory_(ss, id, updateFields);
    
    return {
      success: true,
      message: "저장에 성공했습니다."
    };
    
  } catch (err) {
    return {
      success: false,
      message: "저장 실패: " + err.message
    };
  } finally {
    // 프로세스 성공 여부와 상관없이 무조건 락 해제 보장
    lock.releaseLock();
  }
}

/**
 * 변경 이력을 관리하기 위한 로그 작성 (선택 기능)
 * 스프레드시트에 '수정로그' 시트가 존재할 경우에만 한 행을 덧붙입니다.
 */
function logToHistory_(ss, id, fields) {
  var logSheet = ss.getSheetByName("수정로그");
  if (!logSheet) {
    return; // 시트가 없으면 기본 마스터 직접 쓰기 방식으로 작동하며 통과
  }
  var logMessage = "장소: " + fields["설치장소"] + " | 취급자: " + fields["취급자"] + " | 비고: " + (fields["비고"] || "없음");
  logSheet.appendRow([
    new Date(),
    id,
    logMessage
  ]);
}

/**
 * 모바일 최적화 웹앱 폼 HTML 문자열 생성 함수
 * @param {Object} device - 기기 상세 정보
 * @param {Object} settings - 설정 데이터
 * @return {string} HTML 콘텐츠
 */
function createHtml_(device, settings) {
  var deviceJson = JSON.stringify(device).replace(/<\/script>/gi, '<\\/script>');
  var settingsJson = JSON.stringify(settings).replace(/<\/script>/gi, '<\\/script>');
  
  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>기자재 정보 수정</title>
    <!-- 고급 폰트 적용 (Outfit / Noto Sans KR) -->
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=Noto+Sans+KR:wght@300;400;500;700&display=swap" rel="stylesheet">
    <!-- PDF 저장을 위한 html2pdf.js, html2canvas, jsPDF, qrcode.js 라이브러리 추가 -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js"></script>
    <style>
      :root {
        --bg-color: #0b0f19;
        --card-bg: rgba(22, 28, 45, 0.75);
        --card-border: rgba(255, 255, 255, 0.08);
        --primary-gradient: linear-gradient(135deg, #6366f1, #4f46e5);
        --success-gradient: linear-gradient(135deg, #10b981, #059669);
        --primary-hover: #4338ca;
        --text-color: #f3f4f6;
        --text-muted: #9ca3af;
        --text-label: #cbd5e1;
        --input-bg: #111827;
        --input-border: #374151;
        --input-focus: #6366f1;
        --font-family: 'Outfit', 'Noto Sans KR', sans-serif;
      }
      * {
        box-sizing: border-box;
        margin: 0;
        padding: 0;
      }
      body {
        background-color: var(--bg-color);
        color: var(--text-color);
        font-family: var(--font-family);
        min-height: 100vh;
        display: flex;
        justify-content: center;
        align-items: flex-start;
        padding: 24px 16px;
      }
      .container {
        width: 100%;
        max-width: 480px;
        background: var(--card-bg);
        border: 1px solid var(--card-border);
        border-radius: 28px;
        backdrop-filter: blur(12px);
        padding: 32px 24px;
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
        animation: slideUp 0.6s cubic-bezier(0.16, 1, 0.3, 1);
      }
      @keyframes slideUp {
        from { opacity: 0; transform: translateY(30px); }
        to { opacity: 1; transform: translateY(0); }
      }
      header {
        text-align: center;
        margin-bottom: 28px;
      }
      h1 {
        font-size: 22px;
        font-weight: 700;
        letter-spacing: -0.5px;
        margin-bottom: 8px;
        background: linear-gradient(to right, #a5b4fc, #e0e7ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
      }
      .subtitle {
        font-size: 14px;
        color: var(--text-muted);
      }
      .section-title {
        font-size: 13px;
        font-weight: 600;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 1px;
        margin: 24px 0 12px 0;
        display: flex;
        align-items: center;
      }
      .section-title::after {
        content: "";
        flex-grow: 1;
        height: 1px;
        background: var(--card-border);
        margin-left: 10px;
      }
      .info-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        margin-bottom: 20px;
      }
      .info-item {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.03);
        padding: 12px 16px;
        border-radius: 16px;
      }
      .info-item.full-width {
        grid-column: span 2;
      }
      .info-label {
        font-size: 11px;
        color: var(--text-muted);
        margin-bottom: 4px;
      }
      .info-value {
        font-size: 14px;
        font-weight: 500;
      }
      .info-value input, .info-value select {
        width: 100%;
        background: transparent;
        border: none;
        border-bottom: 1px dashed rgba(255, 255, 255, 0.2);
        color: var(--text-color);
        font-family: var(--font-family);
        font-size: 14px;
        font-weight: 500;
        padding: 2px 0;
        outline: none;
        box-shadow: none;
        border-radius: 0;
        transition: border-color 0.2s;
      }
      .info-value select {
        appearance: none;
        -webkit-appearance: none;
        -moz-appearance: none;
        background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%239ca3af%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.9%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
        background-repeat: no-repeat;
        background-position: right 4px center;
        background-size: 10px auto;
        padding-right: 20px;
      }
      .info-value input:focus, .info-value select:focus {
        border-bottom: 1px solid var(--input-focus);
      }
      .info-value input[type="date"] {
        color-scheme: dark;
      }
      .form-group {
        margin-bottom: 18px;
      }
      label {
        display: block;
        font-size: 13px;
        font-weight: 500;
        color: var(--text-label);
        margin-bottom: 8px;
      }
      select, textarea, input[type="text"], input[type="password"] {
        width: 100%;
        background-color: var(--input-bg);
        border: 1px solid var(--input-border);
        border-radius: 14px;
        padding: 12px 16px;
        color: var(--text-color);
        font-family: var(--font-family);
        font-size: 14px;
        transition: all 0.2s;
        outline: none;
      }
      select:focus, textarea:focus, input[type="text"]:focus, input[type="password"]:focus {
        border-color: var(--input-focus);
        box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.25);
      }
      textarea {
        resize: vertical;
        min-height: 80px;
      }
      /* 하이브리드 입력 추천상자 전용 화살표 아이콘 스타일 */
      input[list] {
        background-image: url("data:image/svg+xml;charset=US-ASCII,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22292.4%22%20height%3D%22292.4%22%3E%3Cpath%20fill%3D%22%239ca3af%22%20d%3D%22M287%2069.4a17.6%2017.6%200%200%200-13-5.4H18.4c-5%200-9.3%201.8-12.9%205.4A17.6%2017.6%200%200%200%200%2082.2c0%205%201.8%209.3%205.4%2012.9l128%20127.9c3.6%203.6%207.8%205.4%2012.8%205.4s9.2-1.8%2012.8-5.4L287%2095c3.5-3.5%205.4-7.8%205.4-12.9%200-5-1.9-9.2-5.5-12.8z%22%2F%3E%3C%2Fsvg%3E");
        background-repeat: no-repeat;
        background-position: right 16px center;
        background-size: 10px auto;
        padding-right: 40px;
      }
      input[list]::-webkit-calendar-picker-indicator {
        display: none !important;
      }
      .btn {
        width: 100%;
        background: var(--primary-gradient);
        border: none;
        padding: 14px 24px;
        border-radius: 14px;
        color: white;
        font-family: var(--font-family);
        font-size: 15px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
        box-shadow: 0 4px 12px rgba(79, 70, 229, 0.35);
        display: flex;
        justify-content: center;
        align-items: center;
      }
      .btn:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 166px rgba(79, 70, 229, 0.45);
      }
      .btn:active {
        transform: translateY(0);
      }
      .btn:disabled {
        opacity: 0.6;
        cursor: not-allowed;
      }
      
      /* 로딩 스피너 */
      .spinner {
        width: 20px;
        height: 20px;
        border: 2px solid rgba(255, 255, 255, 0.3);
        border-radius: 50%;
        border-top-color: white;
        animation: spin 0.8s linear infinite;
        margin-right: 8px;
        display: none;
      }
      @keyframes spin {
        to { transform: rotate(360deg); }
      }
      
      /* 완료 뷰 스타일 */
      .success-view {
        display: none;
        text-align: center;
        animation: fadeIn 0.4s ease-out;
      }
      @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
      }
      .success-icon {
        display: inline-flex;
        justify-content: center;
        align-items: center;
        width: 64px;
        height: 64px;
        border-radius: 50%;
        background: var(--success-gradient);
        margin-bottom: 20px;
        box-shadow: 0 8px 16px rgba(16, 185, 129, 0.3);
      }
      .success-icon svg {
        width: 32px;
        height: 32px;
        fill: white;
      }
      .success-title {
        font-size: 20px;
        font-weight: 700;
        margin-bottom: 12px;
      }
      .summary-card {
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid var(--card-border);
        border-radius: 20px;
        padding: 20px;
        text-align: left;
        margin-bottom: 24px;
      }
      .summary-row {
        display: flex;
        justify-content: space-between;
        padding: 10px 0;
        border-bottom: 1px solid rgba(255, 255, 255, 0.05);
        font-size: 14px;
      }
      .summary-row:last-child {
        border-bottom: none;
      }
      .summary-label {
        color: var(--text-muted);
      }
      .summary-val {
        font-weight: 500;
      }
      
      .btn-secondary {
        background: transparent;
        border: 1px solid var(--card-border);
        color: var(--text-color);
        box-shadow: none;
      }
      .btn-secondary:hover {
        background: rgba(255, 255, 255, 0.05);
        box-shadow: none;
      }
      
      .error-alert {
        display: none;
        background: rgba(239, 68, 68, 0.1);
        border: 1px solid rgba(239, 68, 68, 0.2);
        color: #fca5a5;
        padding: 12px 16px;
        border-radius: 12px;
        font-size: 13px;
        margin-bottom: 18px;
        text-align: left;
      }
      .lock-banner {
        display: none;
        background: rgba(245, 158, 11, 0.1);
        border: 1px solid rgba(245, 158, 11, 0.25);
        color: #fef08a;
        padding: 14px 18px;
        border-radius: 16px;
        font-size: 13.5px;
        line-height: 1.5;
        margin-bottom: 18px;
        align-items: center;
        gap: 12px;
        text-align: left;
        animation: fadeIn 0.4s ease-out;
      }
      
      /* 플로팅 버튼 (FAB) */
      .fab-container {
        position: fixed;
        bottom: 24px;
        right: 24px;
        display: flex;
        flex-direction: column;
        gap: 12px;
        z-index: 1000;
      }
      .fab-btn {
        background: var(--primary-gradient);
        border: none;
        border-radius: 50px;
        color: white;
        padding: 12px 18px;
        font-family: var(--font-family);
        font-size: 13px;
        font-weight: 600;
        cursor: pointer;
        box-shadow: 0 10px 25px rgba(99, 102, 241, 0.4);
        display: flex;
        align-items: center;
        gap: 8px;
        transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        outline: none;
      }
      .fab-btn:hover {
        transform: translateY(-3px);
        box-shadow: 0 15px 30px rgba(99, 102, 241, 0.5);
      }
      .fab-btn.cart-add {
        background: var(--success-gradient);
        box-shadow: 0 10px 25px rgba(16, 185, 129, 0.4);
      }
      .fab-btn.cart-add:hover {
        box-shadow: 0 15px 30px rgba(16, 185, 129, 0.5);
      }
      
      /* 모달 배경 */
      .modal-backdrop {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background: rgba(11, 15, 25, 0.8);
        backdrop-filter: blur(8px);
        display: none;
        justify-content: center;
        align-items: center;
        z-index: 1100;
        opacity: 0;
        transition: opacity 0.25s ease;
      }
      .modal-backdrop.active {
        display: flex;
        opacity: 1;
      }
      /* 모달 카드 */
      .modal-card {
        width: 90%;
        max-width: 400px;
        background: #161c2d;
        border: 1px solid var(--card-border);
        border-radius: 24px;
        padding: 24px;
        box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.7);
        display: flex;
        flex-direction: column;
        max-height: 80vh;
        animation: modalFadeIn 0.3s cubic-bezier(0.16, 1, 0.3, 1);
      }
      @keyframes modalFadeIn {
        from { transform: scale(0.95); opacity: 0; }
        to { transform: scale(1); opacity: 1; }
      }
      .modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 18px;
      }
      .modal-title {
        font-size: 17px;
        font-weight: 700;
        color: var(--text-color);
      }
      .modal-close {
        background: transparent;
        border: none;
        color: var(--text-muted);
        font-size: 24px;
        cursor: pointer;
        outline: none;
      }
      /* 보관함 목록 */
      .cart-list {
        overflow-y: auto;
        flex-grow: 1;
        margin-bottom: 20px;
        max-height: 350px;
      }
      .cart-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        background: rgba(255, 255, 255, 0.03);
        border: 1px solid rgba(255, 255, 255, 0.04);
        border-radius: 14px;
        padding: 12px 16px;
        margin-bottom: 10px;
      }
      .cart-item-info {
        display: flex;
        flex-direction: column;
        gap: 2px;
        cursor: pointer;
        flex-grow: 1;
        text-align: left;
      }
      .cart-item-info:hover {
        opacity: 0.8;
      }
      .cart-item-id {
        font-size: 13.5px;
        font-weight: 600;
        color: #a5b4fc;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .cart-item-id::after {
        content: "🔗";
        font-size: 9px;
      }
      .cart-item-meta {
        font-size: 11.5px;
        color: var(--text-muted);
      }
      .cart-item-delete {
        background: transparent;
        border: none;
        color: #f87171;
        cursor: pointer;
        padding: 4px;
        font-size: 20px;
        outline: none;
      }
      .cart-empty {
        text-align: center;
        color: var(--text-muted);
        padding: 40px 0;
        font-size: 13.5px;
      }
      .modal-footer {
        display: flex;
        gap: 10px;
      }
      
      /* --- 인쇄 전용 그리드 및 인쇄 CSS (PDF 변환 시에도 2x5 레이아웃 유지) --- */
      #print-section {
        display: none;
        position: absolute;
        left: 0;
        top: 0;
        width: 210mm;
        height: 297mm;
        box-sizing: border-box;
        background: white !important;
        color: black !important;
      }
      .print-label {
        border: 1px dashed #ccc !important;
        border-radius: 6px;
        padding: 3mm 4mm;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-sizing: border-box;
        background: white !important;
        overflow: hidden;
      }
      .label-details {
        display: flex;
        flex-direction: column;
        gap: 1mm;
        flex-grow: 1;
        color: black !important;
        text-align: left;
        overflow: hidden;
      }
      .label-title {
        font-size: 11pt;
        font-weight: bold;
        color: #1a202c !important;
        border-bottom: 2px solid #2d3748 !important;
        padding-bottom: 0.5mm;
        margin-bottom: 1mm;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .label-row {
        font-size: 8.5pt;
        display: flex;
        align-items: center;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .label-key {
        font-weight: bold;
        width: 11mm;
        color: #4a5568 !important;
        flex-shrink: 0;
      }
      .label-val {
        color: #2d3748 !important;
        overflow: hidden;
        text-overflow: ellipsis;
      }
      .label-qr {
        width: 24mm;
        height: 24mm;
        display: flex;
        justify-content: center;
        align-items: center;
        margin-left: 2mm;
        flex-shrink: 0;
      }
      .label-qr img, .label-qr canvas {
        width: 100% !important;
        height: 100% !important;
      }
      
      @media print {
        body * {
          visibility: hidden !important;
        }
        #print-section, #print-section * {
          visibility: visible !important;
        }
        #print-section {
          display: block !important;
          position: fixed !important;
          left: 0 !important;
          top: 0 !important;
          width: 210mm !important;
          height: 297mm !important;
          background: white !important;
        }
      }
    </style>
  </head>
  <body>
    <div class="container" id="main-container">
      
      <!-- 수정 입력 폼 화면 -->
      <div id="form-view">
        <header>
          <h1>기자재 정보 확인 및 수정</h1>
          <p class="subtitle">모바일 QR 등록 웹앱</p>
        </header>
        
        <div class="error-alert" id="error-box"></div>
        <div id="lock-banner" class="lock-banner"></div>
        
        <div class="section-title">기기 기본 정보</div>
        <div class="info-grid">
          <div class="info-item">
            <div class="info-label">관리번호</div>
            <div class="info-value" id="disp-id"></div>
          </div>
          <div class="info-item">
            <div class="info-label">종류</div>
            <div class="info-value" id="disp-type-container"></div>
          </div>
          <div class="info-item">
            <div class="info-label">제조사</div>
            <div class="info-value" id="disp-brand-container"></div>
          </div>
          <div class="info-item">
            <div class="info-label">모델명</div>
            <div class="info-value" id="disp-model-container"></div>
          </div>
          <div class="info-item full-width">
            <div class="info-label">도입일자</div>
            <div class="info-value" id="disp-date-container"></div>
          </div>
        </div>
        
        <div class="section-title">수정 입력 정보</div>
        <form id="update-form" onsubmit="submitForm(event)">
          <div class="form-group">
            <label for="location">설치장소</label>
            <input type="text" id="location" list="location-list" placeholder="설치장소 선택 또는 직접 입력" required autocomplete="off">
            <datalist id="location-list"></datalist>
          </div>
          <div class="form-group">
            <label for="manager">취급자</label>
            <input type="text" id="manager" list="manager-list" placeholder="취급자 선택 또는 직접 입력" required autocomplete="off">
            <datalist id="manager-list"></datalist>
          </div>
          <div class="form-group">
            <label for="notes">비고</label>
            <textarea id="notes" placeholder="기타 특이사항이나 상태를 입력해주세요."></textarea>
          </div>
          
          <!-- 조건부 노출 민감 정보 (PC 기기 및 관리자 권한 Y 시) -->
          <div id="sensitive-fields" style="display: none;">
            <div class="form-group">
              <label for="ip">IP 주소</label>
              <input type="text" id="ip" placeholder="192.168.x.x">
            </div>
            <div class="form-group">
              <label for="password1st">비밀번호(1차)</label>
              <input type="password" id="password1st" placeholder="1차 비밀번호 입력" autocomplete="new-password">
            </div>
            <div class="form-group">
              <label for="password2nd">비밀번호(2차)</label>
              <input type="password" id="password2nd" placeholder="2차 비밀번호 입력" autocomplete="new-password">
            </div>
          </div>
          
          <button type="submit" class="btn" id="submit-btn">
            <div class="spinner" id="btn-spinner"></div>
            <span id="btn-text">수정 내용 저장</span>
          </button>
        </form>
      </div>
      
      <!-- 수정 완료 화면 -->
      <div id="success-view" class="success-view">
        <div class="success-icon">
          <svg viewBox="0 0 24 24">
            <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
          </svg>
        </div>
        <div class="success-title">저장 완료</div>
        <p class="subtitle" style="margin-bottom: 24px;">수정 정보가 스프레드시트에 저장되었습니다.</p>
        
        <div class="summary-card">
          <div class="summary-row">
            <span class="summary-label">관리번호</span>
            <span class="summary-val" id="sum-id"></span>
          </div>
          <div class="summary-row">
            <span class="summary-label">설치장소</span>
            <span class="summary-val" id="sum-location"></span>
          </div>
          <div class="summary-row">
            <span class="summary-label">취급자</span>
            <span class="summary-val" id="sum-manager"></span>
          </div>
          <div class="summary-row">
            <span class="summary-label">비고</span>
            <span class="summary-val" id="sum-notes"></span>
          </div>
          <div class="summary-row" id="sum-row-ip" style="display: none;">
            <span class="summary-label">IP 주소</span>
            <span class="summary-val" id="sum-ip"></span>
          </div>
          <div class="summary-row" id="sum-row-pw1" style="display: none;">
            <span class="summary-label">비밀번호(1차)</span>
            <span class="summary-val" id="sum-pw1"></span>
          </div>
          <div class="summary-row" id="sum-row-pw2" style="display: none;">
            <span class="summary-label">비밀번호(2차)</span>
            <span class="summary-val" id="sum-pw2"></span>
          </div>
        </div>
        
        <button class="btn btn-secondary" onclick="showForm()">다시 수정하기</button>
      </div>
    </div>
    
    <!-- 플로팅 버튼 영역 -->
    <div class="fab-container">
      <button class="fab-btn cart-add" id="fab-add-btn" onclick="addToCart()">
        <span>📥 보관함 담기</span>
      </button>
      <button class="fab-btn" id="fab-view-btn" onclick="openCartModal()">
        <span>🛒 보관함 보기 (<span id="cart-count">0</span>)</span>
      </button>
    </div>

    <!-- 보관함 모달창 -->
    <div class="modal-backdrop" id="cart-modal" onclick="closeCartModalOnBackdrop(event)">
      <div class="modal-card" onclick="event.stopPropagation()">
        <div class="modal-header">
          <div class="modal-title">라벨 인쇄 보관함</div>
          <button class="modal-close" onclick="closeCartModal()">&times;</button>
        </div>
        <div class="cart-list" id="cart-items-container"></div>
        <div class="modal-footer" style="flex-wrap: wrap; gap: 8px;">
          <button class="btn btn-secondary" style="flex: 1 1 30%; padding: 10px 8px; font-size: 12.5px;" onclick="clearCart()">전체비우기</button>
          <button class="btn btn-secondary" style="flex: 1 1 30%; padding: 10px 8px; font-size: 12.5px; background: rgba(99, 102, 241, 0.1); border-color: rgba(99, 102, 241, 0.3); color: #a5b4fc;" onclick="saveCartAsPdf(event)">PDF 저장</button>
          <button class="btn" style="flex: 1 1 90%; padding: 10px 8px; font-size: 12.5px;" onclick="printCartLabels()">모아 인쇄</button>
        </div>
      </div>
    </div>

    <!-- 인쇄 전용 그리드 영역 -->
    <div id="print-section"></div>
    
    <script>
      // 서버에서 주입한 기기 정보와 설정 파일 로딩
      const device = ${deviceJson};
      const settings = ${settingsJson};
      
      // 기기 기본 정보 화면 바인딩
      document.getElementById("disp-id").innerText = device["관리번호"] || "-";
      
      const basicEdit = settings.basicInfoEditAllowed;
      
      // 종류 바인딩 (드롭다운 선택상자)
      const typeContainer = document.getElementById("disp-type-container");
      if (basicEdit) {
        let selectHtml = '<select id="type" required>';
        // 설정과 마스터에서 실시간 병합된 종류 목록 순회
        if (settings.deviceTypes && settings.deviceTypes.length > 0) {
          settings.deviceTypes.forEach(t => {
            selectHtml += '<option value="' + t + '">' + t + '</option>';
          });
        } else {
          // 기본 폴백 데이터
          ["PC", "노트북", "프린터", "3D프린터", "TV", "모니터", "태블릿", "기타"].forEach(t => {
            selectHtml += '<option value="' + t + '">' + t + '</option>';
          });
        }
        selectHtml += '</select>';
        typeContainer.innerHTML = selectHtml;
        
        // 현재 기기의 종류를 선택 값으로 매핑 (없으면 기타)
        document.getElementById("type").value = device["종류"] || "기타";
      } else {
        typeContainer.innerHTML = '<span id="disp-type"></span>';
        document.getElementById("disp-type").innerText = device["종류"] || "-";
      }
      
      // 제조사 바인딩
      const brandContainer = document.getElementById("disp-brand-container");
      if (basicEdit) {
        brandContainer.innerHTML = '<input type="text" id="brand" placeholder="제조사 입력">';
        document.getElementById("brand").value = device["제조사"] || "";
      } else {
        brandContainer.innerHTML = '<span id="disp-brand"></span>';
        document.getElementById("disp-brand").innerText = device["제조사"] || "-";
      }
      
      // 모델명 바인딩
      const modelContainer = document.getElementById("disp-model-container");
      if (basicEdit) {
        modelContainer.innerHTML = '<input type="text" id="model" placeholder="모델명 입력">';
        document.getElementById("model").value = device["모델명"] || "";
      } else {
        modelContainer.innerHTML = '<span id="disp-model"></span>';
        document.getElementById("disp-model").innerText = device["모델명"] || "-";
      }
      
      // 도입일자 바인딩
      const dateContainer = document.getElementById("disp-date-container");
      if (basicEdit) {
        dateContainer.innerHTML = '<input type="date" id="date">';
        let rawDate = device["도입일자"] || "";
        let formattedDate = "";
        if (rawDate) {
          const matches = rawDate.toString().match(/\d+/g);
          if (matches && matches.length >= 3) {
            let year = matches[0];
            let month = matches[1];
            let day = matches[2];
            if (year.length === 2) year = "20" + year;
            if (month.length === 1) month = "0" + month;
            if (day.length === 1) day = "0" + day;
            formattedDate = year + "-" + month + "-" + day;
          } else {
            try {
              const d = new Date(rawDate);
              if (!isNaN(d.getTime())) {
                const y = d.getFullYear();
                const m = ("0" + (d.getMonth() + 1)).slice(-2);
                const dayVal = ("0" + d.getDate()).slice(-2);
                formattedDate = y + "-" + m + "-" + dayVal;
              }
            } catch(e) {}
          }
        }
        document.getElementById("date").value = formattedDate;
      } else {
        dateContainer.innerHTML = '<span id="disp-date"></span>';
        document.getElementById("disp-date").innerText = device["도입일자"] || "-";
      }
      
      // 설치장소 datalist 추천 목록 설정
      const locList = document.getElementById("location-list");
      settings.locations.forEach(loc => {
        const opt = document.createElement("option");
        opt.value = loc;
        locList.appendChild(opt);
      });
      document.getElementById("location").value = device["설치장소"] || "";
      
      // 취급자 datalist 추천 목록 설정
      const mgrList = document.getElementById("manager-list");
      settings.managers.forEach(mgr => {
        const opt = document.createElement("option");
        opt.value = mgr;
        mgrList.appendChild(opt);
      });
      document.getElementById("manager").value = device["취급자"] || "";
      
      // 비고란 채우기
      document.getElementById("notes").value = device["비고"] || "";
      
      // 민감 정보 조건부 세팅 및 노출 제어
      const isPc = (device["종류"] && device["종류"].toString().trim().toUpperCase() === "PC");
      const showSensitive = isPc && settings.pcSensitiveEditAllowed;
      
      if (showSensitive) {
        document.getElementById("sensitive-fields").style.display = "block";
        document.getElementById("ip").value = device["IP"] || "";
        document.getElementById("password1st").value = device["비밀번호(1차)"] || "";
        document.getElementById("password2nd").value = device["비밀번호(2차)"] || "";
      }
      
      // 데이터 비동기 전송 및 로딩 제어
      function submitForm(e) {
        e.preventDefault();
        
        const submitBtn = document.getElementById("submit-btn");
        const spinner = document.getElementById("btn-spinner");
        const btnText = document.getElementById("btn-text");
        const errorBox = document.getElementById("error-box");
        
        submitBtn.disabled = true;
        spinner.style.display = "block";
        btnText.innerText = "저장 중...";
        errorBox.style.display = "none";
        
        const id = device["관리번호"];
        const formData = {
          location: document.getElementById("location").value,
          manager: document.getElementById("manager").value,
          notes: document.getElementById("notes").value
        };
        
        if (showSensitive) {
          formData.ip = document.getElementById("ip").value;
          formData.password1st = document.getElementById("password1st").value;
          formData.password2nd = document.getElementById("password2nd").value;
        }
        
        if (basicEdit) {
          formData.deviceType = document.getElementById("type").value;
          formData.manufacturer = document.getElementById("brand").value;
          formData.modelName = document.getElementById("model").value;
          formData.introDate = document.getElementById("date").value;
        }
        
        // Google Apps Script API 비동기 서버 호출
        google.script.run
          .withSuccessHandler(function(response) {
            submitBtn.disabled = false;
            spinner.style.display = "none";
            btnText.innerText = "수정 내용 저장";
            
            if (response && response.success) {
              // 성공 완료 요약 카드 바인딩
              document.getElementById("sum-id").innerText = id;
              document.getElementById("sum-location").innerText = formData.location;
              document.getElementById("sum-manager").innerText = formData.manager;
              document.getElementById("sum-notes").innerText = formData.notes || "-";
              
              if (showSensitive) {
                document.getElementById("sum-row-ip").style.display = "flex";
                document.getElementById("sum-row-pw1").style.display = "flex";
                document.getElementById("sum-row-pw2").style.display = "flex";
                
                document.getElementById("sum-ip").innerText = formData.ip || "-";
                document.getElementById("sum-pw1").innerText = formData.password1st ? "••••••" : "-";
                document.getElementById("sum-pw2").innerText = formData.password2nd ? "••••••" : "-";
              } else {
                document.getElementById("sum-row-ip").style.display = "none";
                document.getElementById("sum-row-pw1").style.display = "none";
                document.getElementById("sum-row-pw2").style.display = "none";
              }
              
              // 로컬 데이터 객체 업데이트 (다시 수정 화면 재진입 시 정합성 보장)
              device["설치장소"] = formData.location;
              device["취급자"] = formData.manager;
              device["비고"] = formData.notes;
              if (showSensitive) {
                device["IP"] = formData.ip;
                device["비밀번호(1차)"] = formData.password1st;
                device["비밀번호(2차)"] = formData.password2nd;
              }
              if (basicEdit) {
                device["종류"] = formData.deviceType;
                device["제조사"] = formData.manufacturer;
                device["모델명"] = formData.modelName;
                device["도입일자"] = formData.introDate;
                
                document.getElementById("type").value = formData.deviceType;
                document.getElementById("brand").value = formData.manufacturer;
                document.getElementById("model").value = formData.modelName;
                document.getElementById("date").value = formData.introDate;
              }
              
              // 화면 전환
              document.getElementById("form-view").style.display = "none";
              document.getElementById("success-view").style.display = "block";
            } else {
              showError(response ? response.message : "저장 중 에러가 발생했습니다.");
            }
          })
          .withFailureHandler(function(error) {
            submitBtn.disabled = false;
            spinner.style.display = "none";
            btnText.innerText = "수정 내용 저장";
            showError("네트워크 에러가 발생했습니다: " + error.message);
          })
          .updateDevice(id, formData);
      }
      
      function showError(msg) {
        const errorBox = document.getElementById("error-box");
        errorBox.innerText = msg;
        errorBox.style.display = "block";
        window.scrollTo(0, 0);
      }
      
      function showForm() {
        document.getElementById("success-view").style.display = "none";
        document.getElementById("form-view").style.display = "block";
      }
      
      // 입력창 클릭/포커스 시 자동으로 추천 목록(드롭다운)을 강제 팝업하는 함수
      function bindAutoPicker(inputId) {
        const inputEl = document.getElementById(inputId);
        if (!inputEl) return;
        
        const triggerPicker = () => {
          if (typeof inputEl.showPicker === "function") {
            try {
              inputEl.showPicker();
            } catch (err) {
              console.log("showPicker failed: ", err);
            }
          }
        };
        
        inputEl.addEventListener("focus", triggerPicker);
        inputEl.addEventListener("click", triggerPicker);
      }
      
      // 설치장소 및 취급자 입력창에 자동 펼침 리스너 바인딩
      bindAutoPicker("location");
      bindAutoPicker("manager");
      
      // 모든 입력 잠금 상태 처리 (프론트엔드 요소 전체 비활성화)
      if (settings.allEditLocked) {
        const banner = document.getElementById("lock-banner");
        if (banner) {
          banner.innerHTML = '<svg viewBox="0 0 24 24" style="width:20px; height:20px; fill:#f59e0b; flex-shrink:0;"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg><span>현재 시스템이 입력 잠금 상태입니다. 정보 조회만 가능하며 정보 수정은 불가능합니다.</span>';
          banner.style.display = "flex";
        }
        
        // 모든 입력 필드 비활성화
        const inputIds = ["location", "manager", "notes", "ip", "password1st", "password2nd", "type", "brand", "model", "date"];
        inputIds.forEach(id => {
          const el = document.getElementById(id);
          if (el) {
            el.disabled = true;
            el.style.opacity = "0.6";
            el.style.cursor = "not-allowed";
          }
        });
        
        // 저장 버튼 비활성화 및 레이블 변경
        const submitBtn = document.getElementById("submit-btn");
        if (submitBtn) {
          submitBtn.disabled = true;
          submitBtn.style.opacity = "0.5";
          submitBtn.style.cursor = "not-allowed";
          submitBtn.style.boxShadow = "none";
          submitBtn.style.background = "var(--input-border)";
        }
        const btnText = document.getElementById("btn-text");
        if (btnText) {
          btnText.innerText = "수정 잠금 활성화됨";
        }
      }
      
      // --- 라벨 보관함 및 모아 인쇄 비즈니스 로직 ---
      const CART_KEY = 'ithings_print_cart';
      
      function getCart() {
        try {
          const raw = localStorage.getItem(CART_KEY);
          return raw ? JSON.parse(raw) : [];
        } catch(e) {
          return [];
        }
      }
      
      function saveCart(cart) {
        try {
          localStorage.setItem(CART_KEY, JSON.stringify(cart));
        } catch(e) {}
        updateCartButtons();
      }
      
      function updateCartButtons() {
        const cart = getCart();
        const countEl = document.getElementById('cart-count');
        if (countEl) {
          countEl.innerText = cart.length;
        }
      }
      
      function addToCart() {
        const cart = getCart();
        const exists = cart.some(item => item['관리번호'] === device['관리번호']);
        if (exists) {
          alert('이미 보관함에 담겨 있는 기기입니다.');
          return;
        }
        
        // 현재 화면에 입력된 최신 필드 값을 가져와 기기 데이터와 합성
        const updatedItem = {
          '관리번호': device['관리번호'],
          '종류': document.getElementById('type') ? document.getElementById('type').value : (device['종류'] || '기타'),
          '설치장소': document.getElementById('location') ? document.getElementById('location').value : (device['설치장소'] || ''),
          '취급자': document.getElementById('manager') ? document.getElementById('manager').value : (device['취급자'] || ''),
          '비고': document.getElementById('notes') ? document.getElementById('notes').value : (device['비고'] || ''),
          '제조사': document.getElementById('brand') ? document.getElementById('brand').value : (device['제조사'] || ''),
          '모델명': document.getElementById('model') ? document.getElementById('model').value : (device['모델명'] || ''),
          '도입일자': document.getElementById('date') ? document.getElementById('date').value : (device['도입일자'] || ''),
          'QR링크': device['QR링크'] || (settings.webAppUrl + '?id=' + encodeURIComponent(device['관리번호']))
        };
        
        cart.push(updatedItem);
        saveCart(cart);
        alert('보관함에 기기를 추가했습니다.');
      }
      
      function openCartModal() {
        const modal = document.getElementById('cart-modal');
        if (modal) {
          renderCartItems();
          modal.classList.add('active');
        }
      }
      
      function closeCartModal() {
        const modal = document.getElementById('cart-modal');
        if (modal) {
          modal.classList.remove('active');
        }
      }
      
      function closeCartModalOnBackdrop(e) {
        if (e.target === document.getElementById('cart-modal')) {
          closeCartModal();
        }
      }
      
      function renderCartItems() {
        const cart = getCart();
        const container = document.getElementById('cart-items-container');
        if (!container) return;
        
        container.innerHTML = '';
        
        if (cart.length === 0) {
          container.innerHTML = '<div class="cart-empty">보관함이 비어 있습니다.</div>';
          return;
        }
        
        cart.forEach(item => {
          const div = document.createElement('div');
          div.className = 'cart-item';
          
          const detailUrl = settings.webAppUrl + '?id=' + encodeURIComponent(item['관리번호']);
          
          div.setAttribute('data-url', detailUrl);
          div.setAttribute('data-id', item['관리번호']);
          
          div.innerHTML = '<div class="cart-item-info">' +
              '<div class="cart-item-id">' + item['관리번호'] + '</div>' +
              '<div class="cart-item-meta">' + item['종류'] + ' | ' + (item['설치장소'] || '-') + '</div>' +
            '</div>' +
            '<button class="cart-item-delete">&times;</button>';
            
          div.querySelector('.cart-item-info').onclick = function() {
            window.location.href = div.getAttribute('data-url');
          };
          div.querySelector('.cart-item-delete').onclick = function() {
            removeFromCart(div.getAttribute('data-id'));
          };
          
          container.appendChild(div);
        });
      }
      
      function removeFromCart(id) {
        let cart = getCart();
        cart = cart.filter(item => item['관리번호'] !== id);
        saveCart(cart);
        renderCartItems();
      }
      
      function clearCart() {
        if (confirm('보관함을 모두 비우시겠습니까?')) {
          saveCart([]);
          renderCartItems();
        }
      }
      
      function printCartLabels() {
        const cart = getCart();
        if (cart.length === 0) {
          alert('출력할 기기가 보관함에 없습니다.');
          return;
        }
        
        const printSec = document.getElementById('print-section');
        if (!printSec) return;
        
        printSec.innerHTML = '';
        
        // 최대 A4 10칸 규격에 맞춰 10개까지만 출력 제한
        const printCount = Math.min(cart.length, 10);
        
        for (let i = 0; i < printCount; i++) {
          const item = cart[i];
          const labelDiv = document.createElement('div');
          labelDiv.className = 'print-label';
          
          // 2열 5행 절대 좌표 계산
          const col = i % 2;
          const row = Math.floor(i / 2);
          const leftPos = 9 + col * (77 + 20);
          const topPos = [19, 69, 121, 173, 225][row];
          
          labelDiv.style.position = 'absolute';
          labelDiv.style.left = leftPos + 'mm';
          labelDiv.style.top = topPos + 'mm';
          labelDiv.style.width = '77mm';
          labelDiv.style.height = '40mm';
          
          const qrData = item['QR링크'] || (settings.webAppUrl + '?id=' + encodeURIComponent(item['관리번호']));
          
          // 로컬 qrcode.js 기반 동적 QR 생성
          const qrContainer = document.createElement('div');
          qrContainer.className = 'label-qr';
          
          new QRCode(qrContainer, {
            text: qrData,
            width: 90,
            height: 90,
            correctLevel: QRCode.CorrectLevel.L
          });
          
          labelDiv.innerHTML = '<div class="label-details">' +
              '<div class="label-title">' + item['관리번호'] + '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">품명:</span>' +
                '<span class="label-val">' + item['종류'] + ' (' + (item['모델명'] || '-') + ')</span>' +
              '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">장소:</span>' +
                '<span class="label-val">' + (item['설치장소'] || '-') + '</span>' +
              '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">담당:</span>' +
                '<span class="label-val">' + (item['취급자'] || '-') + '</span>' +
              '</div>' +
            '</div>';
          
          labelDiv.appendChild(qrContainer);
          printSec.appendChild(labelDiv);
        }
        
        window.print();
      }
      
      function saveCartAsPdf(event) {
        const cart = getCart();
        if (cart.length === 0) {
          alert('출력할 기기가 보관함에 없습니다.');
          return;
        }
        
        const printSec = document.getElementById('print-section');
        if (!printSec) return;
        
        printSec.innerHTML = '';
        
        const printCount = Math.min(cart.length, 10);
        
        for (let i = 0; i < printCount; i++) {
          const item = cart[i];
          const labelDiv = document.createElement('div');
          labelDiv.className = 'print-label';
          
          // 2열 5행 절대 좌표 계산
          const col = i % 2;
          const row = Math.floor(i / 2);
          const leftPos = 9 + col * (77 + 20);
          const topPos = [19, 69, 121, 173, 225][row];
          
          labelDiv.style.position = 'absolute';
          labelDiv.style.left = leftPos + 'mm';
          labelDiv.style.top = topPos + 'mm';
          labelDiv.style.width = '77mm';
          labelDiv.style.height = '40mm';
          
          const qrData = item['QR링크'] || (settings.webAppUrl + '?id=' + encodeURIComponent(item['관리번호']));
          
          // 로컬 qrcode.js 기반 동적 QR 생성
          const qrContainer = document.createElement('div');
          qrContainer.className = 'label-qr';
          
          new QRCode(qrContainer, {
            text: qrData,
            width: 90,
            height: 90,
            correctLevel: QRCode.CorrectLevel.L
          });
          
          labelDiv.innerHTML = '<div class="label-details">' +
              '<div class="label-title">' + item['관리번호'] + '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">품명:</span>' +
                '<span class="label-val">' + item['종류'] + ' (' + (item['모델명'] || '-') + ')</span>' +
              '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">장소:</span>' +
                '<span class="label-val">' + (item['설치장소'] || '-') + '</span>' +
              '</div>' +
              '<div class="label-row">' +
                '<span class="label-key">담당:</span>' +
                '<span class="label-val">' + (item['취급자'] || '-') + '</span>' +
              '</div>' +
            '</div>';
          
          labelDiv.appendChild(qrContainer);
          printSec.appendChild(labelDiv);
        }
        
        const pdfBtn = event.currentTarget;
        const originalText = pdfBtn.innerText;
        pdfBtn.disabled = true;
        pdfBtn.innerText = '생성 중...';
        
        const originalDisplay = printSec.style.display;
        const originalPosition = printSec.style.position;
        const originalZIndex = printSec.style.zIndex;
        const originalLeft = printSec.style.left;
        const originalTop = printSec.style.top;
        
        // 캡처하는 동안 요소를 일시적으로 고정 보임 처리하여 렌더링 누락 차단
        printSec.style.display = 'block';
        printSec.style.position = 'fixed';
        printSec.style.left = '0px';
        printSec.style.top = '0px';
        printSec.style.zIndex = '9999';
        
        // 브라우저가 변경된 fixed 레이아웃을 갱신(Reflow)할 시간을 조금 벌어준 뒤 html2canvas 실행
        setTimeout(() => {
          html2canvas(printSec, {
            scale: 2, 
            useCORS: true, 
            logging: true,
            scrollX: 0,
            scrollY: 0,
            windowWidth: 794,
            windowHeight: 1123,
            width: 794,
            height: 1123,
            backgroundColor: '#ffffff'
          }).then(canvas => {
            // 캡처 후 레이아웃 원상 복구
            printSec.style.display = originalDisplay;
            printSec.style.position = originalPosition;
            printSec.style.zIndex = originalZIndex;
            printSec.style.left = originalLeft;
            printSec.style.top = originalTop;
            
            // Canvas 이미지 데이터 추출
            const imgData = canvas.toDataURL('image/jpeg', 1.0);
            
            // window.jspdf.jsPDF 꺼내오기
            const { jsPDF } = window.jspdf || window;
            if (!jsPDF) {
              alert('PDF 생성 라이브러리가 존재하지 않습니다.');
              pdfBtn.disabled = false;
              pdfBtn.innerText = originalText;
              return;
            }
            
            const doc = new jsPDF({
              orientation: 'portrait',
              unit: 'mm',
              format: 'a4'
            });
            
            // A4 영역에 캔버스 삽입
            doc.addImage(imgData, 'JPEG', 0, 0, 210, 297);
            
            const filename = 'ithings_labels_' + new Date().toISOString().slice(0, 10) + '.pdf';
            doc.save(filename);
            
            pdfBtn.disabled = false;
            pdfBtn.innerText = originalText;
          }).catch(err => {
            console.error('Capture error:', err);
            alert('PDF 캡처 중 오류가 발생했습니다: ' + err.toString());
            printSec.style.display = originalDisplay;
            printSec.style.position = originalPosition;
            printSec.style.zIndex = originalZIndex;
            printSec.style.left = originalLeft;
            printSec.style.top = originalTop;
            pdfBtn.disabled = false;
            pdfBtn.innerText = originalText;
          });
        }, 200);
      }
      
      // 초기 버튼 수 동기화
      updateCartButtons();
    </script>
  </body>
</html>`;
}

/**
 * 접속 및 에러 처리 전용 HTML 에러 카드 화면 생성 함수
 * @param {string} message - 에러 알림 내용
 * @return {string} HTML 콘텐츠
 */
function createErrorHtml_(message) {
  return `<!DOCTYPE html>
<html>
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, shrink-to-fit=no">
    <title>에러 - 기자재 관리 시스템</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;700&family=Noto+Sans+KR:wght@300;400;700&display=swap" rel="stylesheet">
    <style>
      :root {
        --bg-color: #0f172a;
        --card-bg: #1e293b;
        --text-color: #f1f5f9;
        --text-muted: #94a3b8;
        --danger-color: #ef4444;
        --danger-gradient: linear-gradient(135deg, #f87171, #ef4444);
        --font-family: 'Outfit', 'Noto Sans KR', sans-serif;
      }
      body {
        background-color: var(--bg-color);
        color: var(--text-color);
        font-family: var(--font-family);
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 100vh;
        box-sizing: border-box;
      }
      .card {
        background-color: var(--card-bg);
        border-radius: 24px;
        padding: 40px 32px;
        width: 90%;
        max-width: 440px;
        text-align: center;
        box-shadow: 0 20px 25px -5px rgba(0, 0, 0, 0.3), 0 10px 10px -5px rgba(0, 0, 0, 0.2);
        border: 1px solid rgba(255, 255, 255, 0.05);
        animation: fadeIn 0.5s ease-out;
      }
      @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
      }
      .icon {
        display: inline-flex;
        justify-content: center;
        align-items: center;
        width: 72px;
        height: 72px;
        border-radius: 50%;
        background: var(--danger-gradient);
        margin-bottom: 24px;
        box-shadow: 0 8px 16px rgba(239, 68, 68, 0.3);
      }
      .icon svg {
        width: 36px;
        height: 36px;
        fill: white;
      }
      h1 {
        font-size: 22px;
        font-weight: 700;
        margin: 0 0 12px 0;
      }
      p {
        color: var(--text-muted);
        font-size: 15px;
        line-height: 1.6;
        margin: 0 0 24px 0;
        word-break: break-all;
      }
      .footer {
        font-size: 12px;
        color: var(--text-muted);
        opacity: 0.7;
      }
    </style>
  </head>
  <body>
    <div class="card">
      <div class="icon">
        <svg viewBox="0 0 24 24">
          <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/>
        </svg>
      </div>
      <h1>알림 및 오류</h1>
      <p>${message}</p>
      <div class="footer">학교 기자재 관리 시스템</div>
    </div>
  </body>
</html>`;
}

/**
 * [초기 설정] 필요한 시트들과 샘플 헤더 정보를 구성하는 기초 초기화 함수
 * 최초 1회 실행해주어야 시스템이 제대로 작동합니다.
 */
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. [마스터] 시트 초기 구성
  var masterSheet = ss.getSheetByName("마스터");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("마스터");
  }
  
  // 마스터 헤더 정의 (보안 설정에 따른 조건부 수정 필드 및 관리책임자 필드 포함)
  var masterHeaders = [
    "연번", "관리번호", "설치장소", "관리책임자", "취급자", "종류", "제조사", "모델명", "도입일자", "비고", "라벨인쇄", "IP", "비밀번호(1차)", "비밀번호(2차)", "QR링크", "QR이미지", "최종수정일"
  ];
  
  masterSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders])
    .setFontWeight("bold")
    .setBackground("#f3f4f6");
  
  // 테스트용 기초 샘플 데이터 삽입 (기존 데이터가 하나도 없을 때만 삽입)
  if (masterSheet.getLastRow() === 1) {
    masterSheet.appendRow([
      1, "EQ-2026-001", "컴퓨터실", "교사 홍길동", "교사 홍길동", "노트북", "삼성전자", "갤럭시북4", "2026-03-02", "배터리 수명 체크 필요", false, "", "", "", "", "", ""
    ]);
    masterSheet.appendRow([
      2, "EQ-2026-002", "과학실", "교사 김철수", "교사 김철수", "3D프린터", "신도리코", "3DWOX 1", "2025-11-15", "노즐 정비 완료", false, "", "", "", "", "", ""
    ]);
    masterSheet.appendRow([
      3, "EQ-2026-003", "행정실", "교사 이영희", "교사 이영희", "PC", "LG전자", "울트라PC", "2026-01-10", "행정 업무용 PC", false, "192.168.10.22", "admin@2026", "sec#7890", "", "", ""
    ]);
  }

  // 라벨인쇄 열에 체크박스 유효성 검사 적용 (기존 데이터 행 포함 전체 일괄 설정)
  var printColIdx = masterHeaders.indexOf("라벨인쇄") + 1;
  if (printColIdx > 0) {
    var lastRow = masterSheet.getLastRow();
    var checkboxRange = masterSheet.getRange(2, printColIdx, lastRow >= 2 ? lastRow - 1 : 1, 1);
    var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    checkboxRange.setDataValidation(rule);
  }
  
  // 2. [설정] 시트 초기 구성
  var settingsSheet = ss.getSheetByName("설정");
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet("설정");
  }
  
  settingsSheet.clear();
  var settingsHeaders = ["항목", "값"];
  settingsSheet.getRange(1, 1, 1, settingsHeaders.length).setValues([settingsHeaders])
    .setFontWeight("bold")
    .setBackground("#e0e7ff");
    
  var defaultSettings = [
    ["기관명", "비아초등학교"],
    ["웹앱URL", "https://script.google.com/macros/s/배포ID/exec (여기에 실제 배포된 URL을 입력하세요)"],
    ["설치장소목록", "교무실,행정실,1-1,1-2,1-3,과학실,컴퓨터실"],
    ["취급자목록", "홍길동,김철수,이영희,박민수"],
    ["PC민감정보수정허용", "N"],
    ["기본정보수정허용", "N"],
    ["모든입력잠금", "N"],
    ["관리자이름", "교사 이영희"]
  ];
  
  settingsSheet.getRange(2, 1, defaultSettings.length, 2).setValues(defaultSettings);
  settingsSheet.autoResizeColumns(1, 2);
  
  // 3. [기기관리라벨] 시트 초기 구성
  var labelSheet = ss.getSheetByName("기기관리라벨");
  if (!labelSheet) {
    labelSheet = ss.insertSheet("기기관리라벨");
  }
  
  // 무한 실행대기를 예방하기 위해 blocking alert을 non-blocking toast(우측 하단 알림) 및 콘솔 로그로 대체
  ss.toast("초기화 완료: '마스터' 및 '설정' 시트가 생성되었습니다. 설정 값을 확인하신 뒤 웹앱 배포를 진행해주세요.", "시스템 알림");
  console.log("초기화 완료: '마스터' 및 '설정' 시트의 생성이 완료되었습니다.");
}

/**
 * [설정] 시트의 웹앱URL과 마스터 시트의 관리번호를 조합하여 QR링크 및 이미지를 일괄 생성 및 동적 업데이트하는 함수
 */
function refreshQrLinks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("마스터");
  var settings = getSettings_();
  var ui = SpreadsheetApp.getUi();
  
  if (!masterSheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다.");
  }
  
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("마스터 시트에 데이터가 존재하지 않습니다.");
  }
  
  var headerMapping = getHeaderMapping_(masterSheet);
  var idCol = headerMapping["관리번호"];
  var qrLinkCol = headerMapping["QR링크"];
  var qrImgCol = headerMapping["QR이미지"];
  var printCol = headerMapping["라벨인쇄"];
  
  if (!idCol || !qrLinkCol || !qrImgCol) {
    throw new Error("필수 헤더가 누락되었습니다. (관리번호, QR링크, QR이미지 필요)");
  }
  
  var webAppUrl = settings.webAppUrl;
  if (!webAppUrl || webAppUrl.indexOf("http") !== 0 || webAppUrl.indexOf("배포ID") !== -1) {
    throw new Error("[설정] 시트에서 '웹앱URL'을 실제 주소로 정상 수정하여 입력한 뒤 다시 실행해 주세요.");
  }
  
  // 1. 유저 경고 및 모드 결정 팝업
  var response = ui.alert(
    "QR 링크 및 이미지 일괄 갱신",
    "마스터 시트의 [모든 기기]의 QR코드를 처음부터 다시 갱신하시겠습니까?\n\n" +
    "[예] ➔ 전체 기기 일괄 재갱신 (기존 정보 덮어씀)\n" +
    "[아니오] ➔ 신규 기기(QR이 비어 있는 칸) 및 라벨인쇄 체크된 기기만 선별 갱신\n" +
    "[취소] ➔ 작업 취소 및 중단",
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  if (response === ui.Button.CANCEL) {
    ss.toast("QR 코드 갱신 작업이 취소되었습니다.", "시스템 알림");
    return;
  }
  
  var isForceAll = (response === ui.Button.YES);
  
  var idValues = masterSheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  var printValues = printCol ? masterSheet.getRange(2, printCol, lastRow - 1, 1).getValues() : [];
  var qrLinkValues = masterSheet.getRange(2, qrLinkCol, lastRow - 1, 1).getValues();
  
  var updateCount = 0;
  
  // 순회하며 링크 생성 및 IMAGE 폼 수식 적용
  for (var i = 0; i < idValues.length; i++) {
    var id = idValues[i][0].toString().trim();
    if (!id) continue;
    
    var rowNum = i + 2;
    var currentLink = qrLinkValues[i][0].toString().trim();
    var isChecked = printValues.length > 0 ? (printValues[i][0] === true || printValues[i][0].toString().toUpperCase() === "TRUE") : false;
    
    // 강제 전체 갱신이 아닌 경우: 이미 QR링크가 있고, 라벨인쇄 체크가 꺼져 있다면 건너뜀
    if (!isForceAll) {
      if (currentLink && !isChecked) {
        continue;
      }
    }
    
    var finalQrLink = webAppUrl + (webAppUrl.indexOf("?") === -1 ? "?" : "&") + "id=" + encodeURIComponent(id);
    
    // QR 링크 텍스트 셀 업데이트
    masterSheet.getRange(rowNum, qrLinkCol).setValue(finalQrLink);
    
    // QR 이미지 수식 적용 (IMAGE API 사용)
    var cellA1 = masterSheet.getRange(rowNum, qrLinkCol).getA1Notation();
    var formula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & ENCODEURL(' + cellA1 + '))';
    masterSheet.getRange(rowNum, qrImgCol).setFormula(formula);
    
    updateCount++;
  }
  
  var msg = isForceAll 
    ? "전체 기기(" + updateCount + "대)의 QR코드 갱신이 완료되었습니다." 
    : "신규 및 선택된 기기(" + updateCount + "대)의 QR코드 선별 갱신이 완료되었습니다.";
  ss.toast(msg, "시스템 알림");
  console.log(msg);
}

/**
 * 마스터 데이터를 바탕으로 애니라벨 10칸(A4 용지, 88.9mm x 52mm) 규격에 맞춰
 * 구글 스프레드시트 인쇄 기능을 통해 라벨지에 인쇄할 수 있도록 레이아웃을 구성하는 함수
 */
function createLabelSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("마스터");
  var ui = SpreadsheetApp.getUi();
  
  if (!masterSheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다.");
  }
  
  var labelSheet = ss.getSheetByName("기기관리라벨");
  if (labelSheet) {
    labelSheet.clear();
  } else {
    labelSheet = ss.insertSheet("기기관리라벨");
  }
  
  var lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("출력할 기기 데이터가 마스터 시트에 존재하지 않습니다.");
  }
  
  var headerMapping = getHeaderMapping_(masterSheet);
  var idCol = headerMapping["관리번호"];
  var typeCol = headerMapping["종류"];
  var managerCol = headerMapping["관리책임자"];
  var operatorCol = headerMapping["취급자"];
  var qrLinkCol = headerMapping["QR링크"];
  var qrImgCol = headerMapping["QR이미지"];
  var printCol = headerMapping["라벨인쇄"];
  
  if (!idCol || !typeCol || !managerCol || !operatorCol || !qrLinkCol || !qrImgCol) {
    throw new Error("마스터 시트에 필수 컬럼이 부족합니다. (관리번호, 종류, 관리책임자, 취급자, QR링크, QR이미지 필요)");
  }
  
  var data = masterSheet.getRange(2, 1, lastRow - 1, masterSheet.getLastColumn()).getValues();
  var printValues = printCol ? masterSheet.getRange(2, printCol, lastRow - 1, 1).getValues() : [];
  
  // 1. 라벨인쇄가 체크된 대상 기기 선별 필터링
  var targetData = [];
  for (var i = 0; i < data.length; i++) {
    var id = data[i][idCol - 1];
    var isChecked = printValues.length > 0 ? (printValues[i][0] === true || printValues[i][0].toString().toUpperCase() === "TRUE") : false;
    if (id && isChecked) {
      targetData.push({
        rowData: data[i],
        masterRowIndex: i + 2
      });
    }
  }
  
  // 2. 만약 체크된 기기가 하나도 없는 경우: 전체 기기 인쇄 여부 확인
  if (targetData.length === 0) {
    var confirmAll = ui.alert(
      "인쇄 대상 기기 확인",
      "체크된 기기가 없습니다. 마스터 시트의 [전체 기기]를 대상으로 인쇄용 라벨을 생성하시겠습니까?\n\n" +
      "(취소를 선택하면 작업이 중단되며, 기존 라벨 시트의 내용이 보존됩니다.)",
      ui.ButtonSet.OK_CANCEL
    );
    
    if (confirmAll !== ui.Button.OK) {
      ss.toast("라벨 생성이 취소되었습니다. 출력할 기기를 선택하고 다시 실행해 주세요.", "시스템 알림");
      return;
    }
    
    // 전체 기기를 대상으로 라벨링 리스트 채움
    for (var i = 0; i < data.length; i++) {
      var id = data[i][idCol - 1];
      if (id) {
        targetData.push({
          rowData: data[i],
          masterRowIndex: i + 2
        });
      }
    }
  }
  
  var settings = getSettings_();
  var schoolName = settings.schoolName || "비아초등학교";
  var webAppUrl = settings.webAppUrl;
  var hasGeneratedQr = false;
  
  // 3. 인쇄 대상 중 QR링크가 누락된 기기 자동 감지 및 발급
  for (var idx = 0; idx < targetData.length; idx++) {
    var item = targetData[idx];
    var qrLinkVal = item.rowData[qrLinkCol - 1].toString().trim();
    var id = item.rowData[idCol - 1].toString().trim();
    
    if (!qrLinkVal) {
      if (!webAppUrl || webAppUrl.indexOf("http") !== 0 || webAppUrl.indexOf("배포ID") !== -1) {
        throw new Error("[설정] 시트에서 '웹앱URL'을 실제 주소로 정상 수정하여 입력해야 새로운 기기의 QR을 발급할 수 있습니다.");
      }
      
      var finalQrLink = webAppUrl + (webAppUrl.indexOf("?") === -1 ? "?" : "&") + "id=" + encodeURIComponent(id);
      
      // 마스터 시트 즉시 기입
      masterSheet.getRange(item.masterRowIndex, qrLinkCol).setValue(finalQrLink);
      
      var cellA1 = masterSheet.getRange(item.masterRowIndex, qrLinkCol).getA1Notation();
      var formula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" & ENCODEURL(' + cellA1 + '))';
      masterSheet.getRange(item.masterRowIndex, qrImgCol).setFormula(formula);
      
      // 메모리 내 배열 값 동기화
      item.rowData[qrLinkCol - 1] = finalQrLink;
      hasGeneratedQr = true;
    }
  }
  
  if (hasGeneratedQr) {
    ss.toast("신규 기기의 누락된 QR코드를 자동 생성하여 마스터 시트에 반영했습니다.", "시스템 알림");
  }
  
  var activeLabelCount = targetData.length;
  
  // 필요한 행과 열 확보 및 초과 영역 제거하여 Out of Bounds 에러 방지
  var requiredRows = getRequiredRowsForLabels_(activeLabelCount);
  if (requiredRows > 0) {
    var maxRows = labelSheet.getMaxRows();
    if (maxRows < requiredRows) {
      labelSheet.insertRowsAfter(maxRows, requiredRows - maxRows);
    } else if (maxRows > requiredRows) {
      labelSheet.deleteRows(requiredRows + 1, maxRows - requiredRows);
    }
  }
  
  var maxCols = labelSheet.getMaxColumns();
  if (maxCols < 15) {
    labelSheet.insertColumnsAfter(maxCols, 15 - maxCols);
  } else if (maxCols > 15) {
    labelSheet.deleteColumns(16, maxCols - 15);
  }
  
  // 1. 여백 및 격자 너비(Column Width) 정비 (크롬 기본 인쇄 여백 17.78mm 고려한 뺄셈 설계)
  // Column A: 좌측 페이지 여백 (크롬 17.78mm 선적용으로 최소 1px)
  labelSheet.setColumnWidth(1, 1);
  
  // Label 1 (Column B~G): 총 336px (약 88.9mm)
  labelSheet.setColumnWidth(2, 17);  // 좌측 안전 여백 (Margin)
  labelSheet.setColumnWidth(3, 80);  // 종류, 관리번호 라벨 항목명
  labelSheet.setColumnWidth(4, 44);  // 직위 (교사 등)
  labelSheet.setColumnWidth(5, 90);  // 이름
  labelSheet.setColumnWidth(6, 88);  // QR코드
  labelSheet.setColumnWidth(7, 17);  // 우측 안전 여백 (Margin)
  
  // Column H: 라벨 간 가로 갭 (9.2mm ≈ 35px)
  labelSheet.setColumnWidth(8, 35);
  
  // Label 2 (Column I~N): 총 336px (약 88.9mm)
  labelSheet.setColumnWidth(9, 17);  // 좌측 안전 여백 (Margin)
  labelSheet.setColumnWidth(10, 80); // 종류, 관리번호 라벨 항목명
  labelSheet.setColumnWidth(11, 44); // 직위 (교사 등)
  labelSheet.setColumnWidth(12, 90); // 이름
  labelSheet.setColumnWidth(13, 88); // QR코드
  labelSheet.setColumnWidth(14, 17); // 우측 안전 여백 (Margin)
  
  // Column O: 우측 페이지 여백 (크롬 17.78mm 선적용으로 최소 1px)
  labelSheet.setColumnWidth(15, 1);
  
  // 2. 상단 페이지 여백 설정 (공식 20.5mm - 크롬기본 19.05mm = 1.45mm ≈ 5px)
  labelSheet.setRowHeight(1, 5);
  
  var curRow = 2; // 데이터는 2행(상단 여백 아래)부터 작성 시작
  var colLayouts = [
    { start: 2, contentStart: 3, qr: 6 }, // 1열 라벨 (B~G, 내용 C~F, QR F)
    { start: 9, contentStart: 10, qr: 13 } // 2열 라벨 (I~N, 내용 J~M, QR M)
  ];
  var labelCount = 0; // 실제로 생성된 라벨의 개수 추적 카운터
  
  // 그리드 라인 표시 설정 (false를 지정하여 숨김 해제 = 눈금선 표시)
  labelSheet.setHiddenGridlines(false);
  
  for (var i = 0; i < targetData.length; i++) {
    var item = targetData[i];
    var rowData = item.rowData;
    var id = rowData[idCol - 1];
    var type = rowData[typeCol - 1];
    // 설정 시트의 관리자이름(managerName)이 존재할 경우 일괄 덮어쓰기 적용 (오차 보호막)
    var manager = settings.managerName || rowData[managerCol - 1];
    var operator = rowData[operatorCol - 1];
    
    if (!id) continue;
    
    // 10칸(5줄) 단위 페이지 경계 반복 여백 동적 삽입
    if (labelCount > 0 && labelCount % 10 === 0) {
      // 1. 이전 페이지 하단 여백 행 (1페이지 초과 방지를 위해 최소 1px로 여백 제거)
      var bottomPaddingRow = curRow;
      labelSheet.setRowHeight(bottomPaddingRow, 1);
      labelSheet.getRange(bottomPaddingRow, 1, 1, 15).clearFormat();
      
      // 2. 다음 페이지 상단 여백 행 (공식 20.5mm - 크롬기본 19.05mm = 1.45mm ≈ 5px)
      // 2페이지부터 윗 여백 1.8mm (약 7px) 추가 요청 반영: 5px + 7px = 12px -> 3px 추가 후 1px 축소하여 14px -> 7px 축소하여 7px -> 다시 1.5mm(약 6px) 늘려 13px -> 1mm(약 4px) 줄여 9px -> 0.5mm(약 2px) 늘려 11px -> 1px 추가하여 최종 12px -> 1px 미세 조정을 위해 10px로 조정 (pt 반올림 임계값 돌파 시도)
      var nextTopPaddingRow = curRow + 1;
      labelSheet.setRowHeight(nextTopPaddingRow, 10);
      labelSheet.getRange(nextTopPaddingRow, 1, 1, 15).clearFormat();
      
      curRow += 2; // 여백 행 2개 추가분 증가
    }
    
    var layout = colLayouts[labelCount % 2];
    var r = curRow;
    var c = layout.start;
    var cc = layout.contentStart;
    var qc = layout.qr;
    
    // 행 높이 조절
    var isLastRowOfPage = (Math.floor((labelCount % 10) / 2) === 4);
    if (isLastRowOfPage) {
      // 5번째 줄 (9, 10칸) - A4 페이지 초과 밀림 방지를 위해 높이를 45mm(170px), 내부 표를 42mm(159px)로 축소
      labelSheet.setRowHeight(r, 5);     // 상단 안전 여백 (Row Margin)
      labelSheet.setRowHeight(r+1, 24);  // 기관명 (학교명)
      labelSheet.setRowHeight(r+2, 32);  // 종류
      labelSheet.setRowHeight(r+3, 34);  // 관리 번호
      labelSheet.setRowHeight(r+4, 34);  // 관리책임자
      labelSheet.setRowHeight(r+5, 35);  // 취급자
      labelSheet.setRowHeight(r+6, 6);   // 하단 안전 여백 (Row Margin)
    } else {
      // 일반 줄 - 높이 52mm(197px)
      labelSheet.setRowHeight(r, 13);    // 상단 안전 여백 (Row Margin)
      labelSheet.setRowHeight(r+1, 28);  // 기관명 (학교명)
      labelSheet.setRowHeight(r+2, 34);  // 종류
      labelSheet.setRowHeight(r+3, 36);  // 관리 번호
      labelSheet.setRowHeight(r+4, 36);  // 관리책임자
      labelSheet.setRowHeight(r+5, 36);  // 취급자
      labelSheet.setRowHeight(r+6, 14);  // 하단 안전 여백 (Row Margin)
    }
    
    // 1. 물리적 라벨 영역 배경색 및 연한 칼선 테두리 설정 (B~G열, r~r+6행)
    var cardRange = labelSheet.getRange(r, c, 7, 6);
    cardRange.setBorder(true, true, true, true, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    cardRange.setBackground("#ffffff");
    cardRange.setFontFamily("Noto Sans KR");
    
    // 2. 내부 데이터 표 영역 테두리 설정 (C~F열, r+1~r+5행) - 안전 여백 확보
    var contentRange = labelSheet.getRange(r+1, cc, 5, 4);
    contentRange.setBorder(true, true, true, true, true, true, "#475569", SpreadsheetApp.BorderStyle.SOLID);
    
    // 3. 기관명 (학교명)
    var titleCell = labelSheet.getRange(r+1, cc, 1, 4);
    titleCell.merge().setValue(schoolName)
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#f8fafc");
      
    // 4. 종류
    labelSheet.getRange(r+2, cc).setValue("종류")
      .setFontWeight("bold")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#f1f5f9");
    
    var typeCell = labelSheet.getRange(r+2, cc+1, 1, 3); // D~F열 병합
    typeCell.merge().setValue(type)
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // 5. 관리 번호
    labelSheet.getRange(r+3, cc).setValue("관리 번호")
      .setFontWeight("bold")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#f1f5f9");
      
    var idCell = labelSheet.getRange(r+3, cc+1, 1, 2); // D~E열 병합
    idCell.merge().setValue(id)
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // 6. 관리책임자
    labelSheet.getRange(r+4, cc).setValue("관리책임자")
      .setFontWeight("bold")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#f1f5f9");
      
    var mParts = splitNameAndTitle_(manager);
    labelSheet.getRange(r+4, cc+1).setValue(mParts.title)
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
    labelSheet.getRange(r+4, cc+2).setValue(mParts.name)
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // 7. 취급자
    labelSheet.getRange(r+5, cc).setValue("취급자")
      .setFontWeight("bold")
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBackground("#f1f5f9");
      
    var oParts = splitNameAndTitle_(operator);
    labelSheet.getRange(r+5, cc+1).setValue(oParts.title)
      .setFontSize(9)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
    labelSheet.getRange(r+5, cc+2).setValue(oParts.name)
      .setFontSize(10)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // 8. QR코드 (F3:F5 또는 M3:M5 병합하여 차트 API 적용, 여백 chld=L|2 옵션 추가)
    var masterRowIndex = item.masterRowIndex;
    var masterCellA1 = masterSheet.getRange(masterRowIndex, qrLinkCol).getA1Notation();
    var qrFormula = '=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=150x150&ecc=L&margin=2&data=" & ENCODEURL(\'마스터\'!' + masterCellA1 + '))';
    
    var qrCell = labelSheet.getRange(r+3, qc, 3, 1);
    qrCell.merge().setFormula(qrFormula)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
      
    // 두 번째 열 라벨까지 다 그리면 한 행 아래로 줄바꿈 (애니라벨 10칸은 세로 라벨 간 갭이 0mm이므로 간격 행 없이 7행씩 연속 배치)
    if (labelCount % 2 === 1) {
      curRow += 7; // 라벨 카드(7개 행)만큼 증가하여 다음 줄로 이동
    }
    labelCount++;
  }
  
  // 시트 활성화 및 완성 알림 (무한 대기 시간초과 방지 위해 toast 처리)
  labelSheet.activate();
  ss.toast("라벨 시트 생성 완료: '기기관리라벨' 시트에 애니라벨 10칸(88.9mm x 52mm) 규격 및 안전 여백이 적용된 라벨 배치가 완료되었습니다. 구글 스프레드시트의 인쇄 기능(Ctrl + P)을 사용하여 A4 라벨지에 인쇄해 주십시오.", "시스템 알림");
  
  // 4. 출력 완료 후 체크박스 일괄 해제 확인 팝업
  var uncheckConfirm = ui.alert(
    "인쇄 대상 체크박스 초기화",
    "라벨 생성이 완료되었습니다.\n마스터 시트의 인쇄 대상 체크박스를 모두 해제하시겠습니까?",
    ui.ButtonSet.YES_NO
  );
  
  if (uncheckConfirm === ui.Button.YES) {
    uncheckAllPrintLabels();
  } else {
    ss.toast("라벨 인쇄 대상 체크 상태가 유지되었습니다.", "시스템 알림");
  }
}

/**
 * '직위 이름' 형식의 문자열(예: '교사 홍길동')을 직위와 이름으로 쪼개주는 도우미 함수
 * @param {string} value - 원본 관리자/취급자 문자열
 * @return {Object} { title: 직위, name: 이름 }
 */
function splitNameAndTitle_(value) {
  value = (value || "").toString().trim();
  if (!value) {
    return { title: "", name: "" };
  }
  var parts = value.split(/\s+/);
  if (parts.length >= 2) {
    return { title: parts[0], name: parts.slice(1).join(" ") };
  }
  return { title: "교사", name: value }; // 직위가 명시되지 않은 경우 기본값으로 '교사' 사용
}

/**
 * 라벨 개수에 따라 필요한 총 행(row) 수를 계산하는 함수
 * @param {number} numLabels - 출력할 라벨 개수
 * @return {number} 필요한 총 행 수
 */
function getRequiredRowsForLabels_(numLabels) {
  if (numLabels <= 0) return 0;
  var curRow = 2;
  var maxRow = 2;
  for (var idx = 0; idx < numLabels; idx++) {
    if (idx > 0 && idx % 10 === 0) {
      curRow += 2;
    }
    var r = curRow;
    var lastRowOfThisLabel = r + 6;
    if (lastRowOfThisLabel > maxRow) {
      maxRow = lastRowOfThisLabel;
    }
    if (idx % 2 === 1) {
      curRow += 7;
    }
  }
  return maxRow;
}

/**
 * 스프레드시트가 열릴 때 자동으로 실행되어 상단 메뉴바에 커스텀 메뉴를 추가하는 트리거 함수
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("기자재 시스템")
    .addItem("라벨 인쇄 전체 선택", "checkAllPrintLabels")
    .addItem("라벨 인쇄 전체 해제", "uncheckAllPrintLabels")
    .addSeparator()
    .addItem("QR 링크 및 이미지 갱신", "refreshQrLinks")
    .addItem("기기관리 라벨지 출력 생성", "createLabelSheet")
    .addToUi();
}

/**
 * 마스터 시트의 모든 기기에 대해 라벨인쇄 체크박스를 선택(true)하는 함수
 */
function checkAllPrintLabels() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("마스터");
  if (!sheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다.");
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return; // 데이터가 없으면 즉시 종료
  }
  
  var headerMapping = getHeaderMapping_(sheet);
  var printCol = headerMapping["라벨인쇄"];
  if (!printCol) {
    throw new Error("마스터 시트에 '라벨인쇄' 헤더가 존재하지 않습니다.");
  }
  
  sheet.getRange(2, printCol, lastRow - 1, 1).setValue(true);
  ss.toast("모든 기기의 라벨인쇄 대상 체크가 완료되었습니다.", "시스템 알림");
}

/**
 * 마스터 시트의 모든 기기에 대해 라벨인쇄 체크박스를 해제(false)하는 함수
 */
function uncheckAllPrintLabels() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("마스터");
  if (!sheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다.");
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  
  var headerMapping = getHeaderMapping_(sheet);
  var printCol = headerMapping["라벨인쇄"];
  if (!printCol) {
    throw new Error("마스터 시트에 '라벨인쇄' 헤더가 존재하지 않습니다.");
  }
  
  sheet.getRange(2, printCol, lastRow - 1, 1).setValue(false);
  ss.toast("모든 기기의 라벨인쇄 대상 체크가 해제되었습니다.", "시스템 알림");
}

/**
 * 기존 마스터 시트 데이터를 안전하게 보존하면서 '라벨인쇄' 체크박스 열만 추가해주는 마이그레이션 함수
 */
function upgradeMasterSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("마스터");
  var ui = SpreadsheetApp.getUi();
  
  if (!sheet) {
    ui.alert("오류", "마스터 시트를 찾을 수 없습니다. setupSheets()를 먼저 실행하여 초기화해주세요.", ui.ButtonSet.OK);
    return;
  }
  
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  
  // 이미 헤더가 존재하는지 확인
  var printColIdx = -1;
  var notesColIdx = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = headers[i].toString().trim();
    if (h === "라벨인쇄") {
      printColIdx = i + 1;
    } else if (h === "비고") {
      notesColIdx = i + 1;
    }
  }
  
  if (printColIdx > 0) {
    ui.alert("알림", "이미 마스터 시트에 '라벨인쇄' 열이 구성되어 있습니다.", ui.ButtonSet.OK);
    return;
  }
  
  // 비고 열 바로 다음에 라벨인쇄 열 삽입
  var targetCol = notesColIdx > 0 ? notesColIdx + 1 : 11; // 비고가 없으면 기본 K열(11번째)로 설정
  
  // 1. 열 삽입 (기존의 우측 열 데이터들이 안전하게 한 칸씩 오른쪽으로 밀림)
  sheet.insertColumnBefore(targetCol);
  
  // 2. 헤더 작성
  sheet.getRange(1, targetCol).setValue("라벨인쇄")
    .setFontWeight("bold")
    .setBackground("#f3f4f6")
    .setHorizontalAlignment("center");
  
  // 3. 체크박스 규칙 주입
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(2, targetCol, lastRow >= 2 ? lastRow - 1 : 1, 1);
  var rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  range.setDataValidation(rule);
  range.setValue(false); // 기본값 설정
  
  ui.alert("성공", "기존 데이터를 보존한 채 '라벨인쇄' 체크박스 열이 성공적으로 추가되었습니다!\n\n새로고침 없이 상단 메뉴에서 바로 사용 가능합니다.", ui.ButtonSet.OK);
}

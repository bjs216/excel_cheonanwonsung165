// ==========================================
// 🔑 환경 설정 (비밀번호를 원하는 대로 변경하세요)
// ==========================================
var PASSWORD_USER = "1234";       // 일반 사용자 비밀번호
var PASSWORD_ADMIN = "admin5678"; // 관리자 비밀번호

/**
 * 1. 시트가 열릴 때 상단 메뉴를 생성하고 로그인 창을 띄웁니다.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔒 모드 전환 시스템')
    .addItem('👤 로그인 (사용자/관리자)', 'askPassword')
    .addItem('🚪 로그아웃 (시트 잠금)', 'logout')
    .addToUi();
    
  askPassword();
}

function askPassword() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('🔑 모드 선택', '비밀번호를 입력해 주세요:', ui.ButtonSet.OK_CANCEL);
  
  if (result.getSelectedButton() == ui.Button.OK) {
    var inputPw = result.getResponseText();
    var cache = CacheService.getUserCache();
    
    if (inputPw === PASSWORD_ADMIN) {
      cache.put('user_role', 'ADMIN', 21600);
      ui.alert('🔓 관리자 모드로 로그인되었습니다. 전체 수정이 가능합니다.');
    } else if (inputPw === PASSWORD_USER) {
      cache.put('user_role', 'USER', 21600);
      ui.alert('👤 실제 사용자 모드로 로그인되었습니다.\n(고정시간 변경불가 / 타인 데이터 수정불가)');
    } else {
      cache.put('user_role', 'LOCKED', 21600);
      ui.alert('❌ 비밀번호가 올바르지 않습니다. 읽기 전용/잠금 상태로 유지됩니다.');
    }
  }
}

function logout() {
  var cache = CacheService.getUserCache();
  cache.put('user_role', 'LOCKED', 21600);
  SpreadsheetApp.getUi().alert('🔒 로그아웃되었습니다. 시트가 잠금 상태로 전환됩니다.');
}

/**
 * 2. 핵심 권한 제어 엔진 (강제 롤백 및 시트 동기화 처리 적용)
 */
function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // 캐시에서 권한 가져오기
  var cache = CacheService.getUserCache();
  var role = cache.get('user_role');
  
  // 타이틀 및 헤더 영역(1~4행) 및 순번(A열)은 예외 처리
  if (range.getRow() < 5 || range.getColumn() == 1) return;
  
  // [보안 1] 잠금 상태일 때 전체 차단 및 즉시 롤백
  if (!role || role === 'LOCKED') {
    var oldVal = (e.oldValue === undefined || e.oldValue === null) ? "" : e.oldValue;
    range.setValue(oldVal);
    SpreadsheetApp.flush(); // 🔥 시트에 즉시 반영하도록 강제 명령
    Browser.msgBox("⚠️ 잠금 상태: 상단 메뉴의 [모드 전환 시스템]을 통해 비밀번호를 입력해 주세요.");
    return;
  }
  
  // [보안 2] 관리자 모드(ADMIN)는 무제한 패스
  if (role === 'ADMIN') return;
  
  // [보안 3] 실제 사용자 모드(USER) 제어
  if (role === 'USER') {
    var row = range.getRow();
    var col = range.getColumn();
    
    // 🔥 고정 예약 영역(5행~25행) 내에서 시간(B열, 즉 col 2)을 수정하려고 할 때
    if (col == 2 && row <= 25) {
      // 이전 값이 존재하면 그 값으로, 없거나 덮어썼다면 빈칸("")으로 강제 세팅
      var revertVal = (e.oldValue === undefined || e.oldValue === null) ? "" : e.oldValue;
      range.setValue(revertVal);
      
      SpreadsheetApp.flush(); // 🔥 중요: 경고창이 뜨기 전에 시트 화면을 원래대로 먼저 강제 복구시킵니다.
      
      Browser.msgBox("⚠️ 권한 제한: 고정된 미팅 시간은 일반 사용자가 변경할 수 없습니다.");
      return;
    }
    
    // --- 일반 데이터 입력 칸(C열~) 및 하단 긴급 미팅 영역 제어 ---
    var developerMetadata = range.getDeveloperMetadata();
    var originalAuthorToken = "";
    
    var userToken = cache.get('user_token');
    if (!userToken) {
      userToken = "USER_" + Math.random().toString(36).substring(2, 11);
      cache.put('user_token', userToken, 21600);
    }
    
    if (developerMetadata.length > 0) {
      originalAuthorToken = developerMetadata[0].getValue();
    }
    
    // 시나리오 A: 빈 칸에 새로운 정보를 처음 입력할 때
    if (e.oldValue === undefined || e.oldValue === null || e.oldValue === "") {
      if (e.value !== undefined && e.value !== "") {
        range.addDeveloperMetadata("cell_owner", userToken);
      }
    } 
    // 시나리오 B: 이미 내용이 있는 셀을 타인이 수정하거나 지우려고 할 때
    else {
      if (originalAuthorToken !== "" && originalAuthorToken !== userToken) {
        range.setValue(e.oldValue); // 타인이 건드린 경우 이전 데이터로 복구
        SpreadsheetApp.flush();    // 🔥 경고창 뜨기 전 화면 복구 강제 실행
        Browser.msgBox("⚠️ 변경 불가: 다른 사용자가 이미 입력한 칸은 수정하거나 추가할 수 없습니다.");
      } 
      // 내가 썼던 글을 직접 지우는 경우 소유권 초기화
      else if (e.value === undefined || e.value === "") {
        if (developerMetadata.length > 0) {
          developerMetadata[0].remove();
        }
      }
    }
  }
}
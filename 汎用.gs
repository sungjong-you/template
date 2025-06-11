function ConvertDateToyyMMddText_(date, timeDifferenceConsiderationAdditionTime = 0) {

  let y = date.getFullYear();
  let m = date.getMonth();
  let d = date.getDate();
  let h = date.getHours() + timeDifferenceConsiderationAdditionTime;
  let minutes = date.getMinutes();

  return Utilities.formatDate(new Date(y, m, d, h, minutes), "JST", "yyyy-MM-dd");

}//ConvertDateToyyMMddText E

function ConvertDateToyyMMddHHmmssText_(date, timeDifferenceConsiderationAdditionTime = 0) {
  let y = date.getFullYear();
  let m = date.getMonth();
  let d = date.getDate();
  let h = date.getHours();
  let minutes = date.getMinutes();
  let s = date.getSeconds();

  return Utilities.formatDate(new Date(y, m, d, h, minutes, s), "JST", "yyyy-MM-dd HH:mm:ss");
}

function ConvertDataAndTimeToYYYYMMDDHHMMSS_(date,time){
  let y = date.getFullYear();
  let month = date.getMonth();
  let d = date.getDate();
  let h = time.getHours();
  let minutes = time.getMinutes();
  let s = time.getSeconds();

  return Utilities.formatDate(new Date(y, month, d, h, minutes, s), "JST", "yyyy-MM-dd HH:mm:ss");
}

function RepairFomula_(sh,obj){
  // objはクラスのObj
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // 소스 범위 (C11부터 행 끝까지)
  const sourceStartRow = obj.row_obj.FORMULA_R;
  const sourceStartColumn = obj.col_obj.PROPERTY_ID_C; // C열
  const sourceEndColumn = sh.getLastColumn();
  const sourceRange = sh.getRange(sourceStartRow, sourceStartColumn, 1, sourceEndColumn - sourceStartColumn + 1);
  // 붙여넣기 시작 셀 (C14)
  const destinationStartRow = obj.row_obj.SMAE2_R+2;
  const destinationStartColumn = obj.col_obj.PROPERTY_ID_C; // C열

  // 붙여넣기 끝 셀 (Z300)
  const destinationEndRow = sh.getLastRow();
  const destinationEndColumn = sh.getLastColumn(); // Z열

  // 서식 복사 및 붙여넣기
  sourceRange.copyFormatToRange(
    sh,
    destinationStartColumn,
    destinationEndColumn,
    destinationStartRow,
    destinationEndRow
  );
  console.log("書式貼り付け完了")
}

function fetchGETReserves_(hs_id) {
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };

  const reserves_API_id = 'reserves/' + hs_id;
  const url_reserves = URL_API + reserves_API_id;
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };

  try {
    const response = UrlFetchApp.fetch(url_reserves, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_reserve = JSON.parse(responseText);
      return json_reserve
      
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}


function fetchGETHouses_() {
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };
  
  const reserves_API_id = 'houses'
  const url_houses = URL_API + reserves_API_id;
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };

  try {
    const response = UrlFetchApp.fetch(url_houses, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_houses = JSON.parse(responseText);
      return json_houses
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생 : ${e}`);
    return null;
  }
}


function fetchGETRooms_(house_id) {
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };
  
  const sh_db = new ShDataObj(SS)

  const house_API_id = 'rooms?houses=' + house_id;
  const url_house = URL_API + house_API_id;
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };

  try {
    const response = UrlFetchApp.fetch(url_house, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_rooms_obj = JSON.parse(responseText);
      return json_rooms_obj
    }else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}

function fetchGetPaymentMethod_(){
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };
  
  const url = PAYMENTS_METHOD_MASTER_URL
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_payment_method_master = JSON.parse(responseText);
      return json_payment_method_master
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}

function fetchGetRoomType_(){
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };
  
  const url = ROOM_TYPES_MASTER_URL
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_room_types_master = JSON.parse(responseText);
      return json_room_types_master
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}

function fetchGetBillingItems_(){
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };

  const url = URL_API + BILLING_ITEMS_MASTER
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_billing_items_master = JSON.parse(responseText);
      return json_billing_items_master
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}

function fetchGetRoomNumber_(house_id){
  const SYSTEM_API_SS_ID = '16wEKcd8yRunZML2esdFu2YjhxDn7uoovl2zeEy9M3oE'
  const SS_SYSTEM = SpreadsheetApp.openById(SYSTEM_API_SS_ID)
  const HS_API_TOKEN_ID = SS_SYSTEM.getSheetByName("HS_API").getRange(3,1).getValue()
  const HS_API_TOKEN_PW = SS_SYSTEM.getSheetByName("HS_API").getRange(3,2).getValue()
  const BASICAUTH = Utilities.base64Encode(`${HS_API_TOKEN_ID}:${HS_API_TOKEN_PW}`);
  const URL_API ='https://api.hotelsmart.jp/api/v1/';

  const headers = {
    'Authorization': `Basic ${BASICAUTH}`,
    'accept': 'application/json',
    'Content-Type' : 'application/json'
    };

  const url = 'https://api.hotelsmart.jp/api/v1/' + ROOM_NUM_MASTER_URL + house_id
  const options = {
  'method': 'get',
  'headers': headers,
  'muteHttpExceptions': true,
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) { // 200은 정상적인 상태
      const json_room_number_master = JSON.parse(responseText);
      console.log(json_room_number_master)
      return json_room_number_master
    } else {   // 200을 제외한 모드 에러. error code와 error body를 logger로 남김.
      Logger.log(`API 요청 실패: 상태 코드 ${responseCode}`);
      Logger.log(responseText);
      return null;
    }
  } catch (e) {
    Logger.log(`오류 발생: ${e}`);
    return null;
  }
}

function extractSpreadsheetId_(url) {
  const regex = /d\/(.*?)(?=\/edit)/;
  const match = url.match(regex);
  if (match && match[1]) {
    return match[1];
  } else {
    return null; // 또는 다른 방식으로 처리
  }
}
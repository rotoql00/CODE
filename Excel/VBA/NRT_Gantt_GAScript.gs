/**
 * 엑셀 VBA의 Cx_Gantt_SungJun_Choi 함수를 Google Apps Script로 변환한 메인 함수입니다.
 * 이 함수는 Config 시트의 설정을 바탕으로 Data 시트의 간트/스테이지 그리드를 계산하고 출력합니다.
 */
function cxGantt_SungJun_Choi() {
  // 1. 현재 활성화된 구글 스프레드시트 파일 전체 객체를 가져옵니다.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 2. 현재 사용자가 보고 있는(활성화된) 시트를 데이터 시트로 할당합니다.
  const wsData = ss.getActiveSheet();
  
  // 3. "Config"라는 이름의 탭(시트)을 설정 전용 시트로 가져옵니다.
  const wsCfg = ss.getSheetByName("Config");
  
  // 4. Config 시트가 존재하지 않을 경우 에러를 방지하기 위해 체크하고 실행을 중단합니다.
  if (!wsCfg) {
    SpreadsheetApp.getUi().alert("Config 시트를 찾을 수 없습니다.");
    return;
  }

  // 5. 스테이터스 바를 대신하여, 우측 하단에 스크립트가 시작되었음을 알리는 토스트 메시지를 띄웁니다.
  ss.toast("Config-driven Stage grid 계산 시작...", "알림", -1);

  // 6. Config 시트에서 각종 기본 설정값을 가져옵니다. (기본값 설정 포함)
  // VBA의 GetAppConfigValue 커스텀 함수 역할을 하는 헬퍼 함수를 호출합니다.
  const defaultMode = safeTrimUpper(getAppConfigValue(wsCfg, "DefaultMode", "E"));
  
  // 7. 스테이지 문자를 연결할 때 사용할 구분자(Delimiter)를 가져옵니다. (예: "/")
  const joinDelim = getAppConfigValue(wsCfg, "JoinDelimiter", "/");
  
  // 8. 날짜 범위가 한쪽만 열려있을 때(Open-Ended)의 처리 정책을 가져옵니다.
  const openEndedPolicy = safeTrimUpper(getAppConfigValue(wsCfg, "OpenEndedPolicy", "ONE_DAY"));

  // 9. 헤더(날짜)가 시작되는 셀의 주소를 가져옵니다. (기본값: O3)
  const headerStartCellAddr = getAppConfigValue(wsCfg, "HeaderStartCell", "O3");
  
  // 10. 입력 데이터(시작일/종료일 등)가 있는 셀 범위를 가져옵니다. (기본값: D5:M350)
  const inputRangeAddr = getAppConfigValue(wsCfg, "InputRange", "D5:M350");
  
  // 11. 모드(Mode) 값이 들어있는 셀 범위를 가져옵니다. (기본값: B5:B350)
  const modeRangeAddr = getAppConfigValue(wsCfg, "ModeRange", "B5:B350");
  
  // 12. 최종 결과물이 출력될 시작 셀의 주소를 가져옵니다. (기본값: O5)
  const outputStartCellAddr = getAppConfigValue(wsCfg, "OutputStartCell", "O5");

  // 13. 입력 데이터 범위 객체를 구글 시트에서 가져옵니다.
  const rngInput = wsData.getRange(inputRangeAddr);
  
  // 14. 모드 데이터 범위 객체를 구글 시트에서 가져옵니다.
  const rngMode = wsData.getRange(modeRangeAddr);

  // 15. 입력 범위의 총 행(Row) 개수를 산출합니다.
  const nRows = rngInput.getNumRows();

  // 16. 헤더 범위의 마지막 열(Column) 번호를 찾는 헬퍼 함수를 호출합니다.
  const lastHdrCol = findHeaderLastCol(wsData, headerStartCellAddr);
  
  // 17. 헤더 시작 셀의 행(Row)과 열(Column) 번호를 추출합니다.
  const hdrStartCell = wsData.getRange(headerStartCellAddr);
  const hdrRow = hdrStartCell.getRow();
  const hdrStartCol = hdrStartCell.getColumn();
  
  // 18. 헤더 열의 총 개수(시작 열부터 마지막 열까지)를 계산합니다.
  const nCols = lastHdrCol - hdrStartCol + 1;

  // 19. 헤더(날짜) 범위를 객체로 가져옵니다.
  const rngHdr = wsData.getRange(hdrRow, hdrStartCol, 1, nCols);
  
  // 20. 결과물을 한 번에 덮어씌울 출력 범위 객체를 계산하여 가져옵니다. (행 개수 x 열 개수 크기)
  const rngOut = wsData.getRange(wsData.getRange(outputStartCellAddr).getRow(), wsData.getRange(outputStartCellAddr).getColumn(), nRows, nCols);

  // 21. 스프레드시트의 셀 값을 메모리(2차원 배열)로 한 번에 로드합니다. (성능 최적화의 핵심)
  // GAS의 getValues()는 VBA의 Value2와 동일한 역할을 합니다.
  const vInput = rngInput.getValues();
  const vMode = rngMode.getValues();
  const vHdr = rngHdr.getValues()[0]; // 헤더는 1행이므로 첫 번째 배열만 추출합니다.

  // 22. 캘린더 날짜들을 저장할 1차원 배열을 초기화합니다.
  const colDates = [];
  
  // 23. 헤더 배열을 순회하며 각 열의 날짜를 타임스탬프(숫자)로 변환하여 배열에 담습니다.
  for (let c = 0; c < nCols; c++) {
    colDates.push(toDateSerialFast(vHdr[c]));
  }

  // 24. 헤더의 마지막 날짜 타임스탬프를 보관합니다. 정책(TO_END) 처리에 사용됩니다.
  let headerLastDate = 0;
  if (nCols >= 1) {
    headerLastDate = colDates[nCols - 1];
  }

  // 25. 최종 결과를 담을 빈 2차원 배열을 미리 생성하고 구조를 잡습니다. (행 x 열)
  const outArr = [];
  for (let r = 0; r < nRows; r++) {
    outArr.push(new Array(nCols).fill("")); // 빈 문자열로 초기화된 행 배열을 추가합니다.
  }

  // 26. 반복문 내에서 사용할 데이터 캐싱용 객체(딕셔너리 대체)를 생성합니다.
  // VBA의 Scripting.Dictionary를 완벽하게 대체하는 순수 JS 객체입니다.
  const cache = {};

  // 27. 스테이지 설정 데이터(StageConfig)를 메모리에 한 번만 로드하여 속도를 극대화합니다.
  const stageConfigData = getTableData(wsCfg, "StageConfig");
  // 28. 입력 범위의 시작 열(Column) 번호를 저장해둡니다. (상대 위치 계산용)
  const inputStartCol = rngInput.getColumn();

  // 29. 각 행(Row)을 순회하며 스테이지 그리드 계산을 시작합니다.
  for (let r = 0; r < nRows; r++) {
    
    // 30. 현재 행의 모드(Mode) 값을 가져와서 대문자로 변환하고 공백을 제거합니다.
    let modeKey = safeTrimUpper(vMode[r][0]);
    
    // 31. 모드 값이 비어있다면, Config에서 가져온 기본(Default) 모드 값으로 대체합니다.
    if (modeKey.length === 0) modeKey = defaultMode;

    // 32. 캐시를 활용하여 해당 모드의 스테이지 정보를 불러옵니다. (헬퍼 함수 호출)
    const stageInfo = loadStagesForModeCached(modeKey, cache, stageConfigData, inputStartCol);
    
    // 33. 반환된 정보에서 변수들을 추출합니다.
    const stageCount = stageInfo.stageCount;
    const stageName = stageInfo.stageName;
    const relStartIdx = stageInfo.relStartIdx;
    const relFinishIdx = stageInfo.relFinishIdx;

    // 34. 해당 모드에 매칭되는 스테이지가 0개라면, 이미 빈 문자열로 초기화된 상태이므로 다음 행으로 건너뜁니다.
    if (stageCount === 0) continue;

    // 35. 현재 행의 각 스테이지별 시작일과 종료일을 담을 배열을 생성합니다.
    const sArr = new Array(stageCount);
    const fArr = new Array(stageCount);

    // 36. 추출된 스테이지 개수만큼 반복하며 시작/종료일을 파싱합니다.
    for (let i = 0; i < stageCount; i++) {
      let s = 0; // 시작일 타임스탬프 초기화
      let f = 0; // 종료일 타임스탬프 초기화

      // 37. 상대 시작 인덱스가 유효한 범위 내에 있는지 확인합니다.
      if (relStartIdx[i] >= 0 && relStartIdx[i] < vInput[r].length) {
        // 38. 입력 배열에서 해당 인덱스의 값을 타임스탬프로 변환하여 s에 할당합니다.
        s = toDateSerialFast(vInput[r][relStartIdx[i]]);
      }

      // 39. 상대 종료 인덱스가 유효한 범위 내에 있는지 확인합니다.
      if (relFinishIdx[i] >= 0 && relFinishIdx[i] < vInput[r].length) {
        // 40. 입력 배열에서 해당 인덱스의 값을 타임스탬프로 변환하여 f에 할당합니다.
        f = toDateSerialFast(vInput[r][relFinishIdx[i]]);
      }

      // 41. 시작일/종료일 누락 시(Open-Ended) 처리 정책에 따라 날짜를 보정합니다. (객체 반환 방식으로 JS에 맞게 수정)
      const normalized = normalizeRangeWithPolicy(s, f, openEndedPolicy, headerLastDate);
      sArr[i] = normalized.s; // 보정된 시작일 저장
      fArr[i] = normalized.f; // 보정된 종료일 저장
    }

    // 42. 현재 행에 대해, 캘린더의 모든 열(날짜)을 순회하며 그리드를 채웁니다.
    for (let c = 0; c < nCols; c++) {
      // 43. 캘린더의 현재 열 날짜 타임스탬프를 가져옵니다.
      let calN = colDates[c];

      // 44. 캘린더 날짜가 유효하지 않으면 빈 칸을 유지합니다.
      if (calN === 0) {
        outArr[r][c] = "";
      } else {
        // 45. 유효하다면, 조건에 맞는 스테이지 문자를 생성하여 결과 배열에 넣습니다.
        outArr[r][c] = buildStageString(calN, stageCount, stageName, sArr, fArr, joinDelim);
      }
    }
  }

  // 46. 계산이 완료된 거대한 2차원 결과 배열을, 스프레드시트의 출력 범위에 한 번에 뿌려줍니다. (가장 중요한 성능 포인트)
  rngOut.setValues(outArr);

  // 47. 모든 작업이 끝났음을 알리는 완료 팝업(Alert)을 띄웁니다.
  SpreadsheetApp.getUi().alert("완료", "NRT Gantt GAScript Progress completion.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// =========================================================================
// 이하 보조(Helper) 함수들
// =========================================================================

/**
 * 특정 날짜가 주어졌을 때, 포함되는 모든 스테이지 이름을 구분자로 연결하여 반환합니다.
 */
function buildStageString(calN, stageCount, stageName, sArr, fArr, delimiter) {
  let res = ""; // 결과 문자열 초기화
  // 모든 스테이지를 순회합니다.
  for (let i = 0; i < stageCount; i++) {
    // 캘린더 날짜가 해당 스테이지의 시작/종료일 범위 안에 있는지 확인합니다.
    if (isInRange(calN, sArr[i], fArr[i])) {
      // 범위 안이라면 기존 문자열에 구분자를 포함하여 스테이지 이름을 이어붙입니다.
      res = appendStage(res, stageName[i], delimiter);
    }
  }
  return res; // 완성된 문자열을 반환합니다.
}

/**
 * 딕셔너리(캐시)를 확인하여 이미 계산된 모드 정보가 있으면 반환하고, 없으면 새로 계산하여 캐시에 저장합니다.
 */
function loadStagesForModeCached(modeKey, cache, stageConfigData, inputStartCol) {
  // 캐시 객체에 modeKey가 이미 존재하는지 확인합니다.
  if (cache[modeKey]) {
    return cache[modeKey]; // 존재하면 메모리에 저장된 값을 즉시 반환 (속도 향상)
  }
  // 존재하지 않으면 Config 데이터를 뒤져서 새로 파싱합니다.
  const pack = loadStagesFromConfig(modeKey, stageConfigData, inputStartCol);
  // 파싱된 결과를 캐시에 저장해둡니다. (다음 동일 모드 등장 시 재사용)
  cache[modeKey] = pack;
  return pack; // 결과를 반환합니다.
}

/**
 * Config 시트의 데이터 배열에서 특정 모드(modeKey)에 해당하는 스테이지 룰을 추출하고 정렬합니다.
 */
function loadStagesFromConfig(modeKey, stageConfigData, inputStartCol) {
  let stageCount = 0;     // 매칭된 스테이지 개수
  let stageName = [];     // 스테이지 이름 배열
  let relStartIdx = [];   // 상대적 시작 열 인덱스 배열
  let relFinishIdx = [];  // 상대적 종료 열 인덱스 배열
  let seqArr = [];        // 정렬을 위한 시퀀스 번호 배열

  // 헤더를 제외한 데이터 행을 순회합니다. (stageConfigData의 첫 행이 헤더라고 가정)
  for (let r = 1; r < stageConfigData.length; r++) {
    let rowMode = safeTrimUpper(stageConfigData[r][0]); // Mode 열 (A열 가정)
    
    // 현재 행의 모드가 우리가 찾는 모드와 일치하는지 확인합니다.
    if (rowMode === modeKey) {
      stageCount++; // 일치하면 카운트를 1 증가
      
      // 순서(Seq), 스테이지명, 시작열 문자, 종료열 문자를 추출합니다. (열 위치 하드코딩 또는 동적 맵핑 가정)
      seqArr.push(Number(stageConfigData[r][1])); // Seq 열 (B열)
      stageName.push(String(stageConfigData[r][2])); // StageName 열 (C열)
      
      // 열 문자(예: "D", "M")를 숫자로 변환합니다.
      let absS = colLetterToIndex(String(stageConfigData[r][3])); // StartCol 열 (D열)
      let absF = colLetterToIndex(String(stageConfigData[r][4])); // FinishCol 열 (E열)
      
      // 입력 데이터 범위의 시작 위치를 기준으로 상대적 배열 인덱스(0부터 시작)를 계산합니다.
      let relS = absS - inputStartCol; 
      let relF = absF - inputStartCol; 
      
      relStartIdx.push(relS); // 상대 시작 인덱스 저장
      relFinishIdx.push(relF); // 상대 종료 인덱스 저장
    }
  }

  // 추출된 스테이지들을 시퀀스(Seq) 번호를 기준으로 오름차순 정렬합니다. (Bubble Sort 로직 변환)
  for (let a = 0; a < stageCount - 1; a++) {
    for (let b = a + 1; b < stageCount; b++) {
      // 앞의 Seq가 뒤의 Seq보다 크면 자리를 바꿉니다 (스왑).
      if (seqArr[b] < seqArr[a]) {
        // 구조 분해 할당(Destructuring)을 사용하여 JS에서 세련되게 스왑 처리
        [seqArr[a], seqArr[b]] = [seqArr[b], seqArr[a]];
        [stageName[a], stageName[b]] = [stageName[b], stageName[a]];
        [relStartIdx[a], relStartIdx[b]] = [relStartIdx[b], relStartIdx[a]];
        [relFinishIdx[a], relFinishIdx[b]] = [relFinishIdx[b], relFinishIdx[a]];
      }
    }
  }

  // 파싱 및 정렬된 결과들을 객체로 묶어서 반환합니다.
  return { stageCount, stageName, relStartIdx, relFinishIdx };
}

/**
 * 문자열 조합 시 앞에 값이 있으면 구분자를 추가하고, 없으면 이름만 반환합니다.
 */
function appendStage(res, stageName, delimiter) {
  // 기존 결과물이 비어있다면 스테이지 이름만 그대로 반환합니다.
  if (res.length === 0) return stageName;
  // 비어있지 않다면 기존 결과물 + 구분자 + 새로운 스테이지 이름을 결합하여 반환합니다.
  return res + delimiter + stageName;
}

/**
 * 한쪽 날짜만 입력되었을 때의 처리 정책(Policy)을 적용하여 날짜 범위를 보정합니다.
 * VBA의 ByRef(참조에 의한 전달)를 JS 객체 반환 방식으로 대체하였습니다.
 */
function normalizeRangeWithPolicy(s, f, policy, headerLastDate) {
  // 시작일은 있고 종료일은 없는 경우
  if (s !== 0 && f === 0) {
    if (policy === "TO_END") {
      f = headerLastDate; // 정책이 TO_END면 끝까지 채웁니다.
    } else if (policy === "IGNORE") {
      s = 0; f = 0; // 정책이 IGNORE면 무시(0 처리)합니다.
    } else {
      f = s; // 그 외(ONE_DAY 등)의 경우 하루짜리 일정으로 만듭니다.
    }
  } 
  // 시작일은 없고 종료일만 있는 경우
  else if (s === 0 && f !== 0) {
    if (policy === "IGNORE") {
      s = 0; f = 0; // 무시합니다.
    } else {
      s = f; // 하루짜리 일정으로 만듭니다.
    }
  }
  
  // 둘 다 존재하는데, 종료일이 시작일보다 과거인 경우 오류 방지를 위해 둘을 뒤집습니다(스왑).
  if (s !== 0 && f !== 0) {
    if (f < s) {
      let tmp = s;
      s = f;
      f = tmp;
    }
  }
  // 보정된 s와 f를 객체 형태로 반환합니다.
  return { s: s, f: f };
}

/**
 * 캘린더 날짜(calN)가 시작일(s)과 종료일(f) 사이에 포함되는지 논리적으로 검사합니다.
 */
function isInRange(calN, s, f) {
  // 시작일/종료일이 모두 0이면 유효하지 않으므로 false를 반환합니다.
  if (s === 0 && f === 0) return false;
  // 캘린더 날짜가 시작일 이상이고 종료일 이하인지 평가하여 참/거짓을 반환합니다.
  return (calN >= s && calN <= f);
}

/**
 * 날짜 값을 대소 비교가 가능한 숫자형 타임스탬프(밀리초)로 안전하게 변환합니다.
 */
function toDateSerialFast(v) {
  // 값이 비어있거나 유효하지 않으면 0을 반환합니다.
  if (v === "" || v === null || v === undefined) return 0;
  // 구글 시트가 해당 값을 Date 객체로 인식한 경우 getTime()을 통해 타임스탬프를 추출합니다.
  if (v instanceof Date) return v.getTime();
  // 숫자형인 경우 그대로 숫자로 캐스팅하여 반환합니다.
  if (!isNaN(v)) return Number(v);
  // 만약 문자열 형태의 날짜라면 파싱을 시도합니다.
  let parsedDate = new Date(v);
  // 정상적인 날짜 문자열이었다면 타임스탬프를, 아니면 0을 반환합니다.
  if (!isNaN(parsedDate.getTime())) return parsedDate.getTime();
  return 0; // 이도저도 아니면 0 반환
}

/**
 * 문자열을 안전하게 자우 공백 제거 후 대문자로 변환합니다.
 */
function safeTrimUpper(v) {
  // 값이 비어있으면 빈 문자열을 반환합니다.
  if (v === null || v === undefined) return "";
  // 문자열로 강제 변환 후, 양쪽 공백(trim)을 지우고, 대문자(toUpperCase)로 만듭니다.
  return String(v).trim().toUpperCase();
}

/**
 * 엑셀 열 문자(예: "A", "Z", "AA")를 1부터 시작하는 숫자 인덱스로 변환합니다.
 */
function colLetterToIndex(colLetter) {
  // 영문자를 대문자로 정규화하고 공백을 제거합니다.
  colLetter = safeTrimUpper(colLetter);
  // 빈 값이면 0 반환
  if (colLetter.length === 0) return 0;
  
  let res = 0; // 결과값 초기화
  // 문자열의 길이만큼 각 자리의 문자를 순회합니다.
  for (let i = 0; i < colLetter.length; i++) {
    // 26진법 계산 원리를 적용하여 알파벳을 숫자로 환산 누적합니다.
    res = res * 26 + (colLetter.charCodeAt(i) - 64); // 'A'의 charCode는 65이므로 64를 뺍니다.
  }
  return res; // 최종 계산된 열 번호를 반환합니다.
}

/**
 * 지정된 셀부터 우측으로 이동하며 값이 존재하는 마지막 열 번호를 찾아냅니다.
 */
function findHeaderLastCol(ws, headerStartCellAddr) {
  // 헤더 시작 셀 객체를 가져옵니다.
  const startCell = ws.getRange(headerStartCellAddr);
  let r = startCell.getRow(); // 헤더가 위치한 행 번호
  let c = startCell.getColumn(); // 헤더가 시작되는 열 번호
  let lastCol = c; // 마지막 열 번호를 시작 열로 초기화
  
  // 우측으로 최대 시트 끝까지 검색할 수 있도록 1000열 정도 범위를 배열로 긁어옵니다. (셀을 하나씩 읽으면 너무 느림)
  const maxSearchCol = ws.getMaxColumns() - c + 1;
  const headerValues = ws.getRange(r, c, 1, maxSearchCol).getValues()[0];
  
  // 배열을 순회하며 빈 칸이 나올 때까지 탐색합니다.
  for (let i = 0; i < headerValues.length; i++) {
    // 날짜 혹은 데이터가 유효하지 않으면(빈 칸이면) 그만 탐색합니다.
    if (toDateSerialFast(headerValues[i]) === 0) break;
    // 유효하다면 마지막 열 번호를 갱신합니다. (시작 열 c + 현재 인덱스 i)
    lastCol = c + i;
  }
  
  return lastCol; // 탐색이 종료된 최종 마지막 열 번호를 반환합니다.
}

/**
 * (커스텀 헬퍼) 구글 시트에서 Config 시트의 키-값 데이터를 찾아서 반환합니다.
 * AppConfig 테이블이 G열(Key)과 H열(Value)에 위치함.
 */
function getAppConfigValue(wsCfg, keyName, defaultValue) {
  keyName = safeTrimUpper(keyName); // 찾고자 하는 키를 대문자로 정규화
  // G열과 H열의 데이터를 메모리로 가져옵니다. (넉넉하게 100행)
  const data = wsCfg.getRange("G1:H100").getValues(); 
  
  // 가져온 배열을 위에서부터 아래로 순회합니다.
  for (let r = 0; r < data.length; r++) {
    // G열(인덱스 0)의 키 값이 일치하는지 확인합니다.
    if (safeTrimUpper(data[r][0]) === keyName) {
      // 일치하면 H열(인덱스 1)의 값을 문자열로 반환합니다.
      return String(data[r][1]);
    }
  }
  // 반복문을 다 돌아도 찾지 못했다면 기본값을 반환합니다.
  return defaultValue;
}

/**
 * (커스텀 헬퍼) ListObjects 대체용. 특정 표 데이터를 배열로 전체 로드합니다.
 * 이 예제에서는 단순화를 위해 "StageConfig"가 특정 범위(예: F1:J100)에 있다고 가정하고 짰습니다.
 * 실제 사용 시 시트 구조에 맞게 범위를 수정(F1:J100 등)해야 합니다.
 */
function getTableData(wsCfg, tableName) {
  // 구글 시트에서 '이름이 지정된 범위(Named Range)'를 사용했다면 아래 주석을 풉니다.
  // const range = wsCfg.getParent().getRangeByName(tableName);
  // return range.getValues();
  
  // Named Range가 없다면, 임의로 AppConfig는 A~B, StageConfig는 D~H에 있다고 가정하고 가져옵니다.
  // 이 부분은 실제 엑셀 파일의 표(Table) 위치에 맞추어 주소를 바꿔주셔야 합니다.
  if (tableName === "AppConfig") return wsCfg.getRange("G1:H100").getValues();
  if (tableName === "StageConfig") return wsCfg.getRange("A1:E100").getValues();
  
  return []; // 예외 처리용 빈 배열 반환
}

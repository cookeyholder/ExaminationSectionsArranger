# Google Apps Script 效能最佳化評估與重構指南

> **評估日期**: 2025-12-05  
> **評估範圍**: refactor-domain-models 分支相對於 master 分支的重構策略  
> **結論**: 重構方向正確，但需大幅調整以符合 GAS 效能最佳實務

---

## 📖 Google Apps Script 效能最佳化原則

根據 [Google 官方 Best Practices](https://developers.google.com/apps-script/guides/support/best-practices)，以下是讓 GAS 運行更快的核心原則：

### 原則 1: 最小化對外部服務的呼叫（Minimize calls to other services）

> "Using JavaScript operations within your script is considerably faster than calling other services."

**關鍵要點**：
- 純 JavaScript 運算比呼叫 Google 服務（Sheets、Docs、Drive）快得多
- 每次呼叫 `getValues()`、`setValues()` 都需要與 Google 伺服器通訊
- **目標**：將 I/O 呼叫次數降到最低

### 原則 2: 使用批次操作（Use batch operations）

> "Alternating read and write commands is slow. To speed up a script, read all data into an array with one command, perform any operations on the data in the array, and write the data out with one command."

**正確做法**：
```javascript
// ✅ 正確：一次讀取，處理陣列，一次寫入
const data = sheet.getRange('A1:S1000').getValues();
// ... 在記憶體中處理 data 陣列 ...
sheet.getRange('A1:S1000').setValues(data);
```

**錯誤做法**：
```javascript
// ❌ 錯誤：迴圈中逐格讀寫
for (let i = 0; i < 1000; i++) {
  const value = sheet.getRange(i+1, 1).getValue();
  sheet.getRange(i+1, 2).setValue(value * 2);
}
```

### 原則 3: 避免交替讀寫

> "Alternating read and write commands is slow."

**關鍵要點**：
- 將所有讀取操作集中在前面
- 將所有寫入操作集中在後面
- 避免「讀-寫-讀-寫」的交錯模式

### 原則 4: 使用 Cache Service（可選）

> "Use the Cache service to store data between script executions."

**適用場景**：
- 需要跨多次執行保存計算結果
- 讀取昂貴資源（如外部 API）的結果

---

## 📊 試算表結構分析

### 「排入考程的補考名單」工作表（核心資料）

| 欄位索引 | 欄位名稱 | 類型         |
| -------- | -------- | ------------ |
| 0 (A)    | 科別     | 輸入         |
| 1 (B)    | 年級     | 輸入         |
| 2 (C)    | 班級代碼 | 輸入         |
| 3 (D)    | 班級     | 輸入         |
| 4 (E)    | 座號     | 輸入         |
| 5 (F)    | 學號     | 輸入         |
| 6 (G)    | 姓名     | 輸入         |
| 7 (H)    | 科目名稱 | 輸入         |
| 8 (I)    | 節次     | **排程產出** |
| 9 (J)    | 試場     | **排程產出** |
| 10 (K)   | 小袋序號 | **排程產出** |
| 11 (L)   | 小袋人數 | **排程產出** |
| 12 (M)   | 大袋序號 | **排程產出** |
| 13 (N)   | 大袋人數 | **排程產出** |
| 14 (O)   | 班級人數 | **排程產出** |
| 15 (P)   | 時間     | **排程產出** |
| 16 (Q)   | 電腦     | 輸入         |
| 17 (R)   | 人工     | 輸入         |
| 18 (S)   | 任課老師 | 輸入         |

### 資料規模

| 參數             | 值     |
| ---------------- | ------ |
| 節數上限         | 8      |
| 試場數量         | 20     |
| 每間試場人數上限 | 34     |
| 每節可容納學生   | 680    |
| **最大學生人次** | ~5,440 |

---

## 🔍 現況分析

### Master 分支 I/O 模式

```javascript
function runFullSchedulingPipeline() {
  buildFilteredCandidateList();           // 讀+寫
  scheduleCommonSubjectSessions();        // 讀+寫
  scheduleSpecializedSubjectSessions();   // 讀+寫 (使用 buildSessionStatistics)
  assignExamRooms();                      // 讀+寫 (使用 buildSessionStatistics)
  sortFilteredStudentsBySessionRoom();    // 使用 Range.sort()
  allocateBagIdentifiers();               // 讀+寫
  populateSessionTimes();                 // 讀+寫
  updateBagAndClassPopulations();         // 讀+寫
  createExamBulletinSheet();              // 讀+寫
  createProctorRecordSheet();             // 讀+寫
  // ...
}
```

**I/O 次數估計**：約 14-18 次讀寫

### Refactor 分支 I/O 模式（目前）

```javascript
function scheduleCommonSubjectSessions() {
  const exam = createExamFromSheet();  // 讀取
  // ... 處理 ...
  saveExamToSheet(exam);               // 寫入
}

function scheduleSpecializedSubjectSessions() {
  const exam = createExamFromSheet();  // 再次讀取
  // ... 處理 ...
  saveExamToSheet(exam);               // 再次寫入
}
// 每個函式都獨立讀寫！
```

**I/O 次數估計**：約 16-20 次讀寫（更糟！）

---

## 🚨 關鍵問題

### 問題 1: I/O 次數過多

**根本原因**：每個排程函式都獨立呼叫 `createExamFromSheet()` 和 `saveExamToSheet()`

**違反原則**：最小化對外部服務的呼叫

### 問題 2: 未採用「單次讀取-批次處理-單次寫入」模式

**根本原因**：函式設計為獨立單元，而非 Pipeline 組件

**違反原則**：使用批次操作

### 問題 3: 過度設計的抽象層

**根本原因**：`createStatisticsContainer()` 增加複雜度但未減少 I/O

**影響**：程式碼更難理解，但效能未改善

---

## ✅ 最佳化方案

### 方案核心：單次讀取-Pipeline處理-單次寫入

```
┌─────────────────────────────────────────────────────────────┐
│                    runFullSchedulingPipeline()               │
├─────────────────────────────────────────────────────────────┤
│  1. 讀取階段（一次性）                                        │
│     ├── data = getDataRange().getValues()                   │
│     ├── sessionRules = 讀取參數區                            │
│     └── sessionTimes = 讀取節次時間表                        │
├─────────────────────────────────────────────────────────────┤
│  2. 處理階段（純 JavaScript，無 I/O）                        │
│     ├── scheduleCommonSubjects(data, rules)                 │
│     ├── scheduleSpecializedSubjects(data, config)           │
│     ├── assignRooms(data, config)                           │
│     ├── sortStudents(data)                                  │
│     ├── allocateBagIds(data)                                │
│     ├── fillSessionTimes(data, times)                       │
│     └── calculatePopulations(data)                          │
├─────────────────────────────────────────────────────────────┤
│  3. 寫入階段（一次性）                                        │
│     └── setValues(data)                                     │
└─────────────────────────────────────────────────────────────┘
```

### 效能對比

| 指標             | Master | Refactor（目前） | Refactor（優化後） |
| ---------------- | ------ | ---------------- | ------------------ |
| I/O 次數         | 14-18  | 16-20            | **2-4**            |
| 執行時間（預估） | 基準   | +10-20%          | **-60-70%**        |

---

## 📝 重構後的程式碼架構

### 1. 資料存取層（examService.js）

```javascript
/**
 * 一次性讀取所有需要的資料
 * @returns {Object} 包含所有資料的物件
 */
function loadAllData() {
  return {
    // 主要資料
    students: FILTERED_RESULT_SHEET.getDataRange().getValues(),
    
    // 參數設定（一次讀取整個區塊）
    parameters: PARAMETERS_SHEET.getRange('A1:F22').getValues(),
    
    // 節次時間
    sessionTimes: SESSION_TIME_REFERENCE_SHEET.getDataRange().getValues()
  };
}

/**
 * 解析參數
 */
function parseParameters(paramData) {
  return {
    maxSessionCount: paramData[4][1],  // B5
    maxRoomCount: paramData[5][1],     // B6
    maxStudentsPerRoom: paramData[6][1], // B7
    maxSubjectsPerRoom: paramData[7][1], // B8
    sessionCapacity: paramData[8][1],   // B9
    sessionRules: parseSessionRules(paramData)
  };
}

/**
 * 一次性寫回所有資料
 */
function saveAllData(students) {
  const sheet = FILTERED_RESULT_SHEET;
  const lastRow = sheet.getLastRow();
  
  // 清空舊資料
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, students[0].length).clearContent();
  }
  
  // 寫入新資料
  if (students.length > 0) {
    sheet.getRange(2, 1, students.length, students[0].length)
         .setValues(students);
  }
}
```

### 2. 排程純函式（scheduling.js）

```javascript
/**
 * 安排共同科目節次（純函式，無 I/O）
 * @param {Array<Array>} students - 學生資料（不含標題列）
 * @param {Object} sessionRules - 科目到節次的對映
 * @param {Object} columns - 欄位索引
 */
function scheduleCommonSubjectsInternal(students, sessionRules, columns) {
  for (let i = 0; i < students.length; i++) {
    const subject = students[i][columns.subject];
    if (sessionRules[subject]) {
      students[i][columns.session] = sessionRules[subject];
    }
  }
}

/**
 * 安排專業科目節次（純函式，無 I/O）
 */
function scheduleSpecializedSubjectsInternal(students, config, columns) {
  // 建立統計（在記憶體中）
  const stats = buildInMemoryStatistics(students, columns);
  
  // 分配邏輯...
}

/**
 * 安排試場（純函式，無 I/O）
 */
function assignRoomsInternal(students, config, columns) {
  // 分配邏輯...
}
```

### 3. Pipeline 整合（menu.js）

```javascript
/**
 * 一鍵產出公告用補考名單、試場記錄表
 * 採用「單次讀取-Pipeline處理-單次寫入」模式
 */
function runFullSchedulingPipeline() {
  const startTime = new Date();
  
  // ===== 階段 1: 建立候選名單 =====
  buildFilteredCandidateList();
  
  // ===== 階段 2: 一次性讀取所有資料 =====
  const rawData = loadAllData();
  const students = rawData.students.slice(1);  // 去除標題列
  const headerRow = rawData.students[0];
  const params = parseParameters(rawData.parameters);
  const columns = buildColumnIndices(headerRow);
  
  // ===== 階段 3: Pipeline 處理（純 JavaScript，零 I/O）=====
  scheduleCommonSubjectsInternal(students, params.sessionRules, columns);
  scheduleSpecializedSubjectsInternal(students, params, columns);
  assignRoomsInternal(students, params, columns);
  sortStudentsInternal(students, columns);
  allocateBagIdsInternal(students, columns);
  fillSessionTimesInternal(students, rawData.sessionTimes, columns);
  calculatePopulationsInternal(students, columns);
  
  // ===== 階段 4: 一次性寫回 =====
  saveAllData(students);
  
  // ===== 階段 5: 產生報表 =====
  createExamBulletinSheet();
  createProctorRecordSheet();
  composeSmallBagDataset();
  composeBigBagDataset();
  
  // 顯示執行時間
  const elapsed = calculateElapsedSeconds(startTime);
  SpreadsheetApp.getUi().alert('已完成編排，共使用 ' + elapsed + ' 秒');
}
```

---

## 🎯 領域模型的角色調整

### 原本設計的問題

```javascript
// 過度設計：通用統計容器
function createStatisticsContainer(config) {
  // 60+ 行程式碼
  // 動態建立 getter
  // 子容器管理
  // 分配機制
}
```

**問題**：增加複雜度但未減少 I/O

### 簡化後的設計

```javascript
/**
 * 建立記憶體中的統計物件
 * 用於排程演算法的判斷（如科別年級互斥檢查）
 */
function buildInMemoryStatistics(students, columns) {
  const stats = {
    bySession: {},      // session -> { students, deptGrade: {} }
    byDeptGradeSubject: {}  // "科別年級_科目" -> count
  };
  
  students.forEach((student, index) => {
    const session = student[columns.session];
    const deptGrade = student[columns.department] + student[columns.grade];
    const deptGradeSubject = deptGrade + '_' + student[columns.subject];
    
    // 節次統計
    if (!stats.bySession[session]) {
      stats.bySession[session] = { students: [], deptGrade: {} };
    }
    stats.bySession[session].students.push(index);
    stats.bySession[session].deptGrade[deptGrade] = 
      (stats.bySession[session].deptGrade[deptGrade] || 0) + 1;
    
    // 科別年級科目統計
    stats.byDeptGradeSubject[deptGradeSubject] = 
      (stats.byDeptGradeSubject[deptGradeSubject] || 0) + 1;
  });
  
  return stats;
}
```

**優點**：
- 程式碼更簡潔
- 無額外抽象層
- 統計在記憶體中完成，不觸發 I/O

---

## 📋 向後相容策略

為了讓選單中的個別步驟仍可獨立執行，保留包裝函式：

```javascript
/**
 * 安排共同科目節次（公開 API，供選單使用）
 * 內部使用「讀取-處理-寫入」模式
 */
function scheduleCommonSubjectSessions() {
  // 讀取
  const rawData = loadAllData();
  const students = rawData.students.slice(1);
  const headerRow = rawData.students[0];
  const params = parseParameters(rawData.parameters);
  const columns = buildColumnIndices(headerRow);
  
  // 處理
  scheduleCommonSubjectsInternal(students, params.sessionRules, columns);
  
  // 寫入
  saveAllData(students);
}
```

---

## 🔢 預期效能提升

| 場景                                 | 優化前 I/O | 優化後 I/O | 節省比例       |
| ------------------------------------ | ---------- | ---------- | -------------- |
| runFullSchedulingPipeline            | 16-20 次   | 3-4 次     | **80%**        |
| resumePipelineAfterManualAdjustments | 10-12 次   | 2 次       | **80%**        |
| 個別步驟執行                         | 2 次       | 2 次       | 0%（維持相容） |

### 執行時間預估

假設每次 I/O 約 200-500ms：
- **優化前**：16 × 350ms = 5.6 秒（僅 I/O）
- **優化後**：3 × 350ms = 1.05 秒（僅 I/O）
- **節省**：約 4.5 秒

---

## ✅ 總結

### 核心改變

1. **I/O 模式**：從「每函式各自讀寫」改為「單次讀取-Pipeline處理-單次寫入」
2. **領域模型**：從「通用統計容器」簡化為「記憶體中的統計物件」
3. **函式設計**：從「獨立執行單元」改為「Pure Function + 包裝器」

### 符合 Google 最佳實務

- ✅ 最小化對 Spreadsheet 服務的呼叫
- ✅ 使用批次操作讀寫資料
- ✅ 避免交替讀寫
- ✅ 在 JavaScript 陣列中完成所有運算

### 預期效益

- 執行時間減少 60-70%
- 程式碼更簡潔（減少約 200 行）
- 維護更容易（邏輯更直觀）

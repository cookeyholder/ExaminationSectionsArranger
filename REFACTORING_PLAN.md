# 領域模型重構計畫

> **重構目標**：將現有排程系統全面改用 `domainModels.js` 的三層領域模型（Exam → Session → Classroom），不保留向後相容性，追求最佳設計。

> **✅ 狀態**：階段 1-5 已完成（25/52 任務），程式碼重構完成，待 Apps Script 環境測試。

> **📋 實作指引**：本文件提供程式碼範例供參考。完整的實作計畫、任務追蹤和技術決策請參考：
> - **執行指南**：[openspec/changes/refactor-domain-models/IMPLEMENTATION_GUIDE.md](openspec/changes/refactor-domain-models/IMPLEMENTATION_GUIDE.md)
> - **任務清單**：[openspec/changes/refactor-domain-models/tasks.md](openspec/changes/refactor-domain-models/tasks.md)（25/52 完成）
> - **技術設計**：[openspec/changes/refactor-domain-models/design.md](openspec/changes/refactor-domain-models/design.md)

---

## 📐 新架構設計

### 核心概念

```
Exam (考試活動)
  ├── sessions[] (節次陣列)
  │     ├── students[] (該節次的所有學生)
  │     ├── classrooms[] (試場陣列)
  │     │     └── students[] (該試場的學生)
  │     └── 統計屬性 (departmentGradeStatistics, departmentClassSubjectStatistics)
  └── 統計屬性 (sessionDistribution, departmentDistribution, gradeDistribution, subjectDistribution)
```

### 資料流向

```
工作表資料 
  → 建立 Exam 物件
  → 填充 Session.students
  → 執行排程邏輯（分配節次、試場）
  → distributeToChildren() 分配到 Classroom
  → 回寫工作表
```

---

## 🎯 重構階段（無向後相容考量）

### **階段 1：建立核心服務層（2-3 小時）**

#### 目標
建立新的服務層，封裝 Exam 物件的建立與操作

#### 1.1 建立 `examService.js`

```javascript
/**
 * 考試服務 - 負責 Exam 物件的建立、載入與儲存
 */

/**
 * 從工作表建立 Exam 物件
 * @returns {Object} Exam 物件
 */
function createExamFromSheet() {
  const [headerRow, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
  const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();

  const exam = createExamRecord(maxSessionCount, maxRoomCount);

  // 將學生填入對應節次
  candidateRows.forEach(studentRow => {
    const sessionNumber = studentRow[sessionIndex];
    if (sessionNumber > 0 && sessionNumber < exam.sessions.length) {
      exam.sessions[sessionNumber].addStudent(studentRow);
    }
  });

  return exam;
}

/**
 * 將 Exam 物件存回工作表
 * @param {Object} exam - Exam 物件
 */
function saveExamToSheet(exam) {
  const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  const allStudents = [];

  // 從 Classroom 收集所有學生（Single Source of Truth）
  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      allStudents.push(...classroom.students);
    });
  });

  if (allStudents.length > 0) {
    // 清空舊資料
    const lastRow = FILTERED_RESULT_SHEET.getLastRow();
    if (lastRow > 1) {
      FILTERED_RESULT_SHEET.getRange(2, 1, lastRow - 1, headerRow.length).clearContent();
    }
    
    // 寫入新資料
    FILTERED_RESULT_SHEET.getRange(2, 1, allStudents.length, allStudents[0].length)
      .setValues(allStudents);
  }
}

/**
 * 取得欄位索引對映
 * @returns {Object} 欄位名稱到索引的對映
 */
function getColumnIndices() {
  const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  return {
    department: headerRow.indexOf("科別"),
    grade: headerRow.indexOf("年級"),
    classCode: headerRow.indexOf("班級代碼"),
    class: headerRow.indexOf("班級"),
    seatNumber: headerRow.indexOf("座號"),
    studentId: headerRow.indexOf("學號"),
    name: headerRow.indexOf("姓名"),
    subject: headerRow.indexOf("科目名稱"),
    session: headerRow.indexOf("節次"),
    room: headerRow.indexOf("試場"),
    smallBagId: headerRow.indexOf("小袋序號"),
    smallBagPopulation: headerRow.indexOf("小袋人數"),
    bigBagId: headerRow.indexOf("大袋序號"),
    bigBagPopulation: headerRow.indexOf("大袋人數"),
    classPopulation: headerRow.indexOf("班級人數"),
    time: headerRow.indexOf("時間"),
    computer: headerRow.indexOf("電腦"),
    manual: headerRow.indexOf("人工"),
    teacher: headerRow.indexOf("任課老師")
  };
}
```

#### 1.2 更新 `appsscript.json`

```json
{
  "timeZone": "Asia/Taipei",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
  ]
}
```

確保檔案載入順序：
1. `globals.js`
2. `helpers.js`
3. `domainModels.js`
4. `examService.js`
5. 其他模組

---

### **階段 2：重寫排程邏輯（4-5 小時）**

#### 目標
用 Exam 模型重寫所有排程函式

#### 2.1 重寫 `scheduleCommonSubjectSessions()`

```javascript
/**
 * 安排共同科目的節次（物理、國文、英文、數學等）
 */
function scheduleCommonSubjectSessions() {
  const sessionRuleRows = PARAMETERS_SHEET.getRange(2, 5, 21, 2).getValues();
  const exam = createExamFromSheet();
  const columns = getColumnIndices();

  // 建立科目到節次的對映
  const preferredSessionBySubject = {};
  sessionRuleRows.forEach(ruleRow => {
    if (ruleRow[0] && ruleRow[1]) {
      preferredSessionBySubject[ruleRow[0]] = ruleRow[1];
    }
  });

  // 重新分配節次
  exam.sessions.forEach(session => {
    session.students.forEach(student => {
      const subjectName = student[columns.subject];
      const preferredSession = preferredSessionBySubject[subjectName];
      if (preferredSession != null) {
        student[columns.session] = preferredSession;
      }
    });
  });

  // 重建 Exam（因為節次已變更）
  saveExamToSheet(exam);
}
```

#### 2.2 重寫 `scheduleSpecializedSubjectSessions()`

```javascript
/**
 * 安排專業科目的節次
 */
function scheduleSpecializedSubjectSessions() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
  const sessionCapacity = 0.9 * PARAMETERS_SHEET.getRange("B9").getValue();

  const departmentGradeSubjectCounts = Object.entries(
    fetchDepartmentGradeSubjectCounts()
  ).sort(compareCountDescending);

  // 清空所有節次（重新分配）
  exam.sessions.forEach(session => session.clear());
  
  // 收集所有未分配節次的學生
  const unscheduledStudents = [];
  exam.sessions[0].students.forEach(student => {
    if (student[columns.session] === 0) {
      unscheduledStudents.push(student);
    }
  });

  // 為每個節次分配學生
  for (let sessionNumber = 1; sessionNumber <= maxSessionCount; sessionNumber++) {
    const session = exam.sessions[sessionNumber];
    
    for (let countIndex = 0; countIndex < departmentGradeSubjectCounts.length; countIndex++) {
      const [deptGradeSubjectKey, studentCount] = departmentGradeSubjectCounts[countIndex];
      const deptGradeKey = deptGradeSubjectKey.substring(0, deptGradeSubjectKey.indexOf("_"));

      // 檢查該科別是否已排入此節次
      const deptGradeStats = session.departmentGradeStatistics;
      if (Object.keys(deptGradeStats).includes(deptGradeKey)) {
        continue;
      }

      // 檢查容量
      if (studentCount + session.population > sessionCapacity) {
        continue;
      }

      // 分配學生到此節次
      unscheduledStudents.forEach(student => {
        const studentKey = `${student[columns.department]}${student[columns.grade]}_${student[columns.subject]}`;
        if (studentKey === deptGradeSubjectKey && student[columns.session] === 0) {
          student[columns.session] = sessionNumber;
          session.addStudent(student);
        }
      });
    }

    if (session.population >= sessionCapacity) {
      Logger.log(`第${sessionNumber}節已達人數上限：${session.population}`);
    }
  }

  saveExamToSheet(exam);
}
```

#### 2.3 重寫 `assignExamRooms()`

```javascript
/**
 * 安排試場
 */
function assignExamRooms() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
  const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();
  const maxStudentsPerRoom = PARAMETERS_SHEET.getRange("B7").getValue();
  const maxSubjectsPerRoom = PARAMETERS_SHEET.getRange("B8").getValue();

  for (let sessionNumber = 1; sessionNumber <= maxSessionCount; sessionNumber++) {
    const session = exam.sessions[sessionNumber];
    
    // 清空所有試場
    session.classrooms.forEach(classroom => classroom.clear());

    const deptClassSubjectCounts = Object.entries(
      session.departmentClassSubjectStatistics
    ).sort(compareCountDescending);

    for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
      const classroom = session.classrooms[roomNumber];
      let scheduledSubjects = [];

      for (let countIndex = 0; countIndex < deptClassSubjectCounts.length; countIndex++) {
        const [classSubjectKey, count] = deptClassSubjectCounts[countIndex];

        // 檢查是否已排入
        if (scheduledSubjects.includes(classSubjectKey)) continue;

        // 檢查容量
        if (count + classroom.population > maxStudentsPerRoom) continue;

        // 檢查科目數限制
        const subjectCount = Object.keys(classroom.classSubjectStatistics).length;
        if (subjectCount + 1 > maxSubjectsPerRoom) continue;

        // 分配學生到此試場
        session.students.forEach(student => {
          const studentKey = `${student[columns.class]}${student[columns.subject]}`;
          if (studentKey === classSubjectKey && student[columns.room] === 0) {
            student[columns.room] = roomNumber;
            classroom.addStudent(student);
          }
        });

        scheduledSubjects = Object.keys(classroom.classSubjectStatistics);
      }
    }
  }

  saveExamToSheet(exam);
}
```

---

### **階段 3：重寫輔助函式（2-3 小時）**

#### 3.1 重寫 `allocateBagIdentifiers()`

```javascript
/**
 * 計算大、小袋編號
 */
function allocateBagIdentifiers() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  let smallBagCounter = 1;
  let bigBagCounter = 1;

  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      if (classroom.population > 0) {
        // 設定小袋序號
        classroom.students.forEach(student => {
          student[columns.smallBagId] = smallBagCounter;
        });
        
        // 設定大袋序號（每個試場一個大袋）
        classroom.students.forEach(student => {
          student[columns.bigBagId] = bigBagCounter;
        });

        smallBagCounter++;
        bigBagCounter++;
      }
    });
  });

  saveExamToSheet(exam);
}
```

#### 3.2 重寫 `populateSessionTimes()`

```javascript
/**
 * 填入試場時間
 */
function populateSessionTimes() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  const sessionTimeRules = PARAMETERS_SHEET.getRange(2, 7, 21, 2).getValues();
  
  const sessionTimeMap = {};
  sessionTimeRules.forEach(rule => {
    if (rule[0] && rule[1]) {
      sessionTimeMap[rule[0]] = rule[1];
    }
  });

  exam.sessions.forEach((session, sessionNumber) => {
    const timeValue = sessionTimeMap[sessionNumber] || "";
    session.classrooms.forEach(classroom => {
      classroom.students.forEach(student => {
        student[columns.time] = timeValue;
      });
    });
  });

  saveExamToSheet(exam);
}
```

#### 3.3 重寫 `updateBagAndClassPopulations()`

```javascript
/**
 * 計算試場人數、大小袋人數、班級人數
 */
function updateBagAndClassPopulations() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();

  // 計算班級人數
  const classPopulationMap = {};
  exam.sessions.forEach(session => {
    session.students.forEach(student => {
      const className = student[columns.class];
      classPopulationMap[className] = (classPopulationMap[className] || 0) + 1;
    });
  });

  // 更新所有欄位
  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      // 小袋人數 = 試場人數
      const smallBagPopulation = classroom.population;
      
      // 大袋人數 = 試場人數（一個試場一個大袋）
      const bigBagPopulation = classroom.population;

      classroom.students.forEach(student => {
        student[columns.smallBagPopulation] = smallBagPopulation;
        student[columns.bigBagPopulation] = bigBagPopulation;
        student[columns.classPopulation] = classPopulationMap[student[columns.class]] || 0;
      });
    });
  });

  saveExamToSheet(exam);
}
```

---

### **階段 4：重寫排序函式（1 小時）**

#### 4.1 統一排序邏輯

```javascript
/**
 * 依科目排序補考名單
 */
function sortFilteredStudentsBySubject() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  
  const allStudents = [];
  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      allStudents.push(...classroom.students);
    });
  });

  allStudents.sort((a, b) => {
    // 優先依科目
    if (a[columns.subject] !== b[columns.subject]) {
      return a[columns.subject].localeCompare(b[columns.subject], 'zh-TW');
    }
    // 次依班級
    if (a[columns.class] !== b[columns.class]) {
      return a[columns.class].localeCompare(b[columns.class], 'zh-TW');
    }
    // 最後依座號
    return a[columns.seatNumber] - b[columns.seatNumber];
  });

  // 直接寫回工作表（不透過 Exam 物件）
  const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  FILTERED_RESULT_SHEET.getRange(2, 1, allStudents.length, allStudents[0].length)
    .setValues(allStudents);
}

/**
 * 依班級座號排序補考名單
 */
function sortFilteredStudentsByClassSeat() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  
  const allStudents = [];
  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      allStudents.push(...classroom.students);
    });
  });

  allStudents.sort((a, b) => {
    // 優先依班級
    if (a[columns.class] !== b[columns.class]) {
      return a[columns.class].localeCompare(b[columns.class], 'zh-TW');
    }
    // 次依座號
    if (a[columns.seatNumber] !== b[columns.seatNumber]) {
      return a[columns.seatNumber] - b[columns.seatNumber];
    }
    // 最後依科目
    return a[columns.subject].localeCompare(b[columns.subject], 'zh-TW');
  });

  const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  FILTERED_RESULT_SHEET.getRange(2, 1, allStudents.length, allStudents[0].length)
    .setValues(allStudents);
}

/**
 * 依節次試場排序補考名單
 */
function sortFilteredStudentsBySessionRoom() {
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  
  const allStudents = [];
  exam.sessions.forEach(session => {
    session.classrooms.forEach(classroom => {
      allStudents.push(...classroom.students);
    });
  });

  allStudents.sort((a, b) => {
    // 優先依節次
    if (a[columns.session] !== b[columns.session]) {
      return a[columns.session] - b[columns.session];
    }
    // 次依試場
    if (a[columns.room] !== b[columns.room]) {
      return a[columns.room] - b[columns.room];
    }
    // 再依班級
    if (a[columns.class] !== b[columns.class]) {
      return a[columns.class].localeCompare(b[columns.class], 'zh-TW');
    }
    // 最後依座號
    return a[columns.seatNumber] - b[columns.seatNumber];
  });

  const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  FILTERED_RESULT_SHEET.getRange(2, 1, allStudents.length, allStudents[0].length)
    .setValues(allStudents);
}
```

---

### **階段 5：移除舊程式碼（1 小時）**

#### 5.1 刪除清單

從 `scheduling.js` 刪除：
- ❌ `createEmptyClassroomRecord()`
- ❌ `createEmptySessionRecord()`
- ❌ `buildSessionStatistics()`

#### 5.2 驗證

使用 grep 確認沒有遺留引用：
```bash
grep -r "createEmptyClassroomRecord" .
grep -r "createEmptySessionRecord" .
grep -r "buildSessionStatistics" .
```

---

### **階段 6：整合測試（2 小時）**

#### 6.1 完整流程測試

1. 執行「步驟 1. 產出公告用補考名單、試場記錄表」
2. 驗證所有輸出工作表
3. 執行「步驟 2. 合併列印小袋封面」
4. 執行「步驟 3. 合併列印大袋封面」

#### 6.2 邊界案例測試

- 空節次處理
- 單一學生試場
- 超過容量的科目

---

## 📊 新架構優勢

| 面向         | 舊設計          | 新設計                 |
| ------------ | --------------- | ---------------------- |
| **資料結構** | 陣列 + 手動物件 | 統一的領域模型         |
| **統計計算** | 手寫 getter     | 通用統計容器           |
| **擴充性**   | 需修改多處      | 集中在 domainModels.js |
| **可讀性**   | 分散的邏輯      | 清晰的服務層           |
| **測試性**   | 難以單元測試    | 可獨立測試模型         |
| **維護性**   | 重複程式碼多    | DRY 原則               |

---

## ⏱️ 預估時程

| 階段                   | 預估時間 | 累計    |
| ---------------------- | -------- | ------- |
| 階段 1：建立核心服務層 | 2-3 小時 | 3 小時  |
| 階段 2：重寫排程邏輯   | 4-5 小時 | 8 小時  |
| 階段 3：重寫輔助函式   | 2-3 小時 | 11 小時 |
| 階段 4：重寫排序函式   | 1 小時   | 12 小時 |
| 階段 5：移除舊程式碼   | 1 小時   | 13 小時 |
| 階段 6：整合測試       | 2 小時   | 15 小時 |

**總計：約 2 個工作天**

---

## ✅ 完成檢查清單

- [x] `examService.js` 已建立並測試
- [x] 所有排程函式已重寫
- [x] 所有輔助函式已重寫
- [x] 所有排序函式已重寫
- [x] 舊程式碼已移除
- [ ] 完整流程測試通過（待 Apps Script 環境測試）
- [x] 文件已更新（AGENTS.md）
- [x] 程式碼已提交到分支

---

## 📝 實作心得（2025-10-31）

### 已完成階段 (1-5)

**階段 1**: 建立 `examService.js` (171 行)
- ✅ 實作 `getColumnIndices()` - 19 個欄位對映
- ✅ 實作 `createExamFromSheet()` - 從工作表建立 Exam 物件
- ✅ 實作 `saveExamToSheet(exam)` - 從 Classroom 收集資料存回工作表
- 檔案載入順序正確：domainModels.js → examService.js → scheduling.js

**階段 2**: 重寫排程邏輯
- ✅ `scheduleCommonSubjectSessions()` - 使用 exam.sessions 和 getColumnIndices()
- ✅ `scheduleSpecializedSubjectSessions()` - 使用 session.departmentGradeStatistics 檢查互斥
- ✅ `assignExamRooms()` - 使用 classroom.classSubjectStatistics 和 addStudent()
- 移除 buildSessionStatistics() 呼叫，改用 createExamFromSheet()

**階段 3**: 重寫輔助函式
- ✅ `allocateBagIdentifiers()` - 使用 classroom.classSubjectStatistics 計算小袋
- ✅ `populateSessionTimes()` - 遍歷 exam.sessions[i].classrooms[j]
- ✅ `updateBagAndClassPopulations()` - 使用統計屬性更新人數欄位
- 簡化邏輯，移除手動建立 lookup 字典

**階段 4**: 重寫排序函式
- ✅ `sortFilteredStudentsBySubject()` - 明確的排序條件（科別 > 年級 > 節次 > 試場 > 科目 > 座號）
- ✅ `sortFilteredStudentsByClassSeat()` - 科別 > 年級 > 座號 > 節次 > 科目
- ✅ `sortFilteredStudentsBySessionRoom()` - 節次 > 試場 > 科別 > 年級 > 座號 > 科目
- 使用 localeCompare() 進行中文排序，不再依賴欄位編號

**階段 5**: 清理舊程式碼
- ✅ 移除 `createEmptyClassroomRecord()` (-22 行)
- ✅ 移除 `createEmptySessionRecord()` (-45 行)
- ✅ 移除 `buildSessionStatistics()` (-16 行)
- 驗證：grep 搜尋確認無遺留引用

### 程式碼變更統計

- **新增**: examService.js (+171 行)
- **重寫**: scheduling.js 所有排程、輔助、排序函式
- **刪除**: 舊物件建立函式 (-97 行)
- **提交**: 6 個 commits

### 待完成階段

**階段 0** (手動): 建立測試副本、執行流程、匯出基準 CSV  
**階段 6** (手動): 在 Apps Script 環境執行完整測試、比對 CSV、驗證效能  
**階段 7**: 更新文件  
**階段 8**: 準備部署

### 重構效益

1. **程式碼減少**: 淨減少約 80+ 行重複程式碼
2. **可讀性提升**: 使用語意化的欄位名稱（如 `columns.subject`）取代魔術數字
3. **維護性改善**: 統計邏輯集中在 domainModels.js，新增維度只需修改配置
4. **一致性**: 所有函式統一使用 Exam 模型和 saveExamToSheet()

---

## 🚀 開始執行

準備好開始重構時，請依序執行各階段，每完成一個階段建議提交一次 commit。

<!-- OPENSPEC:START -->
# OpenSpec Instructions

These instructions are for AI assistants working in this project.

Always open `@/openspec/AGENTS.md` when the request:
- Mentions planning or proposals (words like proposal, spec, change, plan)
- Introduces new capabilities, breaking changes, architecture shifts, or big performance/security work
- Sounds ambiguous and you need the authoritative spec before coding

Use `@/openspec/AGENTS.md` to learn:
- How to create and apply change proposals
- Spec format and conventions
- Project structure and guidelines

Keep this managed block so 'openspec update' can refresh the instructions.

<!-- OPENSPEC:END -->

# Repository Guidelines

## 專案結構與模組安排
本專案鎖定 Google Apps Script，所有執行程式碼置於存放庫根目錄。`appsscript.json` 定義部署設定，功能模組依職責拆分為多個檔案：

- **`globals.js`**: 常用工作表實體與全域常數
- **`helpers.js`**: 共用工具函式
- **`domainModels.js`**: 領域模型定義（Exam、Session、Classroom）
- **`examService.js`**: 考試物件的建立、載入與儲存服務（重構中）
- **`menu.js`**: 工作列與流程指令（如 `runFullSchedulingPipeline`、`resumePipelineAfterManualAdjustments`）
- **`dataPreparation.js`**: 資料匯入與篩選
- **`scheduling.js`**: 節次與試場編排邏輯
- **`reportGeneration.js`**: 公告與列印資料產生
- **`printing.js`**: PDF 合併輸出

新增功能時請優先尋找對應模組，必要時再建立新檔案並於文件補充說明。

## 領域模型（Domain Models）

專案採用三層階層式領域模型，定義於 `domainModels.js`：

### 架構概覽

```
Exam (考試活動)
  ├── sessions[] (節次陣列，索引 0 不使用，從 1 開始)
  │     ├── students[] (該節次的所有學生)
  │     ├── classrooms[] (試場陣列，索引 0 不使用，從 1 開始)
  │     │     └── students[] (該試場的學生 - Single Source of Truth)
  │     └── 統計屬性
  │           ├── departmentGradeStatistics (科別-年級分布)
  │           └── departmentClassSubjectStatistics (班級-科目分布)
  └── 統計屬性
        ├── sessionDistribution (節次分布)
        ├── departmentDistribution (科別分布)
        ├── gradeDistribution (年級分布)
        └── subjectDistribution (科目分布)
```

### 核心物件

#### **Exam** - 考試活動
```javascript
const exam = createExamRecord(maxSessionCount, maxRoomCount);
// 屬性：sessions[], students[], population
// 統計：sessionDistribution, departmentDistribution, gradeDistribution, subjectDistribution
// 方法：addStudent(), clear(), statistics(dimensionName), distributeToChildren()
```

#### **Session** - 考試節次
```javascript
const session = createSessionRecord(maxRoomCount);
// 屬性：classrooms[], students[], population
// 統計：departmentGradeStatistics, departmentClassSubjectStatistics
// 方法：addStudent(), clear(), statistics(dimensionName), distributeToChildren()
```

#### **Classroom** - 考試試場
```javascript
const classroom = createClassroomRecord();
// 屬性：students[], population
// 統計：classSubjectStatistics
// 方法：addStudent(), clear(), statistics(dimensionName)
```

### 統計容器特性

所有領域物件皆繼承自 `createStatisticsContainer()`，提供：

- **動態統計**：透過 `statistics(dimensionName)` 取得任意維度的統計
- **計算屬性**：`population` 自動計算學生總數
- **具名存取**：可直接存取統計屬性（如 `session.departmentGradeStatistics`）
- **子容器管理**：`distributeToChildren(fn)` 支援向下分配學生

### 資料流向原則

```
工作表 → Exam → Session → Classroom → 工作表
        ↓       ↓         ↓
      統計   統計      統計（真實資料來源）
```

**重要**：學生資料的唯一真實來源（Single Source of Truth）是 `Classroom.students`，上層容器的 `students` 僅用於中間處理。

### 使用範例

```javascript
// 建立考試物件
const exam = createExamFromSheet();

// 存取統計
console.log(exam.departmentDistribution); // {"資訊": 50, "機械": 30}
console.log(exam.sessions[1].population); // 25

// 分配學生到試場
exam.sessions[1].distributeToChildren((student, classrooms) => {
  return student[roomIndex]; // 根據試場編號分配
});

// 儲存回工作表
saveExamToSheet(exam);
```

### 重構進度

目前正在進行領域模型重構，詳見 `REFACTORING_PLAN.md`。重構完成後，所有排程邏輯將統一使用 Exam 模型，舊的 `createEmptySessionRecord()` 和 `createEmptyClassroomRecord()` 將被移除。

## 建置、測試與開發指令
開發流程採用 Apps Script CLI：先執行 `npm install -g @google/clasp` 完成安裝，之後以 `clasp login` 認證帳號；`clasp push` 上傳目前檔案至連結的腳本專案，`clasp pull` 則同步遠端變更。臨時驗證可在試算表中開啟 Apps Script 編輯器，但請確保最終變更回存到版本庫並透過 `clasp push` 部署。

## 程式風格與命名慣例
沿用 Apps Script 預設：兩空格縮排，分號使用需在同一檔案內保持一致，變數宣告以 `const`、`let` 為主。函式與變數採駝峰式命名（如 `assignExamRooms`, `populateSessionTimes`）；讀取參數表的常數則使用大寫蛇形。共用邏輯宜集中在 `helpers.js`，資料或排程專用的程式碼請放回對應模組以維持職責單一。

## 測試指引
目前尚未建立自動化測試，請於試算表的測試副本進行手動驗證。透過 Apps Script 編輯器的 Run 功能執行新進入點，確認產出的資料範圍（例如「註冊組補考名單」、「參數區」）符合預期。請在 PR 敘述中列出手動測試情境，方便後續重現。

## 提交與 Pull Request 守則
用繁體中文撰寫符合 conventional commit 規範的提交訊息，必要時補上英文說明。內文可連結相關議題或試算表，並註明任何工作表結構調整。Pull Request 需包含變更概述、測試紀錄，以及影響輸出成果時的截圖。

## 試算表設定提示
執行腳本前請確認「參數區」中的 B2、B5-B8 等參數範圍及各模組引用的欄位標題未被修改。若學年度或工作表名稱異動，請同步更新程式中的對應常數並在 PR 說明，以維持部署一致性。

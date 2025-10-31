# 實作指南：領域模型重構

## 快速開始

```bash
# 1. 確認在正確分支
git checkout refactor/domain-models

# 2. 建立測試副本試算表（手動操作）
# - 開啟正式試算表
# - 檔案 → 建立副本
# - 重新命名為「補考編排系統_測試_20250127」

# 3. 執行基準測試
# - 在測試副本中執行「步驟 1」
# - 匯出 3 個工作表為 CSV 到 /tmp/baseline/

# 4. 開始階段 1
# 參考 tasks.md 的「階段 1」檢查清單
```

## 階段執行順序

### 階段 0: 準備基準測試資料 ⏱️ 30 分鐘
**目標**：建立可信的比對基準

1. 建立測試副本試算表
2. 執行現有流程並匯出 CSV
3. 記錄執行時間

**輸出**：
- `/tmp/baseline/baseline_scheduled.csv`
- `/tmp/baseline/baseline_bulletin.csv`
- `/tmp/baseline/baseline_classrooms.csv`
- `/tmp/baseline/benchmark.txt`

---

### 階段 1: 建立核心服務層 ⏱️ 2 小時
**目標**：封裝 Exam 物件與工作表之間的轉換

**關鍵檔案**：`examService.js`（新建）

**實作順序**：
1. `getColumnIndices()` - 最簡單，先建立對映
2. `createExamFromSheet()` - 從工作表讀取資料
3. `saveExamToSheet()` - 寫回工作表
4. 測試：執行 `saveExamToSheet(createExamFromSheet())` 應該不改變工作表

**驗收標準**：
```javascript
// 在 Apps Script 編輯器中執行
function testStage1() {
  const exam = createExamFromSheet();
  Logger.log('總人數: ' + exam.population);
  Logger.log('節次 1 人數: ' + exam.sessions[1].population);
  saveExamToSheet(exam);
  Logger.log('完成：工作表應該沒有變化');
}
```

**提交**：`git commit -m "feat(service): 建立 examService.js 封裝 Exam 物件轉換"`

---

### 階段 2: 重寫排程邏輯 ⏱️ 3 小時
**目標**：使用 Exam 模型重寫節次和試場分配

**修改檔案**：`scheduling.js`

**重寫函式**：
1. `scheduleCommonSubjectSessions(exam, sessionIndex)`
2. `scheduleSpecializedSubjectSessions(exam, sessionIndex)`
3. `assignExamRooms(exam, sessionIndex)`

**關鍵變更**：
```javascript
// 舊寫法
const sessions = buildSessionStatistics();
sessions[sessionIndex].students.push(student);

// 新寫法
exam.sessions[sessionIndex].addStudent(student);
```

**驗收標準**：
```javascript
function testStage2() {
  const exam = createExamFromSheet();
  // 假設有 3 個節次
  scheduleCommonSubjectSessions(exam, 1);
  assignExamRooms(exam, 1);
  
  Logger.log('節次 1 統計:');
  Logger.log(exam.sessions[1].departmentGradeStatistics);
  Logger.log('試場數: ' + exam.sessions[1].classrooms.length);
  
  // 比對人數與基準
}
```

**提交**：`git commit -m "refactor(scheduling): 重寫節次和試場分配函式使用 Exam 模型"`

---

### 階段 3: 重寫輔助函式 ⏱️ 2 小時
**目標**：使用 Exam 模型重寫編號和時間填充

**修改檔案**：`scheduling.js`

**重寫函式**：
1. `allocateBagIdentifiers(exam)`
2. `populateSessionTimes(exam)`
3. `updateBagAndClassPopulations(exam)`

**關鍵變更**：
```javascript
// 舊寫法
sessions.forEach(session => {
  session.classrooms.forEach(classroom => {
    // 手動計算統計
  });
});

// 新寫法
exam.sessions.forEach(session => {
  session.classrooms.forEach(classroom => {
    const stats = classroom.classSubjectStatistics;
    // 使用統計屬性
  });
});
```

**驗收標準**：
```javascript
function testStage3() {
  // 執行完整流程
  const exam = createExamFromSheet();
  // ... 執行所有排程函式
  allocateBagIdentifiers(exam);
  populateSessionTimes(exam);
  updateBagAndClassPopulations(exam);
  saveExamToSheet(exam);
  
  // 匯出並比對 CSV
}
```

**提交**：`git commit -m "refactor(scheduling): 重寫輔助函式使用 Exam 模型"`

---

### 階段 4: 重寫排序函式 ⏱️ 1.5 小時
**目標**：使用 Exam 模型重寫排序邏輯

**修改檔案**：`scheduling.js`

**重寫函式**：
1. `sortFilteredStudentsBySubject(exam)`
2. `sortFilteredStudentsByClassSeat(exam)`
3. `sortFilteredStudentsBySessionRoom(exam)`

**驗收標準**：
```javascript
function testStage4() {
  // 執行完整流程，包含排序
  // 比對「公告版補考場次」和「試場記錄表」
}
```

**提交**：`git commit -m "refactor(scheduling): 重寫排序函式使用 Exam 模型"`

---

### 階段 5: 清理舊程式碼 ⏱️ 30 分鐘
**目標**：移除舊的物件建立函式

**修改檔案**：`scheduling.js`

**刪除函式**：
1. `createEmptyClassroomRecord()` (line 410)
2. `createEmptySessionRecord()` (line 432)
3. `buildSessionStatistics()` (line 477)

**驗證命令**：
```bash
# 確認沒有遺留引用
grep -n "createEmptyClassroomRecord\|createEmptySessionRecord\|buildSessionStatistics" *.js

# 應該只出現在 REFACTORING_PLAN.md 和 proposal.md 中
```

**提交**：`git commit -m "refactor(scheduling): 移除舊的物件建立函式"`

---

### 階段 6: 整合測試與效能驗證 ⏱️ 2 小時
**目標**：確保重構後功能完全一致

**測試步驟**：
1. 執行「步驟 1」完整流程
2. 匯出 3 個工作表為 CSV
3. 使用 `diff` 比對

**比對指令**：
```bash
# 在 Terminal 中執行
cd /tmp
diff baseline/baseline_scheduled.csv test/test_scheduled.csv
diff baseline/baseline_bulletin.csv test/test_bulletin.csv
diff baseline/baseline_classrooms.csv test/test_classrooms.csv

# 應該無差異
```

**效能比對**：
```
基準時間：X 秒
測試時間：Y 秒
差異：(Y-X)/X * 100 = Z%
驗收標準：-10% < Z < +10%
```

**重複測試**：執行 3 次以確保穩定性

---

### 階段 7: 文件更新 ⏱️ 1 小時
**目標**：更新所有相關文件

**更新檔案**：
1. `REFACTORING_PLAN.md` - 標記所有階段為「✅ 已完成」
2. `AGENTS.md` - 已包含 Domain Models 章節（無需修改）
3. `openspec/specs/` - 移動規範（optional）

---

### 階段 8: 部署準備 ⏱️ 1 小時
**目標**：合併到 master 並部署

**步驟**：
1. 提交所有變更
2. 建立 Pull Request
3. 最終測試
4. 合併到 master
5. `clasp push` 部署

**Pull Request 範本**：
```markdown
## 變更摘要
重構排程系統，使用統一的領域模型（Exam → Session → Classroom）取代手動物件建立。

## 動機
- 減少程式碼重複（3 個物件建立函式 → 1 個通用工廠）
- 提升可維護性（新增統計維度只需修改配置）
- 統一資料流向（Sheet → Exam → Sheet）

## 變更內容
- ✅ 新增 `examService.js` 封裝 Exam 物件轉換
- ✅ 重寫 `scheduling.js` 所有排程函式
- ✅ 移除 `createEmptySessionRecord()` 等舊函式
- ✅ 更新 `AGENTS.md` 和 `REFACTORING_PLAN.md`

## 測試結果
- ✅ 輸出與基準完全一致（3 次測試）
- ✅ 執行時間差異：+5%（可接受範圍）
- ✅ 無新增錯誤或警告

## 截圖
[附上測試結果截圖]
```

---

## 常見問題

### Q1: 如果測試結果不一致怎麼辦？
**A1**: 
1. 先檢查是否是排序差異（使用 `sort` 排序後再 `diff`）
2. 比對統計數字（總人數、節次人數、試場數）
3. 使用 `Logger.log()` 追蹤資料流向
4. 必要時回退到上一個階段

### Q2: 如何處理 Apps Script 的檔案載入順序？
**A2**: 
Apps Script 按字母順序載入檔案，目前順序為：
```
appsscript.json
dataPreparation.js
domainModels.js      ← 必須在 examService 之前
examService.js       ← 必須在 scheduling 之前
globals.js
helpers.js
menu.js
printing.js
reportGeneration.js
scheduling.js
```
如果遇到 `ReferenceError`，可能需要重新命名檔案（例如 `01_globals.js`）。

### Q3: 如何快速驗證統計結果？
**A3**: 
在 Apps Script 編輯器中執行：
```javascript
function quickCheck() {
  const exam = createExamFromSheet();
  Logger.log('Exam 總人數: ' + exam.population);
  Logger.log('科別分布: ' + JSON.stringify(exam.departmentDistribution));
  
  exam.sessions.forEach((session, i) => {
    if (i === 0) return; // 跳過索引 0
    Logger.log(`節次 ${i} 人數: ${session.population}`);
    Logger.log(`節次 ${i} 試場數: ${session.classrooms.length - 1}`);
  });
}
```

### Q4: 如果執行時間超過 6 分鐘怎麼辦？
**A4**: 
1. 檢查是否有無限迴圈
2. 使用 `Logger.log()` 記錄每個階段的時間
3. 必要時分批處理（例如一次處理一個節次）
4. 目前系統在正常情況下執行時間約 1-2 分鐘，應該不會超過限制

---

## 檢查清單速查

### 階段 1 完成標準
- [ ] `examService.js` 檔案存在
- [ ] `getColumnIndices()` 回傳正確對映
- [ ] `createExamFromSheet()` 建立的 `exam.population` 正確
- [ ] `saveExamToSheet(createExamFromSheet())` 不改變工作表
- [ ] `clasp push` 無錯誤

### 階段 2 完成標準
- [ ] 3 個排程函式已重寫
- [ ] `exam.sessions[i].population` 正確
- [ ] `exam.sessions[i].departmentGradeStatistics` 符合互斥規則
- [ ] 試場容量未超過 `MAX_ROOM_CAPACITY`

### 階段 3 完成標準
- [ ] 大袋、小袋編號連續且無重複
- [ ] 節次時間已填充
- [ ] 班級人數正確
- [ ] 輸出與 `baseline_scheduled.csv` 一致

### 階段 4 完成標準
- [ ] 所有排序函式已重寫
- [ ] 「公告版補考場次」與基準一致
- [ ] 「試場記錄表」與基準一致

### 階段 5 完成標準
- [ ] 舊函式已刪除
- [ ] `grep` 搜尋無遺留引用
- [ ] `clasp push` 無警告

### 階段 6 完成標準
- [ ] 3 次測試結果皆一致
- [ ] 執行時間差異在 ±10% 內
- [ ] 所有 PDF 輸出正確

### 階段 7-8 完成標準
- [ ] 文件已更新
- [ ] Pull Request 已建立
- [ ] 合併到 master
- [ ] 正式環境部署成功

---

## 預估時間

| 階段            | 預估時間      | 累計時間  |
| --------------- | ------------- | --------- |
| 0. 準備基準     | 0.5 小時      | 0.5 小時  |
| 1. 建立服務層   | 2 小時        | 2.5 小時  |
| 2. 重寫排程邏輯 | 3 小時        | 5.5 小時  |
| 3. 重寫輔助函式 | 2 小時        | 7.5 小時  |
| 4. 重寫排序函式 | 1.5 小時      | 9 小時    |
| 5. 清理舊程式碼 | 0.5 小時      | 9.5 小時  |
| 6. 整合測試     | 2 小時        | 11.5 小時 |
| 7. 文件更新     | 1 小時        | 12.5 小時 |
| 8. 部署準備     | 1 小時        | 13.5 小時 |
| **總計**        | **13.5 小時** | -         |

建議分 3-4 個工作天完成，每天專注 3-4 小時。

---

## 下一步

準備好開始了嗎？執行：

```bash
# 檢視詳細任務清單
cat openspec/changes/refactor-domain-models/tasks.md

# 開始階段 0
# 1. 建立測試副本試算表
# 2. 執行現有流程
# 3. 匯出 CSV 到 /tmp/baseline/
```

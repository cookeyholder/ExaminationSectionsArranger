## 1. 建立核心服務層
- [ ] 1.1 建立 `examService.js` 檔案
- [ ] 1.2 實作 `createExamFromSheet()` 函式
- [ ] 1.3 實作 `saveExamToSheet()` 函式
- [ ] 1.4 實作 `getColumnIndices()` 函式
- [ ] 1.5 更新 `appsscript.json` 確保檔案載入順序正確

## 2. 重寫排程邏輯
- [ ] 2.1 重寫 `scheduleCommonSubjectSessions()`
- [ ] 2.2 重寫 `scheduleSpecializedSubjectSessions()`
- [ ] 2.3 重寫 `assignExamRooms()`
- [ ] 2.4 測試節次分配功能
- [ ] 2.5 測試試場分配功能

## 3. 重寫輔助函式
- [ ] 3.1 重寫 `allocateBagIdentifiers()`
- [ ] 3.2 重寫 `populateSessionTimes()`
- [ ] 3.3 重寫 `updateBagAndClassPopulations()`
- [ ] 3.4 測試編號計算正確性
- [ ] 3.5 測試時間填充正確性

## 4. 重寫排序函式
- [ ] 4.1 重寫 `sortFilteredStudentsBySubject()`
- [ ] 4.2 重寫 `sortFilteredStudentsByClassSeat()`
- [ ] 4.3 重寫 `sortFilteredStudentsBySessionRoom()`
- [ ] 4.4 測試所有排序功能

## 5. 清理舊程式碼
- [ ] 5.1 從 `scheduling.js` 移除 `createEmptyClassroomRecord()`
- [ ] 5.2 從 `scheduling.js` 移除 `createEmptySessionRecord()`
- [ ] 5.3 從 `scheduling.js` 移除 `buildSessionStatistics()`
- [ ] 5.4 使用 grep 確認沒有遺留引用
- [ ] 5.5 清理未使用的匯入和註解

## 6. 整合測試
- [ ] 6.1 執行「步驟 1. 產出公告用補考名單、試場記錄表」
- [ ] 6.2 驗證「排入考程的補考名單」工作表
- [ ] 6.3 驗證「公告版補考場次」工作表
- [ ] 6.4 驗證「試場記錄表」工作表
- [ ] 6.5 執行「步驟 2. 合併列印小袋封面」
- [ ] 6.6 執行「步驟 3. 合併列印大袋封面」
- [ ] 6.7 比對重構前後的輸出結果
- [ ] 6.8 記錄執行時間差異

## 7. 文件更新
- [ ] 7.1 更新 `AGENTS.md` 領域模型章節（已完成）
- [ ] 7.2 更新 `REFACTORING_PLAN.md`（已完成）
- [ ] 7.3 建立 OpenSpec 規範文件
- [ ] 7.4 撰寫遷移指南

## 8. 部署準備
- [ ] 8.1 在 `refactor/domain-models` 分支提交所有變更
- [ ] 8.2 建立 Pull Request
- [ ] 8.3 執行完整的手動測試
- [ ] 8.4 合併到 `master` 分支
- [ ] 8.5 執行 `clasp push` 部署到正式環境

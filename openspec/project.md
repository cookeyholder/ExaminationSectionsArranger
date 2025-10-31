# Project Context

## Purpose
高雄高工補考節次與試場編排自動化系統。從註冊組匯出的補考名單自動產生：
1. 節次安排（共同科目與專業科目分配）
2. 試場分配（考量容量、科目數限制）
3. 公告用補考名單
4. 監考記錄表
5. 試卷袋封面（小袋、大袋）的套印資料與 PDF 合併

## Tech Stack
- **Google Apps Script (JavaScript ES5+)**: 主要執行環境
- **Google Sheets**: 資料儲存與使用者介面
- **Google Drive API**: PDF 檔案合併與輸出
- **Clasp CLI**: 本地開發與部署工具

## Project Conventions

### Code Style
- **縮排**: 兩空格
- **分號**: 需在同一檔案內保持一致
- **變數宣告**: 優先使用 `const`、`let`，避免 `var`
- **命名慣例**:
  - 函式與變數：駝峰式（`assignExamRooms`, `sessionNumber`）
  - 全域常數（工作表參照）：大寫蛇形（`FILTERED_RESULT_SHEET`）
  - 參數常數：大寫蛇形（`MAX_SESSION_COUNT`）
- **語言**: 程式碼註解與提交訊息使用繁體中文，程式碼本身使用英文

### Architecture Patterns
- **領域驅動設計**: 使用三層領域模型（Exam → Session → Classroom）
- **服務層模式**: 業務邏輯封裝於服務層（`examService.js`）
- **Repository 模式**: 工作表存取透過 Repository 抽象化
- **單一真實來源**: 學生資料的唯一來源是 `Classroom.students`
- **模組職責分離**:
  - `globals.js`: 全域常數與工作表參照
  - `helpers.js`: 共用工具函式
  - `domainModels.js`: 領域模型定義
  - `examService.js`: Exam 物件的建立、載入與儲存
  - `dataPreparation.js`: 資料匯入與篩選
  - `scheduling.js`: 排程演算法
  - `reportGeneration.js`: 報表產生
  - `printing.js`: PDF 輸出
  - `menu.js`: 使用者介面與流程編排

### Testing Strategy
- **手動測試**: 在試算表測試副本中驗證
- **測試資料**: 保留完整學年度的測試資料集
- **驗證點**:
  1. 節次分配結果（科別不衝突、容量限制）
  2. 試場分配結果（科目數限制、人數限制）
  3. 統計數字正確性（各層級人數加總）
  4. 輸出格式（公告表、記錄表、封面資料）
- **迴歸測試**: 每次修改後執行完整流程，比對關鍵輸出

### Git Workflow
- **分支策略**:
  - `master`: 穩定版本
  - `refactor/*`: 重構工作分支
  - `feature/*`: 新功能開發分支
- **提交訊息**: 遵循 Conventional Commits
  - 格式: `<type>(<scope>): <description>`
  - Type: `feat`, `fix`, `refactor`, `docs`, `test`, `chore`
  - 使用繁體中文描述，必要時補充英文
  - 範例: `feat(scheduling): 新增科目衝突檢查邏輯`
- **Pull Request**: 需包含變更概述、測試紀錄、影響範圍說明

## Domain Context

### 補考流程
1. **資料準備**: 註冊組匯出補考名單 → 課程代碼補完 → 篩選列入考程的科目
2. **節次安排**: 共同科目優先 → 專業科目依科別年級分配 → 避免同科別同節次
3. **試場分配**: 每試場容量限制（參數 B7）→ 科目數限制（參數 B8）→ 優先合併同班級科目
4. **編號計算**: 小袋序號（每班每科一袋）→ 大袋序號（每試場一袋）
5. **資料產出**: 公告表 → 記錄表 → 小袋資料 → 大袋資料 → PDF 合併

### 關鍵概念
- **節次（Session）**: 考試時段，共 9 節（參數 B5）
- **試場（Classroom）**: 考試地點，每節次最多 5 個試場（參數 B6）
- **科別年級**: 同科別同年級不能在同節次（避免作弊）
- **小袋**: 同班級同科目的試卷袋
- **大袋**: 同試場的所有小袋打包

### 資料欄位（共 19 欄）
科別、年級、班級代碼、班級、座號、學號、姓名、科目名稱、節次、試場、小袋序號、小袋人數、大袋序號、大袋人數、班級人數、時間、電腦、人工、任課老師

## Important Constraints

### 技術限制
- **執行時間**: Google Apps Script 單次執行最長 6 分鐘
- **記憶體**: 有限，需避免大型陣列操作
- **API 配額**: Drive API 有每日呼叫次數限制
- **同步執行**: 無法使用 async/await（ES5 環境）

### 業務規則
- **科別年級互斥**: 同科別同年級學生不能在同一節次
- **試場容量**: 每試場最多學生數（參數 B7，預設 35 人）
- **科目數限制**: 每試場最多科目數（參數 B8，預設 3 科）
- **節次容量**: 每節次容量為參數 B9 的 90%
- **索引起始**: Session 和 Classroom 索引從 1 開始（索引 0 不使用）

### 資料完整性
- **課程代碼**: 必須完整才能查詢任課教師
- **姓名遮罩**: 公告表中姓名需部分遮蔽（首尾字保留）
- **參數依賴**: 學年度、學期、補考日期等參數必須正確設定

## External Dependencies

### Google Workspace APIs
- **Spreadsheet Service**: 工作表讀寫、格式設定、篩選器管理
- **Drive Service**: PDF 檔案建立、合併、下載
- **UI Service**: 選單建立、對話框顯示

### 試算表結構
- **工作表**:
  - 「參數區」: 系統參數設定（B2, B5-B9, D2-E22, G2-H22）
  - 「註冊組補考名單」: 原始資料
  - 「開課資料」: 任課教師查詢
  - 「排入考程的補考名單」: 處理結果（主要操作對象）
  - 「公告版補考場次」: 學生查詢用
  - 「試場記錄表」: 監考記錄用
  - 「小袋封面套印用資料」、「大袋封面套印用資料」: PDF 產生用

### Clasp 專案設定
- **Script ID**: 需在 `.clasp.json` 中設定
- **OAuth Scopes**: spreadsheets, drive（定義於 `appsscript.json`）
- **Runtime**: V8 引擎

# OpenSpec 工作流程說明

## 什麼是 OpenSpec？

OpenSpec 是一個規範驅動開發（Spec-Driven Development）的工作流程框架。它幫助我們：

1. **明確需求**：在實作前先定義「要做什麼」
2. **追蹤變更**：將變更提案和規範分開管理
3. **確保品質**：透過情境測試（Scenario）驗證需求
4. **保持同步**：規範與程式碼對應，易於理解系統行為

## 目錄結構

```
openspec/
├── project.md              # 專案慣例與技術棧（已填寫完成）
├── specs/                  # 當前規範 - 「系統現在的樣子」
│   ├── data-models/        # 資料模型規範
│   │   └── spec.md
│   └── scheduling/         # 排程功能規範
│       └── spec.md
└── changes/                # 變更提案 - 「想要改成什麼樣子」
    ├── refactor-domain-models/    # 目前進行中的變更
    │   ├── proposal.md            # 為什麼要改、改什麼、影響範圍
    │   ├── tasks.md               # 實作檢查清單
    │   └── specs/                 # 規範變更（deltas）
    │       ├── data-models/spec.md    # ADDED 新需求
    │       └── scheduling/spec.md     # MODIFIED 修改需求
    └── archive/            # 已完成的變更（歸檔後移至此處）
```

## 三階段工作流程

### 階段 1：建立變更提案

**何時需要建立提案？**
- ✅ 新增功能或能力
- ✅ 破壞性變更（API、資料結構）
- ✅ 架構或模式變更
- ✅ 效能優化（改變行為）

**何時不需要？**
- ❌ Bug 修復（恢復原本的預期行為）
- ❌ 拼字錯誤、格式調整、註解
- ❌ 相依套件更新（非破壞性）
- ❌ 配置變更

**建立步驟：**

1. **檢查現有狀態**
   ```bash
   openspec list          # 查看進行中的變更
   openspec list --specs  # 查看現有規範
   ```

2. **選擇變更 ID**
   - 使用 kebab-case 格式
   - 動詞開頭：`add-`, `update-`, `remove-`, `refactor-`
   - 簡短描述：`refactor-domain-models`

3. **建立目錄結構**
   ```bash
   mkdir -p openspec/changes/[change-id]/specs
   ```

4. **撰寫提案文件**

   **proposal.md** - 說明為什麼、改什麼、影響範圍
   ```markdown
   ## Why
   [1-2 句話說明問題或機會]

   ## What Changes
   - [變更清單]
   - **BREAKING**: [標註破壞性變更]

   ## Impact
   - Affected specs: [受影響的能力]
   - Affected code: [關鍵檔案/系統]
   ```

   **tasks.md** - 實作檢查清單
   ```markdown
   ## 1. 實作
   - [ ] 1.1 建立資料庫結構
   - [ ] 1.2 實作 API 端點
   ```

5. **撰寫規範變更（deltas）**

   在 `specs/[capability]/spec.md` 中使用：
   
   ```markdown
   ## ADDED Requirements
   ### Requirement: 新功能名稱
   The system SHALL provide...

   #### Scenario: 成功案例
   - **WHEN** 使用者執行動作
   - **THEN** 預期結果

   ## MODIFIED Requirements
   ### Requirement: 現有功能名稱
   [完整的修改後需求內容]

   #### Scenario: 修改後的情境
   ...

   ## REMOVED Requirements
   ### Requirement: 舊功能名稱
   **Reason**: [為何移除]
   **Migration**: [如何處理]
   ```

   **重要規則：**
   - ✅ 每個需求標題必須包含 `SHALL` 或 `MUST`
   - ✅ 每個需求至少要有一個 `#### Scenario:`
   - ✅ 使用 4 個 `#` 標記 Scenario（不是 3 個）
   - ✅ MODIFIED 需求要包含完整內容（會完全替換舊的）

6. **驗證提案**
   ```bash
   openspec validate [change-id] --strict
   ```

### 階段 2：實作變更

**執行步驟：**

1. **閱讀提案** - 理解要建構什麼
2. **閱讀設計文件**（如果有）- 檢視技術決策
3. **閱讀 tasks.md** - 取得實作檢查清單
4. **依序實作** - 依照順序完成任務
5. **確認完成** - 確保 tasks.md 中每項都完成
6. **更新檢查清單** - 將所有完成項目標記為 `- [x]`

**批准閘道：**
在開始實作前，提案必須經過審查和批准。

### 階段 3：歸檔變更

部署後，建立獨立的 PR 進行歸檔：

```bash
# 歸檔變更（會更新 specs/ 並移動到 archive/）
openspec archive refactor-domain-models --yes

# 驗證歸檔結果
openspec validate --strict
```

歸檔會：
- 將 `changes/[name]/` 移動到 `changes/archive/YYYY-MM-DD-[name]/`
- 根據 deltas 更新 `specs/`（ADDED、MODIFIED、REMOVED）
- 保留變更歷史記錄

## 常用指令

```bash
# 列出所有進行中的變更
openspec list

# 列出所有現有規範
openspec list --specs

# 查看變更或規範詳情
openspec show [item]

# 驗證變更或規範
openspec validate [item] --strict

# 歸檔已完成的變更
openspec archive <change-id> --yes

# 查看變更的 deltas（除錯用）
openspec show [change] --json --deltas-only
```

## 與 AI 助手協作

### 建立變更提案時

**告訴 AI：**
- "幫我建立變更提案來 [描述功能]"
- "我想要 [描述需求]，請建立 OpenSpec 提案"

**AI 會：**
1. 檢查 `openspec/project.md` 了解專案慣例
2. 執行 `openspec list` 和 `openspec list --specs` 檢查現有狀態
3. 選擇適當的 change-id
4. 建立 proposal.md、tasks.md 和 spec deltas
5. 執行 `openspec validate --strict` 確保格式正確

### 實作變更時

**告訴 AI：**
- "請根據 OpenSpec 提案 [change-id] 開始實作"
- "繼續實作 tasks.md 中的下一個任務"

**AI 會：**
1. 讀取 proposal.md 了解目標
2. 讀取相關的 spec deltas 了解需求
3. 讀取 tasks.md 了解實作步驟
4. 依序完成任務
5. 更新 tasks.md 的完成狀態

### 最佳實踐

1. **一次一個變更** - 避免同時進行多個大型變更
2. **小步前進** - 將大型變更拆分成多個小提案
3. **先寫規範** - 在寫程式前先定義行為
4. **情境測試** - 每個需求至少要有一個可驗證的情境
5. **保持同步** - 完成後記得歸檔，讓 specs/ 反映真實狀態

## 目前進行中的變更

### refactor-domain-models

**目標**：將現有排程系統改用統一的領域模型

**狀態**：
- ✅ domainModels.js 已建立並提交
- ✅ AGENTS.md 已更新
- ✅ REFACTORING_PLAN.md 已建立
- ✅ OpenSpec 提案已建立並驗證通過
- ⬜ examService.js 待建立
- ⬜ 排程函式待重寫

**下一步**：
1. 實作 `examService.js`
2. 重寫排程邏輯
3. 完整測試
4. 部署後歸檔

## 疑問解答

**Q: 何時使用 ADDED vs MODIFIED？**
- ADDED：引入全新的能力或子能力，可獨立作為需求
- MODIFIED：改變現有需求的行為、範圍或驗收標準。必須貼上完整的修改後內容

**Q: Scenario 格式為什麼這麼嚴格？**
- OpenSpec 工具需要精確解析 Scenario 結構
- 必須使用 `####`（4個#）
- 格式：`#### Scenario: 名稱`

**Q: 驗證失敗怎麼辦？**
1. 執行 `openspec validate [change] --strict` 查看詳細錯誤
2. 檢查 Scenario 格式是否正確
3. 確認每個需求都包含 SHALL 或 MUST
4. 使用 `--json --deltas-only` 查看解析結果

**Q: 可以跳過提案直接改程式碼嗎？**
- 小型 bug 修復：可以
- 新功能或破壞性變更：建議先建立提案
- 不確定時：建立提案較安全

---

**記住**：規範（specs/）是真理，變更（changes/）是提案。保持它們同步！

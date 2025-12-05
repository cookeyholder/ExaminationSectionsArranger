## CURRENT Requirements

### Requirement: 共同科目節次安排
The system SHALL 根據參數表的規則，為共同科目（物理、國文、英文、數學等）安排指定的節次。此功能現使用 Exam 領域模型實作。

#### Scenario: 安排共同科目到指定節次
- **GIVEN** 參數表中定義了「數學 → 節次1」的規則
- **AND** 補考名單中有數學科的學生
- **WHEN** 執行 `scheduleCommonSubjectSessions()`
- **THEN** 所有數學科學生的「節次」欄位設為 1
- **AND** 變更寫回工作表

#### Scenario: 忽略未定義規則的科目
- **GIVEN** 參數表中沒有定義「專題製作」的節次規則
- **AND** 補考名單中有專題製作的學生
- **WHEN** 執行 `scheduleCommonSubjectSessions()`
- **THEN** 專題製作學生的節次欄位保持原值（0 或已分配的節次）

### Requirement: 專業科目節次安排
The system SHALL 根據容量限制和科別年級互斥規則，自動為專業科目安排節次。此功能現使用 Exam 領域模型實作。

#### Scenario: 按科別年級科目人數排序分配
- **GIVEN** 節次容量為 30 人
- **AND** 有以下科目待分配：「資訊三_程式設計」20人、「機械二_機械力學」15人
- **WHEN** 執行 `scheduleSpecializedSubjectSessions()`
- **THEN** 人數較多的科目優先分配到節次1
- **AND** 「資訊三_程式設計」分配到節次1
- **AND** 「機械二_機械力學」分配到節次1（因容量足夠且科別不衝突）

#### Scenario: 科別年級互斥檢查
- **GIVEN** 節次1 已有「資訊三_程式設計」20人
- **AND** 待分配「資訊三_電子學」10人
- **WHEN** 執行專業科目節次安排
- **THEN** 「資訊三_電子學」不能分配到節次1（同科別年級已存在）
- **AND** 分配到節次2

#### Scenario: 節次容量限制
- **GIVEN** 節次容量為 30 人（參數 B9 的 90%）
- **AND** 節次1 已有 25 人
- **AND** 待分配科目有 10 人
- **WHEN** 執行專業科目節次安排
- **THEN** 該科目不能分配到節次1（超過容量）
- **AND** 嘗試分配到節次2

### Requirement: 試場分配
The system SHALL 根據容量和科目數限制，將每節次的學生分配到試場。此功能現使用 Exam 領域模型和 `distributeToChildren()` 實作。

#### Scenario: 按班級科目人數排序分配
- **GIVEN** 試場容量為 35 人
- **AND** 節次1 有以下班級科目：「資三甲_數學」20人、「機二乙_物理」15人
- **WHEN** 執行 `assignExamRooms()`
- **THEN** 人數較多的班級科目優先分配
- **AND** 「資三甲_數學」分配到試場1
- **AND** 「機二乙_物理」分配到試場1（容量足夠）

#### Scenario: 試場容量限制
- **GIVEN** 試場容量為 35 人
- **AND** 試場1 已有 30 人
- **AND** 待分配班級科目有 10 人
- **WHEN** 執行試場分配
- **THEN** 該班級科目不能分配到試場1
- **AND** 分配到試場2

#### Scenario: 科目數限制
- **GIVEN** 每試場最多 3 科（參數 B8）
- **AND** 試場1 已有 3 個科目
- **WHEN** 嘗試分配第 4 個科目到試場1
- **THEN** 分配失敗
- **AND** 該科目分配到試場2

#### Scenario: 使用分配機制將學生分配到試場
- **GIVEN** 節次的學生資料已包含試場編號
- **WHEN** 執行試場分配後呼叫 `session.distributeToChildren()`
- **THEN** 學生根據試場編號自動分配到對應的 classroom
- **AND** classroom.students 包含該試場的所有學生

### Requirement: 大小袋編號計算
The system SHALL 為每個試場計算唯一的大袋和小袋編號。此功能現使用 Exam 領域模型實作。

#### Scenario: 計算小袋序號
- **GIVEN** 節次1 試場1 有學生
- **AND** 節次1 試場2 有學生
- **WHEN** 執行 `allocateBagIdentifiers()`
- **THEN** 試場1 的所有學生的「小袋序號」設為相同值（如 1）
- **AND** 試場2 的所有學生的「小袋序號」設為下一個值（如 2）

#### Scenario: 計算大袋序號
- **GIVEN** 每個試場對應一個大袋
- **WHEN** 執行大袋編號計算
- **THEN** 每個試場的所有學生的「大袋序號」相同
- **AND** 大袋序號與小袋序號對應（一對一關係）

### Requirement: 試場時間填充
The system SHALL 根據參數表的節次時間規則，填充每個學生的「時間」欄位。此功能現使用 Exam 領域模型實作。

#### Scenario: 填充節次時間
- **GIVEN** 參數表定義「節次1 → 08:00-09:00」
- **AND** 學生分配在節次1
- **WHEN** 執行 `populateSessionTimes()`
- **THEN** 該學生的「時間」欄位設為「08:00-09:00」

### Requirement: 人數統計更新
The system SHALL 計算並更新小袋人數、大袋人數、班級人數欄位。此功能現使用 Exam 領域模型實作。

#### Scenario: 計算小袋人數
- **GIVEN** 試場1 有 25 個學生
- **WHEN** 執行 `updateBagAndClassPopulations()`
- **THEN** 試場1 所有學生的「小袋人數」設為 25

#### Scenario: 計算大袋人數
- **GIVEN** 試場1 有 25 個學生（一個試場一個大袋）
- **WHEN** 執行人數統計更新
- **THEN** 試場1 所有學生的「大袋人數」設為 25

#### Scenario: 計算班級人數
- **GIVEN** 「資三甲」班級在整場考試中有 50 個學生（分散在不同節次和試場）
- **WHEN** 執行人數統計更新
- **THEN** 所有「資三甲」學生的「班級人數」設為 50

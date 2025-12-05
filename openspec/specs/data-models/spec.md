## CURRENT Requirements

### Requirement: 統計容器建構器
系統 SHALL 提供通用的統計容器建構器，支援動態配置統計維度。

#### Scenario: 建立具有統計功能的容器
- **WHEN** 使用 `createStatisticsContainer()` 並提供統計維度配置
- **THEN** 返回的物件包含 `students` 陣列、`population` 屬性、`addStudent()` 方法
- **AND** 根據配置產生對應的統計屬性

#### Scenario: 存取動態統計
- **GIVEN** 容器已配置 `departmentDistribution` 統計維度
- **WHEN** 存取 `container.departmentDistribution`
- **THEN** 返回科別分布的統計物件（格式: `{"科別名稱": 人數}`）

#### Scenario: 使用通用統計方法
- **GIVEN** 容器已配置多個統計維度
- **WHEN** 呼叫 `container.statistics('dimensionName')`
- **THEN** 返回該維度的統計結果
- **AND** 若維度不存在則拋出錯誤並列出可用維度

### Requirement: 試場（Classroom）領域模型
The system SHALL provide Classroom 領域模型，作為學生資料的唯一真實來源。

#### Scenario: 建立試場記錄
- **WHEN** 呼叫 `createClassroomRecord()`
- **THEN** 返回具有以下屬性的物件：
  - `students`: 空陣列
  - `population`: 0
  - `classSubjectStatistics`: 空物件

#### Scenario: 新增學生到試場
- **GIVEN** 已建立試場記錄
- **WHEN** 呼叫 `classroom.addStudent(studentRow)`
- **THEN** 學生被加入 `students` 陣列
- **AND** `population` 自動更新

#### Scenario: 取得班級科目統計
- **GIVEN** 試場中有多個學生
- **WHEN** 存取 `classroom.classSubjectStatistics`
- **THEN** 返回格式為 `{"班級_科目": 人數}` 的統計物件

### Requirement: 節次（Session）領域模型
The system SHALL provide Session 領域模型，包含多個 Classroom 並提供節次層級的統計。

#### Scenario: 建立節次記錄
- **GIVEN** 最大試場數為 5
- **WHEN** 呼叫 `createSessionRecord(5)`
- **THEN** 返回具有以下屬性的物件：
  - `students`: 空陣列
  - `classrooms`: 包含 6 個 Classroom 的陣列（索引 0-5，0 不使用）
  - `population`: 0

#### Scenario: 取得科別年級統計
- **GIVEN** 節次中有多個學生
- **WHEN** 存取 `session.departmentGradeStatistics`
- **THEN** 返回格式為 `{"科別年級": 人數}` 的統計物件（如 `{"資訊三": 10}`）

#### Scenario: 取得班級科目統計
- **GIVEN** 節次中有多個學生
- **WHEN** 存取 `session.departmentClassSubjectStatistics`
- **THEN** 返回格式為 `{"班級科目": 人數}` 的統計物件（如 `{"資三甲數學": 5}`）

#### Scenario: 分配學生到試場
- **GIVEN** 節次中有學生且學生資料包含試場編號
- **WHEN** 呼叫 `session.distributeToChildren((student) => student[roomIndex])`
- **THEN** 學生根據試場編號分配到對應的 classroom
- **AND** 每個 classroom.students 包含該試場的學生

### Requirement: 考試（Exam）領域模型
The system SHALL provide Exam 領域模型，作為整場考試活動的頂層容器。

#### Scenario: 建立考試記錄
- **GIVEN** 最大節次數為 9，最大試場數為 5
- **WHEN** 呼叫 `createExamRecord(9, 5)`
- **THEN** 返回具有以下屬性的物件：
  - `students`: 空陣列
  - `sessions`: 包含 11 個 Session 的陣列（索引 0-10，0 不使用）
  - 每個 Session 包含 6 個 Classroom

#### Scenario: 取得節次分布統計
- **GIVEN** 考試中有學生分布在不同節次
- **WHEN** 存取 `exam.sessionDistribution`
- **THEN** 返回格式為 `{"節次N": 人數}` 的統計物件

#### Scenario: 取得科別分布統計
- **GIVEN** 考試中有多個科別的學生
- **WHEN** 存取 `exam.departmentDistribution`
- **THEN** 返回格式為 `{"科別": 人數}` 的統計物件

#### Scenario: 取得年級分布統計
- **GIVEN** 考試中有不同年級的學生
- **WHEN** 存取 `exam.gradeDistribution`
- **THEN** 返回格式為 `{"N年級": 人數}` 的統計物件

#### Scenario: 取得科目分布統計
- **GIVEN** 考試中有多個科目
- **WHEN** 存取 `exam.subjectDistribution`
- **THEN** 返回格式為 `{"科目": 人數}` 的統計物件

### Requirement: 子容器分配機制
具有子容器的領域模型 MUST 提供 `distributeToChildren()` 方法，支援向下分配學生。

#### Scenario: 使用分配函式分配學生
- **GIVEN** 容器有學生和初始化的子容器
- **WHEN** 呼叫 `container.distributeToChildren((student, children) => targetIndex)`
- **THEN** 每個學生根據分配函式的回傳值分配到對應的子容器
- **AND** 若索引超出範圍則忽略該學生

#### Scenario: 負載平衡分配
- **GIVEN** 節次有學生需要分配到試場
- **WHEN** 使用負載平衡策略的分配函式
- **THEN** 學生均勻分配到各試場
- **AND** 各試場的 population 儘可能接近

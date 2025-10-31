/**
 * 通用統計容器建構器
 * 提供學生集合管理和動態統計功能
 */
function createStatisticsContainer(config = {}) {
    const { statisticsDimensions = [], children = null } = config;

    // 儲存統計維度設定，供 statistics() 方法使用
    const dimensionsMap = {};
    statisticsDimensions.forEach((dimension) => {
        dimensionsMap[dimension.name] = dimension.keyGenerator;
    });

    const container = {
        students: [],

        /**
         * 計算容器中的學生總數
         */
        get population() {
            return this.students.length;
        },

        /**
         * 新增學生到容器中
         * @param {Array} studentRow - 學生資料列
         */
        addStudent(studentRow) {
            this.students.push(studentRow);
        },

        /**
         * 清空容器中的所有學生
         */
        clear() {
            this.students = [];
        },

        /**
         * 通用統計方法 - 根據指定的維度名稱產生統計
         * @param {string} dimensionName - 統計維度名稱
         * @returns {Object} 統計結果物件
         * @example
         * session.statistics('departmentGradeStatistics')
         * // 返回 {"資訊三": 10, "機械二": 8}
         */
        statistics(dimensionName) {
            const keyGenerator = dimensionsMap[dimensionName];
            if (!keyGenerator) {
                throw new Error(
                    `未定義的統計維度: ${dimensionName}。可用維度: ${Object.keys(
                        dimensionsMap
                    ).join(", ")}`
                );
            }

            const stats = {};
            this.students.forEach((studentRow) => {
                const key = keyGenerator(studentRow);
                stats[key] = (stats[key] || 0) + 1;
            });
            return stats;
        },

        /**
         * 取得所有可用的統計維度名稱
         * @returns {Array<string>} 統計維度名稱陣列
         */
        getAvailableStatistics() {
            return Object.keys(dimensionsMap);
        },
    };

    // 動態建立統計屬性（方便直接存取）
    statisticsDimensions.forEach((dimension) => {
        Object.defineProperty(container, dimension.name, {
            get: function () {
                return this.statistics(dimension.name);
            },
            enumerable: true,
        });
    });

    // 如果有子容器配置，建立子容器陣列
    if (children) {
        container[children.propertyName] = [];

        /**
         * 初始化子容器
         * @param {number} count - 要建立的子容器數量
         */
        container.initializeChildren = function (count) {
            this[children.propertyName] = [];
            for (let i = 0; i < count; i++) {
                this[children.propertyName].push(children.factory());
            }
        };

        /**
         * 將此容器的所有學生分配到子容器（向下分配）
         *
         * @param {Function} distributionFn - 分配函式，決定學生要分配到哪個子容器
         *   簽章: (student, children) => childIndex
         *   參數:
         *     - student: 學生資料列
         *     - children: 子容器陣列（例如 classrooms 或 sessions）
         *   回傳: 子容器的索引編號（0-based）
         *
         * @example
         * // 根據學生資料中的試場編號分配到對應的 classroom
         * session.distributeToChildren((student, classrooms) => {
         *   return student[roomIndex]; // 回傳試場編號作為索引
         * });
         *
         * @example
         * // 使用負載平衡策略分配
         * session.distributeToChildren((student, classrooms) => {
         *   // 找出人數最少的試場
         *   let minIndex = 0;
         *   let minPopulation = classrooms[0].population;
         *   for (let i = 1; i < classrooms.length; i++) {
         *     if (classrooms[i].population < minPopulation) {
         *       minIndex = i;
         *       minPopulation = classrooms[i].population;
         *     }
         *   }
         *   return minIndex;
         * });
         */
        container.distributeToChildren = function (distributionFn) {
            this.students.forEach((student) => {
                const childIndex = distributionFn(
                    student,
                    this[children.propertyName]
                );
                if (
                    childIndex >= 0 &&
                    childIndex < this[children.propertyName].length
                ) {
                    this[children.propertyName][childIndex].addStudent(student);
                }
            });
        };
    }

    return container;
}

/**
 * 建立試場（Classroom）記錄
 * 最底層容器，儲存單一試場的學生資料
 * 統計維度：班級-科目
 */
function createClassroomRecord() {
    return createStatisticsContainer({
        statisticsDimensions: [
            {
                name: "classSubjectStatistics",
                keyGenerator: (studentRow) =>
                    `${studentRow[3]}_${studentRow[7]}`, // 班級_科目
            },
        ],
    });
}

/**
 * 建立節次（Session）記錄
 * 中層容器，包含多個試場
 * 統計維度：科別-年級、班級-科目
 * 子容器：試場（classrooms）
 */
function createSessionRecord(maxRoomCount) {
    const session = createStatisticsContainer({
        statisticsDimensions: [
            {
                name: "departmentGradeStatistics",
                keyGenerator: (studentRow) =>
                    `${studentRow[0]}${studentRow[1]}`, // 科別年級
            },
            {
                name: "departmentClassSubjectStatistics",
                keyGenerator: (studentRow) =>
                    `${studentRow[3]}${studentRow[7]}`, // 班級科目
            },
        ],
        children: {
            propertyName: "classrooms",
            factory: createClassroomRecord,
        },
    });

    // 初始化試場
    if (maxRoomCount !== undefined) {
        session.initializeChildren(maxRoomCount + 1);
    }

    return session;
}

/**
 * 建立考試（Exam）記錄
 * 最高層容器，包含多個節次，代表整場補考活動
 * 統計維度：節次分布、科別分布、年級分布
 * 子容器：節次（sessions）
 *
 * @param {number} maxSessionCount - 最大節次數量
 * @param {number} maxRoomCount - 每個節次的最大試場數量
 * @returns {Object} Exam 記錄物件
 *
 * @example
 * const exam = createExamRecord(9, 5);
 * exam.sessions[1].classrooms[0].addStudent(studentRow);
 * console.log(exam.departmentDistribution); // {"資訊": 50, "機械": 30}
 * console.log(exam.sessionDistribution);    // {"節次1": 20, "節次2": 15}
 */
function createExamRecord(maxSessionCount, maxRoomCount) {
    const exam = createStatisticsContainer({
        statisticsDimensions: [
            {
                name: "sessionDistribution",
                keyGenerator: (studentRow) => {
                    const sessionIndex = 8; // 節次欄位索引
                    return `節次${studentRow[sessionIndex]}`;
                },
            },
            {
                name: "departmentDistribution",
                keyGenerator: (studentRow) => studentRow[0], // 科別
            },
            {
                name: "gradeDistribution",
                keyGenerator: (studentRow) => `${studentRow[1]}年級`, // 年級
            },
            {
                name: "subjectDistribution",
                keyGenerator: (studentRow) => studentRow[7], // 科目
            },
        ],
        children: {
            propertyName: "sessions",
            factory: () => createSessionRecord(maxRoomCount),
        },
    });

    // 初始化節次（索引 0 不使用，從 1 開始）
    if (maxSessionCount !== undefined) {
        exam.initializeChildren(maxSessionCount + 2);
    }

    return exam;
}

/**
 * 建立考試批次（ExamBatch）記錄 - 別名，向後相容
 * @deprecated 請使用 createExamRecord() 替代
 */
function createExamBatchRecord(maxSessionCount, maxRoomCount) {
    return createExamRecord(maxSessionCount, maxRoomCount);
}

// ============================================================================
// 內部純函式 - Pipeline 處理（無 I/O）
// ============================================================================

/**
 * 在記憶體中建立統計物件
 *
 * @param {Array<Array>} students - 學生資料陣列
 * @param {Object} columns - 欄位索引
 * @returns {Object} 統計物件
 */
function buildInMemoryStatistics(students, columns) {
    const stats = {
        bySession: {}, // sessionNumber -> { students: [], deptGrade: {}, deptClassSubject: {} }
        byDeptGradeSubject: {}, // "科別年級_科目" -> count
    };

    students.forEach(function (student, index) {
        const session = student[columns.session];
        const deptGrade = student[columns.department] + student[columns.grade];
        const deptGradeSubject = deptGrade + "_" + student[columns.subject];
        const classSubject = student[columns.class] + student[columns.subject];

        // 初始化節次統計
        if (!stats.bySession[session]) {
            stats.bySession[session] = {
                indices: [],
                population: 0,
                deptGrade: {},
                deptClassSubject: {},
            };
        }

        const sessionStats = stats.bySession[session];
        sessionStats.indices.push(index);
        sessionStats.population++;
        sessionStats.deptGrade[deptGrade] =
            (sessionStats.deptGrade[deptGrade] || 0) + 1;
        sessionStats.deptClassSubject[classSubject] =
            (sessionStats.deptClassSubject[classSubject] || 0) + 1;

        // 全域科別年級科目統計
        stats.byDeptGradeSubject[deptGradeSubject] =
            (stats.byDeptGradeSubject[deptGradeSubject] || 0) + 1;
    });

    return stats;
}

/**
 * 安排共同科目節次（純函式）
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 * @param {Object} sessionRules - 科目到節次的對映
 */
function scheduleCommonSubjectSessionsInternal(
    students,
    columns,
    sessionRules
) {
    students.forEach(function (student) {
        const subjectName = student[columns.subject];
        const preferredSession = sessionRules[subjectName];
        if (preferredSession != null) {
            student[columns.session] = preferredSession;
        }
    });
}

/**
 * 安排專業科目節次（純函式）
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 * @param {Object} params - 參數物件
 * @returns {number} 未排入節次的學生數量
 */
function scheduleSpecializedSubjectSessionsInternal(students, columns, params) {
    const sessionCapacity = 0.9 * params.sessionCapacity;
    const maxSessionCount = params.maxSessionCount;

    // 只統計尚未分配節次的學生（節次 = 0 或空值）
    const deptGradeSubjectCounts = {};
    students.forEach(function (student) {
        // 只處理尚未分配節次的學生
        if (student[columns.session] !== 0 && student[columns.session] !== "") {
            return;
        }
        const key =
            student[columns.department] +
            student[columns.grade] +
            "_" +
            student[columns.subject];
        deptGradeSubjectCounts[key] = (deptGradeSubjectCounts[key] || 0) + 1;
    });

    const sortedCounts = Object.entries(deptGradeSubjectCounts).sort(
        compareCountDescending
    );

    // 節次統計
    const sessionStats = {};
    for (let i = 1; i <= maxSessionCount; i++) {
        sessionStats[i] = { population: 0, deptGrade: {} };
    }

    // 分配學生到節次
    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const session = sessionStats[sessionNumber];

        for (let i = 0; i < sortedCounts.length; i++) {
            const [deptGradeSubjectKey, studentCount] = sortedCounts[i];
            const deptGradeKey = deptGradeSubjectKey.substring(
                0,
                deptGradeSubjectKey.indexOf("_")
            );

            // 檢查互斥規則
            if (Object.keys(session.deptGrade).includes(deptGradeKey)) {
                continue;
            }

            // 檢查容量
            if (studentCount + session.population > sessionCapacity) {
                continue;
            }

            // 分配學生
            students.forEach(function (student) {
                const studentKey =
                    student[columns.department] +
                    student[columns.grade] +
                    "_" +
                    student[columns.subject];
                // 只分配尚未分配節次的學生（節次 = 0 或空值）
                const currentSession = student[columns.session];
                if (
                    studentKey === deptGradeSubjectKey &&
                    (currentSession === 0 || currentSession === "")
                ) {
                    student[columns.session] = sessionNumber;
                    session.population++;
                    session.deptGrade[deptGradeKey] =
                        (session.deptGrade[deptGradeKey] || 0) + 1;
                }
            });
        }
    }

    // 檢查未分配（節次 = 0 或空值）
    const unscheduledCount = students.filter(function (s) {
        const session = s[columns.session];
        return session === 0 || session === "";
    }).length;

    if (unscheduledCount > 0) {
        Logger.log(
            "無法將所有人排入 " +
                maxSessionCount +
                " 節，請檢查是否有某科年級須補考過多科目！"
        );
    }

    return unscheduledCount;
}

/**
 * 安排試場（純函式）
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 * @param {Object} params - 參數物件
 * @returns {boolean} 是否所有學生都已分配
 */
function assignExamRoomsInternal(students, columns, params) {
    const maxSessionCount = params.maxSessionCount;
    const maxRoomCount = params.maxRoomCount;
    const maxStudentsPerRoom = params.maxStudentsPerRoom;
    const maxSubjectsPerRoom = params.maxSubjectsPerRoom;

    // 重置所有試場
    students.forEach(function (student) {
        student[columns.room] = 0;
    });

    let allScheduled = true;

    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        // 取得本節次學生
        const sessionStudents = students.filter(function (s) {
            return s[columns.session] === sessionNumber;
        });

        if (sessionStudents.length === 0) continue;

        // 計算班級科目統計
        const deptClassSubjectCounts = {};
        sessionStudents.forEach(function (student) {
            const key = student[columns.class] + student[columns.subject];
            deptClassSubjectCounts[key] =
                (deptClassSubjectCounts[key] || 0) + 1;
        });

        const sortedCounts = Object.entries(deptClassSubjectCounts).sort(
            compareCountDescending
        );

        // 試場統計
        const roomStats = {};
        for (let r = 1; r <= maxRoomCount; r++) {
            roomStats[r] = { population: 0, classSubject: {} };
        }

        for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
            const room = roomStats[roomNumber];
            const scheduledKeys = [];

            for (let i = 0; i < sortedCounts.length; i++) {
                const [classSubjectKey, count] = sortedCounts[i];

                if (scheduledKeys.includes(classSubjectKey)) continue;
                if (count + room.population > maxStudentsPerRoom) continue;
                if (
                    Object.keys(room.classSubject).length + 1 >
                    maxSubjectsPerRoom
                )
                    continue;

                // 分配學生
                students.forEach(function (student) {
                    if (student[columns.session] !== sessionNumber) return;
                    const studentKey =
                        student[columns.class] + student[columns.subject];
                    if (
                        studentKey === classSubjectKey &&
                        student[columns.room] === 0
                    ) {
                        student[columns.room] = roomNumber;
                        room.population++;
                        room.classSubject[classSubjectKey] =
                            (room.classSubject[classSubjectKey] || 0) + 1;
                    }
                });

                scheduledKeys.push(classSubjectKey);
            }
        }

        // 檢查未分配
        sessionStudents.forEach(function (student) {
            if (student[columns.room] === 0) {
                allScheduled = false;
            }
        });
    }

    return allScheduled;
}

/**
 * 排序學生（純函式）- 依節次試場
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 */
function sortStudentsInternal(students, columns) {
    students.sort(function (a, b) {
        if (a[columns.session] !== b[columns.session])
            return a[columns.session] - b[columns.session];
        if (a[columns.room] !== b[columns.room])
            return a[columns.room] - b[columns.room];
        if (a[columns.department] !== b[columns.department])
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        if (a[columns.grade] !== b[columns.grade])
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        if (a[columns.seatNumber] !== b[columns.seatNumber])
            return a[columns.seatNumber] - b[columns.seatNumber];
        return a[columns.subject].localeCompare(b[columns.subject], "zh-TW");
    });
}

/**
 * 計算大、小袋編號（純函式）
 *
 * 注意：此函式假設學生已依節次、試場排序。
 * 每個試場（節次+試場組合）對應一個大袋，每個班級科目組合對應一個小袋。
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改，需先排序）
 * @param {Object} columns - 欄位索引
 */
function allocateBagIdentifiersInternal(students, columns) {
    if (students.length === 0) return;

    let smallBagCounter = 1;
    let bigBagCounter = 1;
    let lastSession = null;
    let lastRoom = null;
    let smallBagMapping = {};

    students.forEach(function (student) {
        const session = student[columns.session];
        const room = student[columns.room];
        const classSubjectKey =
            student[columns.class] + student[columns.subject];

        // 新試場（節次或試場變更）-> 新大袋
        if (session !== lastSession || room !== lastRoom) {
            // 第一個試場不需要增加計數，之後每個新試場都增加
            if (lastSession !== null) {
                bigBagCounter++;
            }
            smallBagMapping = {};
            lastSession = session;
            lastRoom = room;
        }

        // 新班級科目組合 -> 新小袋
        if (!smallBagMapping[classSubjectKey]) {
            smallBagMapping[classSubjectKey] = smallBagCounter++;
        }

        student[columns.smallBagId] = smallBagMapping[classSubjectKey];
        student[columns.bigBagId] = bigBagCounter;
    });
}

/**
 * 填入節次時間（純函式）
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 * @param {Array<Array>} sessionTimes - 節次時間原始資料
 */
function populateSessionTimesInternal(students, columns, sessionTimes) {
    // 建立節次到時間的對映
    const timeLookup = {};
    sessionTimes.slice(1).forEach(function (row) {
        timeLookup[row[0]] = row[1];
    });

    students.forEach(function (student) {
        const session = student[columns.session];
        student[columns.time] = timeLookup[session] || "";
    });
}

/**
 * 計算各種人數（純函式）
 *
 * @param {Array<Array>} students - 學生資料陣列（會被修改）
 * @param {Object} columns - 欄位索引
 */
function updateBagAndClassPopulationsInternal(students, columns) {
    // 計算班級總人數
    const classPopulationMap = {};
    students.forEach(function (student) {
        const className = student[columns.class];
        classPopulationMap[className] =
            (classPopulationMap[className] || 0) + 1;
    });

    // 計算每個試場內的班級科目人數
    const roomClassSubjectMap = {}; // "session_room_classSubject" -> count
    const roomPopulationMap = {}; // "session_room" -> count

    students.forEach(function (student) {
        const roomKey = student[columns.session] + "_" + student[columns.room];
        const classSubjectKey =
            roomKey + "_" + student[columns.class] + student[columns.subject];

        roomClassSubjectMap[classSubjectKey] =
            (roomClassSubjectMap[classSubjectKey] || 0) + 1;
        roomPopulationMap[roomKey] = (roomPopulationMap[roomKey] || 0) + 1;
    });

    // 更新學生資料
    students.forEach(function (student) {
        const roomKey = student[columns.session] + "_" + student[columns.room];
        const classSubjectKey =
            roomKey + "_" + student[columns.class] + student[columns.subject];

        student[columns.smallBagPopulation] =
            roomClassSubjectMap[classSubjectKey] || 0;
        student[columns.bigBagPopulation] = roomPopulationMap[roomKey] || 0;
        student[columns.classPopulation] =
            classPopulationMap[student[columns.class]] || 0;
    });
}

// ============================================================================
// 公開 API（向後相容）
// ============================================================================

/**
 * 安排共同科目的節次（物理、國文、英文、數學等）
 *
 * 根據「參數區」工作表中的科目-節次對映規則，
 * 將共同科目分配到指定的節次。
 */
function scheduleCommonSubjectSessions() {
    const sessionRuleRows = PARAMETERS_SHEET.getRange(2, 5, 21, 2).getValues();
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    // 建立科目到節次的對映
    const preferredSessionBySubject = {};
    sessionRuleRows.forEach(function (ruleRow) {
        if (ruleRow[0] && ruleRow[1]) {
            preferredSessionBySubject[ruleRow[0]] = ruleRow[1];
        }
    });

    // 重新分配共同科目的節次
    exam.sessions.forEach(function (session) {
        session.students.forEach(function (student) {
            const subjectName = student[columns.subject];
            const preferredSession = preferredSessionBySubject[subjectName];
            if (preferredSession != null) {
                student[columns.session] = preferredSession;
            }
        });
    });

    // 儲存更新後的資料
    saveExamToSheet(exam);
}

/**
 * 安排專業科目的節次
 *
 * 根據科別年級互斥規則和節次容量限制，
 * 將專業科目分配到適當的節次。
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
    exam.sessions.forEach(function (session) {
        session.clear();
    });

    // 收集所有未分配節次的學生
    const allStudents = [];
    exam.sessions[0].students.forEach(function (student) {
        if (student[columns.session] === 0) {
            allStudents.push(student);
        }
    });

    // 為每個節次分配學生
    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const session = exam.sessions[sessionNumber];

        for (
            let countIndex = 0;
            countIndex < departmentGradeSubjectCounts.length;
            countIndex++
        ) {
            const [deptGradeSubjectKey, studentCount] =
                departmentGradeSubjectCounts[countIndex];
            const deptGradeKey = deptGradeSubjectKey.substring(
                0,
                deptGradeSubjectKey.indexOf("_")
            );

            // 檢查該科別年級是否已排入此節次（互斥規則）
            const deptGradeStats = session.departmentGradeStatistics;
            if (Object.keys(deptGradeStats).includes(deptGradeKey)) {
                continue;
            }

            // 檢查容量
            if (studentCount + session.population > sessionCapacity) {
                continue;
            }

            // 分配學生到此節次
            allStudents.forEach(function (student) {
                const studentKey =
                    student[columns.department] +
                    student[columns.grade] +
                    "_" +
                    student[columns.subject];
                if (
                    studentKey === deptGradeSubjectKey &&
                    student[columns.session] === 0
                ) {
                    student[columns.session] = sessionNumber;
                    session.addStudent(student);
                }
            });
        }

        if (session.population >= sessionCapacity) {
            Logger.log(
                "第" + sessionNumber + "節已達人數上限：" + session.population
            );
        }
    }

    // 檢查是否所有學生都已分配
    const unscheduledCount = allStudents.filter(function (s) {
        return s[columns.session] === 0;
    }).length;

    if (unscheduledCount > 0) {
        Logger.log(
            "無法將所有人排入 " +
                maxSessionCount +
                " 節，請檢查是否有某科年級須補考過多科目！"
        );
        SpreadsheetApp.getUi().alert(
            "無法將所有人排入 " +
                maxSessionCount +
                " 節，請檢查是否有某科年級須補考過多科目！"
        );
    }

    saveExamToSheet(exam);
}

/**
 * 安排試場
 *
 * 根據班級科目統計和試場容量限制，
 * 將學生分配到試場。
 */
function assignExamRooms() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();
    const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
    const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();
    const maxStudentsPerRoom = PARAMETERS_SHEET.getRange("B7").getValue();
    const maxSubjectsPerRoom = PARAMETERS_SHEET.getRange("B8").getValue();

    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const session = exam.sessions[sessionNumber];

        // 清空所有試場
        session.classrooms.forEach(function (classroom) {
            classroom.clear();
        });

        const deptClassSubjectCounts = Object.entries(
            session.departmentClassSubjectStatistics
        ).sort(compareCountDescending);

        for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
            const classroom = session.classrooms[roomNumber];
            let scheduledSubjects = [];

            for (
                let countIndex = 0;
                countIndex < deptClassSubjectCounts.length;
                countIndex++
            ) {
                const [classSubjectKey, count] =
                    deptClassSubjectCounts[countIndex];

                // 檢查是否已排入
                if (scheduledSubjects.includes(classSubjectKey)) {
                    continue;
                }

                // 檢查容量
                if (count + classroom.population > maxStudentsPerRoom) {
                    continue;
                }

                // 檢查科目數限制
                const subjectCount = Object.keys(
                    classroom.classSubjectStatistics
                ).length;
                if (subjectCount + 1 > maxSubjectsPerRoom) {
                    continue;
                }

                // 分配學生到此試場
                session.students.forEach(function (student) {
                    const studentKey =
                        student[columns.class] + student[columns.subject];
                    if (
                        studentKey === classSubjectKey &&
                        student[columns.room] === 0
                    ) {
                        student[columns.room] = roomNumber;
                        classroom.addStudent(student);
                    }
                });

                scheduledSubjects = Object.keys(
                    classroom.classSubjectStatistics
                );
            }
        }
    }

    // 檢查是否所有學生都已分配試場
    let allScheduled = true;
    exam.sessions.forEach(function (session) {
        session.students.forEach(function (student) {
            if (student[columns.room] === 0) {
                allScheduled = false;
            }
        });
    });

    if (!allScheduled) {
        Logger.log(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
        SpreadsheetApp.getUi().alert(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
    }

    saveExamToSheet(exam);

    // 檢查是否有學生被安排在第 9 節
    if (exam.sessions[9] && exam.sessions[9].population > 0) {
        SpreadsheetApp.getUi().alert(
            "部分考生被安排在第9節補考，請注意是否需要調整到中午應試！"
        );
    }

    FILTERED_RESULT_SHEET.getRange("I:J").setNumberFormat("#,##0");
    sortFilteredStudentsBySessionRoom();
}

/**
 * 計算大、小袋編號
 *
 * 遍歷所有試場的學生，為每個試場和班級科目組合分配唯一的編號。
 */
function allocateBagIdentifiers() {
    sortFilteredStudentsBySessionRoom();

    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    let smallBagCounter = 1;
    let bigBagCounter = 1;

    exam.sessions.forEach(function (session) {
        session.classrooms.forEach(function (classroom) {
            if (classroom.population > 0) {
                // 為每個班級科目組合分配小袋序號
                const classSubjectKeys = Object.keys(
                    classroom.classSubjectStatistics
                );
                const smallBagMapping = {};

                classSubjectKeys.forEach(function (key) {
                    smallBagMapping[key] = smallBagCounter;
                    smallBagCounter++;
                });

                // 設定小袋和大袋序號
                classroom.students.forEach(function (student) {
                    const classSubjectKey =
                        student[columns.class] + student[columns.subject];
                    student[columns.smallBagId] =
                        smallBagMapping[classSubjectKey];
                    student[columns.bigBagId] = bigBagCounter;
                });

                bigBagCounter++;
            }
        });
    });

    saveExamToSheet(exam);
}

/**
 * 填入節次時間
 *
 * 根據「節次時間表」工作表的對映規則，
 * 填充每個學生的節次時間欄位。
 */
function populateSessionTimes() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    const [timeHeaders, ...sessionTimeRows] =
        SESSION_TIME_REFERENCE_SHEET.getDataRange().getValues();

    const sessionTimeLookup = {};
    sessionTimeRows.forEach(function (timeRow) {
        sessionTimeLookup[timeRow[0]] = timeRow[1];
    });

    exam.sessions.forEach(function (session, sessionNumber) {
        const timeValue = sessionTimeLookup[sessionNumber] || "";
        session.classrooms.forEach(function (classroom) {
            classroom.students.forEach(function (student) {
                student[columns.time] = timeValue;
            });
        });
    });

    saveExamToSheet(exam);
}

/**
 * 計算試場人數、大小袋人數、班級人數
 *
 * 使用 Exam 物件的統計屬性計算各種人數，
 * 並更新到每個學生的對應欄位。
 */
function updateBagAndClassPopulations() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    // 計算班級人數（全考試層級）
    const classPopulationMap = {};
    exam.sessions.forEach(function (session) {
        session.students.forEach(function (student) {
            const className = student[columns.class];
            classPopulationMap[className] =
                (classPopulationMap[className] || 0) + 1;
        });
    });

    // 更新所有欄位
    exam.sessions.forEach(function (session) {
        session.classrooms.forEach(function (classroom) {
            // 小袋人數 = 班級科目組合的人數
            const classSubjectStats = classroom.classSubjectStatistics;

            // 大袋人數 = 試場人數
            const bigBagPopulation = classroom.population;

            classroom.students.forEach(function (student) {
                const classSubjectKey =
                    student[columns.class] + student[columns.subject];
                student[columns.smallBagPopulation] =
                    classSubjectStats[classSubjectKey] || 0;
                student[columns.bigBagPopulation] = bigBagPopulation;
                student[columns.classPopulation] =
                    classPopulationMap[student[columns.class]] || 0;
            });
        });
    });

    saveExamToSheet(exam);
}

/**
 * 依科目排序補考名單
 *
 * 排序優先順序：科別 > 年級 > 節次 > 試場 > 科目 > 座號
 */
function sortFilteredStudentsBySubject() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    const allStudents = [];
    exam.sessions.forEach(function (session) {
        session.classrooms.forEach(function (classroom) {
            allStudents.push(...classroom.students);
        });
    });

    allStudents.sort(function (a, b) {
        // 優先依科別
        if (a[columns.department] !== b[columns.department]) {
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        }
        // 次依年級
        if (a[columns.grade] !== b[columns.grade]) {
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        }
        // 再依節次
        if (a[columns.session] !== b[columns.session]) {
            return a[columns.session] - b[columns.session];
        }
        // 再依試場
        if (a[columns.room] !== b[columns.room]) {
            return a[columns.room] - b[columns.room];
        }
        // 再依科目
        if (a[columns.subject] !== b[columns.subject]) {
            return a[columns.subject].localeCompare(
                b[columns.subject],
                "zh-TW"
            );
        }
        // 最後依座號
        return a[columns.seatNumber] - b[columns.seatNumber];
    });

    // 直接寫回工作表
    const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
    FILTERED_RESULT_SHEET.getRange(
        2,
        1,
        allStudents.length,
        allStudents[0].length
    ).setValues(allStudents);
}

/**
 * 依班級座號排序補考名單
 *
 * 排序優先順序：科別 > 年級 > 座號 > 節次 > 科目
 */
function sortFilteredStudentsByClassSeat() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    const allStudents = [];
    exam.sessions.forEach(function (session) {
        session.classrooms.forEach(function (classroom) {
            allStudents.push(...classroom.students);
        });
    });

    allStudents.sort(function (a, b) {
        // 優先依科別
        if (a[columns.department] !== b[columns.department]) {
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        }
        // 次依年級
        if (a[columns.grade] !== b[columns.grade]) {
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        }
        // 再依座號
        if (a[columns.seatNumber] !== b[columns.seatNumber]) {
            return a[columns.seatNumber] - b[columns.seatNumber];
        }
        // 再依節次
        if (a[columns.session] !== b[columns.session]) {
            return a[columns.session] - b[columns.session];
        }
        // 最後依科目
        return a[columns.subject].localeCompare(b[columns.subject], "zh-TW");
    });

    // 直接寫回工作表
    const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
    FILTERED_RESULT_SHEET.getRange(
        2,
        1,
        allStudents.length,
        allStudents[0].length
    ).setValues(allStudents);
}

/**
 * 依節次試場排序補考名單
 *
 * 排序優先順序：節次 > 試場 > 科別 > 年級 > 座號 > 科目
 */
function sortFilteredStudentsBySessionRoom() {
    const exam = createExamFromSheet();
    const columns = getColumnIndices();

    const allStudents = [];
    exam.sessions.forEach(function (session) {
        session.classrooms.forEach(function (classroom) {
            allStudents.push(...classroom.students);
        });
    });

    allStudents.sort(function (a, b) {
        // 優先依節次
        if (a[columns.session] !== b[columns.session]) {
            return a[columns.session] - b[columns.session];
        }
        // 次依試場
        if (a[columns.room] !== b[columns.room]) {
            return a[columns.room] - b[columns.room];
        }
        // 再依科別
        if (a[columns.department] !== b[columns.department]) {
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        }
        // 再依年級
        if (a[columns.grade] !== b[columns.grade]) {
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        }
        // 再依座號
        if (a[columns.seatNumber] !== b[columns.seatNumber]) {
            return a[columns.seatNumber] - b[columns.seatNumber];
        }
        // 最後依科目
        return a[columns.subject].localeCompare(b[columns.subject], "zh-TW");
    });

    // 直接寫回工作表
    const [headerRow] = FILTERED_RESULT_SHEET.getDataRange().getValues();
    FILTERED_RESULT_SHEET.getRange(
        2,
        1,
        allStudents.length,
        allStudents[0].length
    ).setValues(allStudents);
}

function compareCountDescending(firstEntry, secondEntry) {
    if (firstEntry[1] === secondEntry[1]) {
        return 0;
    }
    return firstEntry[1] < secondEntry[1] ? 1 : -1;
}

function summarizeDepartmentGradeCounts(rowData) {
    const departmentIndex = 0;
    const gradeIndex = 1;
    const aggregateCounts = {};
    for (const row of rowData) {
        const key = row[departmentIndex] + row[gradeIndex];

        if (key in aggregateCounts) {
            aggregateCounts[key] += 1;
        } else {
            aggregateCounts[key] = 1;
        }
    }

    return aggregateCounts;
}

function summarizeDepartmentGradeSubjectCounts(rowData) {
    const departmentIndex = 0;
    const gradeIndex = 1;
    const subjectNameIndex = 7;

    const aggregateCounts = {};
    for (const row of rowData) {
        const key =
            row[departmentIndex] +
            row[gradeIndex] +
            "_" +
            row[subjectNameIndex];

        if (key in aggregateCounts) {
            aggregateCounts[key] += 1;
        } else {
            aggregateCounts[key] = 1;
        }
    }

    return aggregateCounts;
}

function fetchDepartmentGradeCounts() {
    const [, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();

    return summarizeDepartmentGradeCounts(candidateRows);
}

function fetchDepartmentGradeSubjectCounts() {
    const [, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();

    return summarizeDepartmentGradeSubjectCounts(candidateRows);
}

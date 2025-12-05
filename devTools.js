/**
 * 開發測試工具
 * 提供本地開發、偵錯、測試輔助功能
 */

/**
 * 診斷 Pipeline 共同科目節次分配問題
 *
 * 檢查以下項目：
 * 1. sessionRules 的內容和類型
 * 2. 學生資料中的節次欄位值和類型
 * 3. 共同科目是否正確匹配
 */
function diagnosePipelineScheduling() {
    // 載入資料
    const ctx = pipeline_loadData();

    const report = {
        sessionRules: {},
        sessionRulesKeyTypes: {},
        sessionRulesValueTypes: {},
        sampleStudents: [],
        commonSubjectsMatched: [],
        commonSubjectsNotMatched: [],
        sessionColumnType: null,
    };

    // 檢查 sessionRules
    Object.keys(ctx.sessionRules).forEach(function (key) {
        const value = ctx.sessionRules[key];
        report.sessionRules[key] = value;
        report.sessionRulesKeyTypes[key] = typeof key;
        report.sessionRulesValueTypes[key] = typeof value;
    });

    // 檢查前10筆學生資料
    for (let i = 0; i < Math.min(10, ctx.students.length); i++) {
        const student = ctx.students[i];
        report.sampleStudents.push({
            subject: student[ctx.columns.subject],
            session: student[ctx.columns.session],
            sessionType: typeof student[ctx.columns.session],
            room: student[ctx.columns.room],
        });
    }

    // 檢查共同科目匹配狀況
    ctx.students.forEach(function (student) {
        const subjectName = student[ctx.columns.subject];
        if (ctx.sessionRules.hasOwnProperty(subjectName)) {
            report.commonSubjectsMatched.push({
                subject: subjectName,
                ruleSession: ctx.sessionRules[subjectName],
                studentSession: student[ctx.columns.session],
                matches:
                    student[ctx.columns.session] ===
                    ctx.sessionRules[subjectName],
            });
        } else {
            if (
                !report.commonSubjectsNotMatched.some(function (s) {
                    return s.subject === subjectName;
                })
            ) {
                report.commonSubjectsNotMatched.push({
                    subject: subjectName,
                    ruleExists: false,
                });
            }
        }
    });

    Logger.log(JSON.stringify(report, null, 2));
    return report;
}

/**
 * 診斷試場分配問題
 *
 * 檢查以下項目：
 * 1. 各節次的學生數量
 * 2. 試場 0 的學生（未分配）
 */
function diagnoseRoomAssignment() {
    // 先執行共同科目和專業科目排程
    buildFilteredCandidateList();
    let ctx = pipeline_loadData();
    ctx = pipeline_scheduleCommonSubjects(ctx);

    const report = {
        beforeSpecialized: {
            total: ctx.students.length,
            bySession: {},
            session0Count: 0,
            emptySessionCount: 0,
        },
        afterSpecialized: {
            bySession: {},
            session0Count: 0,
        },
        afterRoomAssign: {
            bySession: {},
            room0Students: [],
        },
    };

    // 共同科目分配後
    ctx.students.forEach(function (student) {
        const session = student[ctx.columns.session];
        if (session === 0) {
            report.beforeSpecialized.session0Count++;
        } else if (session === "" || session === null) {
            report.beforeSpecialized.emptySessionCount++;
        } else {
            report.beforeSpecialized.bySession[session] =
                (report.beforeSpecialized.bySession[session] || 0) + 1;
        }
    });

    // 專業科目分配後
    ctx = pipeline_scheduleSpecializedSubjects(ctx);
    ctx.students.forEach(function (student) {
        const session = student[ctx.columns.session];
        if (session === 0 || session === "" || session === null) {
            report.afterSpecialized.session0Count++;
        } else {
            report.afterSpecialized.bySession[session] =
                (report.afterSpecialized.bySession[session] || 0) + 1;
        }
    });

    // 試場分配後
    ctx = pipeline_assignRooms(ctx);
    ctx.students.forEach(function (student) {
        const session = student[ctx.columns.session];
        const room = student[ctx.columns.room];
        if (room === 0) {
            report.afterRoomAssign.room0Students.push({
                class: student[ctx.columns.class],
                name: student[ctx.columns.name],
                subject: student[ctx.columns.subject],
                session: session,
                sessionType: typeof session,
            });
        } else {
            const key = "session" + session;
            report.afterRoomAssign.bySession[key] =
                (report.afterRoomAssign.bySession[key] || 0) + 1;
        }
    });

    Logger.log(JSON.stringify(report, null, 2));
    return report;
}

/**
 * 建立測試資料快照
 */
function createDataSnapshot() {
  try {
    var exam = createExamFromSheet();
    
    return {
      timestamp: new Date().toISOString(),
      exam: {
        population: exam.population,
        sessionCount: exam.sessions.filter(function(s) { return s; }).length - 1,
        statistics: {
          sessions: exam.sessionDistribution,
          departments: exam.departmentDistribution,
          grades: exam.gradeDistribution,
          subjects: exam.subjectDistribution
        },
        sessions: exam.sessions.slice(1).map(function(session, idx) {
          if (!session) return null;
          return {
            sessionNumber: idx + 1,
            population: session.population,
            classroomCount: session.classrooms.filter(function(c) { return c; }).length - 1,
            departmentGradeStats: session.departmentGradeStatistics
          };
        }).filter(function(s) { return s !== null; })
      }
    };
  } catch (error) {
    return {
      error: error.message,
      stack: error.stack
    };
  }
}

/**
 * 將資料快照輸出為 JSON
 */
function serveDataSnapshot() {
  var snapshot = createDataSnapshot();
  return ContentService
    .createTextOutput(JSON.stringify(snapshot, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Web App 進入點
 */
function doGet(e) {
  var mode = e && e.parameter && e.parameter.mode ? e.parameter.mode : 'viewer';
  
  if (mode === 'api') {
    return serveDataSnapshot();
  }
  
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('GAS 開發工具')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 測試用：建立範例資料
 */
function createTestData() {
  var testStudents = [
    ['資訊', '一年級', '資一甲', '張三', '數學'],
    ['資訊', '一年級', '資一甲', '李四', '英文'],
    ['機械', '二年級', '機二乙', '王五', '數學'],
    ['電機', '三年級', '電三丙', '趙六', '物理']
  ];
  
  var sheet = EXAM_SESSIONS_SHEET;
  var startRow = 2;
  
  for (var i = 0; i < testStudents.length; i++) {
    sheet.getRange(startRow + i, 1, 1, 5).setValues([testStudents[i]]);
  }
  
  Logger.log('測試資料建立完成');
}

/**
 * 清除所有節次與試場資料
 */
function clearSchedulingData() {
  var sheet = EXAM_SESSIONS_SHEET;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  if (lastCol > 5) {
    sheet.getRange(2, 6, lastRow - 1, lastCol - 5).clearContent();
  }
  
  Logger.log('排程資料已清除');
}

/**
 * 顯示當前 Exam 物件的除錯資訊
 */
function debugExamObject() {
  var exam = createExamFromSheet();
  
  Logger.log('=== Exam 物件除錯資訊 ===');
  Logger.log('總人數: ' + exam.population);
  Logger.log('節次分布: ' + JSON.stringify(exam.sessionDistribution));
  Logger.log('科別分布: ' + JSON.stringify(exam.departmentDistribution));
  Logger.log('年級分布: ' + JSON.stringify(exam.gradeDistribution));
  Logger.log('科目分布: ' + JSON.stringify(exam.subjectDistribution));
  
  for (var i = 1; i < exam.sessions.length; i++) {
    var session = exam.sessions[i];
    if (session && session.population > 0) {
      Logger.log('\n節次 ' + i + ':');
      Logger.log('  人數: ' + session.population);
      Logger.log('  試場數: ' + (session.classrooms.filter(function(c) { return c; }).length - 1));
      Logger.log('  科別-年級統計: ' + JSON.stringify(session.departmentGradeStatistics));
    }
  }
}

/**
 * 效能測試：測量函式執行時間
 */
function measurePerformance(functionName) {
  var startTime = new Date().getTime();
  
  try {
    var result = this[functionName]();
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    Logger.log(functionName + ' 執行時間: ' + duration + 'ms');
    return { result: result, duration: duration };
  } catch (error) {
    Logger.log(functionName + ' 執行錯誤: ' + error.message);
    throw error;
  }
}

/**
 * 比較兩個 Exam 物件的差異
 */
function compareExamSnapshots(snapshot1, snapshot2) {
  var differences = {
    populationChanged: snapshot1.exam.population !== snapshot2.exam.population,
    sessionCountChanged: snapshot1.exam.sessionCount !== snapshot2.exam.sessionCount,
    statisticsChanges: {}
  };
  
  var keys = ['sessions', 'departments', 'grades', 'subjects'];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var stats1 = JSON.stringify(snapshot1.exam.statistics[key]);
    var stats2 = JSON.stringify(snapshot2.exam.statistics[key]);
    if (stats1 !== stats2) {
      differences.statisticsChanges[key] = {
        before: snapshot1.exam.statistics[key],
        after: snapshot2.exam.statistics[key]
      };
    }
  }
  
  return differences;
}

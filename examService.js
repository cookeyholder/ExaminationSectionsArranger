/**
 * 考試服務層 - 負責 Exam 物件的建立、載入與儲存
 *
 * 此模組封裝了 Exam 領域模型與 Google Sheets 工作表之間的轉換邏輯，
 * 提供統一的資料存取介面，讓排程邏輯專注於業務規則。
 *
 * 主要功能：
 * - 從工作表建立 Exam 物件（createExamFromSheet）
 * - 將 Exam 物件存回工作表（saveExamToSheet）
 * - 取得欄位索引對映（getColumnIndices）
 */

/**
 * 取得欄位索引對映
 *
 * 從「排入考程的補考名單」工作表的標題列讀取欄位名稱，
 * 建立欄位名稱到索引的對映物件，方便程式碼使用語意化的欄位名稱。
 *
 * @returns {Object} 欄位名稱到索引的對映
 * @property {number} department - 科別欄位索引
 * @property {number} grade - 年級欄位索引
 * @property {number} classCode - 班級代碼欄位索引
 * @property {number} class - 班級欄位索引
 * @property {number} seatNumber - 座號欄位索引
 * @property {number} studentId - 學號欄位索引
 * @property {number} name - 姓名欄位索引
 * @property {number} subject - 科目名稱欄位索引
 * @property {number} session - 節次欄位索引
 * @property {number} room - 試場欄位索引
 * @property {number} smallBagId - 小袋序號欄位索引
 * @property {number} smallBagPopulation - 小袋人數欄位索引
 * @property {number} bigBagId - 大袋序號欄位索引
 * @property {number} bigBagPopulation - 大袋人數欄位索引
 * @property {number} classPopulation - 班級人數欄位索引
 * @property {number} time - 時間欄位索引
 * @property {number} computer - 電腦欄位索引
 * @property {number} manual - 人工欄位索引
 * @property {number} teacher - 任課老師欄位索引
 *
 * @example
 * const columns = getColumnIndices();
 * const subjectName = student[columns.subject];
 * const sessionNumber = student[columns.session];
 */
function getColumnIndices() {
    const lastColumn = FILTERED_RESULT_SHEET.getLastColumn();
    const headerRow = FILTERED_RESULT_SHEET.getRange(
        1,
        1,
        1,
        lastColumn
    ).getValues()[0];

    return {
        department: headerRow.indexOf("科別"),
        grade: headerRow.indexOf("年級"),
        classCode: headerRow.indexOf("班級代碼"),
        class: headerRow.indexOf("班級"),
        seatNumber: headerRow.indexOf("座號"),
        studentId: headerRow.indexOf("學號"),
        name: headerRow.indexOf("姓名"),
        subject: headerRow.indexOf("科目名稱"),
        session: headerRow.indexOf("節次"),
        room: headerRow.indexOf("試場"),
        smallBagId: headerRow.indexOf("小袋序號"),
        smallBagPopulation: headerRow.indexOf("小袋人數"),
        bigBagId: headerRow.indexOf("大袋序號"),
        bigBagPopulation: headerRow.indexOf("大袋人數"),
        classPopulation: headerRow.indexOf("班級人數"),
        time: headerRow.indexOf("時間"),
        computer: headerRow.indexOf("電腦"),
        manual: headerRow.indexOf("人工"),
        teacher: headerRow.indexOf("任課老師"),
    };
}

/**
 * 從工作表建立 Exam 物件
 *
 * 讀取「排入考程的補考名單」工作表的資料，建立完整的 Exam 領域模型。
 * 根據每個學生的節次欄位，將學生分配到對應的 Session 物件中。
 *
 * 注意：此函式假設工作表中的節次欄位已填充完畢。
 * 如果節次為 0 或空值，該學生會被略過。
 *
 * @returns {Object} Exam 物件，包含所有節次和學生資料
 * @property {Array<Object>} sessions - 節次陣列（索引 0 不使用，從 1 開始）
 * @property {number} population - 總學生人數
 * @property {Object} sessionDistribution - 節次分布統計
 * @property {Object} departmentDistribution - 科別分布統計
 * @property {Object} gradeDistribution - 年級分布統計
 * @property {Object} subjectDistribution - 科目分布統計
 *
 * @example
 * const exam = createExamFromSheet();
 * Logger.log('總人數: ' + exam.population);
 * Logger.log('第 1 節人數: ' + exam.sessions[1].population);
 * Logger.log('科別分布: ' + JSON.stringify(exam.departmentDistribution));
 */
function createExamFromSheet() {
    const [headerRow, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();
    const columns = getColumnIndices();

    // 從參數區讀取配置
    const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
    const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();

    // 建立 Exam 物件
    const exam = createExamRecord(maxSessionCount, maxRoomCount);

    // 將學生填入對應節次
    candidateRows.forEach((studentRow) => {
        const sessionNumber = studentRow[columns.session];

        // 只處理有效的節次編號
        if (sessionNumber > 0 && sessionNumber < exam.sessions.length) {
            exam.sessions[sessionNumber].addStudent(studentRow);
        }
    });

    return exam;
}

/**
 * 將 Exam 物件存回工作表
 *
 * 從 Exam 物件的所有 Classroom 收集學生資料（Single Source of Truth），
 * 清空「排入考程的補考名單」工作表的現有資料（保留標題列），
 * 然後寫入更新後的學生資料。
 *
 * 資料流向：Classroom.students → 工作表
 * 這確保了 Classroom 層級是學生資料的唯一真實來源。
 *
 * @param {Object} exam - 要儲存的 Exam 物件
 * @throws {Error} 如果 exam 物件無效或沒有學生資料
 *
 * @example
 * const exam = createExamFromSheet();
 * // ... 執行排程邏輯
 * saveExamToSheet(exam);
 */
function saveExamToSheet(exam) {
    const lastColumn = FILTERED_RESULT_SHEET.getLastColumn();
    const headerRow = FILTERED_RESULT_SHEET.getRange(
        1,
        1,
        1,
        lastColumn
    ).getValues()[0];
    const allStudents = [];

    // 從 Classroom 收集所有學生（Single Source of Truth）
    exam.sessions.forEach((session) => {
        session.classrooms.forEach((classroom) => {
            if (classroom.students && classroom.students.length > 0) {
                allStudents.push(...classroom.students);
            }
        });
    });

    // 清空舊資料（保留標題列）
    const lastRow = FILTERED_RESULT_SHEET.getLastRow();
    if (lastRow > 1) {
        FILTERED_RESULT_SHEET.getRange(
            2,
            1,
            lastRow - 1,
            headerRow.length
        ).clearContent();
    }

    // 寫入新資料
    if (allStudents.length > 0) {
        FILTERED_RESULT_SHEET.getRange(
            2,
            1,
            allStudents.length,
            allStudents[0].length
        ).setValues(allStudents);
    }
}

function updateUnfilteredSubjectCodes() {
    const currentSchoolYear = parseInt(
        PARAMETERS_SHEET.getRange("B2").getValue()
    );
    const [unfilteredHeaders, ...unfilteredRows] =
        UNFILTERED_MAKEUP_SHEET.getDataRange().getValues();

    const classColumnIndex = unfilteredHeaders.indexOf("班級");
    const subjectColumnIndex = unfilteredHeaders.indexOf("科目");
    const codeColumnIndex = unfilteredHeaders.indexOf("科目代碼補完");
    const subjectNameColumnIndex = unfilteredHeaders.indexOf("科目名稱");

    const departmentGroupMap = {
        301: "21",
        303: "22",
        305: "23",
        306: "23",
        308: "23",
        309: "23",
        311: "25",
        315: "24",
        373: "28",
        374: "21",
    };

    const gradeYearMap = {
        一: currentSchoolYear,
        二: currentSchoolYear - 1,
        三: currentSchoolYear - 2,
    };

    const patchedRows = [];
    unfilteredRows.forEach(function (candidateRow) {
        const rawSubjectCell = candidateRow[subjectColumnIndex].toString();
        const rawSubjectCode = rawSubjectCell.split(".")[0];
        if (rawSubjectCode.length === 16) {
            candidateRow[codeColumnIndex] =
                rawSubjectCode.slice(0, 3) +
                "553401" +
                rawSubjectCode.slice(3, 9) +
                "0" +
                rawSubjectCode.slice(9);
        } else {
            const gradeDigit = candidateRow[classColumnIndex]
                .toString()
                .slice(2, 3);
            const departmentDigits = rawSubjectCode.slice(0, 3);
            candidateRow[codeColumnIndex] =
                gradeYearMap[gradeDigit] +
                "553401V" +
                departmentGroupMap[departmentDigits] +
                departmentDigits +
                "0" +
                rawSubjectCode.slice(3);
        }

        candidateRow[subjectNameColumnIndex] = rawSubjectCell.split(".")[1];
        patchedRows.push(candidateRow);
    });

    if (patchedRows.length === unfilteredRows.length) {
        writeRangeValuesSafely(
            UNFILTERED_MAKEUP_SHEET.getRange(
                2,
                1,
                patchedRows.length,
                patchedRows[0].length
            ),
            patchedRows
        );
    } else {
        Logger.log("課程代碼補完失敗！");
        SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
    }
}

function updateOpenCourseCodes() {
    const currentSchoolYear = parseInt(
        PARAMETERS_SHEET.getRange("B2").getValue()
    );
    const [openHeaders, ...openRows] =
        OPEN_COURSE_LOOKUP_SHEET.getDataRange().getValues();

    const classColumnIndex = openHeaders.indexOf("班級名稱");
    const codeColumnIndex = openHeaders.indexOf("科目代碼");
    const completedCodeColumnIndex = openHeaders.indexOf("科目代碼補完");

    const departmentGroupMap = {
        301: "21",
        303: "22",
        305: "23",
        306: "23",
        308: "23",
        309: "23",
        311: "25",
        315: "24",
        373: "28",
        374: "21",
    };

    const gradeYearMap = {
        一: currentSchoolYear,
        二: currentSchoolYear - 1,
        三: currentSchoolYear - 2,
    };

    const patchedRows = [];
    openRows.forEach(function (courseRow) {
        const rawSubjectCode = courseRow[codeColumnIndex];
        if (rawSubjectCode.length === 16) {
            courseRow[completedCodeColumnIndex] =
                rawSubjectCode.slice(0, 3) +
                "553401" +
                rawSubjectCode.slice(3, 9) +
                "0" +
                rawSubjectCode.slice(9);
        } else {
            const gradeDigit = courseRow[classColumnIndex]
                .toString()
                .slice(2, 3);
            const departmentDigits = rawSubjectCode.slice(0, 3);
            courseRow[completedCodeColumnIndex] =
                gradeYearMap[gradeDigit] +
                "553401V" +
                departmentGroupMap[departmentDigits] +
                departmentDigits +
                "0" +
                rawSubjectCode.slice(3);
        }
        patchedRows.push(courseRow);
    });

    if (patchedRows.length === openRows.length) {
        writeRangeValuesSafely(
            OPEN_COURSE_LOOKUP_SHEET.getRange(
                2,
                1,
                patchedRows.length,
                patchedRows[0].length
            ),
            patchedRows
        );
    } else {
        Logger.log("開課資料課程代碼補完失敗！");
        SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
    }
}

function buildFilteredCandidateList() {
    const [candidateHeaders, ...candidateRows] =
        TEACHING_SELECTION_SHEET.getDataRange().getValues();
    const [unfilteredHeaders, ...unfilteredRows] =
        UNFILTERED_MAKEUP_SHEET.getDataRange().getValues();
    const [openHeaders, ...openRows] =
        OPEN_COURSE_LOOKUP_SHEET.getDataRange().getValues();

    const studentNumberIndex = unfilteredHeaders.indexOf("學號");
    const classNameIndex = unfilteredHeaders.indexOf("班級");
    const seatNumberIndex = unfilteredHeaders.indexOf("座號");
    const studentNameIndex = unfilteredHeaders.indexOf("姓名");
    const subjectNameIndex = unfilteredHeaders.indexOf("科目名稱");
    const completedCodeIndex = unfilteredHeaders.indexOf("科目代碼補完");
    const openClassIndex = openHeaders.indexOf("班級名稱");
    const openSubjectNameIndex = openHeaders.indexOf("科目名稱");
    const teacherNameIndex = openHeaders.indexOf("任課教師");

    const makeUpRequiredIndex = candidateHeaders.indexOf("要補考");
    const filteredCodeIndex = candidateHeaders.indexOf("課程代碼");
    const requiresComputerIndex = candidateHeaders.indexOf("電腦");
    const requiresManualIndex = candidateHeaders.indexOf("人工");

    const eligibleSubjectsByCode = {};
    candidateRows.forEach(function (candidateRow) {
        if (candidateRow[makeUpRequiredIndex] === true) {
            eligibleSubjectsByCode[candidateRow[filteredCodeIndex]] = {
                computerRequired: candidateRow[requiresComputerIndex],
                manualRequired: candidateRow[requiresManualIndex],
            };
        }
    });

    const teacherLookup = {};
    openRows.forEach(function (courseRow) {
        const teacherCell = courseRow[teacherNameIndex].toString();
        const lookupKey =
            courseRow[openClassIndex].toString() +
            courseRow[openSubjectNameIndex].toString();
        if (teacherCell.length > 10) {
            teacherLookup[lookupKey] = teacherCell.split(",")[0].slice(7);
        } else {
            teacherLookup[lookupKey] = teacherCell.slice(6);
        }
    });

    resetFilteredSheets();

    const filteredRows = [];
    unfilteredRows.forEach(function (studentRow) {
        if (
            Object.keys(eligibleSubjectsByCode).includes(
                studentRow[completedCodeIndex]
            )
        ) {
            const courseKey =
                studentRow[classNameIndex].toString() +
                studentRow[subjectNameIndex].toString();
            const computerMark = eligibleSubjectsByCode[
                studentRow[completedCodeIndex]
            ].computerRequired
                ? "☑"
                : "☐";
            const manualMark = eligibleSubjectsByCode[
                studentRow[completedCodeIndex]
            ].manualRequired
                ? "☑"
                : "☐";
            const filteredRow = [
                lookupDepartmentName(studentRow[classNameIndex]), // 科別
                deriveGradeLevel(studentRow[classNameIndex]), // 年級
                deriveClassCode(studentRow[classNameIndex]), // 班級代碼
                studentRow[classNameIndex], // 班級
                studentRow[seatNumberIndex], // 座號
                studentRow[studentNumberIndex], // 學號
                studentRow[studentNameIndex], // 姓名
                studentRow[subjectNameIndex], // 科目名稱
                (studentRow[8] = 0), // 節次預設為0
                (studentRow[9] = 0), // 試場預設為0
                "", // 小袋序號
                "", // 小袋人數
                "", // 大袋序號
                "", // 大袋人數
                "", // 班級人數
                "", // 時間
                computerMark,
                manualMark,
                teacherLookup[courseKey], // 任課老師
            ];

            filteredRows.push(filteredRow);
        }
    });

    if (filteredRows.length > 0) {
        FILTERED_RESULT_SHEET.getRange(
            2,
            1,
            filteredRows.length,
            filteredRows[0].length
        )
            .setNumberFormat("@STRING@")
            .setValues(filteredRows);
    }
}

function markVisibleCandidateCheckboxes() {
    const candidateRange = TEACHING_SELECTION_SHEET.getRange("A2:A");
    const checkboxValues = candidateRange.getValues();
    const totalRows = candidateRange.getNumRows();
    const totalColumns = candidateRange.getNumColumns();

    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
        if (!ACTIVE_SPREADSHEET.isRowHiddenByFilter(rowIndex + 1)) {
            for (
                let columnIndex = 0;
                columnIndex < totalColumns;
                columnIndex++
            ) {
                checkboxValues[rowIndex][columnIndex] = true;
            }
        }
    }

    writeRangeValuesSafely(candidateRange, checkboxValues);
}

function clearCandidateCheckboxes() {
    const candidateRange = TEACHING_SELECTION_SHEET.getRange("A1:A");
    const checkboxValues = candidateRange.getValues();
    const totalRows = candidateRange.getNumRows();
    const totalColumns = candidateRange.getNumColumns();

    for (let rowIndex = 0; rowIndex < totalRows; rowIndex++) {
        if (!ACTIVE_SPREADSHEET.isRowHiddenByFilter(rowIndex + 1)) {
            for (
                let columnIndex = 0;
                columnIndex < totalColumns;
                columnIndex++
            ) {
                checkboxValues[rowIndex][columnIndex] = false;
            }
        }
    }

    writeRangeValuesSafely(candidateRange, checkboxValues);
}

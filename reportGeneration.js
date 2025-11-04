function createExamBulletinSheet() {
    sortFilteredStudentsByClassSeat();

    const [headerRow, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();
    const classIndex = headerRow.indexOf("班級");
    const studentNumberIndex = headerRow.indexOf("學號");
    const studentNameIndex = headerRow.indexOf("姓名");
    const subjectIndex = headerRow.indexOf("科目名稱");
    const sessionIndex = headerRow.indexOf("節次");
    const roomIndex = headerRow.indexOf("試場");

    BULLETIN_OUTPUT_SHEET.clear();
    if (BULLETIN_OUTPUT_SHEET.getMaxRows() > 5) {
        BULLETIN_OUTPUT_SHEET.deleteRows(
            2,
            BULLETIN_OUTPUT_SHEET.getMaxRows() - 5
        );
    }

    const bulletinRows = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
    candidateRows.forEach(function (examineeRow) {
        let maskedName = "";
        if (examineeRow[studentNameIndex].length === 2) {
            maskedName =
                examineeRow[studentNameIndex].toString().slice(0, 1) + "〇";
        } else {
            const middleMaskLength = examineeRow[studentNameIndex].length - 2;
            maskedName =
                examineeRow[studentNameIndex].toString().slice(0, 1) +
                "〇".repeat(middleMaskLength) +
                examineeRow[studentNameIndex].toString().slice(-1);
        }

        const bulletinRow = [
            examineeRow[classIndex],
            examineeRow[studentNumberIndex],
            maskedName,
            examineeRow[subjectIndex],
            examineeRow[sessionIndex],
            examineeRow[roomIndex],
        ];
        bulletinRows.push(bulletinRow);
    });

    writeRangeValuesSafely(
        BULLETIN_OUTPUT_SHEET.getRange(
            2,
            1,
            bulletinRows.length,
            bulletinRows[0].length
        ),
        bulletinRows
    );
    sortFilteredStudentsBySessionRoom();
    formatBulletinSheet();
}

function formatBulletinSheet() {
    const schoolYearValue = PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue = PARAMETERS_SHEET.getRange("B3").getValue();

    BULLETIN_OUTPUT_SHEET.getRange("A1:F1").mergeAcross();
    BULLETIN_OUTPUT_SHEET.getRange("A1").setValue(
        "高雄高工" +
            schoolYearValue +
            "學年度第" +
            semesterValue +
            "學期補考名單"
    );
    BULLETIN_OUTPUT_SHEET.getRange("A1").setFontSize(20);
    BULLETIN_OUTPUT_SHEET.getRange(
        1,
        1,
        BULLETIN_OUTPUT_SHEET.getMaxRows(),
        BULLETIN_OUTPUT_SHEET.getMaxColumns()
    ).setHorizontalAlignment("center");
    BULLETIN_OUTPUT_SHEET.setFrozenRows(2);
    BULLETIN_OUTPUT_SHEET.getRange("A2:F").createFilter();
    BULLETIN_OUTPUT_SHEET.getRange("A2:F").setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID
    );
}

function createProctorRecordSheet() {
    sortFilteredStudentsBySessionRoom();

    const [headerRow, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();
    const schoolYearValue = PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue = PARAMETERS_SHEET.getRange("B3").getValue();
    const sessionIndex = headerRow.indexOf("節次");
    const roomIndex = headerRow.indexOf("試場");
    const timeIndex = headerRow.indexOf("時間");
    const classIndex = headerRow.indexOf("班級");
    const studentNumberIndex = headerRow.indexOf("學號");
    const studentNameIndex = headerRow.indexOf("姓名");
    const subjectIndex = headerRow.indexOf("科目名稱");
    const classPopulationIndex = headerRow.indexOf("班級人數");

    RECORD_OUTPUT_SHEET.clear();
    if (RECORD_OUTPUT_SHEET.getMaxRows() > 5) {
        RECORD_OUTPUT_SHEET.deleteRows(2, RECORD_OUTPUT_SHEET.getMaxRows() - 5);
    }

    const recordRows = [
        [
            "A表：" +
                schoolYearValue +
                "學年度第" +
                semesterValue +
                "學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "節次",
            "試場",
            "時間",
            "班級",
            "學號",
            "姓名",
            "科目名稱",
            "班級人數",
            "考生到考簽名",
            "違規記錄(打V)",
            "",
            "其他違規\n請簡述",
        ],
        ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""],
    ];

    candidateRows.forEach(function (examineeRow) {
        recordRows.push([
            examineeRow[sessionIndex],
            examineeRow[roomIndex],
            examineeRow[timeIndex],
            examineeRow[classIndex],
            examineeRow[studentNumberIndex],
            examineeRow[studentNameIndex],
            examineeRow[subjectIndex],
            examineeRow[classPopulationIndex],
            "",
            "",
            "",
            "",
        ]);
    });

    writeRangeValuesSafely(
        RECORD_OUTPUT_SHEET.getRange(
            1,
            1,
            recordRows.length,
            recordRows[0].length
        ),
        recordRows
    );

    RECORD_OUTPUT_SHEET.getRange("A1:L1")
        .mergeAcross()
        .setVerticalAlignment("bottom")
        .setFontSize(14)
        .setFontWeight("bold");
    RECORD_OUTPUT_SHEET.getRange("J2:K2").mergeAcross();
    RECORD_OUTPUT_SHEET.getRange("A2:A3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("B2:B3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("C2:C3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("D2:D3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("E2:E3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("F2:F3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("G2:G3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("H2:H3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("I2:I3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("L2:L3")
        .mergeVertically()
        .setVerticalAlignment("middle");

    RECORD_OUTPUT_SHEET.getRange(
        2,
        1,
        recordRows.length + 2,
        recordRows[0].length
    )
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBorder(
            true,
            true,
            true,
            true,
            true,
            true,
            "#000000",
            SpreadsheetApp.BorderStyle.SOLID
        );
}

function composeSmallBagDataset() {
    const schoolYearValue = PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue = PARAMETERS_SHEET.getRange("B3").getValue();
    const [headerRow, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();

    const smallBagIndex = headerRow.indexOf("小袋序號");
    const sessionIndex = headerRow.indexOf("節次");
    const timeIndex = headerRow.indexOf("時間");
    const roomIndex = headerRow.indexOf("試場");
    const classIndex = headerRow.indexOf("班級");
    const subjectIndex = headerRow.indexOf("科目名稱");
    const teacherIndex = headerRow.indexOf("任課老師");
    const smallBagPopulationIndex = headerRow.indexOf("小袋人數");
    const computerIndex = headerRow.indexOf("電腦");
    const manualIndex = headerRow.indexOf("人工");

    SMALL_BAG_DATA_SHEET.clear();
    if (SMALL_BAG_DATA_SHEET.getMaxRows() > 5) {
        SMALL_BAG_DATA_SHEET.deleteRows(
            2,
            SMALL_BAG_DATA_SHEET.getMaxRows() - 5
        );
    }

    const smallBagRows = [
        [
            "學年度",
            "學期",
            "小袋序號",
            "節次",
            "時間",
            "試場",
            "班級",
            "科目名稱",
            "任課老師",
            "小袋人數",
            "電腦",
            "人工",
        ],
    ];
    const processedSmallBags = [];

    candidateRows.forEach(function (examineeRow) {
        if (!processedSmallBags.includes(examineeRow[smallBagIndex])) {
            const datasetRow = [
                schoolYearValue,
                semesterValue,
                examineeRow[smallBagIndex],
                examineeRow[sessionIndex],
                examineeRow[timeIndex],
                examineeRow[roomIndex],
                examineeRow[classIndex],
                examineeRow[subjectIndex],
                examineeRow[teacherIndex],
                examineeRow[smallBagPopulationIndex],
                examineeRow[computerIndex],
                examineeRow[manualIndex],
            ];

            smallBagRows.push(datasetRow);
            processedSmallBags.push(examineeRow[smallBagIndex]);
        }
    });

    writeRangeValuesSafely(
        SMALL_BAG_DATA_SHEET.getRange(
            1,
            1,
            smallBagRows.length,
            smallBagRows[0].length
        ),
        smallBagRows
    );
}

function composeBigBagDataset() {
    const schoolYearValue = PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue = PARAMETERS_SHEET.getRange("B3").getValue();
    const makeUpDateValue = PARAMETERS_SHEET.getRange("B13").getValue();
    const [headerRow, ...candidateRows] =
        FILTERED_RESULT_SHEET.getDataRange().getValues();
    const invigilatorAssignment = transposeMatrix(
        INVIGILATOR_ASSIGNMENT_SHEET.getDataRange().getValues()
    );

    const bigBagIndex = headerRow.indexOf("大袋序號");
    const smallBagIndex = headerRow.indexOf("小袋序號");
    const sessionIndex = headerRow.indexOf("節次");
    const timeIndex = headerRow.indexOf("時間");
    const roomIndex = headerRow.indexOf("試場");
    const bigBagPopulationIndex = headerRow.indexOf("大袋人數");

    BIG_BAG_DATA_SHEET.clear();
    if (BIG_BAG_DATA_SHEET.getMaxRows() > 5) {
        BIG_BAG_DATA_SHEET.deleteRows(2, BIG_BAG_DATA_SHEET.getMaxRows() - 5);
    }

    const bigBagRows = [
        [
            "學年度",
            "學期",
            "大袋序號",
            "節次",
            "試場",
            "補考日期",
            "時間",
            "試卷袋序號",
            "監考教師",
            "各試場人數",
        ],
    ];
    const processedBigBags = [];

    const smallBagRangeByBigBag = {};
    candidateRows.forEach(function (examineeRow) {
        const bigBagKey = "大袋" + examineeRow[bigBagIndex];
        if (!Object.keys(smallBagRangeByBigBag).includes(bigBagKey)) {
            smallBagRangeByBigBag[bigBagKey] = [examineeRow[smallBagIndex]];
        } else {
            smallBagRangeByBigBag[bigBagKey].push(examineeRow[smallBagIndex]);
        }
    });

    candidateRows.forEach(function (examineeRow) {
        if (!processedBigBags.includes(examineeRow[bigBagIndex])) {
            const sessionNum = parseInt(examineeRow[sessionIndex]);
            const roomNum = parseInt(examineeRow[roomIndex]);

            // 如果節次或試場為非數字，無法定位該試場，跳過該筆
            if (isNaN(sessionNum) || isNaN(roomNum)) {
                Logger.log(
                    `無效的節次/試場索引: 節次=${sessionNum}, 試場=${roomNum}`
                );
                return;
            }

            const bagRange =
                smallBagRangeByBigBag["大袋" + examineeRow[bigBagIndex]];

            // 若監考老師資料表缺少該節次或該節次下沒有對應試場，回退為空字串而非跳過整列
            var invigilatorName = "";
            if (invigilatorAssignment && invigilatorAssignment[sessionNum]) {
                invigilatorName =
                    invigilatorAssignment[sessionNum][roomNum] || "";
            }

            const datasetRow = [
                schoolYearValue,
                semesterValue,
                examineeRow[bigBagIndex],
                examineeRow[sessionIndex],
                examineeRow[roomIndex],
                makeUpDateValue,
                examineeRow[timeIndex],
                Math.min(...bagRange) + "-" + Math.max(...bagRange),
                invigilatorName,
                examineeRow[bigBagPopulationIndex],
            ];

            bigBagRows.push(datasetRow);
            processedBigBags.push(examineeRow[bigBagIndex]);
        }
    });

    writeRangeValuesSafely(
        BIG_BAG_DATA_SHEET.getRange(
            1,
            1,
            bigBagRows.length,
            bigBagRows[0].length
        ),
        bigBagRows
    );
}

function transposeMatrix(matrix) {
    if (!matrix || matrix.length === 0) return [];
    return matrix[0].map((_, colIndex) => matrix.map((row) => row[colIndex]));
}

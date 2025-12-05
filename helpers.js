function writeRangeValuesSafely(targetRange, valueMatrix) {
    if (!valueMatrix || valueMatrix.length === 0) {
        Logger.log("writeRangeValuesSafely: 無資料可寫入");
        return;
    }
    if (targetRange.getLastColumn() === valueMatrix[0].length) {
        targetRange.setValues(valueMatrix);
    } else {
        SpreadsheetApp.getUi().alert("欲寫入範圍欄數不足！");
    }
}

function calculateElapsedSeconds(startTimestamp) {
    const endTimestamp = new Date();
    const elapsedMillis = Number(endTimestamp) - Number(startTimestamp);
    return Math.ceil(elapsedMillis / 1000);
}

function measureCallbackSeconds(callback) {
    const startTimestamp = new Date();
    callback();
    const endTimestamp = new Date();
    const elapsedMillis = Number(endTimestamp) - Number(startTimestamp);
    return Math.ceil(elapsedMillis / 1000);
}

function deriveClassCode(className) {
    // 班級代碼查詢
    // 輸入班級中文名稱(4字)，輸出6碼數字代碼
    // 如輸入「機械二丁」，輸出「301204」。

    const departmentCodes = {
        機械: "301",
        汽車: "303",
        資訊: "305",
        電子: "306",
        電機: "308",
        冷凍: "309",
        建築: "311",
        化工: "315",
        圖傳: "373",
        電圖: "374",
    };

    const classAndGradeCode = {
        甲: "01",
        乙: "02",
        丙: "03",
        丁: "04",
        一: "1",
        二: "2",
        三: "3",
    };

    return (
        departmentCodes[className.slice(0, 2)] +
        classAndGradeCode[className.slice(2, 3)] +
        classAndGradeCode[className.slice(-1)]
    );
}

function deriveGradeLevel(className) {
    const gradeMapping = {
        一: "1",
        二: "2",
        三: "3",
    };

    return gradeMapping[className.slice(2, 3)];
}

function lookupDepartmentName(className) {
    const departmentNames = {
        機械: "機械科",
        汽車: "汽車科",
        資訊: "資訊科",
        電子: "電子科",
        電機: "電機科",
        冷凍: "冷凍空調科",
        建築: "建築科",
        化工: "化工科",
        圖傳: "圖文傳播科",
        電圖: "電腦機械製圖科",
    };

    return departmentNames[className.slice(0, 2)];
}

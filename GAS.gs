/**
 * 崇正國樂團智能系統 - GAS 後端終極分區版
 * 整合功能：一問卷一分頁、動態表頭、活動長度計算、EmbedURL 支援
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
// --- API 進入點 ---
function doPost(e) {
    let result;
    try {
        const requestData = JSON.parse(e.postData.contents);
        const action = requestData.action;
        const data = requestData.data;
        checkAndInitSheets();
        switch (action) {
            case 'loginUser': result = loginUser(data); break;
            case 'registerUser': result = registerUser(data); break;
            case 'saveMember': result = saveMember(data); break;
            case 'getAnnouncement': result = getAnnouncement(); break;
            case 'saveAnnouncement': result = saveAnnouncement(data); break;
            case 'getSystemStatus': result = getSystemStatus(data); break;
            case 'submitCheckin': result = submitCheckin(data); break;
            case 'getAttendanceHistory': result = getAttendanceHistory(data); break;
            case 'getEventList': result = getEventList(); break;
            case 'saveEvent': result = saveEvent(data); break;
            case 'activateEvent': result = activateEvent(data); break;
            case 'getRealtimeStatus': result = getRealtimeStatus(data); break;
            case 'getSurveyList': result = getSurveyList(); break;
            case 'saveSurveyConfig': result = saveSurveyConfig(data); break;
            case 'getSurveyStatus': result = getSurveyStatus(data); break;
            case 'submitSurveyResponse': result = submitSurveyResponse(data); break;
            case 'getSurveyResults': result = getSurveyResults(data); break;
            case 'saveSystemSetting': result = saveSystemSetting(data); break;
            case 'deactivateEvent': result = deactivateEvent(); break;
            case 'updateDeviceToken': result = updateDeviceToken(data); break;
            case 'testPush': result = testPush(data); break;
            case 'getStaffList': result = getStaffList(); break;
            case 'getUpcomingEvents': result = getUpcomingEvents(); break;
            case 'submitLeave': result = submitLeave(data); break;
            default: result = { success: false, message: "未知指令" };
        }
    } catch (err) {
        result = { success: false, message: "後端錯誤: " + err.toString() };
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// A. 活動管理函數 (Event Management)
// ==========================================

function getEventList() {
    try {
        const sheet = SS.getSheetByName("Events");
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
        const activeEventId = SS.getSheetByName("SystemConfig").getRange(2, 9).getValue(); // 抓取目前廣播 ID

        const list = data.map(r => ({
            ID: r[0], Name: r[1], Type: r[2], Date: Utilities.formatDate(r[3], "GMT+8", "yyyy-MM-dd"),
            StartTime: formatTimeValue(r[4]), EndTime: formatTimeValue(r[5]),
            RestStart: formatTimeValue(r[8]), RestEnd: formatTimeValue(r[9])
        }));
        return { success: true, data: list.reverse(), activeEventId: activeEventId };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getUpcomingEvents() {
    try {
        const sheet = SS.getSheetByName("Events");
        if (sheet.getLastRow() < 2) return { success: true, data: [] };
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
        const todayStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd");
        
        const list = [];
        data.forEach(r => {
            if (!r[3]) return;
            // 統一格式化成 yyyy-MM-dd 再比較字串，避免時區或物件問題
            const eventDateStr = r[3] instanceof Date ? Utilities.formatDate(r[3], "GMT+8", "yyyy-MM-dd") : String(r[3]);
            if (eventDateStr >= todayStr) {
                list.push({
                    ID: r[0], Name: r[1], Type: r[2], Date: eventDateStr,
                    StartTime: formatTimeValue(r[4]), EndTime: formatTimeValue(r[5])
                });
            }
        });
        return { success: true, data: list };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function deactivateEvent() {
    try {
        const sheet = SS.getSheetByName("SystemConfig");
        sheet.getRange(2, 1).setValue("off");  // 狀態設為關閉 (Cell A2)
        sheet.getRange(2, 9).setValue("");     // 清除活動 ID (Cell I2)
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function saveEvent(data) {
    try {
        const sheet = SS.getSheetByName("Events");
        let totalMin = calculateDiffMinutes(data.startTime, data.endTime);
        if (data.restStart && data.restEnd) {
            totalMin = Math.max(0, totalMin - calculateDiffMinutes(data.restStart, data.restEnd));
        }
        const durStr = Math.floor(totalMin / 60) + "時" + (totalMin % 60) + "分";
        const id = data.id || "EVT-" + Date.now();

        const rowData = [id, data.name, data.type || "團練", data.date, data.startTime, data.endTime, durStr, new Date(), data.restStart || "", data.restEnd || ""];
        const vals = sheet.getDataRange().getValues();
        let rowIdx = -1;
        for (let i = 1; i < vals.length; i++) {
            if (String(vals[i][0]) === String(id)) { rowIdx = i + 1; break; }
        }
        if (rowIdx !== -1) sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
        else sheet.appendRow(rowData);
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function activateEvent(data) {
    try {
        const sheet = SS.getSheetByName("Events");
        const ev = sheet.getDataRange().getValues().find(r => String(r[0]) === String(data.id));
        if (!ev) return { success: false, message: "找不到該活動資料" };

        // 更新系統配置：[狀態A, 活動名B, 類型C, 開始D, 結束E, _F, _G, 日期H, ID(I2), 休始J, 休止K]
        const configData = ["on", ev[1], ev[2], ev[4], ev[5], "", "", ev[3], ev[0], ev[8], ev[9]];
        SS.getSheetByName("SystemConfig").getRange(2, 1, 1, 11).setValues([configData]);
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// B.  改進後的問卷管理 (一問卷一分頁)
// ==========================================

function getSurveyList() {
    try {
        const sheet = SS.getSheetByName("SurveyTemplates");
        const vals = sheet.getDataRange().getValues();
        const activeId = SS.getSheetByName("SystemConfig").getRange(2, 12).getValue();
        const list = vals.slice(1).map(r => ({
            ID: r[0], Title: r[1], Description: r[2], QuestionsJSON: r[3], Active: String(r[0]) === String(activeId), EmbedURL: r[5] || ""
        }));
        return { success: true, data: list.reverse() };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function saveSurveyConfig(data) {
    try {
        const sheet = SS.getSheetByName("SurveyTemplates");
        const id = data.id || "SUR-" + Date.now();
        // 欄位：[ID, Title, Description, QuestionsJSON, CreateDate, EmbedURL]
        const rowData = [id, data.title, data.description, data.questionsJSON, new Date(), data.embedURL || ""];
        const vals = sheet.getDataRange().getValues();
        let rowIdx = -1;
        for (let i = 1; i < vals.length; i++) {
            if (String(vals[i][0]) === String(id)) { rowIdx = i + 1; break; }
        }
        if (rowIdx !== -1) sheet.getRange(rowIdx, 1, 1, 6).setValues([rowData]);
        else sheet.appendRow(rowData);

        // --- 修正邏輯開始 ---
        const configSheet = SS.getSheetByName("SystemConfig");
        if (data.active) {
            // 勾選開啟：設定為目前啟動的問卷 ID
            configSheet.getRange(2, 12).setValue(id);
        } else {
            // 取消勾選：檢查如果目前「正在啟動」的就是這份問卷，才將其清空（關閉）
            const currentActiveId = configSheet.getRange(2, 12).getValue();
            if (String(currentActiveId) === String(id)) {
                configSheet.getRange(2, 12).setValue("");
            }
        }
        // --- 修正邏輯結束 ---
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getSurveyStatus(data) {
    try {
        const sysS = SS.getSheetByName("SystemConfig");
        const activeId = sysS.getRange(2, 12).getValue();
        if (!activeId) return { success: true, active: false };
        const templateSheet = SS.getSheetByName("SurveyTemplates");
        if (!templateSheet) return { success: true, active: false };

        // 找問卷模板
        const template = templateSheet.getDataRange().getValues().find(r => String(r[0]) === String(activeId));
        if (!template) return { success: true, active: false };
        // 預設尚未回覆
        let responded = false;
        let lastResponse = {};

        // v51: 優先從專屬分頁 SRV_ 讀取回覆狀態與歷史
        const surveySheet = SS.getSheetByName("SRV_" + activeId);
        if (surveySheet && surveySheet.getLastRow() > 1 && data && data.email) {
            const resps = surveySheet.getDataRange().getValues();
            const headers = resps[0];
            const email = String(data.email).toLowerCase().trim();

            // 尋找該 Email 的最後一筆 (反轉搜尋)
            const userRow = [...resps].reverse().find(r => r[1] && String(r[1]).toLowerCase().trim() === email);

            if (userRow) {
                responded = true;
                // 從各欄位重建 lastResponse JSON（從第 5 欄開始，跳過組別）
                lastResponse = {};
                for (let i = 4; i < headers.length; i++) {
                    lastResponse[headers[i]] = userRow[i];
                }
            }
        }
        return {
            success: true,
            active: true,
            title: template[1],
            description: template[2],
            questions: JSON.parse(template[3] || "[]"),
            responded: responded,
            lastResponse: lastResponse,
            embedURL: template[5] || ""
        };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

function submitSurveyResponse(data) {
    try {
        const activeSurId = SS.getSheetByName("SystemConfig").getRange(2, 12).getValue();
        if (!activeSurId) return { success: false, message: "無啟動中問卷" };

        // 1. 取得題目定義以建立表頭
        const templates = SS.getSheetByName("SurveyTemplates").getDataRange().getValues();
        const templateRow = templates.find(r => String(r[0]) === String(activeSurId));
        if (!templateRow) return { success: false, message: "找不到問卷定義" };
        const questions = JSON.parse(templateRow[3] || "[]");

        // 2. 獲取或建立專屬分頁
        const surveySheet = getOrCreateSurveySheet(activeSurId, templateRow[1], questions);

        // 3. 準備資料列
        const userEmail = data.email.toLowerCase();
        const userName = data.name || "";
        const userSection = data.section || "";
        const responses = JSON.parse(data.responsesJSON || "{}");
        const headers = surveySheet.getRange(1, 1, 1, surveySheet.getLastColumn()).getValues()[0];
        const rowData = new Array(headers.length).fill("");

        rowData[0] = new Date(); // 時間
        rowData[1] = userEmail;
        rowData[2] = userName;
        rowData[3] = userSection; // 組別

        // 根據題目標籤填入欄位（從第 5 欄開始）
        for (let i = 4; i < headers.length; i++) {
            const label = headers[i];
            if (responses.hasOwnProperty(label)) rowData[i] = responses[label];
        }

        // 4. 重複填寫檢查 (Email 為主)
        const sheetVals = surveySheet.getDataRange().getValues();
        let rowIdx = -1;
        for (let i = 1; i < sheetVals.length; i++) {
            if (String(sheetVals[i][1]).toLowerCase() === userEmail) { rowIdx = i + 1; break; }
        }

        if (rowIdx !== -1) {
            surveySheet.getRange(rowIdx, 1, 1, rowData.length).setNumberFormat('@').setValues([rowData.map(v => v instanceof Date ? Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd HH:mm:ss") : String(v))]);
        } else {
            const newRow = surveySheet.getLastRow() + 1;
            const range = surveySheet.getRange(newRow, 1, 1, rowData.length);
            range.setNumberFormat('@'); // 強制純文字，保留前導 0
            range.setValues([rowData.map(v => v instanceof Date ? Utilities.formatDate(v, "GMT+8", "yyyy-MM-dd HH:mm:ss") : String(v))]);
        }

        // 備份到總表 (選填)
        const master = SS.getSheetByName("SurveyResponses");
        if (master) master.appendRow([new Date(), userEmail, userName, activeSurId, data.responsesJSON]);
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}
function getOrCreateSurveySheet(surveyId, title, questions) {
    const sheetName = "SRV_" + surveyId;
    let sheet = SS.getSheetByName(sheetName);
    if (!sheet) sheet = SS.insertSheet(sheetName);

    // 動態產生表頭：時間 | Email | 姓名 | 組別 | [所有題目標籤...]
    const expectedHeaders = ["填寫時間", "Email", "姓名", "組別"].concat(questions.map(q => q.label));
    const currentHeaders = sheet.getLastColumn() > 0 ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];

    if (JSON.stringify(currentHeaders) !== JSON.stringify(expectedHeaders)) {
        sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders])
            .setFontWeight("bold").setBackground("#f3f4f6").setHorizontalAlignment("center");
        sheet.setFrozenRows(1);
        for (let i = 1; i <= expectedHeaders.length; i++) sheet.autoResizeColumn(i);
    }
    return sheet;
}

function getSurveyResults(data) {
    try {
        const surveyId = data.id;
        const sheet = SS.getSheetByName("SRV_" + surveyId);
        if (!sheet || sheet.getLastRow() <= 1) return { success: true, data: [], rawData: [] };

        const vals = sheet.getDataRange().getValues();
        const headers = vals[0];
        let rows = vals.slice(1);

        // 如果是 Leader 且有指定組別，則過濾
        if (data.role === 'Leader' && data.section) {
            rows = rows.filter(r => String(r[3] || "").trim() === data.section);
        }

        // 1. 統計各題結果 (跳過 填寫時間, Email, 姓名, 組別)
        const results = [];
        for (let col = 4; col < headers.length; col++) {
            const label = headers[col];
            const stats = {};
            rows.forEach(r => {
                const val = String(r[col] || "").trim();
                if (val) {
                    // 如果是多選 (逗號分隔)，拆開計次
                    if (val.includes(',')) {
                        val.split(',').forEach(v => {
                            const subV = v.trim();
                            if (subV) stats[subV] = (stats[subV] || 0) + 1;
                        });
                    } else {
                        stats[val] = (stats[val] || 0) + 1;
                    }
                }
            });
            results.push({ label: label, stats: stats });
        }

        // 2. 獲取原始回覆表格 (用於回覆管理清單)
        const rawData = rows.map(r => {
            let timeStr = "";
            try {
                if (r[0] instanceof Date) {
                    timeStr = Utilities.formatDate(r[0], "GMT+8", "MM/dd HH:mm");
                } else {
                    timeStr = String(r[0] || "");
                }
            } catch (e) { timeStr = String(r[0] || ""); }

            const item = { time: timeStr, email: r[1], name: r[2], section: r[3] || "", answers: {} };
            for (let i = 4; i < headers.length; i++) {
                item.answers[headers[i]] = r[i];
            }
            return item;
        });

        // 3. 若為 Leader，回傳該組全員名單（含填寫狀態）
        let memberList = [];
        if (data.role === 'Leader' && data.section) {
            const memS = SS.getSheetByName("Members");
            const allMembers = memS.getDataRange().getValues().slice(1);
            const sectionMembers = allMembers.filter(m => String(m[2] || "").trim() === data.section);
            // 已填寫的 email 清單
            const filledEmails = rows.map(r => String(r[1] || "").toLowerCase().trim());
            memberList = sectionMembers.map(m => ({
                name: m[1],
                instrument: m[6] || "",
                filled: filledEmails.includes(String(m[3] || "").toLowerCase().trim())
            }));
            // 已填寫排後面，未填寫排前面
            memberList.sort((a, b) => a.filled === b.filled ? 0 : a.filled ? 1 : -1);
        }

        return { success: true, data: results, rawData: rawData.reverse(), memberList: memberList };
    } catch (e) { return { success: false, message: e.toString() }; }
}

// ==========================================
// C. 監控與基礎功能 (Monitoring & Basic)
// ==========================================

function getRealtimeStatus(userData) {
    try {
        const sysS = SS.getSheetByName("SystemConfig");
        if (sysS.getLastColumn() < 13) sysS.insertColumnsAfter(sysS.getLastColumn(), 13 - sysS.getLastColumn());
        const config = sysS.getRange(2, 1, 1, 13).getValues()[0];

        const isSystemOn = config[0] === "on";
        const currentEventName = String(config[1] || "");
        const currentEventDate = config[7] instanceof Date ? Utilities.formatDate(config[7], "GMT+8", "yyyy-MM-dd") : String(config[7] || "");
        const currentEventEndTime = typeof formatTimeValue === 'function' ? formatTimeValue(config[4]) : String(config[4] || "");
        const memS = SS.getSheetByName("Members");
        if (memS.getLastColumn() < 13) memS.insertColumnsAfter(memS.getLastColumn(), 13 - memS.getLastColumn());
        const members = memS.getDataRange().getValues();
        const now = new Date();
        const attSheet = SS.getSheetByName("Attendance");
        let attendance = [];
        if (attSheet && attSheet.getLastRow() > 1) {
            attendance = attSheet.getDataRange().getValues();
        }
        const uniqueEvents = {};
        const userStats = {};
        for (let i = 1; i < attendance.length; i++) {
            // 精準位置：Email(3), 活動名稱(4), Date(5)
            const email = String(attendance[i][3] || "").toLowerCase().trim();
            const attEvent = String(attendance[i][4] || ""); // 真正的活動名稱
            const attDateRaw = attendance[i][5];
            const attDate = attDateRaw instanceof Date ? Utilities.formatDate(attDateRaw, "GMT+8", "yyyy-MM-dd") : String(attDateRaw || "");

            // 紀錄活動總場次 (利用 日期+活動名 當作唯一值)
            if (attDate && attEvent) uniqueEvents[attDate + "_" + attEvent] = true;
            if (!userStats[email]) userStats[email] = { presentCount: 0, leaveCount: 0, today: false, todayStatus: "尚未簽到" };
            
            const attStatus = String(attendance[i][6] || ""); // 判斷是出席還是請假
            if (attStatus === "請假") {
                userStats[email].leaveCount++;
                if (attDate === currentEventDate && attEvent === currentEventName) {
                    userStats[email].today = true;
                    userStats[email].todayStatus = "請假";
                }
            } else {
                userStats[email].presentCount++;
                if (attDate === currentEventDate && attEvent === currentEventName) {
                    userStats[email].today = true;
                    userStats[email].todayStatus = "已簽到";
                }
            }
        }

        // 計算樂團總共辦了幾場活動
        const totalEventsCount = Object.keys(uniqueEvents).length;
        
        // 過濾掉「行政組」，不讓他們出現在出缺席監控中
        const filteredMembers = members.slice(1).filter(m => String(m[2] || "").trim() !== "行政組");
        
        const data = filteredMembers.map(m => {
            const email = String(m[3] || "").toLowerCase().trim();
            const stats = userStats[email] || { presentCount: 0, leaveCount: 0, today: false, todayStatus: "尚未簽到" };

            let statusText = "尚未簽到";
            if (stats.today) {
                statusText = stats.todayStatus;
            } else if (isSystemOn) {
                let endH = 0, endM = 0;
                if (currentEventEndTime.includes('時')) {
                    const parts = currentEventEndTime.split('時');
                    endH = parseInt(parts[0]); endM = parseInt(parts[1]);
                } else if (currentEventEndTime.includes(':')) {
                    const parts = currentEventEndTime.split(':');
                    endH = parseInt(parts[0]); endM = parseInt(parts[1]);
                }
                const endDateTime = new Date(currentEventDate);
                endDateTime.setHours(endH, endM, 0);
                // 如果系統開啟中但已經打鐘了
                if (now > endDateTime) statusText = "缺席";
            } else {
                // 系統已經關閉，但這個人沒簽到
                statusText = "缺席";
            }
            // 【真實缺席數】 = 樂團總共辦了幾場活動 - 出席總次數 - 請假總次數
            let absenceCount = totalEventsCount - stats.presentCount - stats.leaveCount;
            if (absenceCount < 0) absenceCount = 0;
            return {
                Name: m[1], Section: m[2], Instrument: m[6], Email: email,
                status: statusText,
                absenceCount: absenceCount,
                leaveCount: stats.leaveCount,
                Phone: m[9] || "", Birthday: m[10] || "", ID_Number: m[11] || "", PrivacyConsent: m[12] || "NO"
            };
        });
        const stats = { present: data.filter(d => d.status === "已簽到").length, absent: data.filter(d => d.status !== "已簽到").length };
        return { success: true, data, stats, totalEvents: totalEventsCount };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getStaffList() {
    try {
        const memS = SS.getSheetByName("Members");
        if (memS.getLastColumn() < 13) memS.insertColumnsAfter(memS.getLastColumn(), 13 - memS.getLastColumn());
        const members = memS.getDataRange().getValues();
        
        // 只抓取行政組
        const staffs = members.slice(1).filter(m => String(m[2] || "").trim() === "行政組").map(m => ({
            Name: m[1],
            Section: m[2],
            Instrument: m[6], // 職能
            Email: String(m[3] || "").toLowerCase().trim(),
            Phone: m[9] || "",
            Birthday: m[10] || "",
            ID_Number: m[11] || "",
            PrivacyConsent: m[12] || "NO"
        }));
        
        return { success: true, data: staffs };
    } catch (e) { 
        return { success: false, message: e.toString() }; 
    }
}

// --- 核心：自動補齊欄位工具 ---
function ensureSheetWidth(sheet, requiredColumns) {
    if (sheet.getLastColumn() < requiredColumns) {
        sheet.insertColumnsAfter(sheet.getLastColumn(), requiredColumns - sheet.getLastColumn());
    }
}

// 取得系統狀態（包含授權碼）
// 3. 系統狀態：修正讀取範圍並加回授權碼 
function getSystemStatus(data) {
    try {
        const sysS = SS.getSheetByName("SystemConfig");
        if (sysS.getLastColumn() < 13) sysS.insertColumnsAfter(sysS.getLastColumn(), 13 - sysS.getLastColumn());
        const config = sysS.getRange(2, 1, 1, 13).getValues()[0];

        const isSystemOn = config[0] === "on";
        const eventName = String(config[1] || "");
        const type = String(config[2] || "");
        const startTimeStr = typeof formatTimeValue === 'function' ? formatTimeValue(config[3]) : String(config[3] || "");
        const endTimeStr = typeof formatTimeValue === 'function' ? formatTimeValue(config[4]) : String(config[4] || "");
        const eventDate = config[7] instanceof Date ? Utilities.formatDate(config[7], "GMT+8", "yyyy-MM-dd") : String(config[7] || "");
        const eventId = String(config[8] || "");
        const bypassCode = String(config[12] || "");
        let hasCheckedIn = false;
        let hasLeaved = false;
        if (data && data.email && isSystemOn) {
            const attSheet = SS.getSheetByName("Attendance");
            if (attSheet && attSheet.getLastRow() > 1) {
                const attData = attSheet.getDataRange().getValues();
                const email = String(data.email).toLowerCase().trim();
                // 精準位置：Email(索引3), 活動名稱(索引4), Date(索引5), 類別(索引6)
                for (let i = attData.length - 1; i >= 1; i--) {
                    const attEmail = String(attData[i][3] || "").toLowerCase().trim();
                    const attEvent = String(attData[i][4] || ""); // 正確的 EventName 在這裡！
                    const attDateRaw = attData[i][5];
                    const attDate = attDateRaw instanceof Date ? Utilities.formatDate(attDateRaw, "GMT+8", "yyyy-MM-dd") : String(attDateRaw || "");
                    const attType = String(attData[i][6] || "");

                    if (attEmail === email && attDate === eventDate && attEvent === eventName) {
                        hasCheckedIn = true;
                        if (attType === "請假") {
                            hasLeaved = true;
                        }
                        break;
                    }
                }
            }
        }
        return {
            success: true, enabled: isSystemOn, eventName: eventName, type: type,
            eventStart: startTimeStr, eventEnd: endTimeStr, eventDate: eventDate,
            id: eventId, bypassCode: bypassCode, hasCheckedIn: hasCheckedIn, hasLeaved: hasLeaved
        };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function saveSystemSetting(data) {
    try {
        const sysS = SS.getSheetByName("SystemConfig");
        const key = String(data.key);
        const value = String(data.value);

        if (key === 'bypassCode') {
            // 將授權碼保存在 M 欄 (第 13 欄)
            sysS.getRange(2, 13).setValue(value);
            return { success: true };
        }
        return { success: false, message: "無效的設定鍵值" };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// 1. 登入：修正為更穩定的抓取方式
function loginUser(data) {
    try {
        const sheet = SS.getSheetByName("Members");
        const emails = sheet.getRange(1, 4, sheet.getLastRow(), 1).getValues().flat();
        const email = String(data.email).toLowerCase().trim();
        const rowIndex = emails.indexOf(email);
        if (rowIndex === -1) return { success: false, message: "帳號不存在" };

        // 確保寬度至少 13 欄
        ensureSheetWidth(sheet, 13);
        const row = sheet.getRange(rowIndex + 1, 1, 1, 13).getValues()[0];
        if (String(row[4]) !== String(data.password).trim()) return { success: false, message: "密碼錯誤" };

        return {
            success: true,
            userData: {
                ID: row[0], Name: row[1], Section: row[2], Email: row[3], Role: row[5], Instrument: row[6],
                Phone: (function(v){ var s=String(v||""); return (s&&/^\d+$/.test(s)&&!s.startsWith("0"))?"0"+s:s; })(row[9]),
                Birthday: row[10] || "", ID_Number: row[11] || "", PrivacyConsent: row[12] || "NO"
            }
        };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function registerUser(data) {
    const sheet = SS.getSheetByName("Members");
    const email = String(data.email).trim().toLowerCase();
    if (sheet.getDataRange().getValues().some(r => String(r[3]).trim().toLowerCase() === email)) return { success: false, message: "帳號已註冊過" };
    sheet.appendRow(["MEM-" + Date.now(), data.name.trim(), data.section, email, String(data.password).trim(), "Member", data.instrument, new Date(), "系統註冊"]);
    return { success: true };
}

function submitCheckin(data) {
    try {
        const config = SS.getSheetByName("SystemConfig").getRange(2, 1, 1, 11).getValues()[0];
        const eventStart = formatTimeValue(config[3]), eventEnd = formatTimeValue(config[4]);
        const restStart = formatTimeValue(config[9]), restEnd = formatTimeValue(config[10]);
        let totalMinutes = calculateDiffMinutes(eventStart, eventEnd);
        if (restStart && restEnd) totalMinutes -= calculateDiffMinutes(restStart, restEnd);
        const finalDuration = Math.max(0, totalMinutes);
        const durStr = Math.floor(finalDuration / 60) + "時" + (finalDuration % 60) + "分";
        SS.getSheetByName("Attendance").appendRow([new Date(), data.name, data.section, String(data.email).toLowerCase(), data.eventName, data.date, "團練", `${eventStart}~${eventEnd}`, durStr]);
        return { success: true };
    } catch (err) { return { success: false, message: err.toString() }; }
}

function submitLeave(data) {
    try {
        SS.getSheetByName("Attendance").appendRow([
            new Date(), 
            data.name, 
            data.section, 
            String(data.email).toLowerCase(), 
            data.eventName, 
            data.date, 
            "請假", 
            "系統請假", 
            "請假",
            data.reason || ""
        ]);
        return { success: true };
    } catch (err) { 
        return { success: false, message: err.toString() }; 
    }
}

function getAttendanceHistory(data) {
    try {
        const attSheet = SS.getSheetByName("Attendance");
        if (!attSheet || attSheet.getLastRow() <= 1) return { success: true, list: [] };

        const attData = attSheet.getDataRange().getValues();
        const email = String(data.email).toLowerCase().trim();
        const list = [];

        // 精準位置：Email(3), 活動名稱(4), Date(5), 類別(6), Duration即時數(8)
        for (let i = attData.length - 1; i >= 1; i--) {
            const rowEmail = String(attData[i][3] || "").toLowerCase().trim();

            if (rowEmail === email) {
                const eventName = String(attData[i][4] || "一般團練"); // 抓取真正的活動名
                const dateRaw = attData[i][5];
                const dateStr = dateRaw instanceof Date ? Utilities.formatDate(dateRaw, "GMT+8", "yyyy-MM-dd") : String(dateRaw || "");
                const eventType = String(attData[i][6] || "");

                let duration = String(attData[i][8] || "");
                if (eventType === "請假") {
                    duration = "請假";
                } else if (!duration) {
                    duration = "2時30分";
                }
                list.push({
                    status: eventName,
                    date: dateStr,
                    duration: duration
                });
            }
        }
        return { success: true, list: list };
    } catch (e) { return { success: false, message: e.toString() }; }
}

function getAnnouncement() {
    try {
        const data = SS.getSheetByName("Announcements").getRange(2, 1, 1, 2).getValues()[0];
        const time = data[1] instanceof Date ? Utilities.formatDate(data[1], "GMT+8", "yyyy-MM-dd HH:mm") : "";
        return { success: true, content: data[0] || "目前暫無公告。", updateTime: time };
    } catch (e) { return { success: false }; }
}

function saveAnnouncement(data) {
    try {
        SS.getSheetByName("Announcements").getRange(2, 1, 1, 2).setValues([[data.content, new Date()]]);
        // 自動發送推播通知
        sendBroadcastNotification("國樂團公告", data.content);
        return { success: true };
    } catch (e) { return { success: false }; }
}

// ==========================================
// E. FCM 推播核心 (FCM Messaging)
// ==========================================

// 改從「指令碼屬性」讀取金鑰，不把秘密寫在程式碼裡
const SCRIPT_PROP = PropertiesService.getScriptProperties();

const SERVICE_ACCOUNT = {
  "project_id": SCRIPT_PROP.getProperty("FIREBASE_PROJECT_ID") || "baiyang-co",
  "client_email": SCRIPT_PROP.getProperty("FIREBASE_CLIENT_EMAIL"),
  // v43: 重要！自動修補私鑰中的換行符號，防止「引數無效：key」錯誤
  "private_key": (SCRIPT_PROP.getProperty("FIREBASE_PRIVATE_KEY") || "").replace(/\\n/g, '\n')
};



function updateDeviceToken(data) {
  try {
    const sheet = SS.getSheetByName("Members");
    const emails = sheet.getRange(1, 4, sheet.getLastRow(), 1).getValues().flat();
    const email = String(data.email).toLowerCase().trim();
    const rowIndex = emails.indexOf(email);
    if (rowIndex === -1) return { success: false, message: "帳號未找到" };
    
    // 確保存儲 Token 的欄位存在 (假設存在第 14 欄，N 欄)
    ensureSheetWidth(sheet, 14);
    sheet.getRange(rowIndex + 1, 14).setValue(data.token);
    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function sendBroadcastNotification(title, body) {
  try {
    const sheet = SS.getSheetByName("Members");
    const data = sheet.getDataRange().getValues();
    const tokens = [];
    // 收集所有有 Token 的團員 (第 14 欄)
    for (let i = 1; i < data.length; i++) {
        const token = data[i][13]; // 索引 13 是第 14 欄 (N 欄)
        if (token && token.length > 10) tokens.push(token);
    }
    
    if (tokens.length === 0) return;
    
    tokens.forEach(tk => {
      sendFCM(tk, title, body);
    });
  } catch (e) { console.error("Broadcast failed: " + e.toString()); }
}

function sendFCM(targetToken, title, body) {
  const url = `https://fcm.googleapis.com/v1/projects/${SERVICE_ACCOUNT.project_id}/messages:send`;
  const jwtToken = getFCMAuthToken();
  const payload = {
    message: {
      token: targetToken,
      notification: { title: title, body: body },
      data: {
        title: title,
        body: body,
        click_action: "./index.html",
        version: "v41"
      },
      webpush: {
        headers: {
          Urgency: "high"
        },
        notification: {
          title: title,
          body: body,
          icon: "https://fe314343.github.io/0301/icon-192.png",
          badge: "https://fe314343.github.io/0301/icon-192.png",
          requireInteraction: true
        },
        fcm_options: {
          link: "https://fe314343.github.io/0301/index.html"
        }
      }
    }
  };
  
  const options = {
    method: "POST",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + jwtToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const res = UrlFetchApp.fetch(url, options);
  const responseText = res.getContentText();
  Logger.log("FCM Response: " + responseText);
  return responseText;
}

// v42: 診斷用測試推播
function testPush(data) {
  try {
    const sheet = SS.getSheetByName("Members");
    const emails = sheet.getRange(1, 4, sheet.getLastRow(), 1).getValues().flat();
    const email = String(data.email).toLowerCase().trim();
    const rowIndex = emails.indexOf(email);
    if (rowIndex === -1) return { success: false, message: "找不到該帳號" };
    
    const token = sheet.getRange(rowIndex + 1, 14).getValue();
    if (!token) return { success: false, message: "該帳號尚未儲存推播凭證(Token)" };
    
    const response = sendFCM(token, "🔔 系統診斷測試", "如果您看到這則訊息，代表背景通道連線正常！(v42)");
    return { success: true, rawResponse: response };
  } catch (e) {
    return { success: false, message: "診斷失敗: " + e.toString() };
  }
}

function getFCMAuthToken() {
  const header = JSON.stringify({ alg: "RS256", typ: "JWT" });
  const now = Math.floor(Date.now() / 1000);
  const claimSet = JSON.stringify({
    iss: SERVICE_ACCOUNT.client_email,
    scope: "https://www.googleapis.com/auth/cloud-platform",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now
  });
  
  const toSign = Utilities.base64EncodeWebSafe(header) + "." + Utilities.base64EncodeWebSafe(claimSet);
  const signature = Utilities.computeRsaSha256Signature(toSign, SERVICE_ACCOUNT.private_key);
  const jwt = toSign + "." + Utilities.base64EncodeWebSafe(signature);
  
  const res = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });
  
  return JSON.parse(res.getContentText()).access_token;
}

// 2. 儲存：修正為分段儲存，更安全
function saveMember(data) {
    try {
        const sheet = SS.getSheetByName("Members");
        const emails = sheet.getRange(1, 4, sheet.getLastRow(), 1).getValues().flat();
        const email = String(data.email).toLowerCase().trim();
        const rowIndex = emails.indexOf(email);
        if (rowIndex === -1) return { success: false, message: "帳號未找到" };

        ensureSheetWidth(sheet, 13);
        const rowNum = rowIndex + 1;
        if (data.name) sheet.getRange(rowNum, 2).setValue(data.name);
        if (data.instrument) sheet.getRange(rowNum, 7).setValue(data.instrument);
        if (data.password && data.password.trim() !== "") sheet.getRange(rowNum, 5).setValue(data.password.trim());

        if (data.phone !== undefined) {
            let p = String(data.phone).trim();
            // 如果是全數字且沒有 0 開頭，強制補 0
            if (p && /^\d+$/.test(p) && !p.startsWith("0")) p = "0" + p;
            const phoneCell = sheet.getRange(rowNum, 10);
            phoneCell.setNumberFormat('@');  // 強制純文字
            phoneCell.setValue(p); // 寫入字串
        }
        if (data.birthday !== undefined) sheet.getRange(rowNum, 11).setValue(data.birthday);
        if (data.idNumber !== undefined) sheet.getRange(rowNum, 12).setValue(data.idNumber);
        if (data.privacyConsent !== undefined) {
            const consentVal = data.privacyConsent === "YES" ? "YES (" + Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm") + ")" : data.privacyConsent;
            sheet.getRange(rowNum, 13).setValue(consentVal);
        }
        return { success: true };
    } catch (e) { return { success: false, message: e.toString() }; }
}


/*function getSystemStatus(userData) {
  try {
    const config = SS.getSheetByName("SystemConfig").getRange(2, 1, 1, 12).getValues()[0];
    let ed = config[7]; if (ed instanceof Date) ed = Utilities.formatDate(ed, "GMT+8", "yyyy-MM-dd");
    const status = { success: true, enabled: config[0] === "on", eventName: config[1], eventDate: ed, eventStart: formatTimeValue(config[3]), eventEnd: formatTimeValue(config[4]), hasCheckedIn: false };
    if (status.enabled && userData?.email) {
      const att = SS.getSheetByName("Attendance").getDataRange().getValues();
      const email = userData.email.toLowerCase();
      status.hasCheckedIn = att.some(r => {
        let rd = r[5]; if (rd instanceof Date) rd = Utilities.formatDate(rd, "GMT+8", "yyyy-MM-dd");
        return r[3].toLowerCase() === email && String(rd) === String(ed) && String(r[4]) === String(status.eventName);
      });
    }
    return status;
  } catch(e) { return { success: false }; }
}
*/
// ==========================================
// D. 輔助與初始化工具 (Helpers)
// ==========================================

function calculateDiffMinutes(start, end) {
    if (!start || !end) return 0;
    try {
        const [sH, sM] = String(start).split(':').map(Number);
        const [eH, eM] = String(end).split(':').map(Number);
        let diff = (eH * 60 + eM) - (sH * 60 + sM);
        return diff < 0 ? diff + 1440 : diff;
    } catch (e) { return 0; }
}

function formatTimeValue(v) {
    if (v instanceof Date) return Utilities.formatDate(v, "GMT+8", "HH:mm");
    if (typeof v === 'string' && v.includes(':')) return v.substring(0, 5);
    return v ? String(v) : "";
}

function checkAndInitSheets() {
    const sheets = ["Members", "Attendance", "Events", "SystemConfig", "Announcements", "SurveyTemplates", "SurveyResponses"];
    sheets.forEach(n => { if (!SS.getSheetByName(n)) SS.insertSheet(n); });
    const sysS = SS.getSheetByName("SystemConfig");
    if (sysS.getLastRow() === 0) {
        sysS.appendRow(["狀態", "活動名", "類型", "開始", "結束", "_", "_", "日期", "ID", "休始", "休止", "SurveyID"]);
        sysS.appendRow(["off", "", "", "", "", "", "", "", "", "", "", ""]);
    }
}

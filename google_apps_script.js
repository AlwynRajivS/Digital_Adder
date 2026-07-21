/**
 * Google Apps Script Web App Backend for Subtractor Lab
 * 
 * Instructions:
 * 1. Open your Google Sheet.
 * 2. Click Extensions -> Apps Script.
 * 3. Delete any default code and paste this script in.
 * 4. Click Deploy -> New Deployment.
 * 5. Select "Web App" as the type.
 * 6. Set "Execute as" to: "Me" (your email).
 * 7. Set "Who has access" to: "Anyone".
 * 8. Click Deploy, authorize permissions, and copy the Web App URL.
 */

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    // Auto-setup database sheet and headers
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("StudentRegistry");
    if (!sheet) {
      sheet = ss.insertSheet("StudentRegistry");
      const headers = [
        "Roll Number", 
        "Student Name", 
        "Class", 
        "Password", 
        "Completed Modules", 
        "Completion %", 
        "Clicks", 
        "Correct", 
        "Accuracy %", 
        "Last Active", 
        "Arcade Score"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#cfe2f3");
    }

    if (action === "ping") {
      return jsonResponse({ success: true, sheetName: ss.getName() });
    }

    if (action === "pullAll") {
      const data = sheet.getDataRange().getValues();
      const students = [];
      // Row 0 is headers
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[0]) continue; // Skip blank rows
        students.push({
          rollNo: String(row[0]),
          name: String(row[1]),
          classCode: String(row[2]),
          password: String(row[3]),
          completedModules: String(row[4] || ""),
          progress: Number(row[5] || 0),
          clicks: Number(row[6] || 0),
          correct: Number(row[7] || 0),
          accuracy: Number(row[8] || 100),
          lastActive: String(row[9] || ""),
          arcadeScore: Number(row[10] || 0)
        });
      }
      return jsonResponse({ success: true, students: students });
    }

    if (action === "sync") {
      const rollNo = String(payload.rollNo);
      const name = String(payload.name);
      const classCode = String(payload.classCode);
      const completedModules = String(payload.completedModules || "");
      const progress = Number(payload.progress || 0);
      const clicks = Number(payload.clicks || 0);
      const correct = Number(payload.correct || 0);
      const accuracy = Number(payload.accuracy || 100);
      const arcadeScore = Number(payload.arcadeScore || 0);
      const timestamp = new Date().toISOString().replace('T', ' ').substring(0, 16);

      const data = sheet.getDataRange().getValues();
      let studentRowIndex = -1;

      // Search if roll number already exists
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === rollNo) {
          studentRowIndex = i + 1; // 1-indexed row number
          break;
        }
      }

      if (studentRowIndex !== -1) {
        // Update existing student row metrics (keeping their original password)
        sheet.getRange(studentRowIndex, 2).setValue(name);
        sheet.getRange(studentRowIndex, 3).setValue(classCode);
        sheet.getRange(studentRowIndex, 5).setValue(completedModules);
        sheet.getRange(studentRowIndex, 6).setValue(progress);
        sheet.getRange(studentRowIndex, 7).setValue(clicks);
        sheet.getRange(studentRowIndex, 8).setValue(correct);
        sheet.getRange(studentRowIndex, 9).setValue(accuracy);
        sheet.getRange(studentRowIndex, 10).setValue(timestamp);
        
        // Only update arcade score if it exceeds their previous record
        const oldScoreVal = Number(sheet.getRange(studentRowIndex, 11).getValue() || 0);
        if (arcadeScore > oldScoreVal) {
          sheet.getRange(studentRowIndex, 11).setValue(arcadeScore);
        }
      } else {
        // Auto-register new student row (assign default password same as classCode)
        const defaultPassword = classCode.toLowerCase();
        const newRow = [
          rollNo, 
          name, 
          classCode, 
          defaultPassword, 
          completedModules, 
          progress, 
          clicks, 
          correct, 
          accuracy, 
          timestamp, 
          arcadeScore
        ];
        sheet.appendRow(newRow);
      }

      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: false, message: "Unknown action: " + action });

  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

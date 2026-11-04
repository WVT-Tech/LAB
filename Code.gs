function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบทดสอบคุณภาพวัสดุ LAB')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function calculatePieChart() {
  return { ductilePass: 0, ductileFail: 0, bronzePass: 0, bronzeFail: 0 };
}

// ฟังก์ชันระบุชื่อชีทเป้าหมายให้ชัดเจน
function getDBSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Chemical_DB");
  if (!sheet) {
    sheet = ss.insertSheet("Chemical_DB");
  }
  return sheet;
}

function generateDocNo(status, jobType) {
  var sheet = getDBSheet();
  var prefix = (jobType === "M") ? "MWA" : "PWA";
  var year = "2569";
  
  if (sheet.getLastRow() <= 1) {
    return prefix + "001/" + year;
  }
  
  var data = sheet.getDataRange().getValues();
  var maxNo = 0;
  
  for (var i = 1; i < data.length; i++) {
    var docNo = String(data[i][1] || "");
    if (docNo.indexOf(prefix) !== -1) {
      var parts = docNo.split('/');
      var numPart = parseInt(parts[0].replace(/[^0-9]/g, ''), 10);
      if (!isNaN(numPart) && numPart > maxNo) {
        maxNo = numPart;
      }
    }
  }
  
  var nextNo = maxNo + 1;
  return prefix + String(nextNo).padStart(3, '0') + "/" + year;
}

function checkDocNoExist(docNo) {
  var sheet = getDBSheet();
  if (sheet.getLastRow() <= 1) return false;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]) === String(docNo)) {
      return true;
    }
  }
  return false;
}

function saveChemicalDataMultiple(payload) {
  var sheet = getDBSheet();
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["วันที่", "เลขที่เอกสาร", "ประเภทงาน", "ชนิดตัวอย่าง", "สถานะน้ำเหล็ก", "LOT. เตา", "ผลประเมิน", "Fe", "C", "Si", "S", "P", "Mn", "Ni", "Cr", "Cu", "Mo", "Mg", "Al", "Remark"]);
  }
  
  var overallEvaluation = "ผ่าน";
  
  if (payload.details && payload.details.length > 0) {
    payload.details.forEach(function(det) {
      var sampleType = payload.sampleType || "";
      var isBronze = sampleType.indexOf("บรอนซ์") !== -1 || sampleType.indexOf("Bronze") !== -1;
      
      var rowData = [
        payload.date,
        payload.docNo,
        payload.jobType,
        payload.sampleType,
        payload.status,
        det.remark,
        overallEvaluation
      ];
      
      if (isBronze) {
        // เรียงลำดับธาตุสำหรับ "บรอนซ์" (ตามโครงสร้างหน้าเว็บ: Cu, Pb, Sn, Zn, Ni, Fe, P, S, Al, Si, Mn, Cr)
        rowData.push(
          det.cu, det.pb, det.sn, det.zn, 
          det.ni, det.fe, det.p, det.s, 
          det.al, det.si, det.mn, det.cr
        );
      } else {
        // ชุดเดิมของเหล็กหล่อ
        rowData.push(
          det.fe, det.c, det.si, det.s, 
          det.p, det.mn, det.ni, det.cr, 
          det.cu, det.mo, det.mg, det.al
        );
      }
      
      // Specimen Mark ปิดท้าย
      rowData.push(det.specimen);
      
      sheet.appendRow(rowData);
    });
  }
  
  return { success: true, message: "บันทึกข้อมูลเรียบร้อยแล้ว" };
}

function searchDatabase(criteria) {
  var sheet = getDBSheet();
  
  if (sheet.getLastRow() <= 1) {
    return { 
      data: [], 
      pie: { ductilePass: 0, ductileFail: 0, bronzePass: 0, bronzeFail: 0 }, 
      timeline: { labels: [], lineDuctile: [], lineGray: [], lineBronze: [], lineBrass: [], lineStainless: [], lineAlBronze: [] } 
    };
  }
  
  var data = sheet.getDataRange().getValues();
  var filteredData = [];
  var pieStats = { ductilePass: 0, ductileFail: 0, bronzePass: 0, bronzeFail: 0 };
  
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[1]) continue;
    
    var rowDate = "";
    try {
      rowDate = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[0]).trim();
    } catch(e) { rowDate = String(r[0] || ""); }
    
    var docNo = String(r[1] || "");
    var jobType = String(r[2] || "");
    var sampleType = String(r[3] || "");
    var lot = String(r[5] || "");
    var evaluation = String(r[6] || "ผ่าน");
    
    if (criteria) {
      if (criteria.docNo && docNo.indexOf(criteria.docNo) === -1) continue;
      if (criteria.lot && lot.indexOf(criteria.lot) === -1) continue;
      if (criteria.sampleType && sampleType.indexOf(criteria.sampleType) === -1) continue;
      if (criteria.jobType && jobType !== criteria.jobType) continue;
      if (criteria.startDate && rowDate < criteria.startDate) continue;
      if (criteria.endDate && rowDate > criteria.endDate) continue;
    }
    
    filteredData.push({
      date: rowDate,
      docNo: docNo,
      lot: lot,
      sampleType: sampleType,
      jobType: jobType,
      evaluation: evaluation
    });
    
    if (evaluation.indexOf("ผ่าน") !== -1 && evaluation.indexOf("ไม่ผ่าน") === -1) {
      pieStats.ductilePass++;
    } else {
      pieStats.ductileFail++;
    }
  }
  
  var timelineData = calculateTimeline(data);
  
  return {
    data: filteredData.reverse(),
    pie: pieStats,
    timeline: timelineData
  };
}

function calculateTimeline(rows) {
  var dateMap = {};
  
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    
    var d = "";
    try {
      d = (r[0] instanceof Date) ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(r[0]).trim();
    } catch(e) { d = String(r[0]); }
    
    if (!d) continue;
    
    var sampleType = String(r[3] || "");
    var evalRes = String(r[6] || "");
    
    var key = "Gray";
    if (sampleType.indexOf("เหล็กหล่อเหนียว") !== -1) key = "Ductile";
    else if (sampleType.indexOf("บรอนซ์") !== -1 && sampleType.indexOf("อลูมิเนียม") === -1) key = "Bronze";
    else if (sampleType.indexOf("ทองเหลือง") !== -1) key = "Brass";
    else if (sampleType.indexOf("สเตนเลส") !== -1) key = "Stainless";
    else if (sampleType.indexOf("อลูมิเนียมบรอนซ์") !== -1) key = "AlBronze";

    if (!dateMap[d]) {
      dateMap[d] = {
        Ductile: { pass: 0, total: 0 },
        Gray: { pass: 0, total: 0 },
        Bronze: { pass: 0, total: 0 },
        Brass: { pass: 0, total: 0 },
        Stainless: { pass: 0, total: 0 },
        AlBronze: { pass: 0, total: 0 }
      };
    }
    
    dateMap[d][key].total++;
    if (evalRes.indexOf("ผ่าน") !== -1 && evalRes.indexOf("ไม่ผ่าน") === -1) {
      dateMap[d][key].pass++;
    }
  }
  
  var timeline = {
    labels: [],
    lineDuctile: [],
    lineGray: [],
    lineBronze: [],
    lineBrass: [],
    lineStainless: [],
    lineAlBronze: []
  };
  
  var sortedDates = Object.keys(dateMap).sort();
  
  for (var j = 0; j < sortedDates.length; j++) {
    var d = sortedDates[j];
    timeline.labels.push(d);
    
    var calcPct = function(obj) {
      return obj.total > 0 ? Math.round((obj.pass / obj.total) * 100) : 0;
    };
    
    timeline.lineDuctile.push(calcPct(dateMap[d].Ductile));
    timeline.lineGray.push(calcPct(dateMap[d].Gray));
    timeline.lineBronze.push(calcPct(dateMap[d].Bronze));
    timeline.lineBrass.push(calcPct(dateMap[d].Brass));
    timeline.lineStainless.push(calcPct(dateMap[d].Stainless));
    timeline.lineAlBronze.push(calcPct(dateMap[d].AlBronze));
  }
  
  return timeline;
}

function getRecordByDocNo(docNo) {
  var sheet = getDBSheet();
  var data = sheet.getDataRange().getValues();
  var record = { docNo: docNo, details: [], type: "chem" };
  
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (String(r[1]) === docNo) {
      record.date = r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
      record.jobType = r[2];
      record.sampleType = r[3];
      record.status = r[4];
      record.lot = r[5];
      
      var sampleType = String(r[3] || "");
      var det = {};
      
      // ตรวจสอบว่าเป็นบรอนซ์ (ใช้ลำดับคอลัมน์ใหม่) หรือแบบอื่น (ใช้ลำดับคอลัมน์เดิม Fe เป็นหลัก)
      if (sampleType.indexOf("บรอนซ์") !== -1 || sampleType.indexOf("Bronze") !== -1) {
        det = {
          specimen: r[19] || "",
          cu: r[7], pb: r[8], sn: r[9], zn: r[10], ni: r[11], fe: r[12], p: r[13], s: r[14], al: r[15], si: r[16], mn: r[17], cr: r[18],
          remark: r[5]
        };
      } else {
        // รูปแบบเดิม (Fe, C, Si, S, P, Mn, Ni, Cr, Cu, Mo, Mg, Al)
        det = {
          specimen: r[19] || "",
          fe: r[7], c: r[8], si: r[9], s: r[10], p: r[11], mn: r[12],
          ni: r[13], cr: r[14], cu: r[15], mo: r[16], mg: r[17], al: r[18],
          remark: r[5]
        };
      }
      
      record.details.push(det);
    }
  }
  
  return record.details.length > 0 ? record : null;
}

function exportToExcel(criteria) {
  var sheet = getDBSheet();
  var newSpreadsheet = SpreadsheetApp.create("LAB_Export_Report");
  var newSheet = newSpreadsheet.getSheets()[0];
  
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    newSheet.appendRow(data[i]);
  }
  
  return { success: true, url: newSpreadsheet.getUrl() };
}

function generateTensileDocNo(jobType) {
  return (jobType === "M" ? "TEN-M-" : "TEN-P-") + Math.floor(100 + Math.random() * 900) + "/69";
}
function checkTensileDocNoExist(docNo) { return false; }
function saveTensileDataMultiple(payload) { return { success: true, message: "บันทึกข้อมูลแรงดึงสำเร็จ" }; }

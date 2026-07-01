// HẰNG SỐ CẤU HÌNH HỆ THỐNG CỐ ĐỊNH
const MAIN_SPREADSHEET_ID = "1xE1-IB-Mj3L9lsazgXtB_LyNuO-5qWXNpShKqc9Z_qc";
const FILE_IN_CHUNG_ID = "1bUC-WkDBmn1rruCHIWWGG7xolgRkkkfeGxfdpQ7EF-I";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('KẾ HOẠCH DẠY HỌC THEO TUẦN')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Lấy nhanh ID bảng tính giáo viên
function getTargetSsId(nam, truong, gv) {
  var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
  var sheetGiaoVien = ss.getSheetByName("giaovien");
  if (!sheetGiaoVien) return null;
  
  var dataGV = sheetGiaoVien.getDataRange().getValues();
  for (var i = 1; i < dataGV.length; i++) {
    if (dataGV[i][0].toString().trim() === nam && 
        dataGV[i][1].toString().trim() === truong && 
        dataGV[i][2].toString().trim() === gv) {
      return dataGV[i][4] ? dataGV[i][4].toString().trim() : null;
    }
  }
  return null;
}

function getListDataDataFromSheet() { 
  return getListDataFromSheet();
}

// Tối ưu hóa đọc danh mục bằng cách giới hạn vùng Range thay vì đọc toàn bộ Range thừa
function getListDataFromSheet() {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    var mapHeThong = {}; 
    var listNamHoc = [];

    if (sheetGiaoVien) {
      var lastRowGV = sheetGiaoVien.getLastRow();
      if (lastRowGV > 1) {
        var rawDataGV = sheetGiaoVien.getRange(2, 1, lastRowGV - 1, 3).getValues(); // Chỉ lấy 3 cột đầu để dựng bản đồ nhân sự nhanh hơn
        rawDataGV.forEach(function(row) {
          var nam = row[0].toString().trim();
          var truong = row[1].toString().trim();
          var giaovien = row[2].toString().trim();
          if (nam !== "" && truong !== "" && giaovien !== "") {
            if (!mapHeThong[nam]) mapHeThong[nam] = {};
            if (!mapHeThong[nam][truong]) mapHeThong[nam][truong] = [];
            mapHeThong[nam][truong].push({ ten: giaovien });
          }
        });
        listNamHoc = Object.keys(mapHeThong);
      }
    }
    
    var sheetMonHoc = ss.getSheetByName("monhoc");
    var listMonHoc = [];
    var listLopHoc = [];
    if (sheetMonHoc) {
      var lastRowMH = sheetMonHoc.getLastRow();
      if (lastRowMH > 1) {
        var rawDataMH = sheetMonHoc.getRange(2, 1, lastRowMH - 1, 2).getValues();
        rawDataMH.forEach(function(row) {
          if (row[0].toString().trim() !== "") listMonHoc.push(row[0].toString().trim());
          if (row[1].toString().trim() !== "") listLopHoc.push(row[1].toString().trim());
        });
        listMonHoc = listMonHoc.filter(function(item, pos) { return listMonHoc.indexOf(item) === pos; });
        listLopHoc = listLopHoc.filter(function(item, pos) { return listLopHoc.indexOf(item) === pos; });
      }
    }
    
    return { namHoc: listNamHoc, banDoHeThong: mapHeThong, monHoc: listMonHoc, lopHoc: listLopHoc };
  } catch (error) {
    return { namHoc: [], banDoHeThong: {}, monHoc: [], lopHoc: [] };
  }
}

function taiDuLieuNguocVeForm(nam, truong, gv, tuanNum) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    if (!sheetGiaoVien) throw "Không tìm thấy trang tính danh mục giáo viên!";

    var dataGV = sheetGiaoVien.getRange(2, 1, sheetGiaoVien.getLastRow() - 1, 12).getValues();
    var targetRowIndexInGV = -1;
    var targetSsId = "";
    var thongTinThoiGian = { ngayBatDau: "", tetTu: "", tetDen: "", tuan: "", tuanTuNgay: "", tuanDenNgay: "" };

    for (var i = 0; i < dataGV.length; i++) {
      if (dataGV[i][0].toString().trim() === nam && dataGV[i][1].toString().trim() === truong && dataGV[i][2].toString().trim() === gv) {
        targetRowIndexInGV = i + 2; 
        targetSsId = dataGV[i][4].toString().trim(); 
        var tuanHienTai = dataGV[i][8] ? parseInt(dataGV[i][8].toString().trim(), 10) : 1;
        
        if (tuanNum && tuanNum !== "CHUYEN_GV") {
          tuanHienTai = parseInt(tuanNum, 10);
          sheetGiaoVien.getRange(targetRowIndexInGV, 9).setValue(tuanHienTai);
          SpreadsheetApp.flush();
        }
        
        if (isNaN(tuanHienTai) || tuanHienTai < 1) tuanHienTai = 1;
        var startRow = 5 + (tuanHienTai - 1) * 61;
        var totalRowsToLoad = 60;

        var giaTriNgayMoi = sheetGiaoVien.getRange(targetRowIndexInGV, 10, 1, 2).getValues();
        var tz = Session.getScriptTimeZone();
        var dinhDangNgay = function(o) {
          if (o instanceof Date) return Utilities.formatDate(o, tz, "dd/MM/yyyy");
          return o ? o.toString().trim() : "";
        };

        thongTinThoiGian.ngayBatDau = dinhDangNgay(dataGV[i][5]); 
        thongTinThoiGian.tetTu = dinhDangNgay(dataGV[i][6]);     
        thongTinThoiGian.tetDen = dinhDangNgay(dataGV[i][7]);    
        thongTinThoiGian.tuan = tuanHienTai.toString(); 
        thongTinThoiGian.tuanTuNgay = dinhDangNgay(giaTriNgayMoi[0][0]); 
        thongTinThoiGian.tuanDenNgay = dinhDangNgay(giaTriNgayMoi[0][1]); 
        break;
      }
    }

    var matrixGridData = [];
    for(var k=0; k<totalRowsToLoad; k++) {
      matrixGridData.push(["", "", "", "", "", false]); 
    }
    var maTranAnBuoiDuocLuu = Array(12).fill(false);

    if (targetSsId !== "") {
      try {
        var personalSs = SpreadsheetApp.openById(targetSsId);
        var sheetLayLich = personalSs.getSheetByName("laylich");
        if (sheetLayLich !== null) {
          var valuesLayLich = sheetLayLich.getRange(startRow, 10, totalRowsToLoad, 6).getValues();
          for (var r = 0; r < totalRowsToLoad; r++) {
            var rowValues = valuesLayLich[r];
            matrixGridData[r] = [
              rowValues[0] ? rowValues[0].toString().trim() : "",
              rowValues[1] ? rowValues[1].toString().trim() : "",
              rowValues[2] ? rowValues[2].toString().trim() : "",
              rowValues[3] ? rowValues[3].toString().trim() : "",
              rowValues[4] ? rowValues[4].toString().trim() : "",
              (rowValues[5] === true || rowValues[5] === "true" || rowValues[5] === "TRUE")
            ];
          }
        }

        var sheetLuuAn = personalSs.getSheetByName("luu_an");
        if (sheetLuuAn !== null) {
          var giaTriLuuAn = sheetLuuAn.getRange(2, 1, 12, 2).getValues();
          for (var m = 0; m < 12; m++) {
            var valCheck = giaTriLuuAn[m][1];
            maTranAnBuoiDuocLuu[m] = (valCheck === true || valCheck === "true" || valCheck === "TRUE");
          }
        }
      } catch(e) {
        Logger.log("Lỗi: " + e.toString());
      }
    }

    return { 
      thanhCong: true, 
      thongTinThoiGian: thongTinThoiGian, 
      duLieuLuoi: matrixGridData,
      maTranAnBuoi: maTranAnBuoiDuocLuu 
    };
  } catch (err) {
    return { thanhCong: false, tinNhan: err.toString() };
  }
}

function luuBaoGiangHeThong(nam, truong, gv, thongTinThoiGian, dataGrid, mangMaTranAnBuoi) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    if (!sheetGiaoVien) throw "Không tìm thấy sheet 'giaovien'!";

    var tuan = parseInt(thongTinThoiGian.tuan, 10);
    if (isNaN(tuan) || tuan < 1) throw "Tuần học không hợp lệ!";
    var startRow = 5 + (tuan - 1) * 61; 
    var totalRowsToSave = 60; 

    var dataGV = sheetGiaoVien.getDataRange().getValues();
    var targetRowIndexInGV = -1;
    var targetSsId = "";

    for (var i = 1; i < dataGV.length; i++) {
      if (dataGV[i][0].toString().trim() === nam && dataGV[i][1].toString().trim() === truong && dataGV[i][2].toString().trim() === gv) {
        targetRowIndexInGV = i + 1; 
        if (dataGV[i][4]) targetSsId = dataGV[i][4].toString().trim();
        break;
      }
    }

    if (targetRowIndexInGV === -1 || targetSsId === "") throw "Không tìm thấy hồ sơ hoặc cấu hình ID riêng giáo viên!";

    var matrixJK = []; 
    var matrixNO = []; 

    for (var r = 0; r < totalRowsToSave; r++) {
      var rowData = dataGrid[r] || ["", "", "", "", "", "", "", false];
      matrixJK.push([
        rowData[2] !== undefined && rowData[2] !== null ? rowData[2].toString().trim() : "",
        rowData[3] !== undefined && rowData[3] !== null ? rowData[3].toString().trim() : ""
      ]);
      matrixNO.push([
        rowData[6] !== undefined && rowData[6] !== null ? rowData[6].toString().trim() : "",
        (rowData[7] === true || rowData[7] === "true" || rowData[7] === 1 || rowData[7] === "TRUE")
      ]);
    }

    var personalSs = SpreadsheetApp.openById(targetSsId);
    var sheetLayLich = personalSs.getSheetByName("laylich") || personalSs.insertSheet("laylich");
    sheetLayLich.getRange(startRow, 10, totalRowsToSave, 2).setValues(matrixJK); 
    sheetLayLich.getRange(startRow, 14, totalRowsToSave, 2).setValues(matrixNO); 

    var sheetLuuAn = personalSs.getSheetByName("luu_an") || personalSs.insertSheet("luu_an");
    var nhanBuoiCộtA = ["T2_Sang", "T2_Chieu", "T3_Sang", "T3_Chieu", "T4_Sang", "T4_Chieu", "T5_Sang", "T5_Chieu", "T6_Sang", "T6_Chieu", "T7_Sang", "T7_Chieu"];
    var mangGhiLuuAn = [];
    for (var m = 0; m < 12; m++) {
      mangGhiLuuAn.push([nhanBuoiCộtA[m], mangMaTranAnBuoi[m] === true || mangMaTranAnBuoi[m] === "true"]);
    }
    sheetLuuAn.getRange(2, 1, 12, 2).setValues(mangGhiLuuAn);

    sheetGiaoVien.getRange(targetRowIndexInGV, 6, 1, 4).setValues([[
      thongTinThoiGian.ngayBatDau, thongTinThoiGian.tetTu, thongTinThoiGian.tetDen, tuan
    ]]);

    return "Thành công!";
  } catch (err) {
    return "Lỗi thực thi: " + err.toString();
  }
}

function capNhatNghiVaLayKetQua(nam, truong, gv, tuanNum, rowIndex, isChecked) {
  try {
    var targetSsId = getTargetSsId(nam, truong, gv);
    if (!targetSsId) throw "Không tìm thấy ID file giáo viên";

    var tuan = parseInt(tuanNum, 10);
    if (isNaN(tuan) || tuan < 1) tuan = 1;
    
    // Tính toán cấu trúc dòng định vị vị trí tuần học
    var startRowOfWeek = 5 + (tuan - 1) * 61;
    var targetRow = startRowOfWeek + parseInt(rowIndex, 10);

    var personalSs = SpreadsheetApp.openById(targetSsId);
    var sheetLayLich = personalSs.getSheetByName("laylich");
    if (!sheetLayLich) throw "Không tìm thấy sheet 'laylich'";

    // 1. Đồng bộ trạng thái tích Nghỉ (Cột 15 - O) xuống hàng tương ứng trong sheet
    sheetLayLich.getRange(targetRow, 15).setValue(isChecked === true || isChecked === "true" || isChecked === "TRUE");
    SpreadsheetApp.flush(); // Ép hệ thống chạy công thức Excel tịnh tiến bài giảng cho toàn bộ tuần

    // 🚀 NÂNG CẤP ĐỒNG BỘ TOÀN TUẦN: Đọc trọn vẹn 60 dòng của cột Tiết PPCT (Cột 12 - L) và Tên Bài Dạy (Cột 13 - M)
    var rangeToanBoKetQuaTuan = sheetLayLich.getRange(startRowOfWeek, 12, 60, 2).getValues();
    
    // Đóng gói mảng kết quả dữ liệu 60 hàng của tuần vừa được tính toán lại
    var mangKetQuaTuan = [];
    for (var r = 0; r < 60; r++) {
      mangKetQuaTuan.push({
        dongIdx: r,
        tietPpct: rangeToanBoKetQuaTuan[r][0] ? rangeToanBoKetQuaTuan[r][0].toString().trim() : "",
        tenBaiDay: rangeToanBoKetQuaTuan[r][1] ? rangeToanBoKetQuaTuan[r][1].toString().trim() : ""
      });
    }

    return { 
      thanhCong: true, 
      rowIndex: rowIndex, 
      duLieuCapNhatTuan: mangKetQuaTuan // Gửi trọn vẹn mảng cấu trúc tuần về cho client
    };
  } catch (err) {
    return { thanhCong: false, error: err.toString() };
  }
}

function apDungTKBToanNam(nam, truong, gv, tuanHienTai, dataGrid) {
  try {
    var targetSsId = getTargetSsId(nam, truong, gv);
    if (!targetSsId) throw "Không tìm thấy ID file cá nhân giáo viên!";

    var matrixTKB = []; 
    for (var r = 0; r < 60; r++) {
      var rowData = dataGrid[r] || [];
      matrixTKB.push([rowData[2] ? rowData[2].toString().trim() : "", rowData[3] ? rowData[3].toString().trim() : ""]);
    }

    var tuanStart = parseInt(tuanHienTai, 10) + 1;
    if (tuanStart > 35) return "Không có tuần tiếp theo để áp dụng.";

    var fullMatrix = [];
    var blankRow = ["", ""]; 
    for (var w = tuanStart; w <= 35; w++) {
      fullMatrix = fullMatrix.concat(matrixTKB); 
      fullMatrix.push(blankRow); 
    }

    var personalSs = SpreadsheetApp.openById(targetSsId);
    var sheetLayLich = personalSs.getSheetByName("laylich");
    var firstRowOfStartWeek = 5 + (tuanStart - 1) * 61;
    sheetLayLich.getRange(firstRowOfStartWeek, 10, fullMatrix.length, 2).setValues(fullMatrix);

    return "Tuyệt vời! Đã áp dụng Thời khóa biểu thành công từ Tuần " + tuanStart + " đến hết Tuần 35!";
  } catch (err) {
    return "Lỗi hệ thống: " + err.toString();
  }
}

function importPPCTFromServer(nam, truong, gv, monHoc, khoiHocRaw, rawData) {
  try {
    var targetId = getTargetSsId(nam, truong, gv);
    if (!targetId) throw "Không tìm thấy link bảng tính giáo viên!";
    
    var targetSs = SpreadsheetApp.openById(targetId);
    var sheet = targetSs.getSheetByName(monHoc) || targetSs.insertSheet(monHoc);
    var tenKhoiSoHienTai = khoiHocRaw.replace(/[^0-9]/g, ""); 
    var duLieuKhoiKhac = [];
    var lastRowHienTai = sheet.getLastRow();
    
    if (lastRowHienTai > 1) {
      var toanBoDuLieuCu = sheet.getRange(2, 1, lastRowHienTai - 1, 4).getValues();
      for (var k = 0; k < toanBoDuLieuCu.length; k++) {
        var dongCu = toanBoDuLieuCu[k];
        if (dongCu[0].toString().trim() !== "" && dongCu[0].toString().trim() !== tenKhoiSoHienTai) {
          duLieuKhoiKhac.push(dongCu);
        }
      }
    }

    var tempNewRows = [];
    var hasSoTiet = rawData.length > 0 && rawData[0].length >= 3;

    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row[1] || row[1].toString().trim() === "") continue; 

      var colB_TenBai = row[1].toString().trim();
      var colC_SoTiet = row[2] ? row[2].toString().trim() : "";
      var laDongChuDe = (colB_TenBai !== "" && colC_SoTiet === "");

      if (hasSoTiet && !laDongChuDe && !isNaN(colC_SoTiet) && colC_SoTiet !== "" && parseInt(colC_SoTiet) > 1) {
        var soTietLong = parseInt(colC_SoTiet);
        for (var j = 0; j < soTietLong; j++) {
          tempNewRows.push({ tenBai: colB_TenBai, soTiet: 1 });
        }
      } else {
        tempNewRows.push({ tenBai: colB_TenBai, soTiet: laDongChuDe ? "" : 1 });
      }
    }

    var duLieuKhoiMoiFormat = [];
    var sttDem = 1; 

    for (var m = 0; m < tempNewRows.length; m++) {
      var targetRow = tempNewRows[m];
      var mangDongMoi = [
        tenKhoiSoHienTai,
        targetRow.soTiet === "" ? "" : sttDem++,
        targetRow.tenBai,
        targetRow.soTiet
      ];
      duLieuKhoiMoiFormat.push(mangDongMoi);
    }

    var matrixFinalData = duLieuKhoiKhac.concat(duLieuKhoiMoiFormat);
    sheet.clear(); 
    sheet.getRange(1, 1, 1, 4).setValues([["Khối", "STT", "Tên Bài Dạy", "Số tiết"]]);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");

    if (matrixFinalData.length > 0) {
      sheet.getRange(2, 1, matrixFinalData.length, 4).setValues(matrixFinalData);
      sheet.getRange(2, 1, matrixFinalData.length, 2).setHorizontalAlignment("center");
      sheet.getRange(2, 4, matrixFinalData.length, 1).setHorizontalAlignment("center");
    }

    return "Cập nhật hoàn tất! Hệ thống đã rã tiết và đánh số tuần tự cho " + (sttDem - 1) + " bài dạy của Khối " + tenKhoiSoHienTai + ".";
  } catch (e) {
    return "Lỗi máy chủ: " + e.toString();
  }
}

function taiDuLieuPPCTDeXem(nam, truong, gv, monHoc, khoiHocRaw) {
  try {
    var targetId = getTargetSsId(nam, truong, gv);
    if (!targetId) throw "Không tìm thấy tệp giáo viên!";
    
    var targetSs = SpreadsheetApp.openById(targetId);
    var sheet = targetSs.getSheetByName(monHoc);
    if (!sheet) throw "Môn học này hiện chưa được cấu hình dữ liệu PPCT!";
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { thanhCong: true, dataPPCT: [] }; 
    
    var rawData = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var tenKhoiSo = khoiHocRaw.replace(/[^0-9]/g, ""); 
    var dataKetQua = [];
    
    for (var i = 0; i < rawData.length; i++) {
      var row = rawData[i];
      if (row[0].toString().trim() === tenKhoiSo) {
        dataKetQua.push([row[1] ? row[1].toString().trim() : "", row[2] ? row[2].toString().trim() : "", row[3] ? row[3].toString().trim() : ""]);
      }
    }
    return { thanhCong: true, dataPPCT: dataKetQua };
  } catch (e) {
    return { thanhCong: false, error: e.toString() };
  }
}

function xuLyVaDongBoXemIn(nam, truong, gv, tuNgay, tuanVal, mangMaTranAnBuoi, dataGrid) {
  try {
    var ssInChung = SpreadsheetApp.openById(FILE_IN_CHUNG_ID);
    var sheetIn = ssInChung.getSheetByName("in");
    if (!sheetIn) throw "Không tìm thấy trang tính 'in'!";

    sheetIn.getRange("B6").setValue(tuNgay);
    var soTuanSo = parseInt(tuanVal, 10);
    sheetIn.getRange("D2").setValue(isNaN(soTuanSo) ? tuanVal : soTuanSo);

    var matrixIn = [];
    for (var r = 0; r < 60; r++) {
      var rData = dataGrid[r] || ["", "", "", "", "", "", "", false];
      matrixIn.push([
        rData[2] !== undefined && rData[2] !== null ? rData[2].toString().trim() : "", 
        rData[3] !== undefined && rData[3] !== null ? rData[3].toString().trim() : "", 
        rData[4] !== undefined && rData[4] !== null ? rData[4].toString().trim() : "", 
        rData[5] !== undefined && rData[5] !== null ? rData[5].toString().trim() : "", 
        rData[6] !== undefined && rData[6] !== null ? rData[6].toString().trim() : ""  
      ]);
    }
    sheetIn.getRange(4, 4, 60, 5).setValues(matrixIn);
    sheetIn.showRows(4, 60); 
    
    for (var m = 0; m < 12; m++) {
      if (mangMaTranAnBuoi[m] === true || mangMaTranAnBuoi[m] === "true") {
        sheetIn.hideRows(4 + (m * 5), 5); 
      }
    }
    SpreadsheetApp.flush();

    var urlXemIn = "https://docs.google.com/spreadsheets/d/" + FILE_IN_CHUNG_ID + "/htmlembed?gid=1179409360&widget=false&headers=false&chrome=false";
    return { thanhCong: true, urlPreview: urlXemIn };
  } catch (err) {
    return { thanhCong: false, error: err.toString() };
  }
}

function xuLyTaoBanInPDF(nam, truong, gv, tuNgay, tuanVal, mangMaTranAnBuoi, dataGrid) {
  try {
    var ssInChung = SpreadsheetApp.openById(FILE_IN_CHUNG_ID);
    var sheetIn = ssInChung.getSheetByName("in");
    var gidIn = sheetIn.getSheetId(); 

    sheetIn.getRange("B6").setValue(tuNgay);
    var soTuanSo = parseInt(tuanVal, 10);
    sheetIn.getRange("D2").setValue(isNaN(soTuanSo) ? tuanVal : soTuanSo);

    var matrixIn = [];
    for (var r = 0; r < 60; r++) {
      var rData = dataGrid[r] || ["", "", "", "", "", "", "", false];
      matrixIn.push([
        rData[2] !== undefined && rData[2] !== null ? rData[2].toString().trim() : "", 
        rData[3] !== undefined && rData[3] !== null ? rData[3].toString().trim() : "", 
        rData[4] !== undefined && rData[4] !== null ? rData[4].toString().trim() : "", 
        rData[5] !== undefined && rData[5] !== null ? rData[5].toString().trim() : "", 
        rData[6] !== undefined && rData[6] !== null ? rData[6].toString().trim() : ""  
      ]);
    }
    sheetIn.getRange(4, 4, 60, 5).setValues(matrixIn);
    sheetIn.showRows(4, 60); 
    
    for (var m = 0; m < 12; m++) {
      if (mangMaTranAnBuoi[m] === true || mangMaTranAnBuoi[m] === "true") {
        sheetIn.hideRows(4 + (m * 5), 5); 
      }
    }
    SpreadsheetApp.flush(); 

    var url = "https://docs.google.com/spreadsheets/d/" + FILE_IN_CHUNG_ID + "/export" +
              "?format=pdf&gid=" + gidIn + "&portrait=true&size=A4&fitw=true&gridlines=false" +
              "&printtitle=false&sheetnames=false&pagenum=UNDEFINED" +
              "&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5";

    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    return { thanhCong: true, pdfBase64: Utilities.base64Encode(response.getBlob().getBytes()) };
  } catch (err) {
    return { thanhCong: false, error: err.toString() };
  }
}

function xacNhanQuyenIn() {
  UrlFetchApp.fetch("https://www.google.com");
}

function xacThucMatKhauGiaoVien(nam, truong, gv, matKhauNhap) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    var dataGV = sheetGiaoVien.getRange(2, 1, sheetGiaoVien.getLastRow() - 1, 13).getValues(); // 🌟 Đọc đến cột 13 (Cột M)

    for (var i = 0; i < dataGV.length; i++) {
      if (dataGV[i][0].toString().trim() === nam && dataGV[i][1].toString().trim() === truong && dataGV[i][2].toString().trim() === gv) {
        
        var matKhauGoc = dataGV[i][11] ? dataGV[i][11].toString().trim() : "";
        var quyenAdminRaw = dataGV[i][12] ? dataGV[i][12].toString().trim().toLowerCase() : ""; // 🌟 Lấy chữ "x" ở cột M
        
        var dungMatKhau = (matKhauGoc === matKhauNhap);
        var coQuyenAdmin = (quyenAdminRaw === "x"); // 🌟 Kiểm tra điều kiện quyền Admin

        return { 
          hopLe: dungMatKhau, 
          laAdmin: coQuyenAdmin, // 🌟 Gửi cờ phân quyền về cho trình duyệt
          error: dungMatKhau ? "" : "Mật khẩu truy cập không chính xác!" 
        };
      }
    }
    return { hopLe: false, laAdmin: false, error: "Không tìm thấy hồ sơ giáo viên." };
  } catch (err) {
    return { hopLe: false, laAdmin: false, error: err.toString() };
  }
}

function capNhatMatKhauMoi(nam, truong, gv, matKhauCu, matKhauMoi) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    var dataGV = sheetGiaoVien.getRange(2, 1, sheetGiaoVien.getLastRow() - 1, 12).getValues();
    var targetRowIndex = -1;
    var matKhauGoc = "";

    for (var i = 0; i < dataGV.length; i++) {
      if (dataGV[i][0].toString().trim() === nam && dataGV[i][1].toString().trim() === truong && dataGV[i][2].toString().trim() === gv) {
        targetRowIndex = i + 2; 
        matKhauGoc = dataGV[i][11] ? dataGV[i][11].toString().trim() : ""; 
        break;
      }
    }

    if (targetRowIndex === -1) return { thanhCong: false, tinNhan: "Không tìm thấy hồ sơ Giáo viên." };
    if (matKhauGoc !== matKhauCu) return { thanhCong: false, tinNhan: "Mật khẩu hiện tại không khớp!" };

    sheetGiaoVien.getRange(targetRowIndex, 12).setValue(matKhauMoi);
    SpreadsheetApp.flush();
    return { thanhCong: true, tinNhan: "Mật khẩu truy cập đã được làm mới an toàn." };
  } catch (err) {
    return { thanhCong: false, tinNhan: "Lỗi hệ thống: " + err.toString() };
  }
}
/**
 * Tính năng mới: Xử lý nút Cập nhật - Nhân bản file mẫu, di chuyển vào thư mục Năm học/Trường học
 * và thêm giáo viên mới vào hệ thống nếu không trùng tên.
 */
/**
 * Tính năng viết lại: Xử lý nút Cập nhật - Nhân bản file mẫu chuẩn xác vị trí
 * Loại bỏ hoàn toàn các thư mục trùng tên nằm trong Thùng rác (Trash)
 */
function xuLyCapNhatVaNhanBanFile(nam, truong, tenGvMoi) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    if (!sheetGiaoVien) throw "Không tìm thấy sheet 'giaovien'!";

    var dataGV = sheetGiaoVien.getDataRange().getValues();
    
    // Xóa khoảng trắng thừa của toàn bộ dữ liệu đầu vào
    var tenGvMoiTrim = tenGvMoi.toString().trim();
    var namTrim = nam.toString().trim();
    var truongTrim = truong.toString().trim();

    // --- BƯỚC 1: KIỂM TRA TRÙNG TÊN GIÁO VIÊN TRONG CÙNG TRƯỜNG ---
    for (var i = 1; i < dataGV.length; i++) {
      if (dataGV[i][0].toString().trim() === namTrim && 
          dataGV[i][1].toString().trim() === truongTrim && 
          dataGV[i][2].toString().trim().toLowerCase() === tenGvMoiTrim.toLowerCase()) {
        return { thanhCong: false, tinNhan: "Tên giáo viên '" + tenGvMoiTrim + "' đã tồn tại trong trường này!" };
      }
    }

    // --- BƯỚC 2: QUẢN LÝ THƯ MỤC CẤU TRÚC (Bỏ qua Thùng rác) ---
    var folderGoc;
    try {
      // Ưu tiên lấy thư mục cha chứa file tổng này để cấu trúc gọn gàng một chỗ
      var parents = DriveApp.getFileById(MAIN_SPREADSHEET_ID).getParents();
      folderGoc = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    } catch (e) {
      folderGoc = DriveApp.getRootFolder();
    }
    
    // Tìm hoặc tạo thư mục Năm học (Chỉ lấy thư mục đang hoạt động, bỏ qua Thùng rác)
    var foldersNam = folderGoc.getFoldersByName(namTrim);
    var folderNamTarget = null;
    while (foldersNam.hasNext()) {
      var f = foldersNam.next();
      if (!f.isTrashed()) { // Kiểm tra nếu thư mục KHÔNG nằm trong thùng rác
        folderNamTarget = f;
        break;
      }
    }
    if (!folderNamTarget) {
      folderNamTarget = folderGoc.createFolder(namTrim);
    }
    
    // Tìm hoặc tạo thư mục Trường học nằm trong thư mục Năm học (Bỏ qua Thùng rác)
    var foldersTruong = folderNamTarget.getFoldersByName(truongTrim);
    var folderTruongTarget = null;
    while (foldersTruong.hasNext()) {
      var f = foldersTruong.next();
      if (!f.isTrashed()) { // Kiểm tra nếu thư mục KHÔNG nằm trong thùng rác
        folderTruongTarget = f;
        break;
      }
    }
    if (!folderTruongTarget) {
      folderTruongTarget = folderNamTarget.createFolder(truongTrim);
    }

    // --- BƯỚC 3: NHÂN BẢN THẲNG FILE MẪU VÀO THƯ MỤC ĐÍCH ---
    var ID_FILE_MAU = "1GmeTHLhVcR-Hd1K8hfhw8K8W1qEmLLMZnpSy6DPz0Sk"; 
    var fileMau = DriveApp.getFileById(ID_FILE_MAU);
    
    // Nhân bản trực tiếp vào thư mục Trường học đích (không tạo file trung gian ở Root)
    var fileSaoChep = fileMau.makeCopy(tenGvMoiTrim, folderTruongTarget);
    var urlFileMoi = fileSaoChep.getUrl();
    var idFileMoi = fileSaoChep.getId();

    // --- BƯỚC 4: GHI DỮ LIỆU ĐỒNG BỘ VÀO SHEET ---
    var indexDongMoi = sheetGiaoVien.getLastRow() + 1;
    sheetGiaoVien.getRange(indexDongMoi, 1).setValue(namTrim);
    sheetGiaoVien.getRange(indexDongMoi, 2).setValue(truongTrim);
    sheetGiaoVien.getRange(indexDongMoi, 3).setValue(tenGvMoiTrim);
    sheetGiaoVien.getRange(indexDongMoi, 5).setValue(idFileMoi); 
    
    SpreadsheetApp.flush();

    if (typeof ghiLogHanhDong === "function") {
      ghiLogHanhDong(tenGvMoiTrim, "Thêm Giáo Viên", "Khởi tạo file lịch thành công tại thư mục: " + namTrim + "/" + truongTrim);
    }

    return { 
      thanhCong: true, 
      tinNhan: "Xuất sắc! Đã tạo cấu hình và sinh file lịch báo giảng thành công cho giáo viên: " + tenGvMoiTrim,
      url: urlFileMoi
    };

  } catch (err) {
    return { thanhCong: false, tinNhan: "Lỗi xử lý máy chủ: " + err.toString() };
  }
}
/**
 * Tính năng: Đồng bộ trực tiếp Môn học hoặc Lớp học từ lưới về ô tính tương ứng trong sheet laylich
 * Môn học (Cột 2 trên lưới) -> Cột 10 (J) trong sheet laylich
 * Lớp học (Cột 3 trên lưới) -> Cột 11 (K) trong sheet laylich
 */
/**
 * Tính năng: Đồng bộ trực tiếp Môn học/Lớp học và Lấy ngược kết quả Tiết PPCT & Tên bài
 */
function capNhatMonLopRealTime(nam, truong, gv, tuanNum, rowIndex, colIndex, giaTriMoi) {
  try {
    var targetSsId = getTargetSsId(nam, truong, gv);
    if (!targetSsId) return { thanhCong: false, error: "Không tìm thấy ID file giáo viên" };

    var tuan = parseInt(tuanNum, 10);
    if (isNaN(tuan) || tuan < 1) tuan = 1;
    
    // Tính toán dòng bắt đầu của khối 60 dòng trong tuần
    var startRowOfWeek = 5 + (tuan - 1) * 61;
    var targetRow = startRowOfWeek + parseInt(rowIndex, 10);
    var targetCol = (parseInt(colIndex, 10) === 2) ? 10 : 11; // Cột 10 (J) là Môn, Cột 11 (K) là Lớp

    var personalSs = SpreadsheetApp.openById(targetSsId);
    var sheetLayLich = personalSs.getSheetByName("laylich");
    if (!sheetLayLich) return { thanhCong: false, error: "Không tìm thấy trang tính laylich" };

    // 1. Ghi giá trị Môn/Lớp mà giáo viên vừa chọn xuống ô tính đích
    sheetLayLich.getRange(targetRow, targetCol).setValue(giaTriMoi ? giaTriMoi.toString().trim() : "");
    SpreadsheetApp.flush(); // Ép hệ thống chạy công thức Excel hàng loạt cho cả tuần

    // 🚀 GIẢI PHÁP MỚI: Đọc trọn vẹn 60 dòng của cột Tiết PPCT (Cột 12 - L) và Tên Bài Dạy (Cột 13 - M) của cả tuần
    var rangeToanBoKetQuaTuan = sheetLayLich.getRange(startRowOfWeek, 12, 60, 2).getValues();
    
    // Đóng gói mảng kết quả 60 dòng
    var mangKetQuaTuan = [];
    for (var r = 0; r < 60; r++) {
      mangKetQuaTuan.push({
        dongIdx: r,
        tietPpct: rangeToanBoKetQuaTuan[r][0] ? rangeToanBoKetQuaTuan[r][0].toString().trim() : "",
        tenBaiDay: rangeToanBoKetQuaTuan[r][1] ? rangeToanBoKetQuaTuan[r][1].toString().trim() : ""
      });
    }

    return { 
      thanhCong: true, 
      rowIndex: rowIndex,
      duLieuCapNhatTuan: mangKetQuaTuan // Trả mảng 60 dòng về cho trình duyệt nhận diện
    };
  } catch (err) {
    return { thanhCong: false, error: err.toString() };
  }
}
function kiemTraQuyenCapNhat(nam, truong, tenGv) {
  try {
    if (!tenGv) return false;
    
    var ss = SpreadsheetApp.openById("1xE1-IB-Mj3L9lsazgXtB_LyNuO-5qWXNpShKqc9Z_qc");
    var sheetGv = ss.getSheetByName("giaovien");
    if (!sheetGv) return false;
    
    var data = sheetGv.getDataRange().getValues();
    var targetGv = tenGv.toString().trim().toLowerCase(); 
    
    // Bắt đầu quét từ dòng 2 (index 1) trở đi
    for (var i = 1; i < data.length; i++) {
      // Đã sửa thành data[i][2] để lấy dữ liệu Tên giáo viên từ Cột C
      var currentGvName = data[i][2] ? data[i][2].toString().trim().toLowerCase() : "";
      
      // Lấy dữ liệu dấu X từ Cột M
      var processX = data[i][12] ? data[i][12].toString().trim().toLowerCase() : "";
      
      // Khớp tên giáo viên và có chữ "x"
      if (currentGvName === targetGv && processX === "x") { 
        return true; 
      }
    }
    return false;
    
  } catch (err) {
    Logger.log("Lỗi hệ thống phân quyền: " + err.toString());
    return false;
  }
}
/**
 * Tính năng Admin nâng cao: Cấp lại mật khẩu về mặc định 12345678 cho giáo viên được chọn
 * Kiểm tra chéo bảo mật: Chỉ cho phép ghi dữ liệu nếu tài khoản thực hiện có dấu "x" ở cột M
 */
function xuLyCapLaiMatKhauAdmin(nam, truong, tenGvCanReset, tenGvDangNhap) {
  try {
    // Lớp bảo mật 1: Kiểm tra quyền quản trị của tài khoản đang thao tác bấm nút
    var coQuyenAdmin = kiemTraQuyenCapNhat(nam, truong, tenGvDangNhap);
    if (!coQuyenAdmin) {
      return { thanhCong: false, tinNhan: "Từ chối thực thi: Bạn không có quyền quản trị để cấp lại mật khẩu!" };
    }

    var ss = SpreadsheetApp.openById("1xE1-IB-Mj3L9lsazgXtB_LyNuO-5qWXNpShKqc9Z_qc");
    var sheetGv = ss.getSheetByName("giaovien");
    if (!sheetGv) return { thanhCong: false, tinNhan: "Lỗi hệ thống: Không tìm thấy trang tính 'giaovien'!" };

    var data = sheetGv.getDataRange().getValues();
    var targetNam = nam.toString().trim().toLowerCase();
    var targetTruong = truong.toString().trim().toLowerCase();
    var targetGvCanReset = tenGvCanReset.toString().trim().toLowerCase();
    
    var timThayDòngDuLieu = false;

    // Quét dữ liệu từ dòng 2 (index 1) để tìm giáo viên đích
    for (var i = 1; i < data.length; i++) {
      var currentNam = data[i][0] ? data[i][0].toString().trim().toLowerCase() : "";
      var currentTruong = data[i][1] ? data[i][1].toString().trim().toLowerCase() : "";
      var currentGvName = data[i][2] ? data[i][2].toString().trim().toLowerCase() : ""; // Tên giáo viên nằm ở Cột C

      if (currentNam === targetNam && currentTruong === targetTruong && currentGvName === targetGvCanReset) {
        // Thực hiện ghi đè mật khẩu mới "12345678" vào Cột L (Index 11) của giáo viên đó
        sheetGv.getRange(i + 1, 12).setValue("12345678"); 
        timThayDòngDuLieu = true;
        break;
      }
    }

    if (!timThayDòngDuLieu) {
      return { thanhCong: false, tinNhan: "Không tìm thấy thông tin giáo viên cần xử lý trên cơ sở dữ liệu!" };
    }

    SpreadsheetApp.flush(); // Ép hệ thống đồng bộ dữ liệu thực xuống Sheet ngay lập tức
    return { 
      thanhCong: true, 
      tinNhan: "Khởi tạo thành công! Mật khẩu của giáo viên [" + tenGvCanReset + "] đã được đưa về mặc định: 12345678" 
    };

  } catch (err) {
    return { thanhCong: false, tinNhan: "Lỗi xử lý máy chủ: " + err.toString() };
  }
}
// ==========================================================================
// HÀM SERVER: XỬ LÝ CÔ LẬP TÀI KHOẢN VÀ TIÊU HỦY FILE DRIVE GIÁO VIÊN
// ==========================================================================
// ==========================================================================
// HÀM SERVER: XÓA GIÁO VIÊN VÀ TIÊU HỦY FILE DRIVE ĐÚNG CỘT E (CỘT 5)
// ==========================================================================
// ==========================================================================
// HÀM SERVER CHUẨN: XÓA THẲNG DÒNG TRÊN SHEET & TIÊU HỦY FILE DRIVE THEO TÊN
// ==========================================================================
// ==========================================================================
// HÀM SERVER CHUẨN: XÓA THẲNG DÒNG TRÊN SHEET & TIÊU HỦY FILE DRIVE DỰA VÀO CỘT E/D
// ==========================================================================
// ==========================================================================
// HÀM SERVER: TIÊU HỦY RUỘT FILE NGAY LẬP TỨC & XÓA DÒNG TRÊN SHEET GIAOVIEN
// ==========================================================================
// ==========================================================================
// HÀM SERVER: TRỤC XUẤT CƯỠNG CHẾ FILE SANG THƯ MỤC ẨN & XÓA DÒNG TRÊN SHEET
// ==========================================================================
function xuLyXoaGiaoVienAdmin(nam, truong, gvCanXoa, gvAdminThaoTac) {
  try {
    var ss = SpreadsheetApp.openById(MAIN_SPREADSHEET_ID);
    var sheetGiaoVien = ss.getSheetByName("giaovien");
    if (!sheetGiaoVien) return { thanhCong: false, tinNhan: "Lỗi kết nối: Không tìm thấy sheet giaovien!" };

    var lastRow = sheetGiaoVien.getLastRow();
    if (lastRow <= 1) return { thanhCong: false, tinNhan: "Bảng tính nhân sự trống!" };

    var range = sheetGiaoVien.getRange(2, 1, lastRow - 1, 13);
    var dataGV = range.getValues();
    
    var indexDongTrongSheet = -1;
    var idFileKhaiThac = "";

    // 1. Quét tìm dòng chứa thông tin giáo viên cần xóa
    for (var i = 0; i < dataGV.length; i++) {
      var sNam = dataGV[i][0].toString().trim();
      var sTruong = dataGV[i][1].toString().trim();
      var sTen = dataGV[i][2].toString().trim();

      if (sNam === nam && sTruong === truong && (sTen === gvCanXoa || sTen.indexOf(gvCanXoa) === 0)) {
        indexDongTrongSheet = i + 2; 
        idFileKhaiThac = dataGV[i][4] ? dataGV[i][4].toString().trim() : ""; // ID File ở cột E
        break;
      }
    }

    if (indexDongTrongSheet === -1) {
      return { thanhCong: false, tinNhan: "Không tìm thấy hồ sơ giáo viên trên hệ thống!" };
    }

    // 2. 🚀 THỰC THI LUỒNG TIÊU HỦY: DI CHUYỂN FILE SANG THƯ MỤC "THÙNG RÁC NGẦM" CỦA ADMIN
    var thongBaoDrive = "";
    if (idFileKhaiThac !== "" && idFileKhaiThac.length > 10) {
      try {
        var fileDrive = DriveApp.getFileById(idFileKhaiThac);
        
        // Bước A: Tạo hoặc Tìm thư mục ẩn tên là "THƯ MỤC ĐÃ TIÊU HỦY - BÁO GIẢNG" trên Drive của Admin
        var folderAnName = "THƯ MỤC ĐÃ TIÊU HỦY - BÁO GIẢNG";
        var folders = DriveApp.getFoldersByName(folderAnName);
        var targetFolderAn;
        
        if (folders.hasNext()) {
          targetFolderAn = folders.next();
        } else {
          targetFolderAn = DriveApp.createFolder(folderAnName); // Tạo mới nếu chưa có
        }
        
        // Bước B: Đổi tên file để Admin dễ quản lý hậu kiểm
        fileDrive.setName("🔒 [ĐÃ TRỤC XUẤT] - " + gvCanXoa + " (" + truong + ")");
        
        // Bước C: Bơm file vào thư mục ẩn này
        targetFolderAn.addFile(fileDrive);
        
        // Bước D: Quét sạch toàn bộ các thư mục cũ (Thư mục trường) và cắt đứt liên kết
        var parents = fileDrive.getParents();
        while (parents.hasNext()) {
          var parent = parents.next();
          if (parent.getId() !== targetFolderAn.getId()) {
            parent.removeFile(fileDrive); // Ép biến mất khỏi folder trường ngay lập tức!
          }
        }
        
        // Bước E: Khóa quyền riêng tư để giáo viên không thể bấm vào link cũ xem lại lưới
        try {
          fileDrive.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
        } catch(e_lock) {}
        
        thongBaoDrive = "Đã bốc file vứt sang thư mục lưu trữ ẩn của Quản trị viên.";
      } catch (errDrive) {
        thongBaoDrive = "Không thể can thiệp Drive (Có thể file đã bị gỡ hoặc sai cấu hình ID cột E).";
      }
    } else {
      thongBaoDrive = "Không tìm thấy chuỗi ID file tại cột E.";
    }

    // 3. ⚡ GỠ BỎ HOÀN TOÀN DÒNG DỮ LIỆU TRÊN SHEET GIAOVIEN NGAY TỨC THÌ
    sheetGiaoVien.deleteRow(indexDongTrongSheet);

    // Ghi nhận lịch sử quản trị
    if (typeof ghiLogHanhDong === "function") {
      ghiLogHanhDong(gvAdminThaoTac, "TRỤC XUẤT FILE", "Đã gỡ hồ sơ & đẩy file của GV " + gvCanXoa + " vào kho ẩn.");
    }

    return { 
      thanhCong: true, 
      tinNhan: "[QUẢN TRỊ ADMIN] Thực thi lệnh trục xuất dứt điểm thành công!\n\n- " + thongBaoDrive + "\n- Tài khoản đã bị xóa hoàn toàn khỏi Sheet danh mục hệ thống." 
    };

  } catch (err) {
    return { thanhCong: false, tinNhan: "Lỗi lớp Server: " + err.toString() };
  }
}
function kichHoatQuyenDrive_Test() {
  // Lệnh này ép Google phải hỏi quyền tạo thư mục và xóa thư mục
  DriveApp.createFolder("Test_Quyen_Google_Drive").setTrashed(true);
}

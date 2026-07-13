function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : "internal";
  var user = e && e.parameter && e.parameter.user ? e.parameter.user : "";
  
  template.roleMode = page; 
  template.targetVendorUser = user;
  
  return template.evaluate()
      .setTitle('SCRMS Gaskeun')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function saveBufferSettings(heroBuffer, vendorBuffer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var paramSheet = ss.getSheetByName("Parameter_Global");
  if (!paramSheet) {
    paramSheet = ss.insertSheet("Parameter_Global");
    paramSheet.appendRow(["Parameter", "Value"]);
    paramSheet.appendRow(["Hero_Buffer", heroBuffer]);
    paramSheet.appendRow(["Vendor_Buffer", vendorBuffer]);
  } else {
    var data = paramSheet.getDataRange().getValues();
    var heroFound = false, vendorFound = false;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === "Hero_Buffer") { paramSheet.getRange(i+1, 2).setValue(heroBuffer); heroFound = true; }
      if (data[i][0] === "Vendor_Buffer") { paramSheet.getRange(i+1, 2).setValue(vendorBuffer); vendorFound = true; }
    }
    if(!heroFound) paramSheet.appendRow(["Hero_Buffer", heroBuffer]);
    if(!vendorFound) paramSheet.appendRow(["Vendor_Buffer", vendorBuffer]);
  }
  return "Pengaturan berhasil disimpan ke Parameter_Global!";
}

function readData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var paramSheet = ss.getSheetByName("Parameter_Global");
    var heroBuf = 14, vendorBuf = 7;
    
    if(paramSheet) {
      var paramData = paramSheet.getDataRange().getValues();
      for (var p = 0; p < paramData.length; p++) {
        if (paramData[p][0] === "Hero_Buffer") heroBuf = parseInt(paramData[p][1]) || 14;
        if (paramData[p][0] === "Vendor_Buffer") vendorBuf = parseInt(paramData[p][1]) || 7;
      }
    }
    
    var sheet = ss.getSheetByName("Data_Artikel");
    if (!sheet) return { categories: [], articles: [], settings: { hero: heroBuf, vendor: vendorBuf } };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { categories: [], articles: [], settings: { hero: heroBuf, vendor: vendorBuf } };
    
    var categoriesSet = new Set(); 
    var articles = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[2] || row[2].toString().trim() === "") continue; 
      
      var kat = row[1] ? row[1].toString().trim().toUpperCase() : "UN-CATEGORIZED"; 
      categoriesSet.add(kat);
      
      var tsValue = "-";
      if (row[8]) {
        if (row[8] instanceof Date) {
          tsValue = Utilities.formatDate(row[8], "Asia/Jakarta", "dd/MM/yyyy HH:mm");
        } else {
          tsValue = row[8].toString().trim();
        }
      }

      var statusAktif = row[10] ? row[10].toString().trim() : "Aktif";
      var adsVal = row[11] ? parseFloat(row[11]) : 0; 
      var prioFlag = row[12] ? (row[12].toString().trim().toUpperCase() === "PRIORITY") : false;

      articles.push({
        bulan: row[0] ? row[0].toString().trim() : "",        
        kategori: kat,                                        
        sku: row[2].toString().trim(),                        
        nama: row[3] ? row[3].toString().trim() : "-",        
        status: row[4] ? row[4].toString().trim() : "Real Stock", 
        warna: row[5] ? row[5].toString().trim() : "-",       
        size: row[6] ? row[6].toString().trim() : "ALL SIZE", 
        qty: parseInt(row[7]) || 0,                           
        timestamp: tsValue,                                   
        tren: row[9] ? row[9].toString().trim() : "-",        
        isAktif: statusAktif, 
        adsGlobal: adsVal,
        isPriority: prioFlag 
      });
    }
    
    return { categories: Array.from(categoriesSet).sort(), articles: articles, settings: { hero: heroBuf, vendor: vendorBuf } };
  } catch (error) { throw new Error("Gagal membaca database: " + error.message); }
}

function setPriorityStatus(sku, setAsPriority) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Data_Artikel");
    var data = sheet.getDataRange().getValues();
    var targetVal = setAsPriority ? "PRIORITY" : ""; 
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2].toString().trim() === sku.trim()) {
        sheet.getRange(i + 1, 13).setValue(targetVal); 
      }
    }
    return "Status prioritas berhasil diperbarui.";
  } catch(e) { throw new Error(e.message); }
}

function writeLog(sku, nama, size, oldQty, newQty, operator) {
  if (parseInt(oldQty) === parseInt(newQty)) return; 
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Pergerakan_Stok");
    if (!sheet) {
      sheet = ss.insertSheet("Log_Pergerakan_Stok");
      sheet.appendRow(["Timestamp", "SKU", "Nama Artikel", "Size", "Stok Lama", "Stok Baru", "Perubahan", "Operator"]);
    }
    var timestampStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm");
    var diff = newQty - oldQty;
    var perubahanStr = diff > 0 ? "▲ (+" + diff + ")" : "▼ (" + diff + ")";
    sheet.appendRow([timestampStr, sku.trim(), nama.trim(), size.trim(), oldQty, newQty, perubahanStr, operator]);
    cleanOldLogs();
  } catch(e) { console.error("Gagal menulis log: " + e.message); }
}

function getStockLog(sku) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Pergerakan_Stok");
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var rowSku = data[i][1] ? data[i][1].toString().trim() : "";
      if (rowSku === sku.trim()) {
        logs.push({
          timestamp: data[i][0].toString(),
          sku: data[i][1].toString(),
          nama: data[i][2].toString(),
          size: data[i][3].toString(),
          oldQty: data[i][4],
          newQty: data[i][5],
          perubahan: data[i][6].toString(),
          operator: data[i][7].toString()
        });
      }
    }
    return logs;
  } catch(e) { return []; }
}

function cleanOldLogs() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Log_Pergerakan_Stok");
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return;
    var now = new Date();
    var rowsToDelete = [];
    for (var i = 1; i < data.length; i++) {
      var tsStr = data[i][0].toString();
      if (!tsStr || tsStr === "-") continue;
      try {
        var parts = tsStr.split(" ");
        var dateParts = parts[0].split("/");
        var timeParts = parts[1].split(":");
        var logDate = new Date(dateParts[2], dateParts[1] - 1, dateParts[0], timeParts[0], timeParts[1]);
        var ageInDays = (now.getTime() - logDate.getTime()) / (1000 * 60 * 60 * 24);
        if (ageInDays > 5) { rowsToDelete.push(i + 1); }
      } catch(err) { continue; }
    }
    for (var j = rowsToDelete.length - 1; j >= 0; j--) { sheet.deleteRow(rowsToDelete[j]); }
  } catch(e) { console.error("Gagal membersihkan log usang: " + e.message); }
}

// =================================================================
// ENGINE DATABASE FITUR KOTRETAN VENDOR (PERSISTENT SHEET)
// =================================================================

function saveArticleToKotretanDb(vendorUser, sku) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kotretan_Vendor");
    if (!sheet) {
      sheet = ss.insertSheet("Kotretan_Vendor");
      sheet.appendRow(["Username Vendor", "SKU Artikel"]);
    }
    
    var data = sheet.getDataRange().getValues();
    // Cek duplikasi di DB
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === vendorUser.trim() && data[i][1].toString().trim() === sku.trim()) {
        return "SUCCESS: Artikel sudah tercatat.";
      }
    }
    
    sheet.appendRow([vendorUser.trim(), sku.trim()]);
    return "SUCCESS: Artikel ditambahkan ke database.";
  } catch(e) { throw new Error(e.message); }
}

function removeArticleFromKotretanDb(vendorUser, sku) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kotretan_Vendor");
    if (!sheet) return "SUCCESS";
    
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === vendorUser.trim() && data[i][1].toString().trim() === sku.trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    return "SUCCESS";
  } catch(e) { throw new Error(e.message); }
}

function clearVendorKotretanDb(vendorUser) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kotretan_Vendor");
    if (!sheet) return "SUCCESS";
    
    var data = sheet.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][0].toString().trim() === vendorUser.trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    return "SUCCESS";
  } catch(e) { throw new Error(e.message); }
}

function loadSavedKotretanDb() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Kotretan_Vendor");
    if (!sheet) return {};
    
    var data = sheet.getDataRange().getValues();
    var map = {};
    
    for (var i = 1; i < data.length; i++) {
      var vUser = data[i][0].toString().trim();
      var sku = data[i][1].toString().trim();
      if (!vUser || !sku) continue;
      
      if (!map[vUser]) map[vUser] = [];
      map[vUser].push(sku);
    }
    return map; // Return format: { "vendorA": ["SKU1", "SKU2"], "vendorB": [...] }
  } catch(e) { return {}; }
}

// =================================================================

function updateMultipleStock(sku, sizesData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Artikel");
    var data = sheet.getDataRange().getValues(); var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm"); 
    var updatedCount = 0; var sizeMap = {};
    sizesData.forEach(function(item) { sizeMap[item.size.toString().trim()] = parseInt(item.qty); });
    for (var i = 1; i < data.length; i++) {
      var rowSku = data[i][2].toString().trim(); var rowSize = data[i][6].toString().trim();
      if (rowSku === sku.trim() && sizeMap.hasOwnProperty(rowSize)) {
        var newQty = sizeMap[rowSize]; var oldQty = parseInt(data[i][7]) || 0; var newTrend = data[i][9] || "-";
        if (newQty > oldQty) newTrend = "▲"; else if (newQty < oldQty) newTrend = "▼";
        sheet.getRange(i + 1, 8).setValue(newQty); sheet.getRange(i + 1, 9).setValue(timestamp); sheet.getRange(i + 1, 10).setValue(newTrend);
        writeLog(sku, data[i][3].toString(), rowSize, oldQty, newQty, "Staf (Update Varian)");
        updatedCount++;
      }
    }
    return "Berhasil memperbarui " + updatedCount + " ukuran!";
  } catch(e) { throw new Error(e.message); }
}

function toggleStatusAktif(sku, statusBaru) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Artikel"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][2].toString().trim() === sku.trim()) { sheet.getRange(i + 1, 11).setValue(statusBaru); } }
  return "Status diubah menjadi " + statusBaru;
}

function setHeroADS(sku, adsValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Artikel"); var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { if (data[i][2].toString().trim() === sku.trim()) { sheet.getRange(i + 1, 12).setValue(adsValue); } }
  return "Pengaturan Hero Alert tersimpan!";
}

function importBulkData(dataArray) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Artikel");
  var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm"); var newRows = [];
  for (var i = 0; i < dataArray.length; i++) {
    var row = dataArray[i]; var kategori = (row[0] || "").toString().trim().toUpperCase(); var sku = (row[1] || "").toString().trim();
    var nama = (row[2] || "").toString().trim(); var status = (row[3] || "").toString().trim(); var warna = (row[4] || "").toString().trim();
    if (!sku) continue; 
    
    function pushData(size, qtyIndex) { 
      var qty = parseInt(row[qtyIndex]) || 0; 
      newRows.push(["", kategori, sku, nama, status, warna, size, qty, timestamp, "-", "Aktif", 0, ""]); 
      writeLog(sku, nama, size, 0, qty, "Bulk Import Excel");
    }
    if (kategori === "BAG" || kategori === "BAGS") { pushData("ALL SIZE", 5); } else if (kategori === "PANTS" || kategori === "SWEAT PANTS" || kategori === "SWEATPANTS") { pushData("28", 5); pushData("30", 6); pushData("32", 7); pushData("34", 8); } else { pushData("S", 5); pushData("M", 6); pushData("L", 7); pushData("XL", 8); pushData("XXL", 9); }
  }
  if (newRows.length > 0) { var startRow = sheet.getLastRow() + 1; sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows); }
  return "Berhasil mengimpor " + newRows.length + " baris!";
}

function saveOrUpdateSingleArticle(kategori, sku, nama, warna, status, adsGlobal, sizesData, isArtikelBaru) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Artikel");
    var data = sheet.getDataRange().getValues(); var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm");
    var sizeMap = {}; sizesData.forEach(function(item) { sizeMap[item.size.toString().trim()] = { qty: parseInt(item.qty) || 0, isChanged: item.isChanged }; });

    if (!isArtikelBaru) {
      for (var i = 1; i < data.length; i++) {
        var rowSku = data[i][2].toString().trim(); var rowSize = data[i][6].toString().trim();
        if (rowSku === sku.trim() && sizeMap.hasOwnProperty(rowSize)) {
          var targetSizeData = sizeMap[rowSize];
          sheet.getRange(i + 1, 11).setValue("Aktif"); sheet.getRange(i + 1, 9).setValue(timestamp);
          if (targetSizeData.isChanged) {
            var oldQty = parseInt(data[i][7]) || 0; var newQty = targetSizeData.qty; var newTrend = "-";
            if (newQty > oldQty) newTrend = "▲"; else if (newQty < oldQty) newTrend = "▼";
            sheet.getRange(i + 1, 8).setValue(newQty); sheet.getRange(i + 1, 10).setValue(newTrend); 
            writeLog(sku, data[i][3].toString(), rowSize, oldQty, newQty, "Staf (Manual)");
          }
        }
      }
      return "Berhasil memperbarui data SKU: " + sku;
    } else {
      var rowsToAppend = [];
      sizesData.forEach(function(item) { 
        rowsToAppend.push(["", kategori.trim().toUpperCase(), sku.trim(), nama.trim(), status, warna.trim() || "-", item.size, parseInt(item.qty) || 0, timestamp, "-", "Aktif", parseFloat(adsGlobal) || 0, ""]); 
        writeLog(sku.trim(), nama.trim(), item.size, 0, parseInt(item.qty) || 0, "Staf (Artikel Baru)");
      });
      if (rowsToAppend.length > 0) { sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend); }
      return "Berhasil mendaftarkan SKU BARU: " + sku;
    }
  } catch(e) { throw new Error("Gagal memproses data: " + e.message); }
}

function getWebAppUrl() { return ScriptApp.getService().getUrl(); }

function readVendorData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Vendor");
    if (!sheet) { sheet = ss.insertSheet("Data_Vendor"); sheet.appendRow(["Nama Vendor", "Username", "Password (Removed)", "Akses Kategori"]); }
    var data = sheet.getDataRange().getValues(); var vendors = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][1]) continue;
      vendors.push({ nama: data[i][0] ? data[i][0].toString() : "", username: data[i][1] ? data[i][1].toString() : "", akses: data[i][3] ? data[i][3].toString().split(",") : [] });
    }
    return vendors;
  } catch(e) { throw new Error(e.message); }
}

function saveNewVendor(nama, username, listKategori) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Vendor"); var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) { if (data[i][1].toString().trim() === username.trim()) { return "ERROR: Username / ID Link sudah terdaftar!"; } }
    var aksesString = listKategori.join(","); sheet.appendRow([nama.trim(), username.trim(), "-", aksesString]);
    return "SUCCESS: Berhasil mendaftarkan vendor " + nama;
  } catch(e) { throw new Error(e.message); }
}

function updateVendorAccess(nama, username, listKategori) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Vendor"); var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === username.trim()) {
        var aksesString = listKategori.join(","); sheet.getRange(i + 1, 1).setValue(nama.trim()); sheet.getRange(i + 1, 4).setValue(aksesString);
        return "SUCCESS: Berhasil memperbarui akses untuk vendor " + nama;
      }
    }
    return "ERROR: Vendor tidak ditemukan.";
  } catch(e) { throw new Error(e.message); }
}

function deleteVendorAccount(username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); var sheet = ss.getSheetByName("Data_Vendor"); var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString().trim() === username.trim()) { sheet.deleteRow(i + 1); return "Berhasil menghapus vendor."; }
    }
    return "Vendor tidak ditemukan.";
  } catch(e) { throw new Error(e.message); }
}
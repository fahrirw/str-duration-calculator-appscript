function doGet() {
  return HtmlService.createHtmlOutputFromFile('str')
    .setTitle('Form STR Dinamis');
}

// Fungsi untuk popup form saat gambar diklik
function bukaFormSTR() {
  const html = HtmlService.createHtmlOutputFromFile('str')
    .setWidth(500)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Form STR Dinamis');
}

function getDurasiSTR(clusterName, inputDateStr) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STR Simulation");
  const dataRange = sheet.getRange("B5:G").getValues(); // tanpa header

  const inputDate = new Date(inputDateStr);
  inputDate.setHours(0, 0, 0, 0);

  const headers = sheet.getRange("B4:G4").getValues()[0];
  const clusterIndex = headers.findIndex(h => h.toString().toUpperCase() === clusterName.toUpperCase());

  if (clusterIndex === -1) {
    throw new Error("Cluster tidak ditemukan: " + clusterName);
  }

  let tanggalTerbaru = null;
  let durasiTerbaru = null;

  for (let i = 0; i < dataRange.length; i++) {
    const rowTanggalRaw = dataRange[i][0];

    // Coba parse tanggal
    const tanggal = (rowTanggalRaw instanceof Date)
      ? new Date(rowTanggalRaw)
      : new Date(rowTanggalRaw);

    tanggal.setHours(0, 0, 0, 0);

    if (isNaN(tanggal)) continue; // skip jika invalid

    if (tanggal <= inputDate) {
      if (!tanggalTerbaru || tanggal > tanggalTerbaru) {
        tanggalTerbaru = tanggal;
        durasiTerbaru = dataRange[i][clusterIndex];
      }
    }
  }

  if (durasiTerbaru === null) {
    throw new Error("Durasi tidak ditemukan untuk tanggal input: " + inputDateStr);
  }

  return durasiTerbaru;
}


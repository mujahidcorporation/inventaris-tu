function doGet() {
  initSheet();

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('Inventaris Tata Usaha');
}

function initSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let stokSheet = ss.getSheetByName("Stok");

  if (!stokSheet) {

    stokSheet = ss.insertSheet("Stok");

    stokSheet.getRange("A1:D1").setValues([
      ["Kode Barang", "Nama Barang", "Satuan", "Stok"]
    ]);

    stokSheet.getRange("A1:D1")
      .setBackground("#2563eb")
      .setFontColor("white")
      .setFontWeight("bold");

    stokSheet.setFrozenRows(1);
  }

  let trxSheet = ss.getSheetByName("Transaksi");

  if (!trxSheet) {

    trxSheet = ss.insertSheet("Transaksi");

    trxSheet.getRange("A1:F1").setValues([
      ["Tanggal", "Kode Barang", "Nama Barang", "Satuan", "Jenis", "Jumlah"]
    ]);

    trxSheet.getRange("A1:F1")
      .setBackground("#16a34a")
      .setFontColor("white")
      .setFontWeight("bold");

    trxSheet.setFrozenRows(1);
  }
}

function getBarang() {

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Stok");

  const data = sheet.getDataRange().getValues();

  return data.length > 1 ? data.slice(1) : [];
}

function simpanTransaksi(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stokSheet = ss.getSheetByName("Stok");
  const trxSheet = ss.getSheetByName("Transaksi");

  const rows = stokSheet.getDataRange().getValues();

  let ditemukan = false;

  for (let i = 1; i < rows.length; i++) {

    if (rows[i][0] === data.kode) {

      ditemukan = true;

      let stok = Number(rows[i][3]);

      if (data.jenis === "Masuk") {

        stok += Number(data.jumlah);

      } else {

        if (stok < Number(data.jumlah)) {

          return {
            success: false,
            message: "Stok tidak mencukupi"
          };
        }

        stok -= Number(data.jumlah);
      }

      stokSheet.getRange(i + 1, 4).setValue(stok);

      trxSheet.appendRow([
        new Date(),
        data.kode,
        data.nama,
        data.satuan,
        data.jenis,
        data.jumlah
      ]);

      return {
        success: true,
        message: "Transaksi berhasil"
      };
    }
  }

  if (!ditemukan && data.jenis === "Masuk") {

    stokSheet.appendRow([
      data.kode,
      data.nama,
      data.satuan,
      Number(data.jumlah)
    ]);

    trxSheet.appendRow([
      new Date(),
      data.kode,
      data.nama,
      data.satuan,
      data.jenis,
      data.jumlah
    ]);

    return {
      success: true,
      message: "Barang baru berhasil ditambahkan"
    };
  }

  return {
    success: false,
    message: "Barang tidak ditemukan"
  };
}

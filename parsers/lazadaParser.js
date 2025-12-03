const xlsx = require("xlsx");

function parseLazada(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);

  if (data.length === 0) {
    throw new Error("File kosong atau tidak ada data.");
  }

  const filteredData = data.map((row) => {
    // Parsing sellerSku
    // Format: 'BARBIE 533-BIRU TOSCA-UE: 21'
    const skuRaw = row.sellerSku || "";
    const parts = skuRaw.split("-");

    let sku = "";
    let warna = "";
    let size = "";

    if (parts.length >= 3) {
      sku = parts[0].trim().toLowerCase();
      // Ambil kata pertama dari warna (misal: "BIRU TOSCA" -> "biru")
      // Jika ingin mengambil semuanya ("biru tosca"), hapus .split(" ")[0]
      warna = parts[1].trim().toLowerCase().split(" ")[0];
      size = parts[2].replace("UE:", "").trim();
    } else {
      // Fallback jika format tidak sesuai
      sku = skuRaw;
    }

    return {
      tracking_number: row.trackingCode,
      no_pesanan: row.orderNumber,
      pesanan_dibuat: row.createTime,
      skuVarian: `${sku}_${warna}_${size}`,
      sku: sku,
      warna: warna,
      size: size,
      jumlah: 1,
    };
  });

  return filteredData;
}

module.exports = { parseLazada };

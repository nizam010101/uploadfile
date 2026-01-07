const xlsx = require("xlsx");

function parseLazada(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);

  if (data.length === 0) {
    throw new Error("File kosong atau tidak ada data.");
  }

  // --- VALIDATION STEP ---
  const requiredColumns = ["trackingCode", "orderNumber", "createTime", "sellerSku", "variation"];
  const actualColumns = Object.keys(data[0]);
  const missingColumns = requiredColumns.filter(col => !actualColumns.includes(col));

  if (missingColumns.length > 0) {
    throw new Error(`Error: Format file tidak sesuai. Kolom yang tidak ditemukan untuk Lazada: ${missingColumns.join(", ")}`);
  }
  // --- END VALIDATION ---

  const filteredData = data.map((row) => {
    // 1. Ambil Data Mentah
    const skuRaw = row.sellerSku || "";
    const variationRaw = row.variation || "";

    // 2. Logika Parsing sellerSku (Ambil SKU & Warna)
    // Format: "BARBIE 533-BIRU TOSCA"
    const skuParts = skuRaw.split("-");
    
    let sku = "";
    let warna = "";

    if (skuParts.length >= 2) {
      sku = skuParts[0].trim().toLowerCase();
      // Ambil bagian kedua sebagai warna
      warna = skuParts[1].trim().toLowerCase();
    } else {
      sku = skuRaw; // Fallback jika tidak ada strip
    }

    // 3. Logika Parsing variation (Ambil Size)
    // Format: "Family Color: Blue, Size: 36" atau "Size : XL"
    // Logic: Ambil text paling kanan setelah ":"
    let size = "";
    if (variationRaw) {
      const varParts = variationRaw.split(":");
      if (varParts.length > 0) {
        // .pop() mengambil elemen terakhir dari array hasil split
        size = varParts.pop().trim();
      }
    }

    // 4. Generate ID Unik (skuVarian)
    // Gabungan: SKU + Warna + Size
    const skuVarian = [sku, warna, size]
      .filter(Boolean)
      .join("_")
      .toLowerCase();

    // 5. Format Tanggal
    let createDate = "";
    if (row.createTime) {
      const dateStr = String(row.createTime);
      const dateObj = new Date(dateStr);
      if (!isNaN(dateObj.getTime())) {
        const day = String(dateObj.getDate()).padStart(2, "0");
        const month = String(dateObj.getMonth() + 1).padStart(2, "0");
        const year = dateObj.getFullYear();
        createDate = `${year}-${month}-${day}`;
      } else {
        createDate = dateStr.split(" ")[0];
      }
    }

    return {
      no_pesanan: row.orderNumber,
      tracking_number: row.trackingCode,
      pesanan_dibuat: createDate,
      skuVarian: skuVarian,   // Contoh: "barbie 533_biru tosca_36"
      sku: sku,               // Dari sellerSku (kiri)
      warna: warna,           // Dari sellerSku (kanan)
      size: size,             // Dari variation (setelah titik dua)
      variation: variationRaw,// Data asli kolom variation
      jumlah: 1,
    };
  });

  return filteredData;
}

module.exports = { parseLazada };

const xlsx = require("xlsx");

function parseShopee(filePath) {
  const workbook = xlsx.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);

  // --- VALIDATION STEP ---
  if (data.length === 0) {
    throw new Error("The Excel sheet is empty. Nothing to process.");
  }

  const requiredColumns = [
    "tracking_number",
    "order_sn",
    "order_creation_date",
    "product_info",
  ];
  const actualColumns = Object.keys(data[0]);
  const missingColumns = requiredColumns.filter(
    (col) => !actualColumns.includes(col)
  );

  if (missingColumns.length > 0) {
    throw new Error(
      `Error: File format is incorrect. Missing required columns for Shopee: ${missingColumns.join(
        ", "
      )}`
    );
  }
  // --- END VALIDATION ---

  const finalData = [];
  data.forEach((row) => {
    const productInfoString = row.product_info;
    if (productInfoString && typeof productInfoString === "string") {
      const products = productInfoString
        .split(/\[\d+\]/)
        .map((p) => p.trim().replace("Nama Produk:", ""))
        .filter((p) => p);

      products.forEach((product) => {
        const getValue = (regex) => {
          const match = product.match(regex);
          return match && match[1] ? match[1].trim() : null;
        };

        const variasi = getValue(/Nama Variasi:([^;]+)/);
        const jumlah = getValue(/Jumlah:([^;]+)/);
        const sku = getValue(/Nomor Referensi SKU:([^;]+)/);

        let warna = null;
        let size = null;
        if (variasi) {
          const variasiParts = variasi.split(",");
          warna = variasiParts[0] ? variasiParts[0].trim().toLowerCase() : null;
          size = variasiParts[1] ? variasiParts[1].trim().toLowerCase() : null;
        }

        const skuVarian = [sku, warna, size]
          .filter(Boolean) // Menyaring nilai null atau kosong
          .join("_") // Menggabungkan dengan "_"
          .toLowerCase(); // Mengubah menjadi huruf kecil

        // Format order_creation_date to YYYY-MM-DD
        let orderDate = row.order_creation_date;
        if (orderDate) {
          const datePart = String(orderDate).split(" ")[0];
          const parts = datePart.split("-");
          // If format is YYYY-MM-DD, keep as is
          if (parts.length === 3 && parts[0].length === 4) {
            orderDate = datePart;
          } else {
            // Try DD/MM/YYYY format
            const slashParts = datePart.split("/");
            if (slashParts.length === 3) {
              orderDate = `${slashParts[2]}-${slashParts[1]}-${slashParts[0]}`;
            } else {
              orderDate = datePart.replace(/\//g, "-");
            }
          }
        }

        finalData.push({
          no_pesanan: row.order_sn,
          tracking_number: row.tracking_number,
          pesanan_dibuat: orderDate,
          skuVarian: skuVarian,
          sku: sku,
          warna: warna,
          size: size,
          jumlah: jumlah,
        });
      });
    }
  });

  return finalData;
}

module.exports = { parseShopee };

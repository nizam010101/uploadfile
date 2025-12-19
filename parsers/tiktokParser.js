const xlsx = require("xlsx");

function parseTiktok(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  // Convert to array of arrays (header: 1 gives raw array of arrays)
  // defval: "" ensures empty cells are empty strings, avoiding undefined errors
  const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  if (rows.length === 0) {
    throw new Error("File kosong");
  }

  // Get the header (first row)
  const header = rows[0];

  // Define desired columns
  const desiredColumns = [
    "Tracking ID",
    "Order ID",
    "Created Time",
    "Seller SKU",
    "Variation",
    "Quantity",
  ];

  // Find indices of desired columns
  const columnIndices = desiredColumns.map((col) => header.indexOf(col));

  // Check if all columns are found
  const missingColumns = desiredColumns.filter(
    (col, index) => columnIndices[index] === -1
  );

  if (missingColumns.length > 0) {
    throw new Error(
      `Kolom berikut tidak ditemukan: ${missingColumns.join(", ")}`
    );
  }

  // Get all data rows (skipping the header at index 0 and the description at index 1)
  // Row 0: Header
  // Row 1: Description (Skip)
  // Row 2+: Data
  const dataRows = rows.slice(2);

  // Filter data rows to keep only desired columns and create objects for better table display
  const filteredRows = dataRows.map((row) => {
    const getValue = (colName) => {
      const index = desiredColumns.indexOf(colName);
      const originalIndex = columnIndices[index];
      return originalIndex !== -1 ? row[originalIndex] : "";
    };

    const trackingId = getValue("Tracking ID");
    const orderId = getValue("Order ID");
    const createdTime = getValue("Created Time");
    const sellerSku = getValue("Seller SKU");
    const variation = getValue("Variation");
    const quantity = getValue("Quantity");

    let color = "";
    let size = "";

    if (variation) {
      const parts = variation.toString().split(",");
      color = parts[0] ? parts[0].trim() : "";
      size = parts[1] ? parts[1].trim() : "";
    }

    // Create skuVarian (combination of SKU, Color, Size in lowercase)
    const skuVarian = [sellerSku, color, size]
      .filter(Boolean)
      .join("_")
      .toLowerCase()
      .trim();

    // Format createdTime to YYYY-MM-DD
    let createdDate = "";
    if (createdTime) {
      const datePart = String(createdTime).split(" ")[0];
      // Check if already DD/MM/YYYY format - convert to YYYY-MM-DD
      const slashParts = datePart.split("/");
      if (slashParts.length === 3) {
        createdDate = `${slashParts[2]}-${slashParts[1]}-${slashParts[0]}`;
      } else {
        // Try YYYY-MM-DD format - keep as is
        const dashParts = datePart.split("-");
        if (dashParts.length === 3 && dashParts[0].length === 4) {
          createdDate = datePart;
        } else {
          createdDate = datePart;
        }
      }
    }

    return {
      no_pesanan: orderId,
      tracking_number: trackingId,
      pesanan_dibuat: createdDate,
      skuVarian: skuVarian,
      sku: sellerSku,
      warna: color,
      size: size,
      jumlah: quantity,
    };
  });

  return filteredRows;
}

module.exports = { parseTiktok };

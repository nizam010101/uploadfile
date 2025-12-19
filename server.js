const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const os = require("os");

// Import parsers
const { parseShopee } = require("./parsers/shopeeParser");
const { parseLazada } = require("./parsers/lazadaParser");
const { parseTiktok } = require("./parsers/tiktokParser");

const app = express();
const port = process.env.PORT || 3000;

// --- Multer Setup for File Uploads ---
// Use system temp dir on Vercel, and local 'uploads' folder on Windows/Local
const uploadDir = process.env.VERCEL
  ? os.tmpdir()
  : path.join(__dirname, "uploads");

if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const storage = multer.diskStorage({
  destination: function (req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function (req, file, cb) {
    cb(
      null,
      file.fieldname + "-" + Date.now() + path.extname(file.originalname)
    );
  },
});

const upload = multer({
  storage: storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB max per file
    files: 10, // Maximum 10 files at once
  },
  fileFilter: (req, file, cb) => {
    const allowedMimeTypes = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
    ];
    if (allowedMimeTypes.includes(file.mimetype) || file.originalname.endsWith(".xlsx")) {
      cb(null, true);
    } else {
      cb(new Error("Hanya file .xlsx yang diizinkan"), false);
    }
  },
});

// --- Middleware ---
app.use(express.static(path.join(__dirname, "public"))); // Serve static files from 'public' directory

// --- Routes ---
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

app.post("/upload", upload.array("file"), (req, res) => {
  const marketplace = req.body.marketplace;
  const account = req.body.account;
  const files = req.files;

  if (!files || files.length === 0) {
    return res.status(400).json({ message: "No files uploaded." });
  }

  let allProcessedData = [];
  // Initialize filesToClean with all uploaded files immediately to ensure cleanup on error
  const filesToClean = files.map((file) => file.path);

  try {
    for (const file of files) {
      let processedData;

      switch (marketplace) {
        case "shopee":
          processedData = parseShopee(file.path);
          break;
        case "lazada":
          processedData = parseLazada(file.path);
          break;
        case "tiktok":
          processedData = parseTiktok(file.path);
          break;
        default:
          throw new Error("Invalid marketplace selected.");
      }

      if (Array.isArray(processedData)) {
        allProcessedData = allProcessedData.concat(processedData);
      }
    }

    // Add Marketplace and Account to each row with correct column order
    if (Array.isArray(allProcessedData)) {
      allProcessedData = allProcessedData.map((item) => ({
        no_pesanan: item.no_pesanan,
        Marketplace: marketplace,
        Account: account,
        tracking_number: item.tracking_number,
        pesanan_dibuat: item.pesanan_dibuat,
        skuVarian: item.skuVarian,
        sku: item.sku,
        warna: item.warna,
        size: item.size,
        jumlah: item.jumlah,
      }));
    }

    res.json(allProcessedData);
  } catch (error) {
    console.error("Processing error:", error.message);
    // Sanitasi error message untuk production
    const safePatterns = [
      "Missing required columns",
      "tidak ditemukan",
      "kosong",
      "Format file tidak sesuai",
      "Invalid marketplace",
      "Hanya file .xlsx",
    ];
    const isSafeMessage = safePatterns.some((pattern) =>
      error.message.includes(pattern)
    );
    const userMessage = isSafeMessage
      ? error.message
      : "Terjadi kesalahan saat memproses file. Pastikan format file sesuai.";
    res.status(500).json({ message: userMessage });
  } finally {
    // Clean up the uploaded files
    filesToClean.forEach((filePath) => {
      fs.unlink(filePath, (err) => {
        if (err) {
          console.error("Failed to delete uploaded file:", err);
        }
      });
    });
  }
});

// --- Server Start ---
// Export the app for Vercel serverless functions
module.exports = app;

// Only start the server if running locally (not imported as a module)
if (require.main === module) {
  app.listen(port, () => {
    console.log(
      `Server is running on http://localhost:${port}\n` +
      `Please open your browser and navigate to this address.`
    );
  });
}

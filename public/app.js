let currentProcessedData = [];

document
  .getElementById("uploadForm")
  .addEventListener("submit", async function (e) {
    e.preventDefault();

    const form = e.target;
    const formData = new FormData(form);
    const loadingDiv = document.getElementById("loading");
    const outputDiv = document.getElementById("output");
    const tableContainer = document.getElementById("table-container");
    const downloadBtn = document.getElementById("downloadBtn");

    loadingDiv.classList.remove("hidden");
    outputDiv.classList.add("hidden");
    downloadBtn.classList.add("hidden"); // Hide button while processing
    tableContainer.innerHTML = "";

    try {
      const response = await fetch("/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const error = await response.json();
        throw new Error(
          error.message || "An error occurred during processing."
        );
      }

      const data = await response.json();
      currentProcessedData = data; // Store data for download
      displayData(data);
    } catch (error) {
      alert(`Error: ${error.message}`);
    } finally {
      loadingDiv.classList.add("hidden");
    }
  });

document.getElementById("downloadBtn").addEventListener("click", function () {
  if (!currentProcessedData || currentProcessedData.length === 0) {
    alert("No data to download.");
    return;
  }

  const ws = XLSX.utils.json_to_sheet(currentProcessedData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Processed Data");
  XLSX.writeFile(wb, "processed_data.xlsx");
});

function displayData(data) {
  const tableContainer = document.getElementById("table-container");
  const outputDiv = document.getElementById("output");
  const downloadBtn = document.getElementById("downloadBtn");

  if (!data || data.length === 0) {
    tableContainer.innerHTML = "<p>No data to display.</p>";
    outputDiv.classList.remove("hidden");
    return;
  }

  downloadBtn.classList.remove("hidden"); // Show download button

  const headers = Object.keys(data[0]);
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  // Create header row
  const headerRow = document.createElement("tr");
  headers.forEach((headerText) => {
    const th = document.createElement("th");
    th.textContent = headerText.replace(/_/g, " ").toUpperCase();
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  // Create data rows
  data.forEach((row) => {
    const dataRow = document.createElement("tr");
    headers.forEach((header) => {
      const td = document.createElement("td");
      td.textContent = row[header];
      dataRow.appendChild(td);
    });
    tbody.appendChild(dataRow);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  tableContainer.appendChild(table);
  outputDiv.classList.remove("hidden");
}

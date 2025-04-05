// exportExcel.js – Contains the code to export the checklist to Excel.
// Make sure to include these libraries in your HTML:
// <script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js"></script>
// <script src="https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js"></script>

const exportExcel = async () => {
  // Create a new workbook
  const workbook = new ExcelJS.Workbook();

  // Utility: add a worksheet for a given section’s task table
  const addSectionWorksheet = (sectionDiv) => {
    const sectionName = sectionDiv.getAttribute("data-section");
    // Create worksheet using the section name
    const ws = workbook.addWorksheet(sectionName);
    // Add header row
    ws.addRow(["Task", "Description", "Quantity", "Unit", "Price", "Material Cost", "Total Task Cost"]);

    // Get all rows from the section table
    const rows = sectionDiv.querySelectorAll(".section-body .row");
    rows.forEach((rowElem, index) => {
      const task = rowElem.querySelector(".task").value;
      const description = rowElem.querySelector(".description").value;
      const quantity = rowElem.querySelector(".quantity").value || 0;
      const unit = rowElem.querySelector("select").value;
      const price = rowElem.querySelector(".price").value || 0;
      const materialCost = rowElem.querySelector(".material-cost").value || 0;
      // Data row number in Excel (header is row 1)
      const excelRowNum = index + 2;
      // Add row (for now, leave the Total column blank so we can insert a formula)
      const newRow = ws.addRow([task, description, quantity, unit, price, materialCost, null]);
      // Insert formula into the "Total Task Cost" cell (assuming columns: C=3, E=5, F=6)
      newRow.getCell(7).value = { formula: `C${excelRowNum}*E${excelRowNum}+F${excelRowNum}` };
    });

    // Optionally, add some basic styling and number formatting here.
  };

  // Utility: add an attachments worksheet for a section (if there are any photos)
  const addAttachmentsWorksheet = (sectionDiv) => {
    const sectionName = sectionDiv.getAttribute("data-section");
    const photoGroups = sectionDiv.querySelectorAll(".photo-groups .task-photo-group");
    if (!photoGroups.length) return; // Skip if no attachments

    const ws = workbook.addWorksheet(`${sectionName} Attachments`);
    ws.addRow(["Task", "Attachment Type", "Attachment"]);

    photoGroups.forEach((group) => {
      const taskName = group.getAttribute("data-task");
      // For each image in the group:
      const imgs = group.querySelectorAll("img");
      imgs.forEach((img) => {
        // Add a row for this attachment
        const attRow = ws.addRow([taskName, "Photo", "Embedded Image"]);
        // Use ExcelJS to add the image if the src is a base64 data URL
        if (img.src.startsWith("data:image/")) {
          // Determine extension (e.g. "png" or "jpeg")
          const ext = img.src.substring("data:image/".length, img.src.indexOf(";"));
          const imageId = workbook.addImage({
            base64: img.src,
            extension: ext
          });
          // Place the image into the worksheet. This example positions it in column C of the current row.
          ws.addImage(imageId, {
            tl: { col: 2, row: attRow.number - 1 },
            ext: { width: 100, height: 100 }
          });
          // Adjust the row height so the image fits nicely.
          ws.getRow(attRow.number).height = 80;
        }
      });
    });
  };

  // Export Client & Project Details separately.
  const clientDetailsSection = document.querySelector('[data-section="clientDetails"]');
  if (clientDetailsSection) {
    // Create a worksheet for the client details data.
    const clientWS = workbook.addWorksheet("Client Details");
    clientWS.addRow(["Client Name", "Client Address", "OT Report Available", "Drawings Available"]);
    const clientName = document.getElementById("clientName").value;
    const clientAddress = document.getElementById("clientAddress").value;
    const otAvailable = document.getElementById("otReportAvailable").value;
    const drawingsAvailable = document.getElementById("drawingsAvailable").value;
    clientWS.addRow([clientName, clientAddress, otAvailable, drawingsAvailable]);

    // For PDF attachments in Client Details, list them in a separate worksheet.
    const pdfFiles = [];
    const otInput = document.getElementById("otReportAttachment");
    if (otInput && otInput.files.length > 0) {
      Array.from(otInput.files).forEach(file => {
        pdfFiles.push({ type: "OT Report", name: file.name });
      });
    }
    const drawingsInput = document.getElementById("drawingsAttachment");
    if (drawingsInput && drawingsInput.files.length > 0) {
      Array.from(drawingsInput.files).forEach(file => {
        pdfFiles.push({ type: "Drawings", name: file.name });
      });
    }
    if (pdfFiles.length > 0) {
      const attWS = workbook.addWorksheet("Client Attachments");
      attWS.addRow(["Attachment Type", "File Name"]);
      pdfFiles.forEach(pdf => {
        attWS.addRow([pdf.type, pdf.name]);
      });
    }
  }

  // Process all other sections (including Client Details for tasks)
  document.querySelectorAll("[data-section]").forEach((sectionDiv) => {
    // Add a worksheet for the section's checklist tasks
    addSectionWorksheet(sectionDiv);
    // And, if there are attachments (photos), add an attachments worksheet
    addAttachmentsWorksheet(sectionDiv);
  });

  // Optionally, add a worksheet for totals if needed; you could copy overall totals from the DOM.

  // Write workbook to a buffer and trigger download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  saveAs(blob, "bathroom_renovation_checklist.xlsx");
};

document.getElementById("exportExcel").addEventListener("click", exportExcel);

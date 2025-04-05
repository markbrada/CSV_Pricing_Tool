// js/excelExport.js

// Function to export the checklist data to Excel using ExcelJS.
// Each section (with a data-section attribute) becomes its own worksheet.
// Photos/attachments are added as a table within each worksheet.
const exportToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bathroom Renovation Checklist App';
  workbook.created = new Date();

  // Use a Set to track sheet names (normalized to lower case) to avoid duplicates.
  const usedSheetNames = new Set();

  // Iterate over each checklist section.
  const sections = document.querySelectorAll('[data-section]');
  sections.forEach(section => {
    // Determine a raw sheet name: use the h2 text if available; otherwise, the data-section attribute.
    const sectionTitleEl = section.querySelector('h2');
    const rawSectionName = sectionTitleEl
      ? sectionTitleEl.innerText
      : section.getAttribute('data-section') || 'Sheet';

    // Sanitize the sheet name: remove invalid characters and trim whitespace.
    let sheetName = rawSectionName.replace(/[\\\/\*\?\[\]:]/g, '').trim();
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    // Normalize for duplicate checking (Excel is case-insensitive).
    let sheetNameKey = sheetName.toLowerCase();
    const baseName = sheetName;
    let counter = 2;
    while (usedSheetNames.has(sheetNameKey) || !sheetName) {
      sheetName = `${baseName} (${counter++})`;
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
      sheetNameKey = sheetName.toLowerCase();
    }
    usedSheetNames.add(sheetNameKey);

    // Create the worksheet.
    const worksheet = workbook.addWorksheet(sheetName);

    // Define the columns.
    worksheet.columns = [
      { header: 'Task', key: 'task', width: 20 },
      { header: 'Description', key: 'description', width: 30 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit', key: 'unit', width: 10 },
      { header: 'Price', key: 'price', width: 10 },
      { header: 'Material Cost', key: 'materialCost', width: 15 },
      { header: 'Total', key: 'total', width: 10 }
    ];

    // Process task rows from the table within the section.
    const taskRows = section.querySelectorAll('table .row');
    taskRows.forEach((row, index) => {
      const task = row.querySelector('.task') ? row.querySelector('.task').value : '';
      const description = row.querySelector('.description') ? row.querySelector('.description').value : '';
      const quantity = row.querySelector('.quantity') ? row.querySelector('.quantity').value : '';
      const unit = row.querySelector('select') ? row.querySelector('select').value : '';
      const price = row.querySelector('.price') ? row.querySelector('.price').value : '';
      const materialCost = row.querySelector('.material-cost') ? row.querySelector('.material-cost').value : '';
      // Data rows start at row 2 (row 1 is the header)
      const excelRowIndex = index + 2;

      // Add a row with a formula for the Total cell.
      worksheet.addRow({
        task,
        description,
        quantity,
        unit,
        price,
        materialCost,
        total: { formula: `C${excelRowIndex}*E${excelRowIndex}+F${excelRowIndex}` }
      });
    });

    // Special handling for the Client & Project Details section.
    if (section.getAttribute('data-section') === 'clientDetails') {
      worksheet.addRow([]); // Blank row for spacing.
      const clientName = document.getElementById('clientName') ? document.getElementById('clientName').value : '';
      const clientAddress = document.getElementById('clientAddress') ? document.getElementById('clientAddress').value : '';
      worksheet.addRow(['Client Name:', clientName]);
      worksheet.addRow(['Client Address:', clientAddress]);

      // Gather PDF attachment file names for OT Report and Drawings.
      const otInput = document.getElementById('otReportAttachment');
      let otFiles = [];
      if (otInput && otInput.files.length > 0) {
        for (let file of otInput.files) {
          otFiles.push(file.name);
        }
      }
      const drawingsInput = document.getElementById('drawingsAttachment');
      let drawingFiles = [];
      if (drawingsInput && drawingsInput.files.length > 0) {
        for (let file of drawingsInput.files) {
          drawingFiles.push(file.name);
        }
      }
      worksheet.addRow(['OT Report Attachments:', otFiles.join(', ')]);
      worksheet.addRow(['Drawings Attachments:', drawingFiles.join(', ')]);
    }

    // Append photo/attachment data if available.
    const photoGroups = section.querySelector('.photo-groups');
    if (photoGroups) {
      worksheet.addRow([]); // Blank row.
      worksheet.addRow(['Photos/Attachments']);
      // Header row for photo data.
      worksheet.addRow(['Task', 'Photo Count', 'Photo Data (Base64)']);
      const groups = photoGroups.querySelectorAll('.task-photo-group');
      groups.forEach(group => {
        const taskName = group.getAttribute('data-task') || '';
        const images = group.querySelectorAll('img');
        let photoDataList = [];
        images.forEach(img => {
          // Store the data URL (for a real-world app, consider a more efficient approach).
          photoDataList.push(img.src);
        });
        const photoDataCombined = photoDataList.join('\n');
        worksheet.addRow([taskName, images.length, photoDataCombined]);
      });
    }
  });

  // Write the workbook to a buffer and trigger the download.
  workbook.xlsx.writeBuffer().then(buffer => {
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bathroom_renovation_checklist.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  }).catch(err => {
    console.error('Excel export error:', err);
  });
};

// Expose the function globally.
window.exportToExcel = exportToExcel;

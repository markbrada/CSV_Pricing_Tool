// js/excelExport.js

// This function exports the checklist data to an Excel file.
// Each section (with a data-section attribute) becomes its own worksheet.
// Photos and attachments are added in a dedicated table within each sheet.
const exportToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bathroom Renovation Checklist App';
  workbook.created = new Date();

  // Keep track of used sheet names to avoid duplicates.
  const usedSheetNames = new Set();

  // Iterate over each checklist section.
  const sections = document.querySelectorAll('[data-section]');
  sections.forEach(section => {
    // Determine the raw sheet name from the section's <h2> text or data-section attribute.
    const sectionTitleEl = section.querySelector('h2');
    const rawSectionName = sectionTitleEl ? sectionTitleEl.innerText : section.getAttribute('data-section') || 'Sheet';

    // Sanitize the sheet name (remove invalid characters, trim to 31 characters).
    let sheetName = rawSectionName.replace(/[\\\/\*\?\[\]:]/g, '').trim();
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    // Append numeric suffix if needed to ensure uniqueness.
    const baseName = sheetName;
    let counter = 2;
    while (usedSheetNames.has(sheetName) || !sheetName) {
      sheetName = `${baseName} (${counter++})`;
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
    }
    usedSheetNames.add(sheetName);

    // Create the worksheet.
    const worksheet = workbook.addWorksheet(sheetName);

    // Define columns for the tasks.
    worksheet.columns = [
      { header: 'Task', key: 'task', width: 20 },
      { header: 'Description', key: 'description', width: 30 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit', key: 'unit', width: 10 },
      { header: 'Price', key: 'price', width: 10 },
      { header: 'Material Cost', key: 'materialCost', width: 15 },
      { header: 'Total', key: 'total', width: 10 }
    ];

    // Get the task rows (each row in the table with class "row").
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

      // Add a row; the 'total' cell will have a formula (Quantity * Price + Material Cost)
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

      // Get PDF attachment file names for OT Report and Drawings.
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

    // Append photo/attachment data if available in this section.
    const photoGroups = section.querySelector('.photo-groups');
    if (photoGroups) {
      worksheet.addRow([]); // Blank row.
      worksheet.addRow(['Photos/Attachments']);
      // Add a header row for photo data.
      worksheet.addRow(['Task', 'Photo Count', 'Photo Data (Base64)']);
      const groups = photoGroups.querySelectorAll('.task-photo-group');
      groups.forEach(group => {
        const taskName = group.getAttribute('data-task') || '';
        const images = group.querySelectorAll('img');
        let photoDataList = [];
        images.forEach(img => {
          // For simplicity, we store the data URL.
          photoDataList.push(img.src);
        });
        const photoDataCombined = photoDataList.join('\n');
        worksheet.addRow([taskName, images.length, photoDataCombined]);
      });
    }
  });

  // Generate the Excel file and trigger the download.
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

// Expose exportToExcel to the global scope so that it can be called by your HTML.
window.exportToExcel = exportToExcel;

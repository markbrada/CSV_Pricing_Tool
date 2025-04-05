// js/excelExport.js

// Function to export checklist data to an Excel file using ExcelJS.
// This version groups sections by their data-section attribute (case-insensitive)
// so that duplicate sections (e.g. two "plumbing" sections) are merged into one worksheet.
const exportToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bathroom Renovation Checklist App';
  workbook.created = new Date();

  // Group sections by their data-section attribute (normalized to lowercase)
  const sectionMap = new Map();
  // Limit the search to the checklist form to avoid picking up unintended elements
  document.querySelectorAll('#checklistForm div[data-section]').forEach(section => {
    const key = section.getAttribute('data-section').toLowerCase();
    if (!sectionMap.has(key)) {
      sectionMap.set(key, []);
    }
    sectionMap.get(key).push(section);
  });

  // For each group, create one worksheet
  sectionMap.forEach((sectionsArray, key) => {
    // Use the first section's <h2> text (if available) as the base sheet name
    let baseSheetName = sectionsArray[0].querySelector('h2')
      ? sectionsArray[0].querySelector('h2').innerText.trim()
      : key;
    // Sanitize: remove invalid characters and trim to 31 characters
    let sheetName = baseSheetName.replace(/[\\\/\*\?\[\]:]/g, '').trim();
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);

    // Create worksheet for this section group
    const worksheet = workbook.addWorksheet(sheetName);

    // Define columns for task data
    worksheet.columns = [
      { header: 'Task', key: 'task', width: 20 },
      { header: 'Description', key: 'description', width: 30 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit', key: 'unit', width: 10 },
      { header: 'Price', key: 'price', width: 10 },
      { header: 'Material Cost', key: 'materialCost', width: 15 },
      { header: 'Total', key: 'total', width: 10 }
    ];

    let rowIndex = 2; // Row index for data rows (header is row 1)

    // Loop over all sections in this group and append their tasks
    sectionsArray.forEach(section => {
      const taskRows = section.querySelectorAll('table .row');
      taskRows.forEach(row => {
        const task = row.querySelector('.task') ? row.querySelector('.task').value : '';
        const description = row.querySelector('.description') ? row.querySelector('.description').value : '';
        const quantity = row.querySelector('.quantity') ? row.querySelector('.quantity').value : '';
        const unit = row.querySelector('select') ? row.querySelector('select').value : '';
        const price = row.querySelector('.price') ? row.querySelector('.price').value : '';
        const materialCost = row.querySelector('.material-cost') ? row.querySelector('.material-cost').value : '';

        // Add the row with a formula for the Total cell.
        worksheet.addRow({
          task,
          description,
          quantity,
          unit,
          price,
          materialCost,
          total: { formula: `C${rowIndex}*E${rowIndex}+F${rowIndex}` }
        });
        rowIndex++;
      });

      // If this is the Client & Project Details section, add client info.
      if (section.getAttribute('data-section').toLowerCase() === 'clientdetails') {
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
        // Header row for photo data.
        worksheet.addRow(['Task', 'Photo Count', 'Photo Data (Base64)']);
        const groups = photoGroups.querySelectorAll('.task-photo-group');
        groups.forEach(group => {
          const taskName = group.getAttribute('data-task') || '';
          const images = group.querySelectorAll('img');
          let photoDataList = [];
          images.forEach(img => {
            // Include the data URL; for large images, consider storing a reference instead.
            photoDataList.push(img.src);
          });
          const photoDataCombined = photoDataList.join('\n');
          worksheet.addRow([taskName, images.length, photoDataCombined]);
        });
      }
    });
  });

  // Write workbook to a buffer and trigger the file download.
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

// Expose exportToExcel to the global scope.
window.exportToExcel = exportToExcel;

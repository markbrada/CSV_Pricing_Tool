// js/excelExport.js

// Function to export the checklist data to Excel using ExcelJS
const exportToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bathroom Renovation Checklist App';
  workbook.created = new Date();

  // Iterate over each section (each with data-section attribute)
  const sections = document.querySelectorAll('[data-section]');
  sections.forEach(section => {
    // Use the h2 text if available as the sheet name (sanitize for Excel)
    const sectionTitleEl = section.querySelector('h2');
    const sectionName = sectionTitleEl ? sectionTitleEl.innerText : section.getAttribute('data-section');
    let sheetName = sectionName.replace(/[\\\/\*\?\[\]]/g, '');
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
    if (!sheetName) sheetName = 'Sheet';
    const worksheet = workbook.addWorksheet(sheetName);

    // Define columns for tasks
    worksheet.columns = [
      { header: 'Task', key: 'task', width: 20 },
      { header: 'Description', key: 'description', width: 30 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit', key: 'unit', width: 10 },
      { header: 'Price', key: 'price', width: 10 },
      { header: 'Material Cost', key: 'materialCost', width: 15 },
      { header: 'Total', key: 'total', width: 10 }
    ];

    // Get task rows from the table inside this section
    const taskRows = section.querySelectorAll('table .row');
    taskRows.forEach((row, index) => {
      const task = row.querySelector('.task') ? row.querySelector('.task').value : '';
      const description = row.querySelector('.description') ? row.querySelector('.description').value : '';
      const quantity = row.querySelector('.quantity') ? row.querySelector('.quantity').value : '';
      const unit = row.querySelector('select') ? row.querySelector('select').value : '';
      const price = row.querySelector('.price') ? row.querySelector('.price').value : '';
      const materialCost = row.querySelector('.material-cost') ? row.querySelector('.material-cost').value : '';
      // Calculate formula row index (header is row 1, so first data row is 2)
      const excelRowIndex = index + 2;
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

    // Special handling for Client & Project Details section
    if (section.getAttribute('data-section') === 'clientDetails') {
      worksheet.addRow([]);
      const clientName = document.getElementById('clientName') ? document.getElementById('clientName').value : '';
      const clientAddress = document.getElementById('clientAddress') ? document.getElementById('clientAddress').value : '';
      worksheet.addRow(['Client Name:', clientName]);
      worksheet.addRow(['Client Address:', clientAddress]);

      // Get PDF attachment file names for OT Report and Drawings
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

    // Append photo/attachment data if available in this section
    const photoGroups = section.querySelector('.photo-groups');
    if (photoGroups) {
      worksheet.addRow([]);
      worksheet.addRow(['Photos/Attachments']);
      // Add header row for photo data
      worksheet.addRow(['Task', 'Photo Count', 'Photo Data (Base64)']);
      const groups = photoGroups.querySelectorAll('.task-photo-group');
      groups.forEach(group => {
        const taskName = group.getAttribute('data-task') || '';
        const images = group.querySelectorAll('img');
        let photoDataList = [];
        images.forEach(img => {
          // Here we simply include the data URL text; in a real-world scenario, you might
          // store a shorter link or a note if the images are too large.
          photoDataList.push(img.src);
        });
        const photoDataCombined = photoDataList.join('\n');
        worksheet.addRow([taskName, images.length, photoDataCombined]);
      });
    }
  });

  // Generate and trigger download of the Excel file
  workbook.xlsx.writeBuffer().then(buffer => {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bathroom_renovation_checklist.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  });
};

// js/excelExport.js

// This function exports your checklist to an Excel file.
// - The main "Checklist" worksheet contains all tasks (with a column for Section).
// - For each section that has photos, a separate worksheet ("Photos - [Section]")
//   is created, listing the photos grouped by task.
const exportToExcel = () => {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Bathroom Renovation Checklist App';
  workbook.created = new Date();

  // ========================
  // Create Main Checklist Sheet
  // ========================
  const taskSheet = workbook.addWorksheet('Checklist');
  taskSheet.columns = [
    { header: 'Section', key: 'section', width: 15 },
    { header: 'Task', key: 'task', width: 20 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'Unit', key: 'unit', width: 10 },
    { header: 'Price', key: 'price', width: 10 },
    { header: 'Material Cost', key: 'materialCost', width: 15 },
    { header: 'Total', key: 'total', width: 10 }
  ];

  let taskRowIndex = 2; // Data rows start here (row 1 is header)

  // Object to store photo data per section
  const photoDataBySection = {};

  // ========================
  // Iterate over all sections (tasks)
  // ========================
  // Limit to sections inside your checklist form
  const sections = document.querySelectorAll('#checklistForm div[data-section]');
  sections.forEach(section => {
    const sectionKey = section.getAttribute('data-section');
    // Use the h2 text if available as section name; otherwise use the key.
    const sectionName = section.querySelector('h2') ? section.querySelector('h2').innerText.trim() : sectionKey;

    // Process table rows (tasks)
    const taskRows = section.querySelectorAll('table .row');
    taskRows.forEach(row => {
      const task = row.querySelector('.task') ? row.querySelector('.task').value : '';
      const description = row.querySelector('.description') ? row.querySelector('.description').value : '';
      const quantity = row.querySelector('.quantity') ? row.querySelector('.quantity').value : '';
      const unit = row.querySelector('select') ? row.querySelector('select').value : '';
      const price = row.querySelector('.price') ? row.querySelector('.price').value : '';
      const materialCost = row.querySelector('.material-cost') ? row.querySelector('.material-cost').value : '';
      
      // Add a row to the Checklist sheet. We include a formula for Total:
      // Assuming Quantity is column D, Price is column F, and Material Cost is column G.
      taskSheet.addRow({
        section: sectionName,
        task,
        description,
        quantity,
        unit,
        price,
        materialCost,
        total: { formula: `D${taskRowIndex}*F${taskRowIndex}+G${taskRowIndex}` }
      });
      taskRowIndex++;
    });

    // Special handling for Client & Project Details section
    if (sectionKey.toLowerCase() === 'clientdetails') {
      taskSheet.addRow([]);
      const clientName = document.getElementById('clientName') ? document.getElementById('clientName').value : '';
      const clientAddress = document.getElementById('clientAddress') ? document.getElementById('clientAddress').value : '';
      taskSheet.addRow(['Client Name:', clientName]);
      taskSheet.addRow(['Client Address:', clientAddress]);
      taskRowIndex += 3;
      
      // Attachments for OT Report and Drawings
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
      taskSheet.addRow(['OT Report Attachments:', otFiles.join(', ')]);
      taskSheet.addRow(['Drawings Attachments:', drawingFiles.join(', ')]);
      taskRowIndex += 2;
    }

    // ========================
    // Process Photo Groups for this section
    // ========================
    // Look for elements with class "photo-groups" within the section
    const photoGroupContainers = section.querySelectorAll('.photo-groups');
    if (photoGroupContainers.length > 0) {
      if (!photoDataBySection[sectionName]) {
        photoDataBySection[sectionName] = [];
      }
      photoGroupContainers.forEach(container => {
        // Each group should have a data-task attribute indicating the task name
        const groups = container.querySelectorAll('.task-photo-group');
        groups.forEach(group => {
          const taskName = group.getAttribute('data-task') || '';
          const images = group.querySelectorAll('img');
          let photoDataList = [];
          images.forEach(img => {
            photoDataList.push(img.src); // data URL
          });
          photoDataBySection[sectionName].push({
            task: taskName,
            photoCount: images.length,
            photoData: photoDataList.join('\n')
          });
        });
      });
    }
  });

  // ========================
  // Create separate worksheets for photo data per section
  // ========================
  for (const sectionName in photoDataBySection) {
    // Worksheet name: "Photos - [SectionName]"; sanitize and limit length
    let sheetName = `Photos - ${sectionName}`.replace(/[\\\/\*\?\[\]:]/g, '').trim();
    if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
    const photoSheet = workbook.addWorksheet(sheetName);
    photoSheet.columns = [
      { header: 'Task', key: 'task', width: 20 },
      { header: 'Photo Count', key: 'photoCount', width: 15 },
      { header: 'Photo Data (Base64)', key: 'photoData', width: 50 }
    ];
    photoDataBySection[sectionName].forEach(photoRow => {
      photoSheet.addRow(photoRow);
    });
  }

  // ========================
  // Generate and trigger download of the Excel file.
  // ========================
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

window.exportToExcel = exportToExcel;

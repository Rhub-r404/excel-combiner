document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('file-input');
    const rowDeleteInput = document.getElementById('row-delete-input');
    const columnFilterToggle = document.getElementById('column-filter-toggle');
    const columnNumberInput = document.getElementById('column-number-input');
    const previewsContainer = document.getElementById('previews');
    const downloadBtn = document.getElementById('download-btn');
  
    let filesData = [];
  
    // Event Listeners
    fileInput.addEventListener('change', handleFiles);
    rowDeleteInput.addEventListener('input', processFiles);
    columnFilterToggle.addEventListener('change', processFiles);
    columnNumberInput.addEventListener('input', processFiles);
    downloadBtn.addEventListener('click', downloadCombinedFile);
  
    function handleFiles(event) {
      const files = Array.from(event.target.files);
      filesData = files.map(file => ({ file, data: null }));
      processFiles();
    }
  
    async function processFiles() {
      const numRowsToDelete = parseInt(rowDeleteInput.value, 10) || 0;
      const columnFilterEnabled = columnFilterToggle.checked;
      const columnNumber = parseInt(columnNumberInput.value, 10) || 1;
  
      previewsContainer.innerHTML = '';
  
      for (let fileObj of filesData) {
        try {
          const data = await readExcelFile(fileObj.file);
          let processedData = data.slice(numRowsToDelete);
  
          if (columnFilterEnabled) {
            processedData = processedData.filter(row => {
              const cellValue = row[columnNumber - 1];
              return cellValue !== undefined && cellValue !== '';
            });
          }
  
          fileObj.data = processedData;
          displayPreview(fileObj.file.name, processedData);
        } catch (error) {
          console.error(`Error processing file ${fileObj.file.name}:`, error);
          alert(`An error occurred while processing ${fileObj.file.name}.`);
        }
      }
    }
  
    function readExcelFile(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const workbook = XLSX.read(e.target.result, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            resolve(jsonData);
          } catch (error) {
            reject(error);
          }
        };
        reader.onerror = reject;
        reader.readAsBinaryString(file);
      });
    }
  
    function displayPreview(fileName, data) {
      const table = document.createElement('table');
      table.className = 'table-auto w-full text-sm text-left';
  
      // Determine the maximum number of columns
      const maxColumns = data.reduce((max, row) => Math.max(max, row.length), 0);
  
      // Create header row with column numbers
      const headerRow = document.createElement('tr');
      for (let i = 1; i <= maxColumns; i++) {
        const th = document.createElement('th');
        th.className = 'border px-2 py-1 bg-gray-100';
        th.textContent = `Column ${i}`;
        headerRow.appendChild(th);
      }
      table.appendChild(headerRow);
  
      // Populate the table with data rows
      data.forEach((row) => {
        const tr = document.createElement('tr');
        for (let i = 0; i < maxColumns; i++) {
          const td = document.createElement('td');
          td.className = 'border px-2 py-1';
          td.textContent = row[i] !== undefined ? row[i] : '';
          tr.appendChild(td);
        }
        table.appendChild(tr);
      });
  
      const tableContainer = document.createElement('div');
      tableContainer.className = 'overflow-x-auto';
      tableContainer.appendChild(table);
  
      const container = document.createElement('div');
      container.className = 'border p-4';
      const title = document.createElement('h2');
      title.className = 'text-xl font-bold mb-2';
      title.textContent = fileName;
      container.appendChild(title);
      container.appendChild(tableContainer);
  
      previewsContainer.appendChild(container);
    }
  
    function downloadCombinedFile() {
      if (filesData.length === 0) {
        alert('Please upload and process files before downloading.');
        return;
      }
  
      const combinedData = filesData.reduce((acc, fileObj) => {
        return acc.concat(fileObj.data);
      }, []);
  
      const worksheet = XLSX.utils.aoa_to_sheet(combinedData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Combined');
  
      XLSX.writeFile(workbook, 'combined.xlsx');
    }
  });
  
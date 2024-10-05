let data = [];
let filteredData = [];

// Function to load and display the Excel sheet initially
async function loadExcelSheet(fileUrl) {
    try {
        const response = await fetch(fileUrl);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        data = XLSX.utils.sheet_to_json(sheet, { defval: null });
        filteredData = [...data];

        displaySheet(filteredData);
    } catch (error) {
        console.error("Error loading Excel sheet:", error);
    }
}

// Function to display the Excel sheet as an HTML table
function displaySheet(sheetData) {
    const sheetContentDiv = document.getElementById('sheet-content');
    sheetContentDiv.innerHTML = '';

    if (sheetData.length === 0) {
        sheetContentDiv.innerHTML = '<p>No data available</p>';
        return;
    }

    const table = document.createElement('table');
    const headerRow = document.createElement('tr');
    Object.keys(sheetData[0]).forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    table.appendChild(headerRow);

    sheetData.forEach(row => {
        const tr = document.createElement('tr');
        Object.values(row).forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell === null ? 'NULL' : cell;
            tr.appendChild(td);
        });
        table.appendChild(tr);
    });

    sheetContentDiv.appendChild(table);
}

// Apply operation logic remains the same...

// Function to download filtered data as different formats
function downloadExcel() {
    const filename = document.getElementById('filename').value.trim() || 'download';
    const format = document.getElementById('file-format').value;

    if (format === 'xlsx' || format === 'csv') {
        let worksheet = XLSX.utils.json_to_sheet(filteredData);
        let workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Filtered Data');

        if (format === 'xlsx') {
            XLSX.writeFile(workbook, `${filename}.xlsx`);
        } else {
            XLSX.writeFile(workbook, `${filename}.csv`);
        }
    } else if (format === 'jpg' || format === 'jpeg') {
        html2canvas(document.getElementById('sheet-content')).then(canvas => {
            const link = document.createElement('a');
            link.download = `${filename}.${format}`;
            link.href = canvas.toDataURL(`image/${format}`);
            link.click();
        });
    } else if (format === 'pdf') {
        const doc = new jsPDF();
        doc.autoTable({ html: '#sheet-content table' });
        doc.save(`${filename}.pdf`);
    }
}

// Event listeners for apply operation and download modal
document.getElementById('apply-operation').addEventListener('click', applyOperation);
document.getElementById('download-button').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'flex';
});
document.getElementById('confirm-download').addEventListener('click', downloadExcel);
document.getElementById('close-modal').addEventListener('click', () => {
    document.getElementById('download-modal').style.display = 'none';
});

// Load the Excel file on page load
window.onload = function() {
    loadExcelSheet('example.xlsx');
};

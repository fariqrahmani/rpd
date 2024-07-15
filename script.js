let workbook;

document.getElementById('upload').addEventListener('change', (event) => {
  const file = event.target.files[0];
  
  if (file && file.name.split('.').pop().toLowerCase() !== 'xlsx') {
    alert('Please upload a file with .xlsx extension');
    event.target.value = '';  // Reset input file
    return;
  }

  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    workbook = XLSX.read(data, { type: 'array' });
  };
  reader.readAsArrayBuffer(file);
});

function processExcel() {
  if (!workbook) return alert('Please upload a file first');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  const headerRow = json[0];
  const namaRekeningIndex = headerRow.indexOf("NAMA REKENING");

  if (namaRekeningIndex === -1) return alert('Kolom "NAMA REKENING" tidak ditemukan');

  const processedData = json.slice(1).map(row => ({
    "NAMA REKENING": row[namaRekeningIndex]
  }));

  const newSheet = XLSX.utils.json_to_sheet(processedData);
  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'ProcessedData');
  const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
  saveAs(new Blob([wbout], { type: 'application/octet-stream' }), 'processed_data.xlsx');
}
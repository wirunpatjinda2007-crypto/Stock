let stockData = [];

// 1. อ่านไฟล์ Excel
document.getElementById('fileUpload').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        stockData = XLSX.utils.sheet_to_json(firstSheet);
        renderTable();
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

// 2. แสดงผลในตาราง
function renderTable() {
    const tbody = document.querySelector('#stockTable tbody');
    tbody.innerHTML = '';
    stockData.forEach(item => {
        tbody.innerHTML += `<tr>
            <td>${item.ID || ''}</td>
            <td>${item['ชื่อสินค้า'] || ''}</td>
            <td>${item['จำนวน'] || 0}</td>
        </tr>`;
    });
}

// 3. ดาวน์โหลด Excel
function downloadExcel() {
    const ws = XLSX.utils.json_to_sheet(stockData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Stock");
    XLSX.writeFile(wb, "company_stock_export.xlsx");
}

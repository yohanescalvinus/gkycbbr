// Data storage
let tableData = [];
let currentEditingIndex = null;

// DOM Elements
const excelFileInput = document.getElementById('excelFile');
const uploadBtn = document.getElementById('uploadBtn');
const fileNameSpan = document.getElementById('fileName');
const addRowBtn = document.getElementById('addRowBtn');
const exportBtn = document.getElementById('exportBtn');
const saveTemplateBtn = document.getElementById('saveTemplateBtn');
const tableBody = document.getElementById('tableBody');
const rowCountSpan = document.getElementById('rowCount');
const imageModal = document.getElementById('imageModal');
const closeModal = document.querySelector('.close');
const imagePreview = document.getElementById('imagePreview');
const changeImageBtn = document.getElementById('changeImageBtn');
const imageUpload = document.getElementById('imageUpload');

// Event Listeners
uploadBtn.addEventListener('click', () => excelFileInput.click());
excelFileInput.addEventListener('change', handleFileUpload);
addRowBtn.addEventListener('click', addNewRow);
exportBtn.addEventListener('click', exportToExcel);
saveTemplateBtn.addEventListener('click', downloadTemplate);
closeModal.addEventListener('click', () => imageModal.style.display = 'none');
changeImageBtn.addEventListener('click', () => imageUpload.click());
imageUpload.addEventListener('change', handleImageUpload);

// Tutup modal ketika klik di luar
window.addEventListener('click', (e) => {
    if (e.target === imageModal) {
        imageModal.style.display = 'none';
    }
});

// Handle file upload
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    fileNameSpan.textContent = file.name;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Ambil sheet pertama
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        
        // Convert ke JSON
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            defval: ''
        });
        
        // Proses data
        processExcelData(jsonData);
    };
    
    reader.readAsArrayBuffer(file);
}

// Process Excel data
function processExcelData(data) {
    tableData = [];
    
    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row.length === 0) continue;
        
        tableData.push({
            id: Date.now() + i,
            foto: row[0] || '',
            nama: row[1] || '',
            alamat: row[2] || '',
            tanggalLahir: row[3] || ''
        });
    }
    
    // Jika tidak ada data, tambahkan satu baris kosong
    if (tableData.length === 0) {
        addNewRow();
    } else {
        updateTable();
        enableButtons();
    }
}

// Update tabel
function updateTable() {
    tableBody.innerHTML = '';
    
    if (tableData.length === 0) {
        tableBody.innerHTML = `
            <tr class="empty-row">
                <td colspan="6">
                    <i class="fas fa-file-excel"></i>
                    <p>Tidak ada data. Tambah baris baru atau upload file Excel</p>
                </td>
            </tr>
        `;
        return;
    }
    
    tableData.forEach((item, index) => {
        const row = document.createElement('tr');
        
        // Format tanggal untuk display
        let displayDate = item.tanggalLahir;
        if (item.tanggalLahir && !isNaN(new Date(item.tanggalLahir))) {
            const date = new Date(item.tanggalLahir);
            displayDate = date.toLocaleDateString('id-ID');
        }
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>
                ${item.foto ? 
                    `<img src="${item.foto}" alt="Foto" class="photo-preview" 
                         onclick="showImageModal(${index})">` : 
                    `<button class="btn btn-primary" style="padding: 5px 10px; font-size: 12px;"
                         onclick="uploadPhoto(${index})">
                         <i class="fas fa-camera"></i> Upload
                     </button>`
                }
            </td>
            <td><input type="text" class="table-input" value="${item.nama || ''}" 
                 onchange="updateData(${index}, 'nama', this.value)"></td>
            <td><input type="text" class="table-input" value="${item.alamat || ''}" 
                 onchange="updateData(${index}, 'alamat', this.value)"></td>
            <td><input type="date" class="table-input" value="${formatDateForInput(item.tanggalLahir)}" 
                 onchange="updateData(${index}, 'tanggalLahir', this.value)"></td>
            <td class="action-cell">
                <button class="action-btn delete-btn" onclick="deleteRow(${index})">
                    <i class="fas fa-trash"></i> Hapus
                </button>
            </td>
        `;
        
        tableBody.appendChild(row);
    });
    
    rowCountSpan.textContent = `${tableData.length} data`;
}

// Format date untuk input type="date"
function formatDateForInput(dateString) {
    if (!dateString) return '';
    
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return '';
    
    return date.toISOString().split('T')[0];
}

// Tambah baris baru
function addNewRow() {
    tableData.push({
        id: Date.now(),
        foto: '',
        nama: '',
        alamat: '',
        tanggalLahir: ''
    });
    
    updateTable();
    enableButtons();
}

// Update data
function updateData(index, field, value) {
    tableData[index][field] = value;
}

// Hapus baris
function deleteRow(index) {
    if (confirm('Apakah Anda yakin ingin menghapus data ini?')) {
        tableData.splice(index, 1);
        updateTable();
        
        if (tableData.length === 0) {
            disableButtons();
        }
    }
}

// Upload foto
function uploadPhoto(index) {
    currentEditingIndex = index;
    imageUpload.click();
}

// Handle image upload
function handleImageUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    if (!file.type.startsWith('image/')) {
        alert('Harap pilih file gambar!');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        if (currentEditingIndex !== null) {
            tableData[currentEditingIndex].foto = e.target.result;
            updateTable();
        }
        imageUpload.value = '';
    };
    
    reader.readAsDataURL(file);
}

// Show image modal
function showImageModal(index) {
    currentEditingIndex = index;
    const imageUrl = tableData[index].foto;
    
    imagePreview.innerHTML = imageUrl ? 
        `<img src="${imageUrl}" alt="Foto Preview">` :
        `<p>Tidak ada foto</p>`;
    
    imageModal.style.display = 'flex';
}

// Export ke Excel
function exportToExcel() {
    // Siapkan data untuk export
    const exportData = [
        ['Foto', 'Nama', 'Alamat', 'Tanggal Lahir'] // Header
    ];
    
    tableData.forEach(item => {
        exportData.push([
            item.foto || '',
            item.nama || '',
            item.alamat || '',
            item.tanggalLahir || ''
        ]);
    });
    
    // Buat worksheet
    const worksheet = XLSX.utils.aoa_to_sheet(exportData);
    
    // Buat workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    
    // Export ke file
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    XLSX.writeFile(workbook, `data-excel-${timestamp}.xlsx`);
}

// Download template
function downloadTemplate() {
    const templateData = [
        ['Foto', 'Nama', 'Alamat', 'Tanggal Lahir'],
        ['', 'Contoh: Budi Santoso', 'Contoh: Jakarta', '2023-12-31'],
        ['', '', '', '']
    ];
    
    const worksheet = XLSX.utils.aoa_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
    
    XLSX.writeFile(workbook, 'template-data.xlsx');
}

// Enable/disable buttons
function enableButtons() {
    addRowBtn.disabled = false;
    exportBtn.disabled = false;
}

function disableButtons() {
    addRowBtn.disabled = true;
    exportBtn.disabled = true;
}

// Inisialisasi
function init() {
    // Coba load data dari localStorage (opsional)
    const savedData = localStorage.getItem('excelEditorData');
    if (savedData) {
        try {
            tableData = JSON.parse(savedData);
            if (tableData.length > 0) {
                updateTable();
                enableButtons();
                fileNameSpan.textContent = 'Data dari browser';
            }
        } catch (e) {
            console.error('Error loading saved data:', e);
        }
    }
    
    // Auto-save setiap 30 detik (opsional)
    setInterval(() => {
        if (tableData.length > 0) {
            localStorage.setItem('excelEditorData', JSON.stringify(tableData));
        }
    }, 30000);
}

// Jalankan inisialisasi
init();

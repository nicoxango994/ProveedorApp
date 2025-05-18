// Variables globales
let processedData = null;

// Configurar eventos
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('fileInput').addEventListener('change', handleFile);
    document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
    document.getElementById('exportJsonBtn').addEventListener('click', exportToJson);
    document.getElementById('exportCsvBtn').addEventListener('click', exportToCsv);
    setupDragAndDrop();
});

function setupDragAndDrop() {
    const dropArea = document.querySelector('.upload-section');

    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropArea.addEventListener(eventName, highlight, false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropArea.addEventListener(eventName, unhighlight, false);
    });

    dropArea.addEventListener('drop', handleDrop, false);
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight() {
    const dropArea = document.querySelector('.upload-section');
    dropArea.style.borderColor = '#27ae60';
    dropArea.style.backgroundColor = '#e8f5e9';
}

function unhighlight() {
    const dropArea = document.querySelector('.upload-section');
    dropArea.style.borderColor = '#3498db';
    dropArea.style.backgroundColor = '#f8fafc';
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    document.getElementById('fileInput').files = files;
    handleFile({ target: document.getElementById('fileInput') });
}

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        processExcel(workbook);
    };
    reader.readAsArrayBuffer(file);
}

function processExcel(workbook) {
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const jsonData = XLSX.utils.sheet_to_json(firstSheet);
    
    // Procesar los datos
    const result = {};
    let grandTotal = 0;
    let providerCount = 0;
    
    jsonData.forEach(row => {
        const providerId = row.proveedor.toString();
        const folder = row.carpeta || 'Sin Carpeta';
        let balance = row.Saldo;
        
        // Convertir a número si es string (con formato de moneda)
        if (typeof balance === 'string') {
            balance = parseFloat(balance.replace(/[^\d.-]/g, '')) || 0;
        }
        
        if (!result[providerId]) {
            result[providerId] = {
                total: 0,
                folders: {}
            };
            providerCount++;
        }
        
        if (!result[providerId].folders[folder]) {
            result[providerId].folders[folder] = [];
        }
        
        // Agregar pago
        result[providerId].folders[folder].push(balance);
        result[providerId].total += balance;
        grandTotal += balance;
    });
    
    // Ordenar por ID de proveedor
    const sortedResult = {};
    Object.keys(result).sort().forEach(key => {
        sortedResult[key] = result[key];
        
        // Ordenar carpetas por nombre
        const folders = sortedResult[key].folders;
        const sortedFolders = {};
        Object.keys(folders).sort().forEach(folderKey => {
            sortedFolders[folderKey] = folders[folderKey];
        });
        sortedResult[key].folders = sortedFolders;
    });
    
    // Almacenar datos para exportación
    processedData = {
        rawData: jsonData,
        groupedData: sortedResult,
        summary: {
            grandTotal,
            providerCount
        }
    };
    
    // Habilitar botones de exportación
    document.getElementById('exportExcelBtn').disabled = false;
    document.getElementById('exportJsonBtn').disabled = false;
    document.getElementById('exportCsvBtn').disabled = false;
    
    // Mostrar resultados
    displayResults(sortedResult, grandTotal, providerCount);
}

function displayResults(data, grandTotal, providerCount) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';
    
    if (!data || Object.keys(data).length === 0) {
        resultsDiv.innerHTML = '<div class="no-data">No se encontraron datos válidos en el archivo.</div>';
        return;
    }
    
    // Resumen general
    const summary = document.createElement('div');
    summary.className = 'summary';
    summary.innerHTML = `
        <h3>Resumen General</h3>
        <p>Proveedores procesados: <strong>${providerCount}</strong></p>
        <p>Total general: <span class="${grandTotal < 0 ? 'negative' : 'positive'}">${formatCurrency(grandTotal)}</span></p>
    `;
    resultsDiv.appendChild(summary);
    
    // Detalle por proveedor
    for (const providerId in data) {
        const provider = data[providerId];
        const providerCard = document.createElement('div');
        providerCard.className = 'provider-card';
        
        providerCard.innerHTML = `
            <div class="provider-header">
                <span>Proveedor ID: ${providerId}</span>
            </div>
        `;
        
        // Detalle por carpeta (solo pagos)
        for (const folderName in provider.folders) {
            const folder = provider.folders[folderName];
            const folderCard = document.createElement('div');
            folderCard.className = 'folder-card';
            
            folderCard.innerHTML = `
                <div class="folder-header">
                    <span>Carpeta: ${folderName}</span>
                </div>
            `;
            
            // Mostrar pagos
            folder.forEach(pago => {
                const paymentItem = document.createElement('div');
                paymentItem.className = 'payment-item';
                paymentItem.innerHTML = `
                    <span>Monto</span>
                    <span class="${pago < 0 ? 'negative' : 'positive'}">${formatCurrency(pago)}</span>
                `;
                folderCard.appendChild(paymentItem);
            });
            
            providerCard.appendChild(folderCard);
        }
        
        // Total del proveedor
        const providerTotal = document.createElement('div');
        providerTotal.className = `provider-total ${provider.total < 0 ? 'negative' : 'positive'}`;
        providerTotal.innerHTML = `Total Proveedor: ${formatCurrency(provider.total)}`;
        providerCard.appendChild(providerTotal);
        
        resultsDiv.appendChild(providerCard);
    }
}

function formatCurrency(value) {
    return new Intl.NumberFormat('es-AR', {
        style: 'currency',
        currency: 'ARS',
        minimumFractionDigits: 2
    }).format(value);
}

function exportToJson() {
    if (!processedData) {
        alert('Primero carga un archivo Excel');
        return;
    }

    const jsonStr = JSON.stringify(processedData, null, 2);
    const blob = new Blob([jsonStr], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'resumen_proveedores.json';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function exportToExcel() {
    if (!processedData) {
        alert('Primero carga un archivo Excel');
        return;
    }

    // Crear libro de trabajo
    const wb = XLSX.utils.book_new();
    
    // Hoja con resumen detallado
    const detailedData = [];
    
    // Encabezados
    detailedData.push(['Proveedor ID', 'Carpeta', 'Monto']);
    
    // Agregar datos
    for (const providerId in processedData.groupedData) {
        const provider = processedData.groupedData[providerId];
        
        // Agregar pagos por carpeta
        for (const folderName in provider.folders) {
            const folder = provider.folders[folderName];
            
            folder.forEach(pago => {
                detailedData.push([
                    providerId,
                    folderName,
                    pago
                ]);
            });
        }
        
        // Agregar total del proveedor
        detailedData.push([
            providerId,
            'TOTAL PROVEEDOR',
            provider.total
        ]);
        
        // Separador entre proveedores
        detailedData.push(['', '', '']);
    }
    
    // Agregar total general al final
    detailedData.push([
        '',
        'TOTAL GENERAL',
        processedData.summary.grandTotal
    ]);
    
    const ws = XLSX.utils.aoa_to_sheet(detailedData);
    
    // Aplicar formatos
    const range = XLSX.utils.decode_range(ws['!ref']);
    
    // Formato de moneda para la columna Monto (columna C)
    for (let R = 1; R <= range.e.r; R++) {
        const cell = XLSX.utils.encode_col(2) + R;
        if (ws[cell]) {
            ws[cell].z = '"$"#,##0.00;[Red]\-"$"#,##0.00';
        }
    }
    
    // Estilo para los totales
    for (let R = 1; R <= range.e.r; R++) {
        const cellB = XLSX.utils.encode_col(1) + R;
        if (ws[cellB] && ws[cellB].v.includes('TOTAL')) {
            const cellC = XLSX.utils.encode_col(2) + R;
            
            // Negrita para los totales
            if (!ws[cellB].s) ws[cellB].s = {};
            if (!ws[cellC].s) ws[cellC].s = {};
            
            ws[cellB].s.font = { bold: true };
            ws[cellC].s.font = { bold: true };
            
            // Fondo gris para los totales
            ws[cellB].s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
            ws[cellC].s.fill = { patternType: "solid", fgColor: { rgb: "F2F2F2" } };
        }
    }
    
    // Autoajustar columnas
    ws['!cols'] = [
        { wch: 15 }, // Proveedor ID
        { wch: 25 }, // Carpeta
        { wch: 15 }  // Monto
    ];
    
    XLSX.utils.book_append_sheet(wb, ws, "Detalle Proveedores");
    
    // Exportar archivo
    XLSX.writeFile(wb, 'detalle_proveedores.xlsx');
}

function exportToCsv() {
    if (!processedData) {
        alert('Primero carga un archivo Excel');
        return;
    }

    // Crear contenido CSV
    let csvContent = "Proveedor ID,Carpeta,Monto\n";
    
    for (const providerId in processedData.groupedData) {
        const provider = processedData.groupedData[providerId];
        
        // Pagos por carpeta
        for (const folderName in provider.folders) {
            const folder = provider.folders[folderName];
            
            folder.forEach(pago => {
                csvContent += `"${providerId}","${folderName}",${pago}\n`;
            });
        }
        
        // Total del proveedor
        csvContent += `"${providerId}","TOTAL PROVEEDOR",${provider.total}\n\n`;
    }
    
    // Total general al final
    csvContent += `"","TOTAL GENERAL",${processedData.summary.grandTotal}`;
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'detalle_proveedores.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
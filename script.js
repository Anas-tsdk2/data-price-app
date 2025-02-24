// Variables globales
let processedData = [];
let originalFileType = '';
let originalFilters = null;
let originalStyles = {};

// Elements DOM
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileNameDiv = document.getElementById('fileName');
const statusDiv = document.getElementById('status');
const processButton = document.getElementById('processButton');
const exportButton = document.getElementById('exportButton');

// Event Listeners
processButton.addEventListener('click', processFile);
exportButton.addEventListener('click', exportData);

// Gestion du drag & drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length) {
        fileInput.files = files;
        handleFileSelect(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
        handleFileSelect(e.target.files[0]);
    }
});

function handleFileSelect(file) {
    fileNameDiv.textContent = `Fichier sélectionné : ${file.name}`;
    processButton.disabled = false;
    statusDiv.textContent = 'Prêt à traiter';
}

function processFile() {
    const file = fileInput.files[0];

    if (!file) {
        alert('Veuillez sélectionner un fichier.');
        return;
    }

    statusDiv.innerHTML = '<div class="loading"></div>Traitement en cours...';
    processButton.disabled = true;
    exportButton.disabled = true;

    originalFileType = file.name.split('.').pop().toLowerCase();

    if (originalFileType === 'csv') {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const text = e.target.result;
                const data = parseCSV(text);
                processedData = transformData(data);
                statusDiv.textContent = 'Traitement terminé avec succès';
                exportButton.disabled = false;
            } catch (error) {
                console.error('Error processing CSV file:', error);
                statusDiv.textContent = 'Erreur lors du traitement du fichier';
                processButton.disabled = false;
            }
        };
        reader.readAsText(file);
    } else if (['xlsx', 'xls'].includes(originalFileType)) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];

                originalFilters = worksheet['!autofilter'] || null;
                originalStyles = {};
                if (worksheet['!cols']) {
                    originalStyles.cols = worksheet['!cols'];
                }

                const jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
                
                const headers = jsonData[0];
                const excelData = jsonData.slice(1).map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index] || '';
                    });
                    return obj;
                });

                processedData = transformData(excelData);
                statusDiv.textContent = 'Traitement terminé avec succès';
                exportButton.disabled = false;
            } catch (error) {
                console.error('Error processing Excel file:', error);
                statusDiv.textContent = 'Erreur lors du traitement du fichier';
                processButton.disabled = false;
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('Format de fichier non supporté. Utilisez CSV ou Excel.');
        statusDiv.textContent = 'Format de fichier non supporté';
        processButton.disabled = false;
    }
}

function parseCSV(text) {
    const lines = text.split('\n');
    const headers = lines[0].split(';').map(h => h.trim());
    const result = [];

    for(let i = 1; i < lines.length; i++) {
        if(lines[i].trim() === '') continue;
        const values = lines[i].split(';').map(v => v.trim());
        const row = {};
        headers.forEach((header, index) => {
            row[header] = values[index];
        });
        result.push(row);
    }
    return result;
}

function transformData(data) {
    const result = [];
    
    data.forEach(row => {
        if (!row['App tags (raw)']) return;
        
        const appTags = row['App tags (raw)'].split(',');
        const nbrTags = appTags.length;
        const coef = 1 / nbrTags;

        appTags.forEach(tag => {
            const newRow = {...row};
            newRow['Split App tags (raw)'] = tag.trim();
            newRow['Nbr App Tags'] = nbrTags;
            newRow['Coef AppTags'] = coef.toFixed(9);
            result.push(newRow);
        });
    });

    return result;
}

function exportData() {
    if (processedData.length === 0) {
        alert('Aucune donnée à exporter.');
        return;
    }

    statusDiv.innerHTML = '<div class="loading"></div>Exportation en cours...';
    exportButton.disabled = true;

    const filename = `processed_data.${originalFileType}`;

    try {
        if (originalFileType === 'csv') {
            exportCSV(filename);
        } else if (['xlsx', 'xls'].includes(originalFileType)) {
            exportExcel(filename);
        }
        statusDiv.textContent = 'Exportation terminée avec succès';
    } catch (error) {
        console.error('Error during export:', error);
        statusDiv.textContent = 'Erreur lors de l\'exportation';
        exportButton.disabled = false;
    }
}

function exportCSV(filename) {
    const headers = Object.keys(processedData[0]);
    const csvRows = [headers.join(';')];

    processedData.forEach(row => {
        const values = headers.map(header => {
            const value = row[header] || '';
            return `"${value.toString().replace(/"/g, '""')}"`;
        });
        csvRows.push(values.join(';'));
    });

    const csvContent = csvRows.join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    downloadFile(blob, filename);
}

async function exportExcel(filename) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Processed Data');

    // Ajouter les en-têtes
    const headers = Object.keys(processedData[0]);
    worksheet.columns = headers.map(header => ({
        header: header,
        key: header,
        width: 20
    }));

    // Ajouter les données
    worksheet.addRows(processedData);

    // Styler l'en-tête
    worksheet.getRow(1).eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF366092' }
        };
        cell.font = {
            bold: true,
            color: { argb: 'FFFFFFFF' }
        };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });

    // Styler les lignes de données
    for (let i = 2; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        row.eachCell((cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: i % 2 ? 'FFFFFFFF' : 'FFF2F2F2' }
            };
            cell.border = {
                top: { style: 'thin', color: { argb: 'FFE5E5E5' } },
                left: { style: 'thin', color: { argb: 'FFE5E5E5' } },
                bottom: { style: 'thin', color: { argb: 'FFE5E5E5' } },
                right: { style: 'thin', color: { argb: 'FFE5E5E5' } }
            };
        });
    }

    // Ajouter les filtres
    worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headers.length }
    };

    // Générer le fichier
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    downloadFile(blob, filename);
}

function downloadFile(blob, filename) {
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', filename);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
}
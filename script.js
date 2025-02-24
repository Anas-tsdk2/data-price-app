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

function validateDataProcessing(originalData, processedData) {
    const originalVMCount = originalData.length;
    const coefficientSum = processedData.reduce((sum, row) => {
        return sum + parseFloat(row['Coef AppTags']);
    }, 0);
    
    const noAppTagsCount = processedData.filter(row => 
        row['Split App tags (raw)'] === 'No App Tags'
    ).length;
    
    const roundedSum = Math.round(coefficientSum * 1000000) / 1000000;
    
    // Mise à jour de l'interface
    const validationStats = document.getElementById('validationStats');
    const originalVMCountElement = document.getElementById('originalVMCount');
    const coefficientSumElement = document.getElementById('coefficientSum');
    const noAppTagsCountElement = document.getElementById('noAppTagsCount');
    const validationStatusElement = document.getElementById('validationStatus');
    
    validationStats.style.display = 'block';
    originalVMCountElement.textContent = originalVMCount;
    coefficientSumElement.textContent = roundedSum.toFixed(6);
    noAppTagsCountElement.textContent = noAppTagsCount;
    
    if (Math.abs(roundedSum - originalVMCount) > 0.000001) {
        validationStatusElement.textContent = '❌ Erreur de validation';
        validationStatusElement.className = 'validation-status error';
        throw new Error(`Erreur de validation : La somme des coefficients (${roundedSum}) ne correspond pas au nombre de VMs (${originalVMCount})`);
    } else {
        validationStatusElement.textContent = '✅ Validation réussie';
        validationStatusElement.className = 'validation-status success';
    }
    
    return true;
}

function transformData(data) {
    const result = [];
    
    data.forEach(row => {
        // Cas 1: App tags (raw) est vide
        if (!row['App tags (raw)'] || row['App tags (raw)'].trim() === '') {
            // Sous-cas 1: App tags (CT) a une valeur
            if (row['App tags (CT)'] && row['App tags (CT)'].trim() !== '') {
                const appTagsCT = row['App tags (CT)'];
                const tags = appTagsCT.split(',');
                const nbrTags = tags.length;
                const coef = 1 / nbrTags;

                tags.forEach(tag => {
                    const newRow = {...row};
                    newRow['Split App tags (raw)'] = tag.trim();
                    newRow['Nbr App Tags'] = nbrTags;
                    newRow['Coef AppTags'] = coef.toFixed(9);
                    result.push(newRow);
                });
            } 
            // Sous-cas 2: App tags (CT) est aussi vide
            else {
                const newRow = {...row};
                newRow['Split App tags (raw)'] = 'No App Tags';
                newRow['Nbr App Tags'] = 1;
                newRow['Coef AppTags'] = '1.000000000';
                result.push(newRow);
            }
        }
        // Cas 2: App tags (raw) a des valeurs
        else {
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
        }
    });

    return result;
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
                const originalData = parseCSV(text);
                processedData = transformData(originalData);
                
                try {
                    validateDataProcessing(originalData, processedData);
                    statusDiv.textContent = 'Traitement terminé avec succès - Validation OK';
                    exportButton.disabled = false;
                } catch (validationError) {
                    statusDiv.textContent = validationError.message;
                    console.error(validationError);
                    processButton.disabled = false;
                }
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
                const originalData = jsonData.slice(1).map(row => {
                    const obj = {};
                    headers.forEach((header, index) => {
                        obj[header] = row[index] || '';
                    });
                    return obj;
                });

                processedData = transformData(originalData);
                
                try {
                    validateDataProcessing(originalData, processedData);
                    statusDiv.textContent = 'Traitement terminé avec succès - Validation OK';
                    exportButton.disabled = false;
                } catch (validationError) {
                    statusDiv.textContent = validationError.message;
                    console.error(validationError);
                    processButton.disabled = false;
                }
            } catch (error) {
                console.error('Error processing Excel file:', error);
                statusDiv.textContent = 'Erreur lors du traitement du fichier';
                processButton.disabled = false;
            }
        };
        reader.readAsArrayBuffer(file);
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

    const headers = Object.keys(processedData[0]);
    worksheet.columns = headers.map(header => ({
        header: header,
        key: header,
        width: 20
    }));

    worksheet.addRows(processedData);

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

    worksheet.autoFilter = {
        from: { row: 1, column: 1 },
        to: { row: 1, column: headers.length }
    };

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
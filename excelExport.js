const XLSX = require('xlsx');
function exportToExcel(sheetName, headers, data, options = {}) {
    const {
        format = 'csv', // Default to CSV
        delimiter = ',',
        includeHeaders = true,
        customFilename = '',
        timeStamp = true,
        footer = [],
        skipRows = [],
        cellComments = {}
    } = options;

    const contentRows = [];
    let timestamp = '';
    const sheetname = sheetName || generateRandomText(6);

    if (includeHeaders) {
        contentRows.push(headers);
    }

    data.forEach((row, rowIndex) => {
        if (!skipRows.includes(rowIndex)) {
            const formattedRow = row.map(cell => {
                if (typeof cell === 'string') {
                    return `"${cell}"`;
                } else if (typeof cell === 'number') {
                    return cell.toString();
                } else {
                    return '';
                }
            });
            contentRows.push(formattedRow);
        }
    });

    if (footer.length > 0) {
        contentRows.push(footer);
    }

    if (timeStamp) {
        timestamp = new Date().toISOString().replace(/[-:.]/g, ''); // Get a timestamp in a format suitable for a file name
    }

    const baseFilename = customFilename || sheetname;
    const fileName = `${baseFilename}${timeStamp ? '_' + timestamp : ''}`;

    if (format === 'csv') {
        exportToCsv(contentRows, fileName, delimiter);
    } else if (format === 'xlsx') {
        exportToXlsx(contentRows, fileName);
    }
}

function exportToCsv(data, filename, delimiter) {
    const csvContent = data.map(row => row.map(cell => `${cell}`).join(delimiter)).join('\n');

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${filename}.csv`);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function exportToXlsx(data, filename) {
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'blob' });

    const url = URL.createObjectURL(wbout);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${filename}.xlsx`);
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}


function generateRandomText(length) {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    for (let i = 0; i < length; i++) {
        const randomIndex = Math.floor(Math.random() * characters.length);
        result += characters.charAt(randomIndex);
    }
    return result;
}

# ExcelExport
Sample Excel EXport



// Example usage
const headers = ['Name', 'Age', 'Salary'];
const data = [
    ['John', 25, 45000],
    ['Alice', 32, 52000],
    ['Bob', 28, 48000]
];

const options = {
    footer: ['Total', '', '=SUM(C2:C4)'], // Example formula in the footer
    skipRows: [1], // Skip the second row
    cellComments: {
        '2_2': 'This is a comment for cell B2',
        '3_3': 'This is a comment for cell C3'
    }
};

exportToExcel('People', headers, data, options);

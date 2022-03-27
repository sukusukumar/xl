import xlsx from 'node-xlsx';
import fs from 'fs';

// Read data from EXCEL file and convert to JSON
const workSheetsFromFile = xlsx.parse(`${__dirname}/input.xlsx`);

// Copy JSON to a variable named 'data'. Here is how data looks liks
/*
    [
        [ 'Name', 'Subject-1', 'Subject-2' ],
        [ 'Sukumar', 96, 87 ],
        [ 'Surya', 99, 97 ]
    ]
*/
const data = workSheetsFromFile[0].data;

// Insert a new student info
data.push(['Bala', 99, 99]);

/*
    Here is how data looks like after inserting
    [
        [ 'Name', 'Subject-1', 'Subject-2' ],
        [ 'Sukumar', 96, 87 ],
        [ 'Surya', 99, 97 ],
        [ 'Bala', 99, 99 ]
    ]
*/

// Add a new column 'total'
data[0].push('Total');

// Loop through and calculate total for each student

for (let i = 1; i < data.length; i++) {
    let total = 0;
    for (let j = 1; j < data[i].length; j++) {
        total += data[i][j];
    }
    data[i].push(total);
}

// All Done, now convert data back to EXCEL
const buffer = xlsx.build([{name: 'output', data: data}]);

// Write the excel to file 'output.xlsx'
fs.writeFile("output.xlsx", buffer, (err) => {
    if (err) {
        console.log(err);
    } else {
        console.log('DONE');
    }
})
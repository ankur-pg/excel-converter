const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to pad or trim the field according to the specified length
function formatField(value, maxLength) {
    let stringValue = value.toString();
    return stringValue.padEnd(maxLength, ' ');  // Pad the string to ensure fixed width
}

// Convert Excel to DAT using predefined mapping rules
function convertExcelToDat(inputFile, outputFile) {
    const workbook = XLSX.readFile(inputFile);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    // Change header option to 'A' to treat first row as headers based on Excel columns
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 'A', range: 0});

    console.log("Headers:", Object.keys(data[0]));


    console.log("Headers:", Object.keys(data[0]));

    // Define column mappings and lengths manually
    // const mappings = [
    //     { excelColumns: [], length: 1 },
    //     { excelColumns: [], length: 3 },
    //     { excelColumns: [], length: 2 },
    //     { excelColumns: ["CSEQ", "CREF"], length: 30 },
    //     { excelColumns: ["PROPERTY"], length: 10 },
    //     { excelColumns: ["BUILTUPAREA_SQFT"], length: 10 },
    //     { excelColumns: [], length: 10 },
    //     { excelColumns: [], length: 30 },
    //     { excelColumns: ["POSTALADDRESS1"], length: 40 },
    //     { excelColumns: ["POSTALADDRESS2"], length: 40 },
    //     { excelColumns: ["POSTALADDRESS3"], length: 40 },
    //     { excelColumns: ["CITY"], length: 40 },
    //     { excelColumns: ["POSTALCODE"], length: 5 },
    //     { excelColumns: ["MUKIM"], length: 30 },
    //     { excelColumns: ["DAERAH"], length: 20 },
    //     { excelColumns: [], length: 20 },
    //     { excelColumns: ["STATE"], length: 20 },
    //     { excelColumns: [], length: 10 },
    //     { excelColumns: ["LANDAREA_SQFT"], length: 10 },
    //     { excelColumns: ["PROPGROUP"], length: 50 },
    //     { excelColumns: [], length: 1 },
    //     { excelColumns: [], length: 3 },
    //     { excelColumns: [], length: 1 },
    //     { excelColumns: [], length: 8 },
    //     { excelColumns: [], length: 2 },
    //     { excelColumns: [], length: 10 },
    //     { excelColumns: [], length: 9 },
    //     { excelColumns: [], length: 2 },
    //     { excelColumns: ["PURCHASEAMOUNT", "LOANDATE"], length: 8 },
    //     { excelColumns: [], length: 26 },
    //     { excelColumns: [], length: 50 },
    //     { excelColumns: [], length: 1 },
    //     { excelColumns: [], length: 1 }
    // ];

    // const columnLengths = [1, 3, 2, 35, 30, 10, 10, 30, 40, 40, 40, 40, 5, 30, 20, 20, 20, 10, 20, 50, 1, 3, 1, 8, 2, 10, 9, 2, 8, 26, 50, 1, 1]

    const mappings = [
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 3 },
        { excelColumns: [], length: 2 },
        { excelColumns: ["A", "B"], length: 35 }, // Combined CSEQ and CREF (previously A and B)
        { excelColumns: ["C"], length: 30 }, // PROPERTY (previously C)
        { excelColumns: ["O"], length: 10 }, // BUILTUPAREA_SQFT (previously O)
        { excelColumns: [], length: 10 }, 
        { excelColumns: [], length: 30 },
        { excelColumns: ["D"], length: 40 }, // POSTALADDRESS1
        { excelColumns: ["E"], length: 40 }, // POSTALADDRESS2
        { excelColumns: ["F"], length: 40 }, // POSTALADDRESS3
        { excelColumns: ["H"], length: 40 }, // CITY (previously H)
        { excelColumns: ["G"], length: 5 }, // POSTALCODE (previously G)
        { excelColumns: ["K"], length: 30 }, // MUKIM (previously K)
        { excelColumns: ["L"], length: 20 }, // DAERAH (previously L)
        { excelColumns: [], length: 20 },
        { excelColumns: ["I"], length: 20 }, // STATE (previously I)
        { excelColumns: [], length: 10 },
        { excelColumns: ["P"], length: 20 }, // LANDAREA_SQFT (previously P)
        { excelColumns: ["M"], length: 50 }, // PROPGROUP (previously M)
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 3 },
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 8 },
        { excelColumns: [], length: 2 },
        { excelColumns: [], length: 10 },
        { excelColumns: [], length: 9 },
        { excelColumns: [], length: 2 },
        { excelColumns: ["Q", "R"], length: 8 }, // PURCHASEAMOUNT and LOANDATE (previously Q and R)
        { excelColumns: [], length: 26 },
        { excelColumns: [], length: 50 },
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 1 }
    ];

    let outputContent = "";
    data.forEach((row, index) => {
        if (index === 0) {
            return
        }
        let line = "";
        console.log('Processing for ', row)
        mappings.forEach(mapping => {
            if (mapping.excelColumns.length === 0) {
                line += formatField('', mapping.length);
            } else {
                console.log('Check mapping.excelColumns value ', mapping.excelColumns)
                let combinedValue = mapping.excelColumns.map(col => {
                    console.log(`Reading ${col}:`, row[col]); // Log each column value being read
                    return row[col] || '';
                }).join('');
                line += formatField(combinedValue, mapping.length);
            }
        });

        console.log("Formatted line:", line); // Log the formatted line
        outputContent += line + '\n';
    });

    fs.writeFileSync(outputFile, outputContent);
    console.log(`Data written to ${outputFile}`);
}

const inputFile = path.join(__dirname, 'Input_1.xlsx');
const outputFile = path.join(__dirname, 'Input_1.dat');

convertExcelToDat(inputFile, outputFile);

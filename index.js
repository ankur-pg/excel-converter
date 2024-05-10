const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Function to convert square feet to square meters
function convertSqFtToSqM(value) {
    const sqFtToSqMFactor = 0.092903;
    return value * sqFtToSqMFactor;
}

// Function to pad or trim the field according to the specified length
function formatField(value, maxLength, fillChar = ' ') {
    let stringValue = value ? value.toString() : "";
    if (stringValue.length > maxLength) {
        stringValue = stringValue.substring(0, maxLength); // Truncate to max length
    }
    return stringValue.padEnd(maxLength, fillChar); // Pad with the specified fill character
}

// Convert Excel to DAT using predefined mapping rules
function convertExcelToDat(inputFile, outputFile) {
    const workbook = XLSX.readFile(inputFile);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    // Change header option to 'A' to treat first row as headers based on Excel columns
    const data = XLSX.utils.sheet_to_json(worksheet, {header: 'A', range: 0});

    // const columnLengths = [1, 3, 2, 35, 30, 10, 10, 30, 40, 40, 40, 40, 5, 30, 20, 20, 20, 10, 20, 50, 1, 3, 1, 8, 2, 10, 9, 2, 8, 26, 50, 1, 1]

    const mappings = [
        { excelColumns: [], length: 1, dummyChar: 'D' },
        { excelColumns: [], length: 3, dummyChar: 'D' },
        { excelColumns: [], length: 2, dummyChar: 'D' },
        { excelColumns: ["A", "B"], length: 35 }, // Combined CSEQ and CREF (previously A and B) - Collateral ID
        { excelColumns: ["C"], length: 30 }, // PROPERTY (previously C) - Property Type
        { excelColumns: ["O"], length: 10, convert: convertSqFtToSqM }, // BUILTUPAREA_SQFT (previously O) - B/U (sqm)
        { excelColumns: [], length: 10 }, 
        { excelColumns: [], length: 30 },
        { excelColumns: ["D"], length: 40 }, // POSTALADDRESS1
        { excelColumns: ["E"], length: 40 }, // POSTALADDRESS2
        { excelColumns: ["F"], length: 40 }, // POSTALADDRESS3
        { excelColumns: ["H"], length: 40 }, // CITY (previously H)
        { excelColumns: ["G"], length: 5 }, // POSTALCODE (previously G)
        { excelColumns: ["K"], length: 30 }, // MUKIM (previously K)
        { excelColumns: ["L"], length: 20 }, // DAERAH (previously L) - District
        { excelColumns: [], length: 20 },
        { excelColumns: ["I"], length: 20 }, // STATE (previously I)
        { excelColumns: [], length: 10 },
        { excelColumns: ["P"], length: 20, convert: convertSqFtToSqM }, // LANDAREA_SQFT (previously P) - L/A (sqm)
        { excelColumns: ["M"], length: 50 }, // PROPGROUP (previously M) - Land Use
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 3 },
        { excelColumns: [], length: 1 },
        { excelColumns: [], length: 8 },
        { excelColumns: [], length: 2 },
        { excelColumns: [], length: 10 },
        { excelColumns: [], length: 9 },
        { excelColumns: [], length: 2 },
        { excelColumns: ["Q", "R"], length: 8 }, // PURCHASEAMOUNT and LOANDATE (previously Q and R) - Client Value
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
                line += formatField('', mapping.length, mapping.dummyChar || ' ');
            } else {
                let combinedValue = mapping.excelColumns.map(col => {
                    let value = row[col] || '';
                    if (mapping.convert) {
                        value = mapping.convert(parseFloat(value)).toFixed(2); // Convert and format the value
                    }
                    return value;
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

const inputFile = path.join(__dirname, 'Sample_6.xlsx');
const outputFile = path.join(__dirname, 'Sample_6.dat');

convertExcelToDat(inputFile, outputFile);

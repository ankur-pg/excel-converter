const XLSX = require('xlsx')
const fs = require('fs')
const path = require('path')

const convertExcelToDat = (inputFile, outputFile) => {
    if (!fs.existsSync(inputFile)) {
        console.error("File not found:", inputFile)
        return
    }

    const workbook = XLSX.readFile(inputFile)
    const sheetName = workbook.SheetNames[0]
    const worksheet = workbook.Sheets[sheetName]
    const data = XLSX.utils.sheet_to_json(worksheet)

    let formattedData = ''
    data.forEach(row => {
        formattedData += Object.values(row).join(',') + '\n'
    })

    fs.writeFileSync(outputFile, formattedData)
    console.log(`Data written to ${outputFile}`)
}

// Usage: Replace 'Test1.xlsx' and 'Test1.dat' with your actual file paths
const inputFile = path.join(__dirname, 'Test1.xlsx')
const outputFile = path.join(__dirname, 'Test1.dat')
convertExcelToDat(inputFile, outputFile)

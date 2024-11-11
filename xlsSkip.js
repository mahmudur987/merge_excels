const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Define the folder containing .xls files and output path
const folderPath = "./All files";
const outputFilePath = "./Output/Robi+airtel.xlsx";

// Ensure the output directory exists
if (!fs.existsSync("./Output")) {
  fs.mkdirSync("./Output");
}

// Array to hold all data from each .xls file
let allData = [];

// Define how many rows to skip before merging data (e.g., skip 3 rows)
const rowsToSkip = 7;

// Read each .xls file in the folder
fs.readdirSync(folderPath).forEach((file) => {
  if (path.extname(file) === ".xls") {
    const filePath = path.join(folderPath, file);
    console.log(`Reading file: ${filePath}`);

    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
    });

    if (sheetData.length > 0) {
      // Skip the first 3 rows and get the rest of the data
      const dataStartingIndex = rowsToSkip;

      // Add the data from the current sheet to the allData array
      const data = sheetData.slice(dataStartingIndex); // Skip first 3 rows and take rest of the data
      allData = allData.concat(data); // Merge data
    } else {
      console.log(`No data found in ${file}`);
    }
  }
});

// Verify if allData array has been populated
if (allData.length === 0) {
  console.log("No data found in any files.");
} else {
  console.log("Data to be written to merged file:", allData);
}

// Convert merged data to worksheet and create a new workbook
const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet(allData); // Use aoa_to_sheet for consistent ordering
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "MergedData");

// Save the merged workbook to a new file
XLSX.writeFile(newWorkbook, outputFilePath);
console.log("Files have been merged successfully!");

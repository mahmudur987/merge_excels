const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Define the folder containing .xls files and output path
const folderPath = "./All files";
const outputFilePath = "./Output/GP.xlsx";

// Ensure the output directory exists
if (!fs.existsSync("./Output")) {
  fs.mkdirSync("./Output");
}

// Array to hold all data from each .xls file
let allData = [];

// Read each .xls file in the folder
fs.readdirSync(folderPath).forEach((file) => {
  if (path.extname(file) === ".xls") {
    const filePath = path.join(folderPath, file);
    console.log(`Reading file: ${filePath}`);

    // Read the .xls file
    const workbook = XLSX.readFile(filePath);

    // Assuming data is in the first sheet
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    if (sheetData.length === 0) {
      console.log(`No data found in ${file}`);
    } else {
      console.log(`Data read from ${file}:`, sheetData);
    }

    allData = allData.concat(sheetData); // Add data to the array
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
const newWorksheet = XLSX.utils.json_to_sheet(allData);
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "MergedData");

// Save the merged workbook to a new file
XLSX.writeFile(newWorkbook, outputFilePath);
console.log("Files have been merged successfully!");

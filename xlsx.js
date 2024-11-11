const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

// Define the folder containing Excel files and output path
const folderPath = "./All files";
const outputFilePath = "./Output/GP.xlsx";

// Ensure the output directory exists
if (!fs.existsSync("./Output")) {
  fs.mkdirSync("./Output");
}

// Array to hold all data from each Excel file
let allData = [];

// Read each Excel file in the folder
fs.readdirSync(folderPath).forEach((file) => {
  if (path.extname(file) === ".xlsx") {
    const filePath = path.join(folderPath, file);
    console.log(`Reading file: ${filePath}`);
    const workbook = XLSX.readFile(filePath);

    // Check if the workbook has sheets
    if (workbook.SheetNames.length === 0) {
      console.log(`No sheets found in ${file}`);
      return;
    }

    const sheetName = workbook.SheetNames[0]; // Assuming data is in the first sheet
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    if (sheetData.length === 0) {
      console.log(`No data found in the first sheet of ${file}`);
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

const ExcelJS = require("exceljs");
const fs = require("fs");
const xcelFileName = "./en-ar.xlsx";
const sheetName = "Sheet1";
const enBundleFilePath = "./en/app-strings.json";
const arBundleFilePath = "./ar/app-strings.json";

function readToTxtFile(filePath, data) {
  // Your data to be written to the text file
  data = JSON.stringify(data);

  // The path to the text file you want to write to
  // Write the data to the text file
  fs.writeFile(filePath, data, (err) => {
    if (err) {
      console.error("Error writing to the file:", err);
    } else {
      console.log("Data has been written to the file:", filePath);
    }
  });
}

function formatStringsList(stringsList) {
  const enResult = {};
  const arResult = {};

  stringsList.forEach((inputString) => {
    // Remove special characters and replace spaces with underscores
    const formattedKey = inputString.en
      .replace(/[^\w\s]/g, "")
      .replace(/ /g, "_")
      .toLowerCase();

    enResult[formattedKey] = inputString.en;
    enResult[`@${formattedKey}`] = { description: "" };

    arResult[formattedKey] = inputString.ar;
    arResult[`@${formattedKey}`] = { description: "" };
  });

  return { enResult: enResult, arResult: arResult };
}
// Create a new workbook
const workbook = new ExcelJS.Workbook();

// Load the Excel file
workbook.xlsx
  .readFile(xcelFileName)
  .then(() => {
    // Use the workbook or specific worksheet
    const worksheet = workbook.getWorksheet(sheetName);

    // Initialize an array to store the data objects
    const data = [];

    // Iterate through rows (excluding the header row)
    for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
      const enCell = worksheet.getCell(`A${rowNumber}`); // The first column contains en values
      const arCell = worksheet.getCell(`B${rowNumber}`); // The second column contains the second lang, for me it's ar

      // Extract the text content from rich text if present
      const enValue =
        enCell.text ||
        (enCell.value && enCell.value.richText
          ? enCell.value.richText[0].text
          : "");
      const arValue =
        arCell.text ||
        (arCell.value && arCell.value.richText
          ? arCell.value.richText[0].text
          : "");

      // Create an object to store the data with "en" and "ar" keys
      const rowData = {
        en: enValue.trim(),
        ar: arValue.trim(),
      };

      // Add the rowData object to the data array
      data.push(rowData);
    }
    let translationBundle = formatStringsList(data);
    readToTxtFile(enBundleFilePath, translationBundle.enResult);
    readToTxtFile(arBundleFilePath, translationBundle.arResult);
    // console.log(translationBundle.arResult);
    // Now the 'data' array contains objects for each row
    // console.log("Data:", data);
  })
  .catch((err) => {
    console.error("Error reading the Excel file:", err);
  });

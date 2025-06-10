const ExcelJS = require('exceljs');
const path = require('path');

const OUTPUT = path.resolve('college_basic_details_merged.xlsx');
const SHEET_NAME = 'Colleges';

// ✅ List of primary key fields for composite uniqueness
const PRIMARY_KEYS = ['College ID', 'City', 'area_name'];

async function consolidateExcelFiles() {
  const mergedWorkbook = new ExcelJS.Workbook();
  const mergedSheet = mergedWorkbook.addWorksheet(SHEET_NAME);
  const seenKeys = new Set();
  let headers = [];

  async function loadAndAppend(filePath) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(SHEET_NAME);

    if (!sheet) {
      console.warn(`❌ Sheet "${SHEET_NAME}" not found in: ${filePath}`);
      return;
    }

    console.log(`📘 Reading: ${filePath}`);
    const headerRow = sheet.getRow(1);
    const currentHeaders = headerRow.values.slice(1); // Remove Excel's first empty cell

    // Initialize header and set merged sheet structure
    if (headers.length === 0) {
      headers = currentHeaders;
      mergedSheet.columns = headers.map(h => ({ header: h, key: h }));
    }

    const keyIndexes = PRIMARY_KEYS.map(pk => {
      const index = headers.indexOf(pk);
      if (index === -1) {
        throw new Error(`❌ Primary key '${pk}' not found in headers.`);
      }
      return index;
    });

    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // Skip header row
      const values = row.values.slice(1); // Remove Excel's internal offset
      const keyParts = keyIndexes.map(i => values[i] ?? '').join('|');
      
      if (!seenKeys.has(keyParts)) {
        seenKeys.add(keyParts);
        const rowObject = {};
        headers.forEach((h, i) => {
          rowObject[h] = values[i];
        });
        mergedSheet.addRow(rowObject);
      } else {
        console.log(`⚠️ Duplicate entry skipped: ${keyParts}`);
      }
    });
  }



  (async () => {
    const statesAndUTs = [
      "maharashtra",
      "karnataka",
      "tamil-nadu",
      "uttar-pradesh",
      "kerala",
      "gujarat",
      "west-bengal",
      "haryana",
      "madhya-pradesh",
      "rajasthan",
      "andhra-pradesh",
      "telangana",
      "punjab",
      "odisha",
      "uttarakhand",
      "bihar",
      "assam",
      "chhattisgarh",
      "jharkhand",
      "himachal-pradesh",
      "jammu-and-kashmir",
      "chandigarh",
      "puducherry",
      "goa",
      "meghalaya",
      "nagaland",
      "arunachal-pradesh",
      "tripura",
      "manipur",
      "sikkim",
      "mizoram",
      "andaman-and-nicobar-islands",
      "daman-and-diu",
      "dadra-and-nagar-haveli"
    ];
    for (const state of statesAndUTs) {
      const fileName = `${state}-colleges.xlsx`;
      console.log(`Processing colleges for state: ${fileName}`);  
      await loadAndAppend(fileName);
    }

      // await loadAndAppend(FILE1);

    await mergedWorkbook.xlsx.writeFile(OUTPUT);
    console.log(`✅ Merged ${seenKeys.size} unique rows into: ${OUTPUT}`);
  })();

}

consolidateExcelFiles().catch(console.error);

const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  

async function updateCollegeSheetWithMetadata(state) {
  const fileName = `${state}-colleges.xlsx`;
  const filePath = path.resolve(fileName);
  console.log("file path: ", filePath);
  const workbook = new ExcelJS.Workbook();

  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Colleges');
  // Ensure new columns are added if not present
  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values.map(h => h && h.toString().toLowerCase());

  const areaNameCol = headers.includes('area_name') ? headers.indexOf('area_name') : headerRow.cellCount + 1;
  const universityTypeCol = headers.includes('university_type') ? headers.indexOf('university_type') : headerRow.cellCount + 2;
  const nirfRankingCol = headers.includes('nirf_ranking') ? headers.indexOf('nirf_ranking') : headerRow.cellCount + 3;

  if (!headers.includes('area_name')) headerRow.getCell(areaNameCol).value = 'area_name';
  if (!headers.includes('university_type')) headerRow.getCell(universityTypeCol).value = 'university_type';
  if (!headers.includes('nirf_ranking')) headerRow.getCell(nirfRankingCol).value = 'nirf_ranking';

  headerRow.commit();

  const urlColIndex = headers.findIndex(h => h === 'url');
//   console.log("urlColIndex", urlColIndex, worksheet.rowCount);
  if (urlColIndex === -1) {
    console.error("❌ 'url' column not found in sheet.");
    return;
  }

    for (let i = 2; i <= worksheet.rowCount; i++) {
        console.log("urlColIndex", urlColIndex, worksheet.rowCount);

        const row = worksheet.getRow(i);
        const url = row.getCell(urlColIndex)?.value;
        console.log("URL: ", url);
        if (!url || typeof url !== 'string') continue;

        const apiUrl = `https://collegedunia.com/web-api/${url}`;
        try {
            const response = await axios.get(apiUrl, {
                headers: {
                'accept': '*/*',
                'api-request': 'true',
                'referer': `https://collegedunia.com/${url}`,
                'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'
                }
            });

            const data = response.data;
            // console.log("RESPONSE: ",data.basic_info.area_name, data.basic_info.university_type, data.basic_info.ranking);

            const nirfRankings = (data.basic_info.ranking || [])
                .filter(r => r.agency?.toLowerCase().includes('nirf'))

            row.getCell(areaNameCol).value = data.basic_info.area_name || '';
            row.getCell(universityTypeCol).value = data.basic_info.university_type || '';
            row.getCell(nirfRankingCol).value = JSON.stringify(nirfRankings) || '';

            row.commit();
            console.log(`✅ Updated row ${i} → ${url}`);
        } catch (err) {
            console.error(`⚠️ Error at row ${i} (${url}): ${err.message}`);
        }

        if ((i - 1) % 10 === 0) {
            console.log('⏸️ Sleeping for 5 seconds...');
            await sleep(5000);
        }
    }

  await workbook.xlsx.writeFile(filePath);
  console.log('✅ Excel file updated with university metadata.');
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
    console.log(`Processing colleges for state: ${state}`);  
    await updateCollegeSheetWithMetadata(state);
  }
  })();
  
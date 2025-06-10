const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');

const inputFile = 'college_india_list_with_logo.xlsx';
const inputSheet = 'courses_with_stream';
const outputSheet = 'course_fee_data';

function base64Encode(json) {
  return Buffer.from(JSON.stringify(json)).toString('base64');
}

async function fetchSlugData(college_id, headoneData) {
  const payload = {
    college_id: parseInt(college_id),
    entity_type: "fees",
    ...headoneData,
  };
  console.log("HEADONE DATA: ", payload);
  const encoded = base64Encode(payload);
  const headers = {
    'accept': '*/*',
    'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
    'cache-control': 'max-age=0',
    'referer': 'https://collegedunia.com/university/25446-all-india-institute-of-medical-sciences-aiims-new-delhi/courses-fees',
    'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36'
  }
  const url = `https://collegedunia.com/web-api/get-college-slug-data?data=${encoded}`;

  try {
    const response = await axios.get(url, { headers });
    // console.log("fetchSlugData", response.data);
    if (typeof response.data !== 'object') {
        console.warn(`⚠️ Invalid JSON for college_id: ${college_id}`);
        return null;
      }
  
    return response.data;
  } catch (error) {
    console.error(`Slug API error (college_id: ${college_id}):`, error.message);
    return null;
  }
}

async function fetchCourseData(course_id, college_id) {
  const payload = {
    course_id: course_id.toString(),
    college_id: college_id.toString(),
    entity_type: "fees"
  };

  const headers= {
    'accept': '*/*',
    'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
    'api-request': 'true',
    'cache-control': 'max-age=0',
    'priority': 'u=1, i',
    'referer': 'https://collegedunia.com/university/54797-amity-university-noida/courses-fees',
    'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"macOS"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36'
  }

  const encoded = base64Encode(payload);
  const url = `https://collegedunia.com/web-api/get-college-course-data?data=${encoded}`;

  try {
    const response = await axios.get(url,{
        headers
    });
    if (typeof response.data !== 'object') {
        console.warn(`⚠️ Invalid JSON for college_id: ${college_id}`);
        return null;
      }  
    // console.log("fetchCourseData", response.data);
    return response.data;
  } catch (error) {
    console.error(`Course API error (college_id: ${college_id}, course_id: ${course_id}):`, error.message);
    return null;
  }
}

async function main() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(inputFile);
    const input = workbook.getWorksheet(inputSheet);
    let output = workbook.getWorksheet(outputSheet);
    if (!output) output = workbook.addWorksheet(outputSheet);
    output.columns = [
        { header: 'college_id', key: 'college_id', width: 15 },
        { header: 'course_id', key: 'course_id', width: 15 },
        { header: 'feesData', key: 'feesData', width: 20 },
    ];

    for (let i = 2; i <= input.rowCount; i++) {
        const row = input.getRow(i);
        const course_id = row.getCell('B').value;
        const college_id = row.getCell('A').value;
        const streamsCell = row.getCell('J').value;
        const headoneCell = row.getCell('I').value;

        if (!college_id || !course_id) continue;

        let feesData = null;
        try{
            // CASE 1: If 'streams' has data, use get-college-course-data
            if (streamsCell && typeof streamsCell === 'string' && JSON.parse?.(streamsCell)?.length > 0) {
                try {
                    const streams = JSON.parse(streamsCell);
                    // console.log("streams", streams);
                    if (Array.isArray(streams) && streams.length > 0) {
                        for (const stream of streams) {
                            const streamId = stream.id;
                            const res = await fetchCourseData(streamId, college_id);
                            feesData = res?.course_fees || res?.course_fees || null;
                            if (!Object.keys(feesData)) break;
                            if (feesData !== null) {
                                output.addRow({ college_id, course_id: stream.id, feesData });
                                console.log(`✅ [${i}] course_id: ${streamId}`);
                            } else {
                                console.log(`⚠️  [${i}] course_id: ${course_id} | fees: not found`);
                            }
                    
                        }
                        
                    }
                } catch (e) {
                    console.warn(`Row ${i}: Failed to parse streams JSON.`);
                }
            }
            // CASE 2: If no valid streams, fallback to headone and use get-college-slug-data
            else if (headoneCell && typeof headoneCell === 'string' && Object.keys(JSON.parse(headoneCell)).length>0) {
                try {
                    const headoneData = JSON.parse(headoneCell);
                    // console.log("headoneData: ", headoneData);
                    const res = await fetchSlugData(college_id, headoneData);
                    feesData = res?.feesData || res?.fees_data || null;
                    if (feesData !== null) {
                        output.addRow({ college_id, course_id, feesData });
                        console.log(`✅ [${i}] course_id: ${course_id}`);
                    } else {
                        console.log(`⚠️  [${i}] course_id: ${course_id} | fees: not found`);
                    }
                    // await workbook.xlsx.writeFile(inputFile);
                } catch (e) {
                    console.warn(`Row ${i}: Failed to parse headone JSON.`);
                }
            }
            // CASE 3: If no streams or headone data, skip this row
            else {
                console.warn(`Row ${i}: No valid streams or headone data found.`);
                continue;
            }

            await workbook.xlsx.writeFile(inputFile);
            // Delay to avoid rate limits
            await new Promise(resolve => setTimeout(resolve, 1000));
        } catch (e) {
            console.warn(`Row ${i}: Failed to parse streams JSON.`);
            continue;
        }
    }

  console.log(`✅ Output written to sheet "${outputSheet}" in ${inputFile}`);
}

main().catch(err => console.error(err));

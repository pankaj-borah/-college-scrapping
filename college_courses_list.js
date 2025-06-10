const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs');
const path = require('path');

async function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function extractCourses() {
  const filePath = path.resolve('college_india_list_with_logo.xlsx');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const collegeSheet = workbook.getWorksheet('Colleges');
  if (!collegeSheet) {
    console.error("❌ 'Colleges' sheet not found.");
    return;
  }

  const collegeIdColIndex = collegeSheet.getRow(1).values.findIndex(v => v === 'College ID');
  const urlColIndex = collegeSheet.getRow(1).values.findIndex(v => v === 'url');
  
  if (collegeIdColIndex === -1) {
    console.error("❌ 'College ID' column not found.");
    return;
  }

  if (urlColIndex === -1) {
    console.error("❌ 'URL' column not found.");
    return;
  }


  const requiredHeaders = ['college_id', 'course_id', 'display_name', 'duration', 'eligibility', 'level', 'short_head', 'type', 'headone_data', 'streams' ];

  let courseSheet = workbook.getWorksheet('courses_with_stream1');
  
  if (!courseSheet) {
    // Create sheet and set headers
    courseSheet = workbook.addWorksheet('courses_with_stream1');
    courseSheet.columns = requiredHeaders.map(header => ({
      header,
      key: header
    }));
  } else {
    // Ensure headers are correct and keys are bound
    const existingHeaderRow = courseSheet.getRow(1);
    const existingHeaders = existingHeaderRow.values.slice(1); // ExcelJS row.values[0] is null
    const headersMatch = requiredHeaders.every((h, i) => existingHeaders[i] === h);
  
    if (!headersMatch) {
      // Fix header row if incorrect
      existingHeaderRow.values = [null, ...requiredHeaders];
      existingHeaderRow.commit();
    }
  
    // Re-bind keys so addRow({key: value}) works
    courseSheet.columns = requiredHeaders.map(header => ({
      header,
      key: header
    }));
  }
    
    for (let i = 2; i <= collegeSheet.rowCount; i++) {
        const collegeId = collegeSheet.getRow(i).getCell(collegeIdColIndex).value;
        const url = collegeSheet.getRow(i).getCell(urlColIndex).value;
        if (!collegeId) continue;
    
        const collegeSlug = `https://collegedunia.com/college/${url}`; // Update dummy-name if needed
        const headers = {
            'accept': '*/*',
            'accept-language': 'en-GB,en-US;q=0.9,en;q=0.8',
            'cache-control': 'max-age=0',
            'referer': collegeSlug,
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
            'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'api-request': 'true',
            'cookie': 'idu=your-cookies-here'
        };
    
        let totalCoursesFetched = 0;
    
        try {
            console.log("URL: ", url);
            // ✅ Step 1: Fetch Page 1 via direct endpoint
            const page1Res = await axios.get(
                `https://collegedunia.com/web-api/${url}/courses-fees`,
                { headers }
            );
            const data = page1Res.data; 
            //https://collegedunia.com/web-api/university/25446-all-india-institute-of-medical-sciences-aiims-new-delhi/courses-fees

    
            const courses = data?.course_data?.courses || [];
            // console.log(courses)
            for (const course of courses) {
                const course_id = course.course_id;

                courseSheet.addRow({
                    college_id: collegeId.toString(),
                    course_id: course_id.toString() || "",
                    display_name: course?.display_name?.toString() || "",
                    duration: course?.duration?.toString() || "",
                    eligibility: course?.eligibility?.toString() || "",
                    level: course?.level?.toString() || "",
                    short_head: course?.short_head?.toString() || "",
                    type: course?.type?.toString() || "",
                    headone_data: course?.headone_data ? JSON.stringify(course?.headone_data) : "",
                    streams: course?.streams ? JSON.stringify(course?.streams) : "[]",
                });
                totalCoursesFetched++;
            }
            await workbook.xlsx.writeFile(filePath);
            console.log(`✅ Page 1: Fetched & wrote ${courses.length} courses for College ID: ${collegeId}`);
        } catch (err) {
            console.error(`❌ Failed to fetch page 1 for College ID: ${collegeId} → ${err.message}`);
            continue;
        }
    
        // ✅ Step 2: Continue with paginated pages (2+)
        let coursePage = 2;
        while (true) {
            const payload = {
                id: collegeId.toString(),
                course_page: coursePage.toString(),
                page: "1",
                college: collegeId.toString(),
                tab: "courses-fees",
                source: "college"
            };
    
            const encodedPayload = Buffer.from(JSON.stringify(payload)).toString('base64');
            const url = `https://collegedunia.com/web-api/college/courses-list?data=${encodedPayload}`;
    
            try {
                const response = await axios.get(url, { headers });
                const courses = response?.data?.courses || [];
                if (courses.length === 0) break;
    
                for (const course of courses) {
                    const course_id = course.course_id;

                    courseSheet.addRow({
                        college_id: collegeId,
                        course_id: course_id || "",
                        display_name: course.display_name || "",
                        duration: course.duration || "",
                        eligibility: course.eligibility || "",
                        level: course.level || "",
                        short_head: course.short_head || "",
                        type: course.type || "",
                        headone_data: course?.headone_data ? JSON.stringify(course?.headone_data) : "",
                        streams: course.streams ? JSON.stringify(course.streams) : "",
                    });
                    totalCoursesFetched++;
    
                }
    
                await workbook.xlsx.writeFile(filePath);
                console.log(`✅ Page ${coursePage}: Fetched & wrote ${courses.length} courses for College ID: ${collegeId}`);
                coursePage++;
                await sleep(5000); // Rate-limiting
    
            } catch (err) {
                console.error(`⚠️ Error fetching Page ${coursePage} for College ID: ${collegeId} → ${err.message}`);
                break;
            }
        }
    
        if (totalCoursesFetched === 0) {
            console.warn(`⚠️ No courses found for College ID: ${collegeId}`);
        }
    }

  await workbook.xlsx.writeFile(filePath);
  console.log('✅ All course data written to the "courses" sheet.');
}

extractCourses();

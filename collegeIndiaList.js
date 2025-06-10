const express = require('express');
const { Builder, By, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const axios = require('axios');
const ExcelJS = require('exceljs');
const cheerio = require('cheerio');
const fs = require("fs");
const path = require("path");

const app = express();

const USERNAME = 'degreestech_f4gEmO'; //replace with your username
const KEY = 'oDhFHHpqp3eXV7xQmNNU'; //replace with your accesskey
const GRID_URL = 'hub-cloud.browserstack.com/wd/hub';

const browserstackUrl = `https://${USERNAME}:${KEY}@${GRID_URL}`;
 
const options = new chrome.Options();
options.addArguments('--headless'); // headless mode
options.addArguments('--disable-gpu'); // optional: needed on Windows
options.addArguments('--no-sandbox');  // optional: needed in some environments

const capabilities = {
    "bstack:options": {
      os: "Windows",
      osVersion: "10",
      local: "false",
      seleniumVersion: "4.0.0",
      userName: USERNAME,
      accessKey: KEY,
      buildName: "browserstack-build",
      sessionName: "college scraping1"
    },
    browserName: "Chrome",
    browserVersion: "latest"
  };

const HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Referer": "https://collegedunia.com/",
    "Upgrade-Insecure-Requests": "1",
    "Connection": "keep-alive"
  };

// const gridUrl = 'https://' + USERNAME + ':' + KEY + '@' + GRID_URL;


// Requires: npm install selenium-webdriver axios exceljs dotenv


app.listen(3811, () => {
    console.log('Server is running on http://localhost:3811');
});

const statesAndUTs = [
  "delhi-ncr",
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

app.get('/lambdatest', async (request, response) => {
    console.log("test")
    try {
      for (const state of statesAndUTs) {
        console.log(`Processing colleges for state: ${state}`);  
        const data = await scrape(state);
        response.status(200).json(data);
      }
    } catch (error) {
      console.error("Scraper failed:", error);  // ✅ Add this line
      response.status(500).json({
        message: 'Server error occurred',
      });
    }
  });

function encodePayload(page, state) {
  const payload = {
    url: state,
    page,
    view: "table",
    has_text_ranking: false,
  };
  const jsonStr = JSON.stringify(payload);
  return Buffer.from(jsonStr).toString('base64');
}

async function scrape(state) {
    const fileName = `all_colleges_basic_info.xlsx`;
    const filePath = path.resolve(fileName);
    const workbook = new ExcelJS.Workbook();
    let sheet;
    // If file exists, load and continue appending
    if (fs.existsSync(filePath)) {
      await workbook.xlsx.readFile(filePath);
      sheet = workbook.getWorksheet("Colleges");
  
    } else {
      sheet = workbook.addWorksheet("Colleges");
      sheet.columns = [
        { header: "College ID", key: "college_id" },
        { header: "College Name", key: "college_name" },
        { header: "College Short Form", key: "college_short_form" },
        { header: "City", key: "city" },
        { header: "State", key: "state" },
        { header: "Approved by", key: "approvals" },
        { header: "url", key: "url"},
        { header: "average_salary", key: "average_salary" },
        { header: "highest_salary", key: "highest_salary" },
        { header: "course_count", key: "course_count" },
        { header: "Logo", key: "logo" },
        { header: "Cover", key: "cover" },
        { header: "NAAC Grading", key: "naac_grading" },
        { header: "area_name", key: "area_name" },
        { header: "university_type", key: "university_type" },
        { header: "nirf_ranking", key: "nirf_ranking" },
      ];
    }
  
    const driver = await new Builder().forBrowser('chrome').setEdgeOptions(options).build();
    try {
        let page = 0;
        while(true) {
            console.log(`Fetching page ${page} for state: ${state}`);
            const encoded = encodePayload(page, state);
            const apiUrl = `https://collegedunia.com/web-api/listing?data=${encoded}`;
            const res = await axios.get(apiUrl, { headers: HEADERS });
            const colleges = res.data.colleges || [];
    
            if (!colleges.length) break;
    
            for (const college of colleges) {
              console.log("College Name:", college.college_name);
              // Add and immediately write the row
              sheet.addRow({
                college_id: college.college_id,
                college_name: college.college_name,
                college_short_form: college.college_short_form,
                state: college.state,
                city: college.college_city,
                approvals: college.approvals,
                url: college.url,
                // university_type: ownership,
                average_salary: college.placement?.average_salary,
                highest_salary: college.placement?.highest_salary,
                course_count: college.courseCount,
                logo: college.logo,
                cover: college.cover,
                naac_grading: college.naac_grading
              });
      
              // Write to file after each college to persist progress
            //   await workbook.xlsx.writeFile(filePath);
              // break;
              console.log(`✅ Saved: ${college.college_name} in college_list1.xlsx`);
              
              // await new Promise(r => setTimeout(r, 300)); // Short delay between colleges
            }
              await workbook.xlsx.writeFile(filePath);

            // break;
            // await new Promise(r => setTimeout(r, 60000)); // 1-minute delay between pages
            if(page%5===0){
                await driver.sleep(5000); // Short delay between colleges
            }
            page++;
        }
  
      console.log("✅ Final data saved to Excel.");
    } finally {
      console.log("driver quitting.");
      await driver.quit();
    }
  }
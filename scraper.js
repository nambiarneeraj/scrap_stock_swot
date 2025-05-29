import puppeteer from 'puppeteer';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const CONFIG = {
  headless: 'new',
  timeout: 30000,
  userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
  popupSelectors: ['#wzrk-cancel', '.modal-close', '.close', '.btn-close', '.overlay-close'],
  swotSelectors: [
    'a[href*="swot-analysis"]',
    'a[data-target="#swot"]',
    'a[onclick*="swot"]',
    '.swot_tab a',
    '.swot_analysis a',
    'a[title*="SWOT"]',
    'a:-webkit-any-link:contains("SWOT")'
  ],
  sheetId: '1XHkvjD7uZuQKc-Q8tc3TW2rIQeLh3HQTPdyLfdqV-3o',
  sheetName: 'StockData',
  serviceAccount: {
    email: 'stockmarketdata@stock-data-461213.iam.gserviceaccount.com',
    privateKey: '-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQCossofJJiuSPO2\nOJNpdjwklekzYvD9Huh0GfOXbxoNCoecDJ7PeEr2pcLyLRbo+6DJ18/E8HGmLr4+\nWMrnL5UYCANoSGHkqJaC9jMvpJdrflOZEgfn1k7VeSAxYa0Ja7p3eIqfILVbKwq3\nnyKala/HzyhN2/JCWe5enooK/OAAuEFRDyy9j7653d/xVf9mxxuPy4kGgbkYDfxc\n+d9nkqhNoWtuKMV2LC+1PtJcTP5npjR/xIpy8y8QF/Xzta43d+8lGIoZArr4mAxa\nuI8YtHvgTdHRkmoODTYJiWbEjUbkkej7FXqJSWNu3W8erCZnvoSGItzhHaHWkjBv\n+ZSOfYJlAgMBAAECggEADmwQ9FsJyoYy5g/6jXhhAxrl1h5ntKjg/cjLygh8VUN7\n/UIYjNG5Gh9o7F26da4wotDjOSNgZyRCRplxgnsCgT/Xurr/BrAuN4GaaQ1//Hdo\ntWj7NnDCsJfFKGbNukTbEVeqBWY9QU8+s7Zt6OiFhjMWb2f8U0iiixt9EoKbGLSE\nLqmBuLeUBrceda4NT1U1m+QbynbkyfrYn1E21NsmN3Wq743X+A2tUNx5VvuilWJK\niAps/fi93GfBY1n2UoH47Xy4W5klr3+hc5Ru1zHTADd5gcCDQh8qxMGQku8ZB8Up\nQDEO6y2ehWPUoxSBXSYT3btxa+Z7zA74tK4npXNaiQKBgQDlkW2ka6WbB0zJiQV+\nWLXjqKBIDVQ4uwbPGG8kJ62Wp19AjvUAzttY9TqPrAyqQ8/MzYjp4G9UWup8bc3B\nZeHtk0c3vg4dh2K5wVOGPE8QHJmO3EZ2GEJvnH7r74GyjlMrozeoIQB8Pulb3kst\nAKHNCGYZcxpYZUlgLdRFW1zL3wKBgQC8HzaiMHtw2QqcEGrU/aZB4vXBKEggwVi3\nuT7NwrN9i3hk74GAgUMIRfT9iVUcVONwrFKJEZqwfcMmqU0Cax0v7OfrK2LM9+a+\nqwludYqyPLEOuPb+bAD6s1FCUQe+NJbSl4cokhO2tVHejhfrwA5SwmXAMDABVb45\nT01Wo1Y6OwKBgHI1M3LFCxJhQ1ZQEKeWwoaL8ZFm8Ct5AB4vbbty8e0tPzoC5OiO\nAJn1BjlLwtFCAzNEXYTc3wX8ZQOaLO62HPvwdVHJ/4O5QuhewYrangrJ76se8v71\nerfEB3ChKskF/WKMRLgkEvW85qOJp6Sv188FCqZGmSi42xQ6OIx4s2XJAoGAXo8Z\n+SCBi9GtAZFHAdSVs1yPxw2mY8CMBZ15sheB/UMTuzigUaWnugrgAGj9fQY2ZLZZ\nrkhJBxnP9Cj5apPI0gQ09wKR4RFizMhQL1Op6bmUDiBvFqfXPizQVZNBXxw0C5ra\n90ul2Rr/Ee0+nOOmz3ajip0uJB2jRk9UQo5Lk20CgYAlTYu0NoY+yFSrdLiLnf+I\nhWJnG3wRgARdqdXht6mWYBMMwhJWnIBQ3h7wUCyWQ/Xk7BiTO42Xv/1m7vlmkULP\nuR25KA8Yn6rc+w0sL/M/AM54G8PByqKACdHZBMKi3d8KzCiUvVfU6cK6C3FgR40Q\nFIDsaOSNMwF6LBpQUoWWCw==\n-----END PRIVATE KEY-----\n'
  }
};

function extractNumbers(text) {
  if (!text) return null;
  const match = text.match(/\d+/);
  return match ? parseInt(match[0], 10) : null;
}

async function createBrowser() {
  return await puppeteer.launch({
    headless: CONFIG.headless,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-accelerated-2d-canvas',
      '--disable-gpu',
      '--window-size=1920,1080',
      '--single-process',
      '--no-zygote'
    ],
    
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined
  });
}

async function saveToGoogleSheets(data) {
  const serviceAccountAuth = new JWT({
    email: CONFIG.serviceAccount.email,
    key: CONFIG.serviceAccount.privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet(CONFIG.sheetId, serviceAccountAuth);
  await doc.loadInfo();

  let sheet = doc.sheetsByTitle[CONFIG.sheetName];
  if (!sheet) {
    sheet = await doc.addSheet({
      title: CONFIG.sheetName,
      headerValues: [
        'Company',
        'MC Essentials',
        'Strengths Count',
        'Weaknesses Count',
        'Opportunities Count',
        'Threats Count',
        'Date',
      ],
    });
  }

  await sheet.addRows(data);
  console.log('Data successfully appended to Google Sheets');
}

async function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function scrapeCompanyDataWithRetry(page, url, companyName, retries = 2) {
  let attempt = 0;
  while (attempt <= retries) {
    try {
      return await scrapeCompanyData(page, url, companyName);
    } catch (error) {
      attempt++;
      console.log(`Attempt ${attempt} failed for ${companyName}:`, error.message);
      if (attempt <= retries) {
        await delay(3000);
        console.log(`Retrying ${companyName}...`);
      }
    }
  }
  return {
    company: companyName,
    essentials: null,
    strengths: null,
    weaknesses: null,
    opportunities: null,
    threats: null,
    timestamp: new Date().toLocaleDateString('en-GB')
  };
}

async function scrapeCompanyData(page, url, companyName) {
  try {
    console.log(`Navigating to ${url}`);
    await page.goto(url, {
      waitUntil: 'networkidle2',
      timeout: CONFIG.timeout
    });

    await handlePopups(page);
    await page.waitForSelector('#common_header', { timeout: 15000 });

    const essentials = await scrapeMCEssentials(page);
    const swot = await scrapeSWOTAnalysis(page);

    const dateObj = new Date();
    const day = String(dateObj.getDate()).padStart(2, '0');
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const year = dateObj.getFullYear();
    const formattedDate = `${day}/${month}/${year}`;

    return {
      company: companyName,
      essentials: extractNumbers(essentials),
      strengths: extractNumbers(swot.strengths),
      weaknesses: extractNumbers(swot.weaknesses),
      opportunities: extractNumbers(swot.opportunities),
      threats: extractNumbers(swot.threats),
      timestamp: formattedDate
    };
  } catch (error) {
    console.error(`Error scraping ${url}:`, error.message);
    throw error;
  }
}

async function scrapeMCEssentials(page) {
  try {
    await page.waitForSelector('.bx_mceti .esbx', { timeout: 8000 });
    return await page.evaluate(() => {
      const essentialsElement = document.querySelector('.bx_mceti .esbx');
      return essentialsElement ? essentialsElement.textContent.trim() : null;
    });
  } catch (error) {
    console.log('MC Essentials not found');
    return null;
  }
}

async function scrapeSWOTAnalysis(page) {
  try {
    console.log('Looking for SWOT tab...');
    await delay(2000); 

    const swotTab = await findClickableElement(page, CONFIG.swotSelectors);

    if (!swotTab) {
      console.log('SWOT tab not found with standard selectors, trying alternative approach');
      const links = await page.$$('a');
      for (const link of links) {
        const text = await page.evaluate(el => el.textContent?.toLowerCase(), link);
        if (text && text.includes('swot')) {
          console.log('Found SWOT link by text content');
          await link.click();
          break;
        }
      }
    } else {
      console.log('Found SWOT tab with selector');
      await swotTab.click();
    }

    try {
      await page.waitForSelector('.swot_feature', { visible: true, timeout: 15000 });
      console.log('SWOT content loaded successfully');
    } catch (e) {
      console.log('SWOT content not found after clicking tab');
    }

    return await page.evaluate(() => {
      const container = document.querySelector('.swot_feature');
      if (!container) return {
        strengths: null,
        weaknesses: null,
        opportunities: null,
        threats: null
      };

      const getCount = (className) => {
        const element = container.querySelector(`.${className}`);
        if (!element) return null;
        const text = element.textContent || '';
        const match = text.match(/\d+/);
        return match ? match[0] : null;
      };

      return {
        strengths: getCount('swli1') || getCount('strengths'),
        weaknesses: getCount('swli2') || getCount('weaknesses'),
        opportunities: getCount('swli3') || getCount('opportunities'),
        threats: getCount('swli4') || getCount('threats')
      };
    });
  } catch (error) {
    console.log('SWOT Analysis error:', error.message);
    return {
      strengths: null,
      weaknesses: null,
      opportunities: null,
      threats: null
    };
  }
}

async function findClickableElement(page, selectors) {
  for (const selector of selectors) {
    try {
      await page.waitForSelector(selector, { visible: true, timeout: 2000 });
      const element = await page.$(selector);
      if (element) {
        const isVisible = await page.evaluate(el => {
          const style = window.getComputedStyle(el);
          return style && style.display !== 'none' && style.visibility !== 'hidden' && el.offsetHeight > 0;
        }, element);

        if (isVisible) {
          console.log(`Found visible element with selector: ${selector}`);
          return element;
        }
      }
    } catch (err) {
      continue;
    }
  }
  return null;
}

async function handlePopups(page) {
  console.log('Checking for popups...');
  for (const selector of CONFIG.popupSelectors) {
    try {
      await page.waitForSelector(selector, { visible: true, timeout: 3000 });
      await page.click(selector);
      console.log(`Closed popup with selector: ${selector}`);
      await delay(500);
    } catch (err) {
      
    }
  }
}

(async () => {
  try {
    const companies = [
      { category: "refineries", name: 'Reliance Industries', slug: 'relianceindustries', sectorCode: 'RI' },
      { category: "computers-software", name: 'TATA Consultancies', slug: 'tataconsultancyservices', sectorCode: 'TCS' },
      { category: "computers-software", name: 'Infosys Limited', slug: 'infosys', sectorCode: 'IT' },
      { category: "banks-private-sector", name: 'HDFC Bank Ltd', slug: 'hdfcbank', sectorCode: 'HDF01' },
      { category: "banks-private-sector", name: 'ICICI Bank Limited', slug: 'icicibank', sectorCode: 'ICI02' },
      { category: "personal-care", name: 'Hindustan Unilever Limited', slug: 'hindustanunilever', sectorCode: 'HU' },
      { category: "banks-public-sector", name: 'State Bank of India', slug: 'statebankindia', sectorCode: 'SBI' },
      { category: "banks-private-sector", name: 'Kotak Mahindra Bank Ltd', slug: 'kotakmahindrabank', sectorCode: 'KMB' },
      { category: "diversified", name: 'ITC Ltd', slug: 'itc', sectorCode: 'ITC' },
      { category: "infrastructure-general", name: 'Larsen & Toubro Ltd.', slug: 'larsentoubro', sectorCode: 'LT' }
    ];

    const browser = await createBrowser();

   
    const scrapePromises = companies.map(async (company) => {
      const page = await browser.newPage(); 
      await page.setUserAgent(CONFIG.userAgent);
      await page.setDefaultTimeout(CONFIG.timeout);

      const url = `https://www.moneycontrol.com/india/stockpricequote/${company.category}/${company.slug}/${company.sectorCode}`;
      console.log(`\nFetching data for ${company.name}...`);

      try {
        const companyData = await scrapeCompanyDataWithRetry(page, url, company.name);
        console.log(`Completed ${company.name}`);
        return companyData;
      } catch (error) {
        console.error(`Failed to scrape ${company.name}:`, error.message);
        
        return {
          company: company.name,
          essentials: null,
          strengths: null,
          weaknesses: null,
          opportunities: null,
          threats: null,
          timestamp: new Date().toLocaleDateString('en-GB')
        };
      } finally {
        await page.close();
      }
    });

  
    const settledResults = await Promise.allSettled(scrapePromises);

    
    const results = settledResults
      .filter(result => result.status === 'fulfilled')
      .map(result => result.value);

    await browser.close();

    const sheetData = results.map(r => ({
      'Company': r.company,
      'MC Essentials': r.essentials,
      'Strengths Count': r.strengths,
      'Weaknesses Count': r.weaknesses,
      'Opportunities Count': r.opportunities,
      'Threats Count': r.threats,
      'Date': r.timestamp
    }));

    await saveToGoogleSheets(sheetData);

    console.log('\nScraping completed. Data saved to Google Sheets');
  } catch (error) {
    console.error('Fatal error:', error);
    process.exit(1);
  }
})();

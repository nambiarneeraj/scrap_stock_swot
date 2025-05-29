import puppeteer from 'puppeteer';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';

const CONFIG = {
  headless: 'new',
  timeout: 60000,
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
    email: 'stockmarketdata@stock-data-461213.iam.gserviceAccount.com',
    privateKey: '-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCwueY5d0Weo+/Y\nbniiFVce0IT9kBfWynbMw87AHmG0B/SwsJHAvWDSywXFpV/v/sNCEKEJK3tP1J0X\nqzG5ELwESv1I4Xxm/a9u+E79Y803GR83dQoEe/9XJxWWnyo7oo0KKnTUcdQg4WSC\n0oJzrHXz59oYOuRohYSPo+ukpfIXFbK/JebHJIm/+FZr4ftrq1TVH/eScP0ah8sr\nR7rdUPxcZ/lU573W+31zf1VGk2YmN/q3agFq8QG8mJT/fZUEofnbtNCEVbFZwWQM\nkTJ3EBZAnwAtxE603K/gPbiwwDYuZj9JaI4gmdwuKNy7pNeaV99KQ/jANOxB5lk4\nwUF+dFZ3AgMBAAECggEAAUHsC8qiedosHGZ6K0UVmp2HOhAAFheDYuTzIADXVyLw\nsHUr2arp+SCtXiuqvsLEUZWxX6b/OKEzAPZx4yGQWkN3q+tsKYYdQnU33VZhdnE8\n/KBePYvtqYlt/jq3Cozsjf72rTBQ1G8Qz7F/G3hFr+zywKtinAUfMq/K6cqt1KY1\ng557kRcy6gu1PZde/vv8rQTVVKLN9q7zcIwaDmTHvFDD11wJ3ISfqXQXTNuqxV03\n0xXUBF8B6EuvGNuN8E9Crqyv8kPyum8bAeHTd9zTzptvECDcu2ldlVdhGx6iE3fB\nXSs1bz/UZItNvlA287nN95P4Uf4NK9yic04RN+G68QKBgQDcrzZMpPcO5rDkoAPh\nTGuLAWsiXfpjADSSMvb6N316UrJtMugA+gHun5jPRRZ88x+y/krFgGaE36tpSahy\nF5IPKzUV4UGbEbGTxPA0XOM7f5rOeoW6HzFPLRF7zHS2qpgnfeT+ob/HWzQ3u8uh\nHRfcANE47DAOYtHMzTwIeKmRqQKBgQDNAdj+GJuiye64oaCpY7Ioy2AZwDEaebVr\nYlWLfgZOv7C0I4dkuTV6qB54T0aVjphEoPmqVNfr9AY1HH8+cuZiCH8FsCDmT9ap\nfkAuw9NfLn5tmDQLMYWAhIgJA+l9/t2OaarBMyICUe7QZRBUWVXGJFYTeWtAAB/n\nw56NxiP7HwKBgCy0Ya+NC289VEA8Gg0dyftSwj0oBHzhocSsBlQRwZ1x+ysb0NvB\nyXppYi86s5+EMLu1v7falun71WFyxmi2VaQ1AH/6LawYHXztvCsfVfjLlLSXJVfa\n0cZUPuJxPIN0c3YsjqL2aT8dPqq7pDhzCE5M7BU341RGuHFgcfTVXKRhAoGABO+o\nc+XPyYmnL9bkcW+vGIBdHgGcrRCFJ8LEYIl2SWsgLBY26lvzR7LImQj/oBZA4FYn\n7MwCLvI/PAQlpDFMDsw5kr9860680nPxw65/ZmlOLgFeL27P0hpe1Ci99ISwfP9a\nVzCN/xRN9cKZNA66m/y//dQMmwvluMTjCnLc5u0CgYEApmTlmGYK+VpiTBYbqDsZ\neu+EMFFUVGhRUTQ1qU/GrkNOVkyUE2gULaN7h0yxEfT7i8VnPxChJ9HvMp55AtKj\nvFtDjzR/utN8paLl4KQdSpX/X4nxn3rGslzKHzJCI2kLUeaAiXTUpXCKy1AZV+y7\ncBH8CCkMOtHrvigXeH9jduY=\n-----END PRIVATE KEY-----\n'
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
    executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
    protocolTimeout: 180000
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
}

async function scrapeMCEssentials(page) {
  try {
    await page.waitForSelector('.bx_mceti .esbx', { timeout: 8000 });
    return await page.evaluate(() => {
      const essentialsElement = document.querySelector('.bx_mceti .esbx');
      return essentialsElement ? essentialsElement.textContent.trim() : null;
    });
  } catch (error) {
    console.log('MC Essentials not found or timed out:', error.message);
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
      console.log('SWOT content not found after clicking tab or timed out:', e.message);
      return {
        strengths: null,
        weaknesses: null,
        opportunities: null,
        threats: null
      };
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
    console.log('SWOT Analysis error (outer catch):', error.message);
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
      console.log(`Selector ${selector} not found or timed out: ${err.message}`);
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
      console.log(`Popup selector ${selector} not found or timed out: ${err.message}`);
    }
  }
}

(async () => {
  try {
    const companies = [
      { category: "refineries", name: 'Reliance Industries', slug: 'relianceindustries', sectorCode: 'RI' }
     // { category: "computers-software", name: 'TATA Consultancies', slug: 'tataconsultancyservices', sectorCode: 'TCS' },
      //{ category: "computers-software", name: 'Infosys Limited', slug: 'infosys', sectorCode: 'IT' },
      //{ category: "banks-private-sector", name: 'HDFC Bank Ltd', slug: 'hdfcbank', sectorCode: 'HDF01' },
      //{ category: "banks-private-sector", name: 'ICICI Bank Limited', slug: 'icicibank', sectorCode: 'ICI02' },
      //{ category: "personal-care", name: 'Hindustan Unilever Limited', slug: 'hindustanunilever', sectorCode: 'HU' },
      //{ category: "banks-public-sector", name: 'State Bank of India', slug: 'statebankindia', sectorCode: 'SBI' },
      //{ category: "banks-private-sector", name: 'Kotak Mahindra Bank Ltd', slug: 'kotakmahindrabank', sectorCode: 'KMB' },
      //{ category: "diversified", name: 'ITC Ltd', slug: 'itc', sectorCode: 'ITC' },
      //{ category: "infrastructure-general", name: 'Larsen & Toubro Ltd.', slug: 'larsentoubro', sectorCode: 'LT' }
    ];

    const browser = await createBrowser();
    const page = await browser.newPage();
    await page.setUserAgent(CONFIG.userAgent);
    await page.setDefaultTimeout(CONFIG.timeout);

    const results = [];

    for (const company of companies) {
      const url = `https://www.moneycontrol.com/india/stockpricequote/${company.category}/${company.slug}/${company.sectorCode}`;
      console.log(`\nFetching data for ${company.name}...`);

      const companyData = await scrapeCompanyData(page, url, company.name);
      results.push(companyData);

      console.log(`Completed ${company.name}`);
    }

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
    console.error('Fatal error in main execution block:', error);
    process.exit(1);
  }
})();

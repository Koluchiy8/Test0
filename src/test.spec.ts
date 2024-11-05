import { chromium } from 'playwright';
import config from './config.json';
import * as fs from 'fs';
import * as path from 'path';
import { format } from 'date-fns';



const logToFile = (message: string) => {
    const logFilePath = path.join(__dirname, 'log.txt'); // Path to log-file
    const timestamp = new Date().toISOString(); // Getting today's date
    const logMessage = `${timestamp} - ${message}\n`; // Format log message
    fs.appendFileSync(logFilePath, logMessage); // Write message to file
};


describe('Excel Online TODAY() Function Test', () => {
  test('Verify TODAY() function in Excel Online', async () => {
    
    // open Chrome, allow copying to clipboard
    const browser = await chromium.launch({ headless: false, channel: 'chrome' });
    const page = await browser.newPage();
    await browser.contexts()[0].grantPermissions(['clipboard-read']);

    // authorization on the website MS
    logToFile('Opening the website https://office.com')
    await page.goto('https://office.com');
    await page.waitForLoadState();
    logToFile('Authorization with login and password...')
    await page.click('text=Sign in');
    await page.waitForLoadState();
    await page.fill('input[type="email"]', config.email);
    await page.waitForLoadState();
    await page.click('text=Next');
    await page.waitForTimeout(1000);
    await page.fill('input[type="password"]', config.password);
    await page.waitForLoadState();
    await page.click('text=Sign in');
    await page.waitForLoadState();
    // window "Stay signed"
    await page.click('text=Yes');
    await page.waitForLoadState();
    
    // go to Excel page
    logToFile('Opening Excel page...')
    await page.goto('https://www.office.com/launch/Excel/');
    await page.waitForLoadState();
    
    // opening new workbook
    logToFile('Opening new workbook...')
    await page.click('text="Blank workbook"');
    await page.waitForTimeout(3000);


    // opening Excel Online in a new tab
    const newPage = browser.contexts()[0].pages()[1]; 
    await newPage.waitForTimeout(10000);
    
    // Input the current date function from the keyboard, copy the value from this cell.
    await newPage.keyboard.type('=TODAY()');
    await newPage.keyboard.press('Enter');
    await newPage.keyboard.press('ArrowUp');
    await newPage.waitForTimeout(3000);
    await newPage.keyboard.press('Control+C');
    await newPage.waitForTimeout(3000);

    // getting the value from the clipboard
    const cellValue = await newPage.evaluate(async () => {
      try {
        return await navigator.clipboard.readText();
      } catch (err) {
        logToFile('Error reading from the clipboard: ' + err);
        return '';
      }
    });

    logToFile('cellValue: ' + cellValue);
    
    // close browser
    await browser.close();

    // getting today's date in the desired format
    const today = format(new Date(), 'M/d/yyyy');

    logToFile('today_date: ' + today);
    expect(cellValue).toBe(today);

  });
});
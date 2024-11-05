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



//const today = format(new Date(), 'MM/dd/yyyy');
//logToFile(today);



/*

//logToFile('Opened https://example.com');



//import { format } from 'date-fns';


(async () => {
  const browser = await chromium.launch({ headless: false, channel: 'chrome' });
  //const context = await browser.newContext();
  const page = await browser.newPage();
  await browser.contexts()[0].grantPermissions(['clipboard-read']);




  await page.goto('https://office.com');
  //await page.screenshot({ path: 'screenshot_home.png' });
  //console.log("asdfasdf");
  await page.waitForTimeout(1000);
  await page.click('text=Sign in');
  await page.waitForTimeout(1000);
  //await page.screenshot({ path: 'screenshot1.png' });
  await page.waitForTimeout(1000);
  await page.fill('input[type="email"]', config.email);
  //await page.screenshot({ path: 'screenshot2.png' });
  await page.click('text=Next');
  await page.waitForTimeout(1000);
  //await page.screenshot({ path: 'screenshot3.png' });
  await page.fill('input[type="password"]', config.password);
  await page.click('text=Sign in');
  await page.waitForTimeout(1000);
  //Stay signed
  await page.click('text=Yes');
  await page.waitForTimeout(10000);
  await page.goto('https://www.office.com/launch/Excel/');
  await page.waitForTimeout(10000);
  await page.click('text="Blank workbook"');
  await page.waitForTimeout(20000);
  const newPage = browser.contexts()[0].pages()[1]; //доступ ко второй открытой вкладке
  logToFile('Excel URL: ' + newPage.url());
  //const content = await newPage.content();
  //logToFile('Content: ' + content);
  //await page.click('div[aria-label="A1"]', { force: true });
  await newPage.keyboard.type('=TODAY()');
  await newPage.keyboard.press('Enter');
  await newPage.keyboard.press('ArrowUp');
  await newPage.keyboard.press('Control+C');

  await newPage.waitForTimeout(5000);


















  const copiedValue = await newPage.evaluate(async () => {
    try {
      return await navigator.clipboard.readText();
    } catch (err) {
      console.error('Ошибка при чтении из буфера обмена:', err);
      return '';
    }
  });
  await newPage.waitForTimeout(10000);
  //await newPage.click('text=Allow');


  
  logToFile('Скопированное значение даты: ' + copiedValue);



  
  await newPage.waitForTimeout(10000);


  //const page2 = await context.newPage();
  //await page2.goto('https://playwright.dev');
  //await page2.waitForTimeout(20000);
    // Получаем все открытые вкладки (страницы)
    //const contexts = browser.contexts();

    //logToFile('Открытые вкладки и окна во всех контекстах:');
    //contexts.forEach((context, contextIndex) => {
   //   const pages = browser.contexts()[0].pages();
   //   pages.forEach((page, pageIndex) => {
   //     logToFile(`  Вкладка ${pageIndex + 1}: ${page.url()}`);
   //   });
    //});



  


 



 


  

  logToFile('Finished');

  
  //await page.screenshot({ path: 'screenshot3.png' });


  

  




  await browser.close();
  
})();
*/
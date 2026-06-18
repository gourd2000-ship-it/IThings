const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  const url = 'https://script.google.com/macros/s/AKfycbzxXfdOgCebpPlb6fEHJZg8-smpOSE-uRVuMKterrsBuUCqsN0XnUn7kTWNTlekGQsZAQ/exec?id=EQ-2026-001';
  console.log('Navigating to:', url);
  
  try {
    await page.goto(url, { waitUntil: 'networkidle' });
    console.log('Page loaded.');
    
    const title = await page.title();
    console.log('Page Title:', title);
    
    const content = await page.textContent('body');
    console.log('Page Content:', content.substring(0, 1000));
  } catch (e) {
    console.error('Error during test:', e);
  } finally {
    await browser.close();
  }
})();

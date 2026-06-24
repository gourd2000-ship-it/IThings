const { chromium } = require('playwright');

(async () => {
  console.log('Starting Playwright API Lock Test...');
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  const baseUrl = 'https://script.google.com/macros/s/AKfycbzxXfdOgCebpPlb6fEHJZg8-smpOSE-uRVuMKterrsBuUCqsN0XnUn7kTWNTlekGQsZAQ/exec';
  const debugUrl = `${baseUrl}?id=DEBUG_TEST`;
  const deviceUrl = `${baseUrl}?id=PC-1`;
  
  // iframe 내부에서 렌더링된 진짜 텍스트 콘텐츠를 안전하게 긁어오는 함수
  async function getDebugDataFromFrames(page) {
    const debugPrefix = 'DEBUG_DATA:';
    for (let i = 0; i < 15; i++) {
      for (const frame of page.frames()) {
        const text = await frame.evaluate(() => document.body ? document.body.textContent : '').catch(() => '');
        const index = text.indexOf(debugPrefix);
        if (index !== -1) {
          const jsonStr = text.substring(index + debugPrefix.length).trim();
          try {
            return JSON.parse(jsonStr);
          } catch (err) {
            console.log('Failed to parse frame JSON, retrying...', err.message);
          }
        }
      }
      await new Promise(r => setTimeout(r, 1000));
    }
    throw new Error('DEBUG_DATA not found in any frame after timeout.');
  }
  
  try {
    // 1. 잠금 해제 (unlock) 호출
    console.log('\n--- Step 1: Unlocking inputs ---');
    await page.goto(`${debugUrl}&action=unlock`, { waitUntil: 'networkidle' });
    console.log('Unlock action triggered. Waiting for 3 seconds for sheets commit...');
    await new Promise(r => setTimeout(r, 3000));
    
    // 2. 잠금 해제 상태 확인
    console.log('\n--- Step 2: Verifying unlocked status ---');
    await page.goto(debugUrl, { waitUntil: 'networkidle' });
    let debugData = await getDebugDataFromFrames(page);
    console.log('allEditLocked status:', debugData.settings.allEditLocked);
    console.log('Detailed settings values:', JSON.stringify(debugData.settings, null, 2));
    if (debugData.settings.allEditLocked !== false) {
      throw new Error('Failed to unlock edit lock.');
    }
    
    // 3. 잠금 해제 상태에서 API 호출 테스트 (기기: PC-1)
    console.log('\n--- Step 3: Testing updateDevice when UNLOCKED ---');
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    let targetFrame = null;
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasGoogleScript = await frame.evaluate(() => {
          return typeof google !== 'undefined' && typeof google.script !== 'undefined' && typeof google.script.run !== 'undefined';
        }).catch(() => false);
        
        if (hasGoogleScript) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find Google Apps Script frame.');
    }
    
    let apiResponse = await targetFrame.evaluate(async () => {
      return new Promise((resolve) => {
        google.script.run
          .withSuccessHandler((res) => resolve(res))
          .withFailureHandler((err) => resolve({ success: false, message: err.toString() }))
          .updateDevice('PC-1', {
            location: '컴퓨터실',
            manager: 'Kristina Pernar',
            notes: 'Playwright API Lock Test - Unlocked State'
          });
      });
    });
    console.log('Update Device (Unlocked State) Response:', JSON.stringify(apiResponse, null, 2));
    
    // 4. 모든 입력 잠금 (lock) 호출
    console.log('\n--- Step 4: Locking inputs ---');
    await page.goto(`${debugUrl}&action=lock`, { waitUntil: 'networkidle' });
    console.log('Lock action triggered. Waiting for 3 seconds for sheets commit...');
    await new Promise(r => setTimeout(r, 3000));
    
    // 5. 잠금 설정 상태 확인
    console.log('\n--- Step 5: Verifying locked status ---');
    await page.goto(debugUrl, { waitUntil: 'networkidle' });
    debugData = await getDebugDataFromFrames(page);
    console.log('allEditLocked status:', debugData.settings.allEditLocked);
    console.log('Detailed settings values:', JSON.stringify(debugData.settings, null, 2));
    if (debugData.settings.allEditLocked !== true) {
      throw new Error('Failed to lock edit lock.');
    }
    
    // 6. 잠금 설정 상태에서 API 호출 테스트 (기기: PC-1)
    console.log('\n--- Step 6: Testing updateDevice when LOCKED ---');
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    targetFrame = null;
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasGoogleScript = await frame.evaluate(() => {
          return typeof google !== 'undefined' && typeof google.script !== 'undefined' && typeof google.script.run !== 'undefined';
        }).catch(() => false);
        
        if (hasGoogleScript) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find Google Apps Script frame.');
    }
    
    apiResponse = await targetFrame.evaluate(async () => {
      return new Promise((resolve) => {
        google.script.run
          .withSuccessHandler((res) => resolve(res))
          .withFailureHandler((err) => resolve({ success: false, message: err.toString() }))
          .updateDevice('PC-1', {
            location: '컴퓨터실',
            manager: 'Kristina Pernar',
            notes: 'Playwright API Lock Test - Locked State'
          });
      });
    });
    console.log('Update Device (Locked State) Response:', JSON.stringify(apiResponse, null, 2));
    
    // 7. 잠금 해제로 원상 복구 (테스트 후 청소)
    console.log('\n--- Step 7: Cleaning up (Unlocking inputs) ---');
    await page.goto(`${debugUrl}&action=unlock`, { waitUntil: 'networkidle' });
    console.log('Cleanup completed.');
    
  } catch (e) {
    console.error('Error during test execution:', e);
  } finally {
    await browser.close();
    console.log('Browser closed. Test finished.');
  }
})();

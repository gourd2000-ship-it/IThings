const { chromium } = require('playwright');

(async () => {
  console.log('Starting Playwright UI Lock Test...');
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  const baseUrl = 'https://script.google.com/macros/s/AKfycbzxXfdOgCebpPlb6fEHJZg8-smpOSE-uRVuMKterrsBuUCqsN0XnUn7kTWNTlekGQsZAQ/exec';
  const debugUrl = `${baseUrl}?id=DEBUG_TEST`;
  const deviceUrl = `${baseUrl}?id=PC-1`;
  
  try {
    // ----------------------------------------
    // 시나리오 A: 잠금 활성화(LOCKED) 상태 테스트
    // ----------------------------------------
    console.log('\n=== Scenario A: Locking inputs & Verifying UI ===');
    
    // 1. 잠금 활성화
    console.log('Triggering lock...');
    await page.goto(`${debugUrl}&action=lock`, { waitUntil: 'networkidle' });
    await new Promise(r => setTimeout(r, 3000)); // commit 대기
    
    // 2. 기기 상세 페이지 접속
    console.log('Navigating to Device Edit Page:', deviceUrl);
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    // 3. iframe 탐색 (구글 샌드박스 내부 구조)
    console.log('Searching for Apps Script user content frame...');
    let targetFrame = null;
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasBanner = await frame.evaluate(() => {
          return document.getElementById('lock-banner') !== null;
        }).catch(() => false);
        
        if (hasBanner) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find frame containing #lock-banner.');
    }
    
    console.log('Apps Script frame found. Checking UI element states...');
    
    // 4. 잠금 UI 속성 검증
    const lockUiState = await targetFrame.evaluate(() => {
      const banner = document.getElementById('lock-banner');
      const location = document.getElementById('location');
      const manager = document.getElementById('manager');
      const notes = document.getElementById('notes');
      const submitBtn = document.getElementById('submit-btn');
      const btnText = document.getElementById('btn-text') ? document.getElementById('btn-text').innerText : '';
      
      const isBannerVisible = banner && window.getComputedStyle(banner).display === 'flex';
      
      return {
        bannerText: banner ? banner.innerText : '',
        isBannerVisible,
        locationDisabled: location ? location.disabled : false,
        managerDisabled: manager ? manager.disabled : false,
        notesDisabled: notes ? notes.disabled : false,
        submitBtnDisabled: submitBtn ? submitBtn.disabled : false,
        btnText
      };
    });
    
    console.log('UI State (LOCKED):');
    console.log(JSON.stringify(lockUiState, null, 2));
    
    // 검증 어설션
    if (!lockUiState.isBannerVisible) {
      throw new Error('Lock banner should be visible when locked.');
    }
    if (!lockUiState.bannerText.includes('입력 잠금')) {
      throw new Error('Lock banner should contain edit-lock message.');
    }
    if (!lockUiState.locationDisabled || !lockUiState.managerDisabled || !lockUiState.notesDisabled) {
      throw new Error('All input fields must be disabled.');
    }
    if (!lockUiState.submitBtnDisabled) {
      throw new Error('Submit button must be disabled.');
    }
    if (lockUiState.btnText !== '수정 잠금 활성화됨') {
      throw new Error('Submit button text should be "수정 잠금 활성화됨".');
    }
    
    console.log('>> SCENARIO A PASSED SUCCESSFULLY!');
    
    // ----------------------------------------
    // 시나리오 B: 잠금 비활성화(UNLOCKED) 상태 테스트
    // ----------------------------------------
    console.log('\n=== Scenario B: Unlocking inputs & Verifying UI ===');
    
    // 1. 잠금 해제
    console.log('Triggering unlock...');
    await page.goto(`${debugUrl}&action=unlock`, { waitUntil: 'networkidle' });
    await new Promise(r => setTimeout(r, 3000)); // commit 대기
    
    // 2. 기기 상세 페이지 다시 접속
    console.log('Navigating to Device Edit Page:', deviceUrl);
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    // 3. iframe 탐색
    console.log('Searching for Apps Script user content frame...');
    targetFrame = null;
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasLocation = await frame.evaluate(() => {
          return document.getElementById('location') !== null;
        }).catch(() => false);
        
        if (hasLocation) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find frame containing #location.');
    }
    
    // 4. 잠금 해제 UI 속성 검증
    const unlockUiState = await targetFrame.evaluate(() => {
      const banner = document.getElementById('lock-banner');
      const location = document.getElementById('location');
      const manager = document.getElementById('manager');
      const notes = document.getElementById('notes');
      const submitBtn = document.getElementById('submit-btn');
      const btnText = document.getElementById('btn-text') ? document.getElementById('btn-text').innerText : '';
      
      const isBannerHidden = !banner || window.getComputedStyle(banner).display === 'none';
      
      return {
        isBannerHidden,
        locationEnabled: location ? !location.disabled : false,
        managerEnabled: manager ? !manager.disabled : false,
        notesEnabled: notes ? !notes.disabled : false,
        submitBtnEnabled: submitBtn ? !submitBtn.disabled : false,
        btnText
      };
    });
    
    console.log('UI State (UNLOCKED):');
    console.log(JSON.stringify(unlockUiState, null, 2));
    
    // 검증 어설션
    if (!unlockUiState.isBannerHidden) {
      throw new Error('Lock banner should be hidden when unlocked.');
    }
    if (!unlockUiState.locationEnabled || !unlockUiState.managerEnabled || !unlockUiState.notesEnabled) {
      throw new Error('Input fields should be enabled.');
    }
    if (!unlockUiState.submitBtnEnabled) {
      throw new Error('Submit button should be enabled.');
    }
    if (unlockUiState.btnText !== '수정 내용 저장') {
      throw new Error('Submit button text should be "수정 내용 저장".');
    }
    
    console.log('>> SCENARIO B PASSED SUCCESSFULLY!');
    
  } catch (e) {
    console.error('Error during UI lock test execution:', e);
  } finally {
    // 5. 청소
    console.log('\nCleaning up (Unlocking inputs)...');
    await page.goto(`${debugUrl}&action=unlock`, { waitUntil: 'networkidle' });
    await browser.close();
    console.log('Browser closed. Test finished.');
  }
})();

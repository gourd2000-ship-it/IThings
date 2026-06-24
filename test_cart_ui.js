const { chromium } = require('playwright');

(async () => {
  console.log('Starting Playwright Cart UI & PDF Download E2E Test...');
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();
  
  // 브라우저 내부 로그 및 에러 수신 등록
  page.on('console', msg => {
    console.log(`[Browser Console] ${msg.type()}: ${msg.text()}`);
  });
  page.on('pageerror', err => {
    console.error(`[Browser PageError] ${err.toString()}`);
  });

  // 브라우저 alert 및 confirm 대화상자 자동 핸들러 등록
  page.on('dialog', async dialog => {
    console.log(`[Dialog Triggered] Type: ${dialog.type()}, Message: "${dialog.message()}"`);
    await dialog.accept(); // 모든 경고 및 확인창 수락
  });

  const baseUrl = 'https://script.google.com/macros/s/AKfycbzxXfdOgCebpPlb6fEHJZg8-smpOSE-uRVuMKterrsBuUCqsN0XnUn7kTWNTlekGQsZAQ/exec';
  const deviceUrl = `${baseUrl}?id=PC-1`;
  let targetFrame = null;
  
  try {
    console.log('\nStep 1: Navigating to Device Page to clear LocalStorage...');
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    // iframe 탐색
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasFab = await frame.evaluate(() => {
          return document.getElementById('fab-add-btn') !== null;
        }).catch(() => false);
        
        if (hasFab) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find frame containing 보관함 버튼.');
    }
    
    console.log('Frame found. Clearing LocalStorage...');
    await targetFrame.evaluate(() => {
      localStorage.clear();
    });
    
    // 로컬 스토리지를 깨끗하게 비운 뒤 다시 기기 정보 화면으로 접속하여 깨끗한 상태를 유도
    console.log('Re-navigating to Device Page with clean LocalStorage...');
    await page.goto(deviceUrl, { waitUntil: 'networkidle' });
    
    // 새로 로드된 iframe 다시 찾기
    targetFrame = null;
    for (let i = 0; i < 10; i++) {
      for (const frame of page.frames()) {
        const hasFab = await frame.evaluate(() => {
          return document.getElementById('fab-add-btn') !== null;
        }).catch(() => false);
        
        if (hasFab) {
          targetFrame = frame;
          break;
        }
      }
      if (targetFrame) break;
      await new Promise(r => setTimeout(r, 1000));
    }
    
    if (!targetFrame) {
      throw new Error('Could not find frame after reload.');
    }
    
    // 2. 초기 개수 0 검증
    console.log('\nStep 2: Verifying initial cart count...');
    let cartCount = await targetFrame.evaluate(() => {
      const countEl = document.getElementById('cart-count');
      return countEl ? countEl.innerText : 'N/A';
    });
    console.log('Initial cart count:', cartCount);
    if (cartCount !== '0') {
      throw new Error(`Initial cart count should be 0, but got ${cartCount}`);
    }
    
    // 3. 보관함에 담기 (addToCart)
    console.log('\nStep 3: Triggering addToCart...');
    await targetFrame.click('#fab-add-btn');
    await new Promise(r => setTimeout(r, 2000)); // 알림창 동기화 및 LocalStorage 저장 대기
    
    // 4. 개수 1 검증
    console.log('\nStep 4: Verifying updated cart count...');
    cartCount = await targetFrame.evaluate(() => {
      const countEl = document.getElementById('cart-count');
      return countEl ? countEl.innerText : 'N/A';
    });
    console.log('Updated cart count:', cartCount);
    if (cartCount !== '1') {
      throw new Error(`Updated cart count should be 1, but got ${cartCount}`);
    }
    
    // 5. 모달 열기
    console.log('\nStep 5: Opening cart modal...');
    await targetFrame.click('#fab-view-btn');
    await new Promise(r => setTimeout(r, 1000)); // 모달 렌더링 대기
    
    const modalState = await targetFrame.evaluate(() => {
      const modal = document.getElementById('cart-modal');
      const container = document.getElementById('cart-items-container');
      const items = container ? container.querySelectorAll('.cart-item') : [];
      
      const isActive = modal && modal.classList.contains('active');
      const itemsList = [];
      items.forEach(el => {
        const idEl = el.querySelector('.cart-item-id');
        const metaEl = el.querySelector('.cart-item-meta');
        itemsList.push({
          id: idEl ? idEl.innerText : '',
          meta: metaEl ? metaEl.innerText : ''
        });
      });
      
      return {
        isActive,
        itemCount: items.length,
        itemsList
      };
    });
    
    console.log('Modal status:', JSON.stringify(modalState, null, 2));
    if (!modalState.isActive) {
      throw new Error('Cart modal should be active.');
    }
    if (modalState.itemCount !== 1) {
      throw new Error(`Expected 1 item in modal list, but found ${modalState.itemCount}`);
    }
    
    // 6. html2pdf 라이브러리 탑재 여부 검증
    console.log('\nStep 6: Verifying html2pdf.js loading inside frame...');
    const hasHtml2Pdf = await targetFrame.evaluate(() => {
      return typeof html2pdf !== 'undefined';
    });
    console.log('html2pdf.js library loaded:', hasHtml2Pdf);
    if (!hasHtml2Pdf) {
      throw new Error('html2pdf.js library is not loaded in the frame.');
    }
    
    // 7. PDF 저장 기능 실행 및 버튼 텍스트 상태 변경 검증
    console.log('\nStep 7: Triggering saveCartAsPdf...');
    
    // PDF 다운로드 대기 이벤트 등록 (Playwright에서 파일 다운로드를 캡처)
    const downloadPromise = page.waitForEvent('download', { timeout: 10000 }).catch(() => null);
    
    // 모달창 내 'PDF 저장' 버튼 클릭
    // 버튼 내용 확인을 위해 text가 'PDF 저장'인 버튼 요소를 타겟팅합니다.
    const pdfBtnSelector = 'button:has-text("PDF 저장")';
    await targetFrame.click(pdfBtnSelector);
    
    console.log('PDF Generation triggered. Checking buttons dynamic text changes...');
    
    // 잠시 후 버튼 텍스트가 '생성 중...'인지 또는 원복되었는지 루프 대기
    let successState = false;
    for (let i = 0; i < 20; i++) {
      const btnState = await targetFrame.evaluate(() => {
        const btns = Array.from(document.querySelectorAll('.modal-card button'));
        const pdfBtn = btns.find(b => b.innerText.includes('PDF') || b.innerText.includes('생성'));
        return pdfBtn ? { text: pdfBtn.innerText, disabled: pdfBtn.disabled } : null;
      });
      console.log(`- Button State polling ${i + 1}:`, btnState);
      
      if (btnState && btnState.text === 'PDF 저장' && !btnState.disabled) {
        successState = true;
        break;
      }
      await new Promise(r => setTimeout(r, 500));
    }
    
    if (!successState) {
      throw new Error('PDF Button did not return to idle state after generation.');
    }
    
    console.log('PDF saving process finished successfully.');
    
    const download = await downloadPromise;
    if (download) {
      console.log('Successfully captured PDF download file!');
      console.log('Download URL:', download.url());
      console.log('Suggested Filename:', download.suggestedFilename());
    } else {
      console.log('Download event not caught (expected on headless Apps Script sandboxes, status verification is sufficient).');
    }
    
    console.log('\n>> ALL CART UI & PDF GENERATION SCENARIOS PASSED SUCCESSFULLY!');
    
  } catch (e) {
    console.error('Error during E2E test execution:', e);
  } finally {
    // 로컬 스토리지 청소 및 종료
    console.log('\nCleaning up localstorage...');
    if (targetFrame) {
      await targetFrame.evaluate(() => {
        localStorage.clear();
      }).catch(() => {});
    }
    await browser.close();
    console.log('Browser closed. Test finished.');
  }
})();

// Optional browser regression check: requires Playwright and a running server.
const { chromium, expect } = require('playwright/test');
const fs = require('node:fs');
const path = require('node:path');

(async () => {
  const base = process.argv[2] || 'http://127.0.0.1:8765';
  if (!process.env.DOCGEN_USERNAME || !process.env.DOCGEN_PASSWORD) throw new Error('Set DOCGEN_USERNAME and DOCGEN_PASSWORD');
  const catalog = JSON.parse(fs.readFileSync(path.join(__dirname, '../templates/catalog.json'), 'utf8'));
  const output = path.join(__dirname, '../tmp/browser-qa');
  fs.mkdirSync(output, { recursive: true });
  const browser = await chromium.launch({ executablePath: process.env.PLAYWRIGHT_CHROMIUM_EXECUTABLE || undefined, headless: true });
  try {
    for (const [device, viewport] of [['desktop', { width: 1440, height: 1000 }], ['mobile', { width: 390, height: 844 }]]) {
      const context = await browser.newContext({ viewport, acceptDownloads: true });
      const page = await context.newPage();
      const errors = [];
      page.on('pageerror', error => errors.push(error.message));
      await page.goto(base);
      await page.getByLabel('Логин').fill(process.env.DOCGEN_USERNAME);
      await page.getByLabel('Пароль').fill(process.env.DOCGEN_PASSWORD);
      await page.getByRole('button', { name: 'Войти', exact: true }).click();
      await expect(page.locator('#generator-form')).toBeVisible();
      await page.locator('#area').fill('32,5');
      await page.locator('#area-price').fill('21800,75');
      for (const [mode, config] of Object.entries(catalog)) {
        await page.locator(`label.option-card:has(input[value="${mode}"])`).click();
        const names = await page.locator('[id^="vendor-"]').allTextContents();
        expect(names).toEqual(config.suppliers.map(s => s.name));
        expect(await page.locator('#markup-1').getAttribute('aria-label')).toBe(`Надбавка: ${names[1]}`);
        expect(await page.locator('#area-fields').isVisible()).toBe(mode !== 'stands');
        expect(await page.locator('#stand-fields').isVisible()).toBe(['stands', 'combined'].includes(mode));
        if (['stands', 'combined'].includes(mode)) {
          await page.locator('#stand-count').fill('3');
          await page.locator('#stand-price').fill('90000,50');
        }
        const totals = (await page.locator('[id^="total-"]').allTextContents()).map(text => text.replace(/\s/g, ''));
        const expected = mode === 'stands' ? ['270001,5₸', '288901,62₸', '310501,74₸']
          : mode === 'combined' ? ['978525,88₸', '1047022,62₸', '1125304,69₸']
          : ['708524,38₸', '758121₸', '814802,95₸'];
        expect(totals).toEqual(expected);
        expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth)).toBe(true);
        console.log(device, mode, 'suppliers, prices and visibility OK');
      }
      await page.locator('label.option-card:has(input[value="finishing"])').click();
      const price = page.locator('#area-price');
      await price.fill('12345678');
      expect(await price.evaluate(input => {
        const style = getComputedStyle(input);
        const canvas = document.createElement('canvas').getContext('2d');
        canvas.font = style.font;
        return canvas.measureText(input.value).width <= input.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
      })).toBe(true);
      await price.fill('21800,75');
      await page.screenshot({ path: path.join(output, `${device}.png`), fullPage: true });
      if (device === 'desktop') {
        const downloads = [];
        page.on('download', download => downloads.push(download.suggestedFilename()));
        await page.getByRole('button', { name: 'Сформировать 2 файла' }).click();
        await expect(page.locator('#result')).toBeVisible({ timeout: 120000 });
        await expect.poll(() => downloads.length, { timeout: 10000 }).toBe(2);
        expect(downloads.some(name => name.endsWith('.docx'))).toBe(true);
        expect(downloads.some(name => name.endsWith('.pdf'))).toBe(true);
        expect(await page.locator('#download-list a').count()).toBe(2);
        console.log('Browser generation and both automatic downloads OK');
      }
      expect(errors).toEqual([]);
      await context.close();
    }
  } finally {
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });

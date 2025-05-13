import { test, expect } from '@playwright/test';

test('test', async ({ page }) => {
  await page.goto('https://www.google.com/sorry/index?continue=https://www.google.com/search%3Fq%3Dzebra%2Breporting%2Bportal%26oq%3Dzebra%2Breporting%2Bportal%26gs_lcrp%3DEgZjaHJvbWUyBggAEEUYOdIBCDgyNzFqMGoyqAIAsAIB%26sourceid%3Dchrome%26ie%3DUTF-8%26sei%3Dg4wKaPS5H43JptQPxJie8Qc&q=EgSuWV6MGIOZqsAGIjCqSIt1dhYeizy6aQCwrzvA7sO2SxSHqtPR3LZZNNOlTQXoQ1BfsEsw6zJhWAPXq0QyAnJSWgFD');
  await page.locator('iframe[name="a-61e8idc2np09"]').contentFrame().getByRole('checkbox', { name: 'I\'m not a robot' }).click();
  await page.locator('form').filter({ hasText: 'Select...Select A Report' }).locator('svg').click();
  await page.locator('#react-select-2-option-1-0').click();
  await page.getByRole('checkbox').check();
  await page.getByRole('button', { name: 'Generate Report' }).click();
});
class TimesheetPage {
  constructor(page) {
    this.page = page;
  }

  async open() {
    //await this.page.waitForURL(/\/home$/);

    await this.page.locator('text=Timesheets').click({ force: true });
    await this.page.waitForTimeout(10000);

    await this.page.locator('.tabs-container .tab').nth(1).click();
    await this.page.waitForTimeout(8000);

    // AG-Grid root wait
    await this.page.waitForSelector('.ag-root', { timeout: 15000 });
    console.log('All Timesheet LOADED');
  }


async applyAgGridSetFilter(filterName, values) {
  if (!values?.length) return;

  console.log(`Applying ${filterName}: ${values.join(', ')}`);

  try {
    // Open filter panel
    const sidebarBtn = this.page.locator('[ref="eSideBarButton"], .ag-icon-menu');
    await sidebarBtn.first().click({ force: true });

    const panel = this.page.locator('.ag-filter-toolpanel');
    await panel.waitFor({ timeout: 10000 });

    // Search filter name
    const searchInput = panel.locator('.ag-filter-toolpanel-search input');
    await searchInput.fill(filterName);
    await this.page.waitForTimeout(1000);

    const filterHeader = panel.locator(`.ag-group-title-bar:has-text("${filterName}")`);
    await filterHeader.click({ force: true });
    await this.page.waitForTimeout(1000);

    const filterGroup = filterHeader.locator('..');

    const miniFilter = filterGroup.locator(
      '.ag-mini-filter input[aria-label="Search filter values"]'
    );
    await miniFilter.waitFor({ timeout: 5000 });

    for (const value of values) {
      console.log(`   🔍 Searching value: ${value}`);

      await miniFilter.fill(value);

      const loading = filterGroup.locator('.ag-filter-loading');
      await loading.waitFor({ state: 'hidden', timeout: 5000 });

      const valueRow = filterGroup
        .locator('.ag-virtual-list-item[role="option"]')
        .filter({ hasText: value })
        .first();

      await valueRow.waitFor({ timeout: 5000 });
      await valueRow.click({ force: true });

      console.log(`   ✅ ${value} SELECTED`);
    }

    const applyBtn = filterGroup.locator('button:has-text("Apply")');
    if (await applyBtn.isVisible()) {
      await applyBtn.click({ force: true });
    }

    console.log(`✅ ${filterName} APPLIED`);

  } catch (err) {
    console.log(`❌ ${filterName} FAILED: ${err.message}`);
  } finally {
    // 🔧 CHANGE 6: Close panel after work done
    await this.page.keyboard.press('Escape');
    await this.page.waitForTimeout(500);
  }
}


async applyAgGridDateFilter(filterName, operator, startDate, endDate) {
  console.log(`📅 Applying Date Filter → ${filterName} | ${operator}`);

  try {
    // 1️⃣ Open Filters panel
    const sidebarBtn = this.page.locator('[ref="eSideBarButton"], .ag-icon-menu');
    await sidebarBtn.first().click({ force: true });

    const panel = this.page.locator('.ag-filter-toolpanel');
    await panel.waitFor({ timeout: 10000 });

    // 2️⃣ Search filter
    const searchInput = panel.locator('.ag-filter-toolpanel-search input');
    await searchInput.fill(filterName);
    await this.page.waitForTimeout(500);

    // 3️⃣ Open Date filter group
    const filterHeader = panel.locator(
      `.ag-group-title-bar:has-text("${filterName}")`
    );
    await filterHeader.click({ force: true });

    const filterGroup = filterHeader.locator('..');

    // 4️⃣ Focus operator dropdown
    const operatorWrapper = filterGroup.locator(
      '.ag-filter-select .ag-picker-field-wrapper'
    );
    await operatorWrapper.waitFor();
    await operatorWrapper.click();

    // 🔥 KEYBOARD BASED SELECTION (MOST STABLE)
    const operatorOrder = [
      'Equals',
      'Does not equal',
      'Before',
      'After',
      'Between',
      'Blank',
      'Not blank'
    ];

    const targetIndex = operatorOrder.indexOf(operator);
    if (targetIndex === -1) {
      throw new Error(`Unsupported operator: ${operator}`);
    }

    // Reset to first option
    await this.page.keyboard.press('Home');
    await this.page.waitForTimeout(200);

    // Move down to desired option
    for (let i = 0; i < targetIndex; i++) {
      await this.page.keyboard.press('ArrowDown');
      await this.page.waitForTimeout(150);
    }

    await this.page.keyboard.press('Enter');
    console.log(`   🔽 Operator selected: ${operator}`);

console.log('🔍 Finding date inputs...');
await this.page.waitForTimeout(1000);  // Wait for Between inputs to render

// 🔥 EXACT SELECTORS FROM YOUR HTML
const dateInput1 = filterGroup.locator('input[placeholder="yyyy-mm-dd"]').nth(0);
const dateInput2 = filterGroup.locator('input[placeholder="yyyy-mm-dd"]').nth(1);

console.log(`📝 Input 1 count: ${(await dateInput1.count())}`);
console.log(`📝 Input 2 count: ${(await dateInput2.count())}`);
// 5️⃣ BYPASS AG-GRID UI - Direct Filter Model Set
console.log('🔧 Setting AG-Grid filter model directly...');

if (operator === 'Between') {
const input1 = filterGroup.locator('.ag-filter-from input[placeholder*="yyyy"]');
await input1.evaluate((el, date) => {
  el.type = 'text';
  el.value = date;  // ← DYNAMIC: this.convertToISO(startDate)
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
}, this.convertToISO(startDate));  // ← EXCEL startDate
console.log(`✅ FROM input SET: ${this.convertToISO(startDate)}`);

// 🔥 INPUT 2: ag-filter-to 
const input2 = filterGroup.locator('.ag-filter-to input[placeholder*="yyyy"]');
await input2.waitFor({ timeout: 1000 });
await input2.evaluate((el, date) => {
  el.type = 'text';
  el.value = date;  // ← DYNAMIC: this.convertToISO(endDate)
  el.dispatchEvent(new Event('input', { bubbles: true }));
  el.dispatchEvent(new Event('change', { bubbles: true }));
}, this.convertToISO(endDate));  // ← EXCEL endDate
console.log(`✅ TO input SET: ${this.convertToISO(endDate)}`);
  
  console.log(`   📅 HARDCODE TEST: 2025-12-01 → 2025-12-09`);
} else if (!['Blank', 'Not blank'].includes(operator)) {
  await this.page.evaluate((filterName, dateValue) => {
    const gridApi = window.agGrid.api;
    if (gridApi) {
      gridApi.setFilterModel({
        [filterName]: {
          filterType: 'date',
          type: operator.toLowerCase(),
          dateFrom: dateValue
        }
      });
      gridApi.onFilterChanged();
    }
  }, filterName, this.convertToISO(startDate));
}

await this.page.waitForTimeout(2000);  // Wait for filter to apply
console.log(`✅ Date Filter Applied via AG-Grid API`);



    // 6️⃣ Apply
    const applyBtn = filterGroup.locator('button:has-text("Apply")');
    await applyBtn.click({ force: true });

    console.log(`✅ Date Filter Applied Successfully`);
  } catch (err) {
    console.log(`❌ Date Filter FAILED: ${err.message}`);
  } finally {
    await this.page.keyboard.press('Escape');
    await this.page.waitForTimeout(500);
  }
}



convertToISO(dateValue) {
  if (!dateValue) return '';

  // ✅ Case 1: Proper JS Date object
  if (dateValue instanceof Date) {
    const yyyy = dateValue.getFullYear();
    const mm = String(dateValue.getMonth() + 1).padStart(2, '0');
    const dd = String(dateValue.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }

  // ✅ Case 2: Excel string → 01-12-2025
  if (typeof dateValue === 'string' && /^\d{2}-\d{2}-\d{4}$/.test(dateValue)) {
    const [dd, mm, yyyy] = dateValue.split('-');
    return `${yyyy}-${mm}-${dd}`;
  }

  // ✅ Case 3: JS Date string → Mon Dec 01 2025 05:30:00 GMT+0530
  if (typeof dateValue === 'string') {
    const parsed = new Date(dateValue);
    if (!isNaN(parsed.getTime())) {
      const yyyy = parsed.getFullYear();
      const mm = String(parsed.getMonth() + 1).padStart(2, '0');
      const dd = String(parsed.getDate()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}`;
    }
  }

  throw new Error(`Unsupported date format: ${dateValue}`);
}





async exportData() {
  console.log('Exporting CSV...');

  const exportBtn = this.page.locator('text=Export');

  const [download] = await Promise.all([
    this.page.waitForEvent('download'), // CSV download
    exportBtn.click({ force: true })
  ]);

  console.log('CSV Download triggered:', await download.suggestedFilename());
  return download;
}




}

module.exports = { TimesheetPage };

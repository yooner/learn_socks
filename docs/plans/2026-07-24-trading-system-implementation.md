# Excel Trading System Implementation Plan

**Goal:** Generate a clean Excel trading-management template, a synthetic-data test workbook, and a Markdown verification report in the repository root.

**Architecture:** A Python generator built with `openpyxl` creates five styled worksheets with prefilled formulas, validations, tables, and conditional formatting. Pure calculation helpers define the metric contract, while workbook integration tests use both `openpyxl` structure inspection and LibreOffice headless recalculation to verify cached formula results.

**Tech Stack:** Python 3, openpyxl, unittest, LibreOffice headless

---

### Task 1: Establish dependencies and calculation contract

**Files:**
- Create: `requirements.txt`
- Create: `tests/test_trading_workbook.py`
- Create: `trading_workbook.py`

**Step 1: Add a dependency declaration**

Declare `openpyxl>=3.1,<4`.

**Step 2: Write failing unit tests**

Create tests for:

- position sizing rounds down to 100-share lots
- invalid stop prices return no position
- realized P&L includes both fees
- win/loss/flat/open transactions aggregate at the correct grain
- average loss is stored as a positive magnitude
- expectancy subtracts the loss contribution
- compound return uses win and loss counts
- compound-loss ceiling handles missing and boundary win rates
- target trade count rounds upward and rejects non-positive expectancy

**Step 3: Run tests to verify RED**

Run: `python3 -m unittest tests.test_trading_workbook.CalculationTests -v`

Expected: FAIL because calculation helpers do not exist.

**Step 4: Implement minimal pure calculation helpers**

Add functions for position size, P&L, return rate, summary metrics, compound-loss ceiling, account balance, monthly realized loss, and required target trades.

**Step 5: Run tests to verify GREEN**

Run: `python3 -m unittest tests.test_trading_workbook.CalculationTests -v`

Expected: all calculation tests pass.

**Step 6: Commit**

Commit dependency, tests, and calculation helpers.

### Task 2: Generate the workbook structure

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing workbook-structure tests**

Assert that the generated workbook has exactly these sheets:

- `单次交易`
- `买入理由`
- `多次统计数据`
- `账户数据`
- `目标收益`

Also assert required headers, formulas, formula/input fills, frozen panes, filters, tables, data validations, conditional formatting, percentage/date/currency formats, and automatic calculation settings.

**Step 2: Run tests to verify RED**

Run: `python3 -m unittest tests.test_trading_workbook.WorkbookStructureTests -v`

Expected: FAIL because workbook generation is missing.

**Step 3: Implement workbook generation**

Create five sheets, 200 preformatted transaction rows, 500 reason-log rows, a statistics panel, an account panel, and a target-return panel. Add:

- per-trade account and risk-rate snapshots
- formulas guarded against blank inputs and zero denominators
- stage dropdowns
- numeric/date/custom validations
- green/red result formatting and risk warnings
- Chinese instructions and metric definitions

**Step 4: Run tests to verify GREEN**

Run: `python3 -m unittest tests.test_trading_workbook.WorkbookStructureTests -v`

Expected: all structure tests pass.

**Step 5: Commit**

Commit workbook construction and structure tests.

### Task 3: Add synthetic integration data and recalculation checks

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing integration tests**

Use six synthetic trades covering two wins, two losses, one flat close, and one open position. Include buy/sell fees and different holding periods. Assert recalculated values for:

- each position size, P&L, return rate, and holding period
- completed/win/loss/flat counts
- win and loss probability
- average win/loss, expectancy, and compound return
- current account balance and negative current-month loss
- target profit and rounded-up required trade count

**Step 2: Run tests to verify RED**

Run: `python3 -m unittest tests.test_trading_workbook.IntegrationTests -v`

Expected: FAIL because sample generation or cached formula calculation is missing.

**Step 3: Implement sample generation and LibreOffice recalculation**

Add deterministic synthetic fixtures and a CLI that writes:

- `交易管理系统.xlsx`
- `交易管理系统_测试版.xlsx`

Use LibreOffice headless in the test to open and save the test workbook, then inspect formula results with `data_only=True`.

**Step 4: Run tests to verify GREEN**

Run: `python3 -m unittest tests.test_trading_workbook.IntegrationTests -v`

Expected: all integration tests pass.

**Step 5: Commit**

Commit sample generation and integration checks.

### Task 4: Generate deliverables and report

**Files:**
- Create: `交易管理系统.xlsx`
- Create: `交易管理系统_测试版.xlsx`
- Create: `交易管理系统测试报告.md`
- Modify: `docs/plans/2026-07-24-trading-system-design.md`

**Step 1: Update the design**

Record that each transaction snapshots both account balance and risk percentage so later account-setting changes cannot alter historical position sizes.

**Step 2: Generate both workbooks**

Run: `python3 trading_workbook.py`

Expected: both `.xlsx` files are created.

**Step 3: Recalculate the test workbook**

Run LibreOffice headless against a temporary output directory and replace the test copy with the recalculated result.

**Step 4: Run the full verification suite**

Run: `python3 -m unittest discover -v`

Expected: all tests pass.

**Step 5: Inspect the final files**

Run:

- `soffice --headless --convert-to pdf --outdir <temp> 交易管理系统_测试版.xlsx`
- render or inspect representative PDF pages
- `git status --short`

Expected: the workbook is readable, sheets are laid out correctly, and only intended files are changed.

**Step 6: Write the test report**

Document environment, synthetic cases, expected values, actual recalculated values, and pass/fail status.

**Step 7: Commit**

Commit the final deliverables, report, updated design, and implementation plan.

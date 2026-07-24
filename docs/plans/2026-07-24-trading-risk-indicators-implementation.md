# Trading Risk and Indicator Improvements Implementation Plan

**Goal:** Add indicator dropdowns backed by a new worksheet, separate suggested and actual buy shares, add portfolio-level monthly risk warnings, and prove the workbook through 100 sequential append/recalculate/inspect cycles.

**Architecture:** Keep `trading_workbook.py` as the single workbook generator and calculation-contract module. Expand the trade sheet from 28 to 30 columns, add a named range backed by a sixth worksheet, and expose pure-Python risk helpers so Excel formulas and independent test expectations share the same metric definitions. A separate progressive validator will append exactly one trade per step, invoke LibreOffice after every append, inspect cached formula values, and publish the final 100-row workbook plus an auditable Markdown report.

**Tech Stack:** Python 3, openpyxl 3.1, unittest, LibreOffice headless

---

### Task 1: Lock the new calculation contract with failing tests

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing pure calculation tests**

Add tests for:

```python
def test_open_risk_uses_only_unsold_actual_shares():
    trades = [
        {"buy_price": 10, "stop_price": 9, "actual_buy_shares": 500, "sell_date": None},
        {"buy_price": 20, "stop_price": 18, "actual_buy_shares": 200, "sell_date": date(2026, 7, 20)},
    ]
    self.assertEqual(tw.calculate_open_theoretical_loss(trades), 500)

def test_risk_status_includes_candidate_without_double_counting_actual_position():
    self.assertEqual(tw.calculate_opening_risk_status(4_100, 900, 5_000), "禁止开仓")
    self.assertEqual(tw.calculate_opening_risk_status(4_000, 900, 5_000), "允许开仓")
    self.assertEqual(tw.calculate_opening_risk_status(5_100, 0, 5_000), "禁止开仓")
```

The threshold is inclusive: total risk equal to the maximum is prohibited.

**Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.CalculationTests -v
```

Expected: errors because the two risk helpers do not exist.

**Step 3: Implement minimal pure helpers**

Implement:

```python
def calculate_open_theoretical_loss(trades):
    return sum(
        max(float(trade["buy_price"]) - float(trade["stop_price"]), 0)
        * int(trade["actual_buy_shares"])
        for trade in trades
        if trade.get("sell_date") is None
        and trade.get("actual_buy_shares")
        and trade.get("buy_price") is not None
        and trade.get("stop_price") is not None
    )

def calculate_opening_risk_status(open_risk, candidate_risk, monthly_limit):
    if monthly_limit is None or monthly_limit <= 0:
        return None
    return "禁止开仓" if open_risk + candidate_risk >= monthly_limit else "允许开仓"
```

**Step 4: Run tests and verify GREEN**

Run the CalculationTests command again and require zero failures.

### Task 2: Add the technical-indicator worksheet and dropdown contract

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing workbook-structure tests**

Assert:

```python
self.assertEqual(
    workbook.sheetnames,
    ["单次交易", "买入理由", "多次统计数据", "账户数据", "目标收益", "技术指标"],
)
self.assertEqual(workbook["技术指标"]["A2"].value, "蜡烛图")
self.assertIn("MACD", [workbook["技术指标"].cell(row, 1).value for row in range(2, 30)])
self.assertIn("技术指标列表", workbook.defined_names)
indicator_validations = [
    item for item in workbook["买入理由"].data_validations.dataValidation
    if item.formula1 == "=技术指标列表"
]
self.assertEqual(len(indicator_validations), 1)
self.assertIn("F2:H501", str(indicator_validations[0].sqref))
```

**Step 2: Run the structure test and verify RED**

Expected: only five worksheets and no named range/dropdown.

**Step 3: Implement the worksheet**

Add `TECHNICAL_INDICATORS` rows with at least:

`蜡烛图、趋势线、MACD、移动平均线、成交量、RSI、KDJ、布林带、支撑位、压力位、缺口、形态突破、均线金叉、均线死叉、量价背离`.

Create `技术指标` with columns `技术指标、分类、说明`, table name `TechnicalIndicatorCatalog`, and a workbook name:

```python
DefinedName("技术指标列表", attr_text="'技术指标'!$A$2:$A$201")
```

Apply one list validation with `formula1="=技术指标列表"` to `F2:H501`.

**Step 4: Run the structure test and verify GREEN**

Require the named range, three-column dropdown coverage, and six-sheet order.

### Task 3: Migrate the trade-sheet columns and formulas

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing formula and style tests**

Use this exact column contract:

```text
A 交易编号             P 卖出费用
B 股票代码             Q 买入价的由来
C 买入时账户金额       R 止损价的由来
D 本次允许亏损比例     S 期望卖出价的由来
E 买入价               T 实际卖出价的由来
F 买入建议股数         U 期望盈利比例
G 开仓风险告警         V 期望止损比例
H 实际买入股数         W 盈亏比
I 买入日期             X 实际盈亏金额
J 期望卖出价           Y 实际收益率
K 实际卖出价           Z 与平均盈利比例差值
L 卖出股数             AA 持有天数
M 卖出日期             AB 复利容许平均亏损上限
N 止损价               AC 复利风险判断
O 买入费用             AD 交易打分评价
```

Assert `F2` contains `ROUNDDOWN`, `H2` is an input cell, `L2` defaults to `H2`, and `X2`, `Y2`, and the statistics average-trade formula use `H`, never `F`.

**Step 2: Run structure tests and verify RED**

Expected: header mismatch and formulas still reference the old automatic buy-share column.

**Step 3: Implement the column migration**

Use the following formula rules:

```excel
F2 = IF(valid snapshots and E2>N2, ROUNDDOWN((C2*D2)/(E2-N2)/100,0)*100, "")
G2 = IF(required inputs missing, "",
        IF('账户数据'!$B$9 + IF(H2="",F2*(E2-N2),0) >= '账户数据'!$B$7,
           "禁止开仓","允许开仓"))
L2 = IF(K2="","",H2)
X2 = IF(required sell and actual-share inputs missing, "",
        K2*L2-E2*H2-IF(O2="",0,O2)-IF(P2="",0,P2))
Y2 = X2/(E2*H2+buy fee)
```

Move all date, percentage, currency, input-fill, validation, conditional-formatting, table and header-comment references to the new columns. Add red/green conditional formatting to `G2:G201`.

**Step 4: Run all structure tests and verify GREEN**

Require no references to the suggested-share column in realized or aggregate metrics.

### Task 4: Add account-level open-risk visibility

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Step 1: Write failing account-formula tests**

Assert:

```python
self.assertEqual(account["A6"].value, "当月允许最大亏损比例")
self.assertEqual(account["A7"].value, "当月允许最大亏损金额")
self.assertEqual(account["A9"].value, "当前未平仓理论亏损")
self.assertEqual(account["A10"].value, "当月剩余可开仓风险额度")
self.assertIn("SUMPRODUCT", account["B9"].value)
self.assertIn("'单次交易'!H2:H201", account["B9"].value)
self.assertEqual(account["B10"].value, '=IF(OR(B7="",B9=""),"",B7-B9)')
```

**Step 2: Run and verify RED**

Expected: missing rows and old cumulative-risk labels.

**Step 3: Implement the account formulas**

Use:

```excel
B9 = IFERROR(
       SUMPRODUCT(('单次交易'!M2:M201="")*
                  ('单次交易'!H2:H201>0)*
                  ('单次交易'!E2:E201>'单次交易'!N2:N201)*
                  '单次交易'!H2:H201*
                  ('单次交易'!E2:E201-'单次交易'!N2:N201)),0)
B10 = IF(OR(B7="",B9=""),"",B7-B9)
```

Format B9/B10 as currency and B10 red when negative.

**Step 4: Run and verify GREEN**

Confirm account and trade-warning formulas reference the same open-risk and monthly-limit cells.

### Task 5: Build and test the 100-step progressive scenario

**Files:**
- Create: `progressive_workbook_validation.py`
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`
- Modify: `交易管理系统测试报告.md`

**Step 1: Write failing generator and append tests**

Assert that `generate_sample_transactions(as_of_date, 100)` returns 100 deterministic, unique trades containing:

- wins, losses, flats, and open positions
- actual shares below suggested shares
- fees and varied holding periods
- all preset indicator categories
- enough simultaneous open risk to produce both `允许开仓` and `禁止开仓`

For each prefix from 1 through 100, append one trade to a clean workbook and compare pure-Python expected counts, P&L, balance, open risk, and risk status.

**Step 2: Run and verify RED**

Expected: missing 100-row generator and append API.

**Step 3: Implement deterministic sample generation and one-row append**

Create:

```python
generate_sample_transactions(as_of_date: date, count: int = 100)
append_trade_to_workbook(wb: Workbook, item: Mapping[str, Any], row: int)
append_reason_to_workbook(wb: Workbook, item: Mapping[str, Any], row: int)
```

Actual shares must be explicit multiples of 100 between 20% and 100% of the suggested position. The final sample workbook is created by repeatedly calling the append functions, never by assigning a complete matrix.

**Step 4: Implement the LibreOffice progressive validator**

For each step 1..100:

1. append exactly one trade and one reason row
2. save the current workbook
3. convert it with headless LibreOffice to force formula recalculation
4. reopen with `data_only=True`
5. compare the new row and aggregate cells against independent Python calculations
6. append an audit row containing step, trade id, completed/win/loss/open counts, suggestion, actual shares, P&L, balance, open risk, monthly limit, risk status, and PASS
7. print progress immediately

Use the recalculated 100th workbook as `交易管理系统_测试版.xlsx` and write all 100 audit rows into `交易管理系统测试报告.md`.

**Step 5: Run the 100-step validator**

Run:

```bash
.venv/bin/python progressive_workbook_validation.py \
  --output-dir . \
  --as-of-date 2026-07-24 \
  --steps 100
```

Expected: 100 progress lines, 100 PASS audit rows, zero Excel errors, and regenerated test workbook/report.

### Task 6: Regenerate deliverables and perform full verification

**Files:**
- Modify: `交易管理系统.xlsx`
- Modify: `交易管理系统_测试版.xlsx`
- Modify: `交易管理系统测试报告.md`

**Step 1: Generate the clean workbook**

Run:

```bash
.venv/bin/python trading_workbook.py
```

**Step 2: Run the complete test suite**

Run:

```bash
.venv/bin/python -m unittest discover -v
```

Expected: all tests pass with zero failures.

**Step 3: Verify file integrity**

Run:

```bash
unzip -t 交易管理系统.xlsx
unzip -t 交易管理系统_测试版.xlsx
```

Expected: no archive errors.

**Step 4: Export representative worksheets to PDF and inspect**

Use headless LibreOffice to export the final sample workbook, render representative pages, and verify:

- no clipped or overlapping headers
- indicator dropdown columns and risk warning remain legible
- no `#VALUE!`, `#DIV/0!`, `#REF!`, or `#NAME?`
- red/green warning states are visually distinguishable

**Step 5: Commit and integrate**

Commit implementation and regenerated artifacts, rerun the complete verification command, then merge the feature branch into `main` and remove the worktree.

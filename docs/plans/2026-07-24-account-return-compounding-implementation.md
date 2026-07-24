# Account Return Compounding Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve position-level returns while adding account-level returns and use the latter for expectancy, exact trade-by-trade compounding, account validation, and target-trade estimates.

**Architecture:** Keep workbook generation and pure metric contracts in `trading_workbook.py`. Add one calculated trade column at `AE`, compute account statistics from that column, and expose pure helpers so integration tests independently reproduce Excel results. Update the progressive validator to carry the same account-return field.

**Tech Stack:** Python 3, openpyxl 3.1, unittest, LibreOffice headless

## Global Constraints

- Preserve all six existing worksheets and current risk controls.
- Keep the existing position-return formula for trade review.
- Use the manually saved buy-time account snapshot as the account-return denominator.
- Do not keep generated test workbooks in the repository.

---

### Task 1: Add account-return calculation contracts

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Interfaces:**
- Produces: `calculate_account_return(pnl: float, account_snapshot: float) -> float | None`
- Produces: `calculate_compound_return(account_returns: Iterable[float | None]) -> float | None`
- Updates: `summarize_trades(trades)` to consume `account_return`

- [ ] **Step 1: Write failing pure calculation tests**

```python
def test_account_return_uses_pre_trade_total_equity(self):
    self.assertEqual(calculate_account_return(1_000, 100_000), 0.01)
    self.assertAlmostEqual(
        calculate_account_return(-2_500, 101_000),
        -2_500 / 101_000,
    )
    self.assertIsNone(calculate_account_return(100, 0))

def test_compound_return_multiplies_actual_account_returns(self):
    returns = [0.01, -2_500 / 101_000]
    self.assertAlmostEqual(calculate_compound_return(returns), -0.015)
    self.assertIsNone(calculate_compound_return([None]))
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.CalculationTests -v
```

Expected: import failure because the two helpers do not exist.

- [ ] **Step 3: Implement the minimal helpers and summary change**

```python
def calculate_account_return(pnl, account_snapshot):
    if account_snapshot <= 0:
        return None
    return pnl / account_snapshot

def calculate_compound_return(account_returns):
    factor = 1.0
    count = 0
    for value in account_returns:
        if value is None:
            continue
        factor *= 1 + value
        count += 1
    return factor - 1 if count else None
```

Change `summarize_trades` so `average_win`, `average_loss`, `expectancy`, and `compound_return` use each record's `account_return`; calculate `compound_return` from the ordered closed records instead of powers of conditional averages.

- [ ] **Step 4: Run calculation tests and verify GREEN**

Run the command from Step 2 and require zero failures.

### Task 2: Add the workbook account-return column and formulas

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Interfaces:**
- Consumes: `单次交易!C` account snapshot and `单次交易!X` realized P&L
- Produces: `单次交易!AE` actual account return

- [ ] **Step 1: Write failing workbook-structure assertions**

Assert the 31-column contract, `Y1 == "单笔仓位收益率"`, `AE1 == "实际账户收益率"`, and:

```python
self.assertEqual(
    trade["AE2"].value,
    '=IF(OR(X2="",C2="",C2<=0),"",X2/C2)',
)
self.assertIn("'单次交易'!AE2:AE201", stats["B8"].value)
self.assertIn(
    "SUMPRODUCT(IFERROR(LN(1+'单次交易'!AE2:AE201),0))",
    stats["B14"].value,
)
```

- [ ] **Step 2: Run the structure test and verify RED**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.WorkbookStructureTests -v
```

Expected: the `AE` column and account-return statistics are absent.

- [ ] **Step 3: Implement the workbook formulas**

- Append `实际账户收益率` to `TRADE_HEADERS`.
- Rename `实际收益率` to `单笔仓位收益率`.
- Rename the difference column and calculate it from `AE`.
- Expand `TradeRecords` from `A:AD` to `A:AE`.
- Format and conditionally color `AE` as a percentage.
- Change statistics `B8:B9` to average `AE`.
- Change `B14` to the logarithmic equivalent of a product, which avoids range implicit intersection in Excel and LibreOffice:

```excel
=IF(COUNT('单次交易'!AE2:AE201)=0,"",
 EXP(SUMPRODUCT(IFERROR(LN(1+'单次交易'!AE2:AE201),0)))-1)
```

- [ ] **Step 4: Run structure tests and verify GREEN**

Require zero failures.

### Task 3: Add account validation and coherent target math

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `trading_workbook.py`

**Interfaces:**
- Produces: `账户数据!B11` actual cumulative account return
- Updates: `目标收益!B7` to consume target rate and account expectancy

- [ ] **Step 1: Write failing assertions**

```python
self.assertEqual(account["A11"].value, "账户实际累计收益率")
self.assertEqual(
    account["B11"].value,
    '=IF(OR(B2="",B3="",B2<=0),"",B3/B2-1)',
)
self.assertIn("LN(1+B3)/LN(1+B6)", target["B7"].value)
self.assertNotIn("B5*B6", target["B7"].value)
```

- [ ] **Step 2: Verify RED**

Run the structure test command and confirm failures only cover the missing formulas.

- [ ] **Step 3: Implement the account and target formulas**

Append the account cumulative-return row with percentage format. Replace the target count formula with:

```excel
=IF(OR(B3="",B6="",B3<=0,B6<=0),"暂不可计算",
 ROUNDUP(LN(1+B3)/LN(1+B6),0))
```

- [ ] **Step 4: Verify GREEN**

Run the structure tests and require zero failures.

### Task 4: Update sample metrics and integration verification

**Files:**
- Modify: `tests/test_trading_workbook.py`
- Modify: `progressive_workbook_validation.py`
- Modify: `trading_workbook.py`

**Interfaces:**
- Consumes: sample `account_snapshot`
- Produces: `account_return` in sample metric dictionaries

- [ ] **Step 1: Extend integration expectations**

Check every closed row's `AE` against `pnl / account_snapshot`, confirm blank rows stay blank, confirm statistics `B14` matches pure `calculate_compound_return`, and confirm `账户数据!B11 == 当前金额 / 初始金额 - 1`.

- [ ] **Step 2: Run integration tests and verify RED**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.IntegrationTests -v
```

Expected: account-return fields and formulas are missing.

- [ ] **Step 3: Update sample and progressive metric records**

Add:

```python
account_return = calculate_account_return(pnl, item["account_snapshot"])
```

to `sample_trade_metrics` and the progressive validator's `_metric_record`, and include `account_snapshot` and `account_return` in returned records.

- [ ] **Step 4: Run the complete test suite**

Run:

```bash
.venv/bin/python -m unittest -v
```

Expected: all tests pass, including LibreOffice recalculation.

### Task 5: Regenerate, inspect, commit, and push

**Files:**
- Modify: `交易管理系统.xlsx`
- Verify: all tracked files

- [ ] **Step 1: Regenerate the clean workbook**

Run:

```bash
.venv/bin/python trading_workbook.py .
```

Confirm only `交易管理系统.xlsx` is retained as the generated deliverable.

- [ ] **Step 2: Run fresh verification**

Run:

```bash
.venv/bin/python -m unittest -v
git diff --check
git status --short
```

Expected: zero test failures, no whitespace errors, and only intended files changed.

- [ ] **Step 3: Commit and push**

```bash
git add trading_workbook.py progressive_workbook_validation.py tests/test_trading_workbook.py docs/plans/2026-07-24-account-return-compounding-design.md docs/plans/2026-07-24-account-return-compounding-implementation.md 交易管理系统.xlsx
git commit -m "feat: calculate compounding from account returns"
git push origin main
```

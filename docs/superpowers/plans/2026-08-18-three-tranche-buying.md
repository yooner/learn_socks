# Fixed Three-Tranche Buying Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a no-macro, fixed three-batch buying workflow to the trading workbook, with one-time full sale and strict lock/single-trade/account risk controls.

**Architecture:** Extend “单次交易” from AJ to AR and keep each position on one row. Pure Python helpers define the batch arithmetic and rule ordering; Excel formulas mirror those helpers, while the existing tracking, account, statistics, lock, and expectation modules consume cumulative shares and weighted cost. Upgrade V4 into a separate V5 file without modifying historical input cells.

**Tech Stack:** Python 3, openpyxl 3.1.5, `unittest`, Excel `.xlsx` formulas/data validation/conditional formatting

**Spec:** `docs/superpowers/specs/2026-08-18-three-tranche-buying-design.md`

## Global Constraints

- Keep the daily workflow in “单次交易”; do not add a worksheet.
- Support at most three buy batches and one full-position sale.
- Use no VBA, Office Script, or macro-enabled workbook format.
- Treat the existing E/H/I/O fields as first price, first shares, first date, and total buy fees.
- Require 100-share lots, complete price/share pairs, and batch 2 before batch 3.
- Forbid additions while `_LockStatus` is `已锁仓` or `解锁信息不完整`.
- Reject a completed batch when position risk exceeds `C×D` or all open risk exceeds `账户数据!B7`.
- Compute risk from the highest effective stop in `持仓跟踪!S`, falling back to N, and never count negative risk.
- Preserve the existing 100 historical trades and produce `交易管理系统_V5_三批买入.xlsx` from V4.

## File Structure

- Modify `trading_workbook.py`: pure calculations, workbook formulas, validations, V5 upgrade entry point, and sample integration.
- Modify `tests/test_trading_workbook.py`: arithmetic, workbook structure, validation, formula dependency, compatibility, and upgrade tests.
- Modify `progressive_workbook_validation.py`: validate progressive samples against cumulative position values.
- Create `交易管理系统_V5_三批买入.xlsx`: final user-facing workbook artifact.

---

### Task 1: Define and test the three-batch calculation contract

**Files:**
- Modify: `tests/test_trading_workbook.py:69-160`
- Modify: `trading_workbook.py:103-156`

**Interfaces:**
- Produces: `calculate_tranche_position(tranches, effective_stop, is_closed=False) -> dict[str, float | int | None]`
- Produces: `check_tranche_rules(trade_id, tranches, is_locked, position_risk, single_trade_limit, total_open_risk, account_risk_limit) -> str`
- Consumes: no workbook state; inputs are plain Python values.

- [ ] **Step 1: Write failing arithmetic tests**

Add tests that establish single-, double-, and triple-batch behavior:

```python
def test_tranche_position_calculates_weighted_cost_shares_and_risk(self):
    result = tw.calculate_tranche_position(
        [(10.0, 1000), (8.0, 500), (12.0, 500)],
        effective_stop=9.0,
    )
    self.assertEqual(result["total_shares"], 2000)
    self.assertAlmostEqual(result["buy_amount"], 20_000.0)
    self.assertAlmostEqual(result["weighted_buy_price"], 10.0)
    self.assertAlmostEqual(result["current_risk"], 2500.0)

def test_tranche_position_treats_stop_above_a_batch_as_zero_risk(self):
    result = tw.calculate_tranche_position(
        [(10.0, 100), (8.0, 100), (None, None)],
        effective_stop=9.0,
    )
    self.assertEqual(result["current_risk"], 100.0)

def test_closed_tranche_position_has_zero_current_risk(self):
    result = tw.calculate_tranche_position(
        [(10.0, 100), (8.0, 100), (None, None)],
        effective_stop=7.0,
        is_closed=True,
    )
    self.assertEqual(result["current_risk"], 0.0)
```

- [ ] **Step 2: Run the arithmetic tests and verify failure**

Run:

```bash
.venv/bin/python -m unittest \
  tests.test_trading_workbook.CalculationTests.test_tranche_position_calculates_weighted_cost_shares_and_risk \
  tests.test_trading_workbook.CalculationTests.test_tranche_position_treats_stop_above_a_batch_as_zero_risk \
  tests.test_trading_workbook.CalculationTests.test_closed_tranche_position_has_zero_current_risk -v
```

Expected: all three tests fail with `AttributeError: module 'trading_workbook' has no attribute 'calculate_tranche_position'`.

- [ ] **Step 3: Implement the minimal batch calculator**

Add this function beside the existing pure calculation functions:

```python
def calculate_tranche_position(
    tranches: Iterable[tuple[float | None, int | None]],
    effective_stop: float | None,
    is_closed: bool = False,
) -> dict[str, float | int | None]:
    valid = [
        (float(price), int(shares))
        for price, shares in tranches
        if price is not None and shares is not None and shares > 0
    ]
    total_shares = sum(shares for _, shares in valid)
    buy_amount = sum(price * shares for price, shares in valid)
    weighted = buy_amount / total_shares if total_shares else None
    stop = float(effective_stop) if effective_stop is not None else None
    risk = 0.0 if is_closed or stop is None else sum(
        max(price - stop, 0.0) * shares for price, shares in valid
    )
    return {
        "total_shares": total_shares,
        "buy_amount": buy_amount,
        "weighted_buy_price": weighted,
        "current_risk": risk,
    }
```

- [ ] **Step 4: Add failing rule-priority tests**

Cover every rule outcome and the exact Chinese strings used later by AR:

```python
def test_tranche_rules_enforce_pair_order_lots_lock_and_risk(self):
    cases = [
        (None, [(None, None)] * 3, False, 0, 1000, 0, 5000, ""),
        ("T1", [(10, None), (None, None), (None, None)], False, 0, 1000, 0, 5000, "违规：第一批价格或股数缺失"),
        ("T1", [(10, 100), (9, None), (None, None)], False, 0, 1000, 0, 5000, "违规：第二批价格与股数须成对填写"),
        ("T1", [(10, 100), (None, None), (8, 100)], False, 0, 1000, 0, 5000, "违规：必须先完成第二批"),
        ("T1", [(10, 150), (None, None), (None, None)], False, 0, 1000, 0, 5000, "违规：股数须为100股整数倍"),
        ("T1", [(10, 100), (9, 100), (None, None)], True, 900, 1000, 900, 5000, "违规：锁仓期间禁止加仓"),
        ("T1", [(10, 100), (9, 100), (None, None)], False, 1001, 1000, 1001, 5000, "违规：超过单笔风险上限"),
        ("T1", [(10, 100), (9, 100), (None, None)], False, 900, 1000, 1200, 1199, "违规：超过账户风险上限"),
        ("T1", [(10, 100), (9, 100), (None, None)], False, 900, 1000, 1200, 5000, "通过"),
    ]
    for trade_id, tranches, locked, position_risk, one_limit, total_risk, account_limit, expected in cases:
        with self.subTest(expected=expected):
            self.assertEqual(
                tw.check_tranche_rules(
                    trade_id,
                    tranches,
                    locked,
                    position_risk,
                    one_limit,
                    total_risk,
                    account_limit,
                ),
                expected,
            )
```

- [ ] **Step 5: Implement the rule checker in the tested priority order**

Implement `check_tranche_rules` so it returns blank for no trade ID, validates the first pair, validates each optional pair, forbids batch 3 before batch 2, checks positive 100-share lots, applies the lock only when batch 2 or 3 exists, accepts risk exactly equal to a limit, and returns the strings asserted above.

- [ ] **Step 6: Run calculation tests**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.CalculationTests -v
```

Expected: all calculation tests pass.

- [ ] **Step 7: Commit the calculation contract**

```bash
git add trading_workbook.py tests/test_trading_workbook.py
git commit -m "feat: define three-tranche position rules"
```

---

### Task 2: Add the V5 columns and formula dependencies

**Files:**
- Modify: `tests/test_trading_workbook.py:483-728`
- Modify: `trading_workbook.py:103-120, 663-915, 1203-1339, 2016-2191`

**Interfaces:**
- Consumes: `calculate_tranche_position` and the V4 workbook layout through AJ.
- Produces: `THREE_TRANCHE_HEADERS`, `_three_tranche_formulas(row, trade_end_row) -> dict[int, str]`, and `apply_three_tranche_buying(wb, trade_end_row=201) -> Workbook`.

- [ ] **Step 1: Write failing workbook-structure tests**

Add assertions for renamed first-batch headers, AK:AR, formula cells, formats, and the expanded table:

```python
def test_three_tranche_columns_and_formulas_use_cumulative_position(self):
    trade = self.workbook["单次交易"]
    self.assertEqual(trade["E1"].value, "第一批买入价")
    self.assertEqual(trade["H1"].value, "第一批买入股数")
    self.assertEqual(trade["I1"].value, "首次买入日期")
    self.assertEqual(trade["O1"].value, "买入总费用")
    self.assertEqual(
        [trade.cell(1, column).value for column in range(37, 45)],
        ["第二批买入价", "第二批买入股数", "第三批买入价", "第三批买入股数", "实际加权买入价", "累计买入股数", "当前持仓风险", "分仓规则检查"],
    )
    self.assertIn("E2*H2", trade["AO2"].value)
    self.assertIn("AK2*AL2", trade["AO2"].value)
    self.assertIn("AM2*AN2", trade["AO2"].value)
    self.assertIn("H2+IF(AL2", trade["AP2"].value)
    self.assertIn("MAXIFS", trade["AQ2"].value)
    self.assertIn("'持仓跟踪'!$S$2:$S$501", trade["AQ2"].value)
    self.assertIn("违规：超过单笔风险上限", trade["AR2"].value)
    self.assertEqual(next(iter(trade.tables.values())).ref, "A1:AR201")
```

Update the existing formula-contract assertions so X uses all three batch amounts, Y uses that same total cost, and U/V use AO rather than E.

- [ ] **Step 2: Run the new structure test and verify failure**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.WorkbookStructureTests.test_three_tranche_columns_and_formulas_use_cumulative_position -v
```

Expected: fail because AK:AR and `apply_three_tranche_buying` do not exist.

- [ ] **Step 3: Implement headers, styles, and formulas**

Add:

```python
THREE_TRANCHE_HEADERS = (
    "第二批买入价",
    "第二批买入股数",
    "第三批买入价",
    "第三批买入股数",
    "实际加权买入价",
    "累计买入股数",
    "当前持仓风险",
    "分仓规则检查",
)
```

Implement `_three_tranche_formulas` with these formula contracts:

- AO: blank without H; otherwise `(E*H + AK*AL + AM*AN) / AP`, with blank optional batches treated as zero.
- AP: blank without H; otherwise `H + AL + AN`, with blank optional shares treated as zero.
- AQ: zero for rows with M, blank for unused rows, otherwise sum `MAX(batch price - effective stop, 0) × batch shares`; effective stop is `MAX(N, MAXIFS(持仓跟踪!S, type="每日跟踪", trade ID=A))` with `IFERROR` fallback.
- AR: the exact priority and messages tested in Task 1; single risk compares `AQ>C*D`, and account risk compares `SUM($AQ$2:$AQ$201)>'账户数据'!$B$7`.

In `apply_three_tranche_buying`:

```python
renamed = {
    "E1": "第一批买入价",
    "H1": "第一批买入股数",
    "I1": "首次买入日期",
    "O1": "买入总费用",
}
```

Write AK:AR headers, set AK:AN to input fill, AO:AR to formula fill, use currency formats for AK/AM/AO/AQ, integer format for AL/AN/AP, widths of 16–24, and extend `TradeRecords` plus its auto-filter to `A1:AR201`.

- [ ] **Step 4: Replace dependent trade/account/statistics formulas**

For every trade row, rewrite:

```text
L = sold quantity lookup already supplied by tracking
U = (J-AO)/AO
V = (AO-N)/AO
X = K*L-(E*H+AK*AL+AM*AN)-O-P
Y = X/((E*H+AK*AL+AM*AN)+O)
```

Use guarded Excel formulas so blank rows and invalid denominators return blank. Replace `账户数据!B9` with `=IFERROR(SUM('单次交易'!AQ2:AQ201),0)` and replace `多次统计数据!B12` with the average of `AO×AP` over rows having AP.

- [ ] **Step 5: Apply V5 to newly built workbooks**

Call `apply_three_tranche_buying(wb)` immediately after `apply_trade_expectation_fields(wb)` in `build_workbook`. Update structure-test expectations from V4’s AJ boundary to V5’s AR boundary while preserving all earlier lock, stop, and expectation assertions.

- [ ] **Step 6: Run structure tests**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.WorkbookStructureTests -v
```

Expected: every workbook structure and formula contract test passes.

- [ ] **Step 7: Commit the workbook formula model**

```bash
git add trading_workbook.py tests/test_trading_workbook.py
git commit -m "feat: add three-tranche workbook formulas"
```

---

### Task 3: Enforce the input workflow and full-sale synchronization

**Files:**
- Modify: `tests/test_trading_workbook.py:483-728`
- Modify: `trading_workbook.py:764-877, 957-1158, 1920-2178`

**Interfaces:**
- Consumes: AK:AR formulas from Task 2, `_LockStatus`, `_TrackType`, `_TrackTradeId`, and `_TrackRule`.
- Produces: strict validations for H/AK:AN, AR conditional formatting, and tracking Z lookup of AP.

- [ ] **Step 1: Write failing validation and synchronization tests**

Add:

```python
def test_tranche_validations_enforce_entry_order_lock_and_risk(self):
    trade = self.workbook["单次交易"]
    validations = trade.data_validations.dataValidation
    second_price = next(item for item in validations if "AK2:AK201" in str(item.sqref))
    second_shares = next(item for item in validations if "AL2:AL201" in str(item.sqref))
    third_price = next(item for item in validations if "AM2:AM201" in str(item.sqref))
    third_shares = next(item for item in validations if "AN2:AN201" in str(item.sqref))
    self.assertIn("ISNUMBER(AK2)", second_price.formula1)
    self.assertIn("AK2<>\"\"", second_shares.formula1)
    self.assertIn("MOD(AL2,100)=0", second_shares.formula1)
    self.assertIn("_LockStatus", second_shares.formula1)
    self.assertIn("AQ2<=C2*D2", second_shares.formula1)
    self.assertIn("SUM($AQ$2:$AQ$201)<='账户数据'!$B$7", second_shares.formula1)
    self.assertIn("AND(AK2<>\"\",AL2<>\"\")", third_price.formula1)
    self.assertIn("MOD(AN2,100)=0", third_shares.formula1)

def test_tracking_full_sale_uses_cumulative_shares(self):
    tracking = self.workbook["持仓跟踪"]
    self.assertIn("'单次交易'!$AP$2:$AP$201", tracking["Z2"].value)
    self.assertNotIn("'单次交易'!$H$2:$H$201", tracking["Z2"].value)
```

- [ ] **Step 2: Run the tests and verify failure**

Run:

```bash
.venv/bin/python -m unittest \
  tests.test_trading_workbook.WorkbookStructureTests.test_tranche_validations_enforce_entry_order_lock_and_risk \
  tests.test_trading_workbook.WorkbookStructureTests.test_tracking_full_sale_uses_cumulative_shares -v
```

Expected: validation lookup fails and tracking Z still references H.

- [ ] **Step 3: Add practical no-macro validations**

Use positive-decimal validation for AK and AM. AK allows blank or a positive number; AM additionally requires a complete second pair when nonblank. Use custom validation for AL and AN so the normal input order is price first, shares second:

```text
AL blank OR (AK present AND positive 100-share integer AND unlocked AND AQ<=C*D AND SUM(AQ)<=账户数据!B7)
AN blank OR (AK/AL complete AND AM present AND positive 100-share integer AND unlocked AND AQ<=C*D AND SUM(AQ)<=账户数据!B7)
```

Replace the H validation with the existing expectation/stop/lock requirements plus positive 100-share lots and the same single/account risk limits after H is entered. Configure `showErrorMessage=True` with errors that tell the user which prerequisite or risk cap failed.

The price cell is allowed to exist briefly without shares so sequential entry remains possible; AR must immediately show the incomplete-pair violation until the share is supplied. This is the non-macro mechanism that combines usable entry with persistent audit enforcement.

- [ ] **Step 4: Add AR audit formatting and comments**

Add red conditional formatting for `LEFT($AR2,2)="违规"`, green formatting for `$AR2="通过"`, and comments on AK1, AM1, and AR1 explaining price-first entry, 100-share lots, and paste-bypass auditing.

- [ ] **Step 5: Synchronize the full sale to AP**

Change `_tracking_row_formulas` column 26 to look up `单次交易!AP2:AP201`. Keep X/Y/AA as the existing sale action, price, and fee sources. Confirm that future K/L/M/P lookup formulas in “单次交易” remain unchanged because tracking Z now supplies the cumulative quantity.

- [ ] **Step 6: Run all structure tests**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.WorkbookStructureTests -v
```

Expected: all tests pass, including the earlier lock, stop-plan, and mandatory-expectation validations.

- [ ] **Step 7: Commit the input and sale workflow**

```bash
git add trading_workbook.py tests/test_trading_workbook.py
git commit -m "feat: enforce tranche entry and cumulative sale"
```

---

### Task 4: Add the V4-to-V5 upgrade and prove historical compatibility

**Files:**
- Modify: `tests/test_trading_workbook.py:729-898`
- Modify: `trading_workbook.py:2179-2215`

**Interfaces:**
- Consumes: `apply_three_tranche_buying` from Task 2.
- Produces: `upgrade_workbook_with_three_tranche_buying(source, destination) -> Path`.

- [ ] **Step 1: Write the failing V5 upgrade integration test**

```python
def test_three_tranche_upgrade_preserves_v4_history(self):
    source = Path(__file__).resolve().parents[1] / "交易管理系统_V4_交易预期.xlsx"
    before = load_workbook(source, data_only=False)
    with tempfile.TemporaryDirectory() as temporary:
        destination = Path(temporary) / "交易管理系统_V5_三批买入.xlsx"
        result = tw.upgrade_workbook_with_three_tranche_buying(source, destination)
        after = load_workbook(result, data_only=False)
        trade_before = before["单次交易"]
        trade_after = after["单次交易"]
        for cell in ("A2", "E2", "H2", "I2", "O2", "AG2", "AJ2", "AF38"):
            self.assertEqual(trade_after[cell].value, trade_before[cell].value)
        self.assertIsNone(trade_after["AK2"].value)
        self.assertIsNone(trade_after["AL2"].value)
        self.assertEqual(trade_after["E1"].value, "第一批买入价")
        self.assertEqual(trade_after["AR1"].value, "分仓规则检查")
        self.assertEqual(next(iter(trade_after.tables.values())).ref, "A1:AR201")
```

- [ ] **Step 2: Run the integration test and verify failure**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.IntegrationTests.test_three_tranche_upgrade_preserves_v4_history -v
```

Expected: fail because `upgrade_workbook_with_three_tranche_buying` is undefined.

- [ ] **Step 3: Implement the V5 upgrade function**

```python
def upgrade_workbook_with_three_tranche_buying(
    source: str | Path,
    destination: str | Path,
) -> Path:
    source_path = Path(source)
    destination_path = Path(destination)
    workbook = load_workbook(source_path, data_only=False)
    apply_three_tranche_buying(workbook)
    workbook.save(destination_path)
    return destination_path
```

Guard `apply_three_tranche_buying` with `AK1` so rerunning cannot overwrite V5 data. Do not clear AK:AN, and do not rewrite historical input cells E/H/I/O beyond changing their headers.

- [ ] **Step 4: Add boundary tests for formulas and historical rows**

Assert all existing nonblank trade IDs remain in the same rows, AK:AN are blank for all 100 historical records, AO/AP/AQ/AR contain formulas, and the workbook retains the same sheet names and active sheet except for the intentional V5 active sheet “单次交易”.

- [ ] **Step 5: Run integration tests that do not require LibreOffice**

Run:

```bash
.venv/bin/python -m unittest \
  tests.test_trading_workbook.IntegrationTests.test_lock_upgrade_preserves_the_existing_trade_history \
  tests.test_trading_workbook.IntegrationTests.test_dynamic_stop_upgrade_preserves_history_and_syncs_future_sales \
  tests.test_trading_workbook.IntegrationTests.test_trade_expectation_upgrade_preserves_v3_history_and_tracking \
  tests.test_trading_workbook.IntegrationTests.test_three_tranche_upgrade_preserves_v4_history -v
```

Expected: all four versioned upgrade paths pass.

- [ ] **Step 6: Commit the V5 upgrade path**

```bash
git add trading_workbook.py tests/test_trading_workbook.py
git commit -m "feat: add V5 three-tranche upgrade"
```

---

### Task 5: Update sample and progressive validation coverage

**Files:**
- Modify: `tests/test_trading_workbook.py:392-482, 806-1070`
- Modify: `trading_workbook.py:1523-1658, 2192-2233`
- Modify: `progressive_workbook_validation.py:329-400`

**Interfaces:**
- Consumes: AO/AP/AQ/AR formula model and existing `generate_sample_transactions` records.
- Produces: sample workbooks whose historical records remain valid one-batch positions, plus progressive checks using cumulative values.

- [ ] **Step 1: Write failing sample compatibility tests**

Add assertions that sample input remains in E/H, optional batches stay empty, and expected single-batch AO/AP values can be derived from the source record:

```python
def test_sample_data_is_a_valid_single_batch_position(self):
    workbook = tw.build_workbook(with_sample_data=True, as_of_date=self.AS_OF_DATE)
    trade = workbook["单次交易"]
    items = tw.generate_sample_transactions(self.AS_OF_DATE)
    for row, item in enumerate(items, start=2):
        self.assertEqual(trade.cell(row, 5).value, item["buy_price"])
        self.assertEqual(trade.cell(row, 8).value, item["actual_buy_shares"])
        self.assertIsNone(trade.cell(row, 37).value)
        self.assertIsNone(trade.cell(row, 38).value)
        self.assertIn(f"E{row}*H{row}", trade.cell(row, 41).value)
        self.assertIn(f"H{row}", trade.cell(row, 42).value)
```

- [ ] **Step 2: Run the sample test and verify failure**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook.IntegrationTests.test_sample_data_is_a_valid_single_batch_position -v
```

Expected: fail until V5 formulas are present in the sample workbook.

- [ ] **Step 3: Update calculation-grain sample metrics**

Keep generated sample input as one batch, but make `sample_trade_metrics` expose `buy_amount`, `weighted_buy_price`, `total_shares`, and `current_risk` by calling `calculate_tranche_position([(buy_price, shares), (None, None), (None, None)], stop_price, is_closed=sell_date is not None)`. Preserve existing metric keys so older tests remain valid.

- [ ] **Step 4: Update progressive validation**

In `run_progressive_validation`, compare position amount and share metrics against `buy_amount` and `total_shares`, and add inspected formula/value coordinates AO, AP, AQ, and AR. The report should name these fields `weighted_buy_price`, `cumulative_shares`, `current_position_risk`, and `tranche_rule_check`.

- [ ] **Step 5: Run all non-LibreOffice tests**

Run:

```bash
.venv/bin/python -m unittest tests.test_trading_workbook -v
```

If the known LibreOffice/XPC recalculation test cannot launch, rerun the suite excluding only `IntegrationTests.test_100_trade_sample_workbook_recalculates_to_expected_metrics`; every other test must pass.

- [ ] **Step 6: Commit sample and progressive validation support**

```bash
git add trading_workbook.py tests/test_trading_workbook.py progressive_workbook_validation.py
git commit -m "test: cover cumulative tranche workbook metrics"
```

---

### Task 6: Generate and verify the V5 workbook artifact

**Files:**
- Create: `交易管理系统_V5_三批买入.xlsx`
- Verify: `交易管理系统_V4_交易预期.xlsx`
- Verify: `交易管理系统_V5_三批买入.xlsx`

**Interfaces:**
- Consumes: `upgrade_workbook_with_three_tranche_buying` and the checked-in V4 workbook.
- Produces: the final V5 workbook delivered to the user.

- [ ] **Step 1: Generate V5 from V4**

Run:

```bash
.venv/bin/python -c 'from trading_workbook import upgrade_workbook_with_three_tranche_buying; upgrade_workbook_with_three_tranche_buying("交易管理系统_V4_交易预期.xlsx", "交易管理系统_V5_三批买入.xlsx")'
```

Expected: the V5 file is created while V4 remains unchanged.

- [ ] **Step 2: Run ZIP and openpyxl integrity checks**

Run:

```bash
unzip -t "交易管理系统_V5_三批买入.xlsx"
.venv/bin/python -c 'from openpyxl import load_workbook; p="交易管理系统_V5_三批买入.xlsx"; w=load_workbook(p,data_only=False); t=w["单次交易"]; assert t["AR1"].value=="分仓规则检查"; assert next(iter(t.tables.values())).ref=="A1:AR201"; print(w.sheetnames, t.max_column, t.max_row)'
```

Expected: ZIP reports no errors; openpyxl prints seven sheets, 44 columns, and at least 201 rows.

- [ ] **Step 3: Run regression tests and formula scans**

Run all tests that are available in the environment, then scan workbook XML and Python source for stale single-batch dependencies in final P&L, account open risk, average trade amount, and tracking sale shares. Accept H references only where they intentionally mean first-batch input or opening validation.

- [ ] **Step 4: Render a visual preview**

Use macOS Quick Look to render `交易管理系统_V5_三批买入.xlsx` and inspect the resulting preview for readable AK:AR headers, correct input/formula colors, and no broken sheet layout. Record the known LibreOffice/XPC limitation if recalculation remains unavailable.

- [ ] **Step 5: Confirm artifact isolation**

Compare SHA-256 hashes of V4 before and after generation and assert the V4 hash is unchanged. Confirm V2, V3, and the original workbook still exist and were not overwritten.

- [ ] **Step 6: Commit implementation sources, not generated history**

If all source changes were already committed in Tasks 1–5, do not create an empty commit. Leave the generated V5 workbook as the user-facing artifact unless this repository’s established policy tracks binary workbook outputs.

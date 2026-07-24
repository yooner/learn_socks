import math
import importlib
import shutil
import subprocess
import tempfile
import unittest
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
import trading_workbook as tw
from trading_workbook import (
    calculate_compound_loss_ceiling,
    calculate_monthly_loss,
    calculate_position_size,
    calculate_realized_pnl,
    calculate_required_trades,
    calculate_return_rate,
    summarize_trades,
)

EXPECTED_SHEETS = [
    "单次交易",
    "买入理由",
    "多次统计数据",
    "账户数据",
    "目标收益",
    "技术指标",
]

TRADE_HEADERS = [
    "交易编号",
    "股票代码",
    "买入时账户金额",
    "本次允许亏损比例",
    "买入价",
    "买入建议股数",
    "开仓风险告警",
    "实际买入股数",
    "买入日期",
    "期望卖出价",
    "实际卖出价",
    "卖出股数",
    "卖出日期",
    "止损价",
    "买入费用",
    "卖出费用",
    "买入价的由来",
    "止损价的由来",
    "期望卖出价的由来",
    "实际卖出价的由来",
    "期望盈利比例",
    "期望止损比例",
    "盈亏比",
    "实际盈亏金额",
    "实际收益率",
    "与平均盈利比例差值",
    "持有天数",
    "复利容许平均亏损上限",
    "复利风险判断",
    "交易打分评价",
]


class CalculationTests(unittest.TestCase):
    def test_position_size_rounds_down_to_a_share_lot(self):
        self.assertEqual(
            calculate_position_size(
                account_snapshot=100_000,
                risk_rate=0.01,
                buy_price=10,
                stop_price=9,
            ),
            1_000,
        )
        self.assertEqual(
            calculate_position_size(
                account_snapshot=99_780,
                risk_rate=0.01,
                buy_price=25,
                stop_price=23.5,
            ),
            600,
        )

    def test_position_size_rejects_invalid_risk_inputs(self):
        invalid_cases = (
            (0, 0.01, 10, 9),
            (100_000, 0, 10, 9),
            (100_000, 0.01, 10, 10),
            (100_000, 0.01, 10, 11),
        )
        for account, risk_rate, buy_price, stop_price in invalid_cases:
            with self.subTest(
                account=account,
                risk_rate=risk_rate,
                buy_price=buy_price,
                stop_price=stop_price,
            ):
                self.assertIsNone(
                    calculate_position_size(
                        account,
                        risk_rate,
                        buy_price,
                        stop_price,
                    )
                )

    def test_realized_pnl_includes_buy_and_sell_fees(self):
        pnl = calculate_realized_pnl(
            buy_price=10,
            buy_shares=1_000,
            sell_price=11,
            sell_shares=1_000,
            buy_fee=5,
            sell_fee=5,
        )
        self.assertEqual(pnl, 990)
        self.assertAlmostEqual(
            calculate_return_rate(pnl, 10, 1_000, 5),
            990 / 10_005,
        )

    def test_summary_uses_only_closed_non_flat_trades_for_win_rate(self):
        trades = [
            {
                "pnl": 990,
                "return_rate": 990 / 10_005,
                "hold_days": 5,
                "trade_amount": 10_000,
                "sell_date": date(2026, 5, 10),
            },
            {
                "pnl": -1_210,
                "return_rate": -1_210 / 20_005,
                "hold_days": 3,
                "trade_amount": 20_000,
                "sell_date": date(2026, 6, 4),
            },
            {
                "pnl": 1_790,
                "return_rate": 1_790 / 15_005,
                "hold_days": 8,
                "trade_amount": 15_000,
                "sell_date": date(2026, 6, 18),
            },
            {
                "pnl": -732,
                "return_rate": -732 / 18_006,
                "hold_days": 2,
                "trade_amount": 18_000,
                "sell_date": date(2026, 7, 3),
            },
            {
                "pnl": 0,
                "return_rate": 0,
                "hold_days": 1,
                "trade_amount": 18_000,
                "sell_date": date(2026, 7, 6),
            },
            {
                "pnl": None,
                "return_rate": None,
                "hold_days": None,
                "trade_amount": 20_000,
                "sell_date": None,
            },
        ]

        result = summarize_trades(trades)

        self.assertEqual(result["completed_count"], 5)
        self.assertEqual(result["win_count"], 2)
        self.assertEqual(result["loss_count"], 2)
        self.assertEqual(result["flat_count"], 1)
        self.assertEqual(result["win_hold_days"], 13)
        self.assertEqual(result["loss_hold_days"], 5)
        self.assertEqual(result["win_rate"], 0.5)
        self.assertEqual(result["loss_rate"], 0.5)
        self.assertAlmostEqual(
            result["average_win"],
            ((990 / 10_005) + (1_790 / 15_005)) / 2,
        )
        self.assertAlmostEqual(
            result["average_loss"],
            ((1_210 / 20_005) + (732 / 18_006)) / 2,
        )
        self.assertGreater(result["average_loss"], 0)
        expected_expectancy = (
            result["win_rate"] * result["average_win"]
            - result["loss_rate"] * result["average_loss"]
        )
        self.assertAlmostEqual(result["expectancy"], expected_expectancy)
        expected_compound = (
            (1 + result["average_win"]) ** result["win_count"]
            * (1 - result["average_loss"]) ** result["loss_count"]
            - 1
        )
        self.assertAlmostEqual(result["compound_return"], expected_compound)
        self.assertAlmostEqual(
            result["average_trade_amount"],
            (10_000 + 20_000 + 15_000 + 18_000 + 18_000 + 20_000) / 6,
        )

    def test_compound_loss_ceiling_handles_boundaries(self):
        self.assertIsNone(calculate_compound_loss_ceiling(None, 0.5))
        self.assertIsNone(calculate_compound_loss_ceiling(0.1, 0))
        self.assertIsNone(calculate_compound_loss_ceiling(0.1, 1))
        expected = 1 - (1 + 0.1) ** (-0.5 / (1 - 0.5))
        self.assertAlmostEqual(
            calculate_compound_loss_ceiling(0.1, 0.5),
            expected,
        )

    def test_monthly_loss_preserves_negative_sign(self):
        trades = [
            {"pnl": -732, "sell_date": date(2026, 7, 3)},
            {"pnl": 400, "sell_date": date(2026, 7, 10)},
            {"pnl": -100, "sell_date": date(2026, 6, 30)},
            {"pnl": None, "sell_date": None},
        ]
        self.assertEqual(
            calculate_monthly_loss(trades, as_of_date=date(2026, 7, 24)),
            -732,
        )

    def test_required_trades_rounds_up_and_rejects_non_positive_inputs(self):
        self.assertEqual(calculate_required_trades(1_001, 10_000, 0.025), 5)
        self.assertIsNone(calculate_required_trades(1_000, 0, 0.025))
        self.assertIsNone(calculate_required_trades(1_000, 10_000, 0))
        self.assertIsNone(calculate_required_trades(1_000, 10_000, -0.01))

    def test_sample_generator_produces_100_diverse_sequential_trades(self):
        items = tw.generate_sample_transactions(
            as_of_date=date(2026, 7, 24),
            count=100,
        )
        self.assertEqual(len(items), 100)
        self.assertEqual(len({item["trade_id"] for item in items}), 100)
        self.assertTrue(any(item["sell_date"] is None for item in items))
        self.assertTrue(any(item["actual_buy_shares"] is None for item in items))
        self.assertTrue(
            any(
                item["actual_buy_shares"] is not None
                and item["actual_buy_shares"] < item["suggested_shares"]
                for item in items
            )
        )
        outcomes = {item["outcome"] for item in items}
        self.assertEqual(outcomes, {"win", "loss", "flat", "open", "candidate"})
        indicators = {
            indicator
            for item in items
            for indicator in item["indicators"]
        }
        self.assertIn("蜡烛图", indicators)
        self.assertIn("趋势线", indicators)
        self.assertIn("MACD", indicators)
        sample_summary = summarize_trades(
            tw.sample_trade_metrics(date(2026, 7, 24))
        )
        self.assertGreater(sample_summary["expectancy"], 0)

    def test_append_trade_writes_only_one_requested_row(self):
        workbook = tw.build_workbook(with_sample_data=False)
        item = tw.generate_sample_transactions(
            as_of_date=date(2026, 7, 24),
            count=1,
        )[0]

        tw.append_trade_to_workbook(workbook, item, row=2)
        tw.append_reason_to_workbook(workbook, item, row=2)

        trade = workbook["单次交易"]
        reasons = workbook["买入理由"]
        self.assertEqual(trade["A2"].value, item["trade_id"])
        self.assertEqual(trade["H2"].value, item["actual_buy_shares"])
        self.assertIsNone(trade["A3"].value)
        self.assertEqual(reasons["A2"].value, item["trade_id"])
        self.assertEqual(reasons["F2"].value, item["indicators"][0])
        self.assertIsNone(reasons["A3"].value)

    def test_open_risk_uses_only_unsold_actual_shares(self):
        trades = [
            {
                "buy_price": 10,
                "stop_price": 9,
                "actual_buy_shares": 500,
                "sell_date": None,
            },
            {
                "buy_price": 20,
                "stop_price": 18,
                "actual_buy_shares": 200,
                "sell_date": date(2026, 7, 20),
            },
            {
                "buy_price": 30,
                "stop_price": 31,
                "actual_buy_shares": 300,
                "sell_date": None,
            },
        ]
        self.assertEqual(tw.calculate_open_theoretical_loss(trades), 500)

    def test_risk_status_includes_candidate_at_inclusive_limit(self):
        self.assertEqual(
            tw.calculate_opening_risk_status(4_100, 900, 5_000),
            "禁止开仓",
        )
        self.assertEqual(
            tw.calculate_opening_risk_status(4_000, 900, 5_000),
            "允许开仓",
        )
        self.assertEqual(
            tw.calculate_opening_risk_status(5_100, 0, 5_000),
            "禁止开仓",
        )
        self.assertIsNone(
            tw.calculate_opening_risk_status(0, 0, None),
        )


class WorkbookStructureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.workbook = tw.build_workbook(with_sample_data=False)

    def test_workbook_has_the_six_agreed_sheets(self):
        self.assertEqual(self.workbook.sheetnames, EXPECTED_SHEETS)

    def test_trade_sheet_contains_required_headers_and_guarded_formulas(self):
        ws = self.workbook["单次交易"]
        self.assertEqual(
            [ws.cell(1, column).value for column in range(1, 31)],
            TRADE_HEADERS,
        )
        self.assertIn("ROUNDDOWN", ws["F2"].value)
        self.assertIn("IF(OR(", ws["F2"].value)
        self.assertIn("'账户数据'!$B$9", ws["G2"].value)
        self.assertIn("'账户数据'!$B$7", ws["G2"].value)
        self.assertIn('"禁止开仓"', ws["G2"].value)
        self.assertEqual(ws["L2"].value, '=IF(K2="","",H2)')
        self.assertIn('IF(O2="",0,O2)', ws["X2"].value)
        self.assertIn('IF(P2="",0,P2)', ws["X2"].value)
        self.assertIn("E2*H2", ws["X2"].value)
        self.assertNotIn("E2*F2", ws["X2"].value)
        self.assertNotIn("{row}", ws["X2"].value)
        self.assertIn('IF(OR(X2="",E2="",H2=""),"",IF(', ws["Y2"].value)
        self.assertIn('OR(A2=""', ws["AB2"].value)
        self.assertIn("'多次统计数据'!$B$8", ws["AB2"].value)
        self.assertIn("'多次统计数据'!$B$9", ws["AC2"].value)

    def test_input_and_formula_cells_use_distinct_fills(self):
        ws = self.workbook["单次交易"]
        self.assertEqual(ws["C2"].fill.fgColor.rgb, "00DDEBF7")
        self.assertEqual(ws["D2"].fill.fgColor.rgb, "00DDEBF7")
        self.assertEqual(ws["F2"].fill.fgColor.rgb, "00E7E6E6")
        self.assertEqual(ws["G2"].fill.fgColor.rgb, "00E7E6E6")
        self.assertEqual(ws["H2"].fill.fgColor.rgb, "00DDEBF7")
        self.assertEqual(ws["X2"].fill.fgColor.rgb, "00E7E6E6")

    def test_trade_and_reason_logs_are_filterable_and_frozen(self):
        trade = self.workbook["单次交易"]
        reasons = self.workbook["买入理由"]
        self.assertEqual(trade.freeze_panes, "A2")
        self.assertEqual(reasons.freeze_panes, "A2")
        self.assertEqual(len(trade.tables), 1)
        self.assertEqual(len(reasons.tables), 1)
        self.assertGreater(len(trade.data_validations.dataValidation), 5)
        self.assertGreater(len(reasons.data_validations.dataValidation), 2)

    def test_indicator_dropdowns_reference_the_technical_indicator_sheet(self):
        indicators = self.workbook["技术指标"]
        reasons = self.workbook["买入理由"]
        self.assertEqual(indicators["A1"].value, "技术指标")
        self.assertEqual(indicators["A2"].value, "蜡烛图")
        indicator_values = [
            indicators.cell(row, 1).value
            for row in range(2, indicators.max_row + 1)
        ]
        self.assertIn("趋势线", indicator_values)
        self.assertIn("MACD", indicator_values)
        self.assertIn("技术指标列表", self.workbook.defined_names)
        indicator_validations = [
            item
            for item in reasons.data_validations.dataValidation
            if item.formula1 == "=技术指标列表"
        ]
        self.assertEqual(len(indicator_validations), 1)
        self.assertIn("F2:H501", str(indicator_validations[0].sqref))

    def test_statistics_account_and_target_formulas_follow_metric_contract(self):
        stats = self.workbook["多次统计数据"]
        account = self.workbook["账户数据"]
        target = self.workbook["目标收益"]
        self.assertIn("COUNTIF", stats["B3"].value)
        self.assertIn("-AVERAGEIF", stats["B9"].value)
        self.assertIn("B6*B8-B7*B9", stats["B13"].value)
        self.assertIn("POWER(1+B8,B3)", stats["B14"].value)
        self.assertIn("'单次交易'!E2:E201", stats["B12"].value)
        self.assertIn("'单次交易'!H2:H201", stats["B12"].value)
        self.assertNotIn("'单次交易'!F2:F201", stats["B12"].value)
        self.assertIn("SUM('单次交易'!X2:X201)", account["B3"].value)
        self.assertIn("SUMIFS", account["B8"].value)
        self.assertIn("EOMONTH(TODAY()", account["B8"].value)
        self.assertEqual(account["A6"].value, "当月允许最大亏损比例")
        self.assertEqual(account["A7"].value, "当月允许最大亏损金额")
        self.assertEqual(account["A9"].value, "当前未平仓理论亏损")
        self.assertEqual(account["A10"].value, "当月剩余可开仓风险额度")
        self.assertIn("SUMPRODUCT", account["B9"].value)
        self.assertIn("'单次交易'!H2:H201", account["B9"].value)
        self.assertEqual(
            account["B10"].value,
            '=IF(OR(B7="",B9=""),"",B7-B9)',
        )
        self.assertIn('"暂不可计算"', target["B7"].value)
        self.assertIn("ROUNDUP", target["B7"].value)

    def test_workbook_has_validations_risk_formatting_and_auto_calculation(self):
        trade = self.workbook["单次交易"]
        self.assertGreater(len(trade.conditional_formatting), 3)
        self.assertEqual(trade["U2"].number_format, "0.00%")
        self.assertEqual(trade["I2"].number_format, "yyyy-mm-dd")
        self.assertEqual(trade["X2"].number_format, '¥#,##0.00;[Red]-¥#,##0.00')
        self.assertEqual(self.workbook.calculation.calcMode, "auto")
        self.assertTrue(self.workbook.calculation.fullCalcOnLoad)
        self.assertTrue(self.workbook.calculation.forceFullCalc)


class IntegrationTests(unittest.TestCase):
    AS_OF_DATE = date(2026, 7, 24)

    def _recalculate_with_libreoffice(self, source: Path, output_dir: Path) -> Path:
        soffice = shutil.which("soffice")
        self.assertIsNotNone(soffice, "LibreOffice is required for formula QA")
        output_dir.mkdir(parents=True)
        profile_dir = output_dir / "lo-profile"
        profile_dir.mkdir()
        command = [
            soffice,
            "--headless",
            f"-env:UserInstallation={profile_dir.resolve().as_uri()}",
            "--convert-to",
            "xlsx",
            "--outdir",
            str(output_dir),
            str(source),
        ]
        completed = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=60,
        )
        self.assertEqual(
            completed.returncode,
            0,
            msg=f"LibreOffice failed: {completed.stdout}\n{completed.stderr}",
        )
        recalculated = output_dir / source.name
        self.assertTrue(
            recalculated.exists(),
            msg=f"LibreOffice did not create {recalculated}: {completed.stdout}",
        )
        return recalculated

    def test_100_trade_sample_workbook_recalculates_to_expected_metrics(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / "source" / "交易管理系统_测试版.xlsx"
            source.parent.mkdir()
            workbook = tw.build_workbook(
                with_sample_data=True,
                as_of_date=self.AS_OF_DATE,
            )
            workbook.save(source)
            recalculated = self._recalculate_with_libreoffice(
                source,
                root / "recalculated",
            )
            values = load_workbook(recalculated, data_only=True)

            trade = values["单次交易"]
            stats = values["多次统计数据"]
            account = values["账户数据"]
            target = values["目标收益"]

            expected_items = tw.generate_sample_transactions(
                self.AS_OF_DATE,
                count=100,
            )
            expected_trades = tw.sample_trade_metrics(self.AS_OF_DATE)
            for row, item, metric in zip(
                range(2, 102),
                expected_items,
                expected_trades,
            ):
                self.assertEqual(
                    trade.cell(row, 6).value,
                    item["suggested_shares"],
                )
                self.assertEqual(
                    trade.cell(row, 8).value,
                    item["actual_buy_shares"],
                )
                if metric["pnl"] is None:
                    self.assertIsNone(trade.cell(row, 24).value)
                else:
                    self.assertAlmostEqual(
                        trade.cell(row, 24).value,
                        metric["pnl"],
                    )
            for column in (6, 7, 12, 21, 22, 23, 24, 25, 26, 27, 28, 29):
                self.assertIsNone(
                    trade.cell(102, column).value,
                    msg=f"blank row formula column {column} must stay blank",
                )

            summary = summarize_trades(expected_trades)
            self.assertEqual(stats["B2"].value, summary["completed_count"])
            self.assertEqual(stats["B3"].value, summary["win_count"])
            self.assertEqual(stats["B4"].value, summary["loss_count"])
            self.assertEqual(stats["B5"].value, summary["flat_count"])
            self.assertAlmostEqual(stats["B6"].value, summary["win_rate"])
            self.assertAlmostEqual(stats["B7"].value, summary["loss_rate"])
            self.assertAlmostEqual(stats["B8"].value, summary["average_win"])
            self.assertAlmostEqual(stats["B9"].value, summary["average_loss"])
            self.assertEqual(stats["B10"].value, summary["win_hold_days"])
            self.assertEqual(stats["B11"].value, summary["loss_hold_days"])
            self.assertAlmostEqual(
                stats["B12"].value,
                summary["average_trade_amount"],
            )
            self.assertAlmostEqual(stats["B13"].value, summary["expectancy"])
            self.assertAlmostEqual(
                stats["B14"].value,
                summary["compound_return"],
            )

            current_balance = 100_000 + sum(
                trade["pnl"]
                for trade in expected_trades
                if trade["pnl"] is not None
            )
            self.assertAlmostEqual(account["B3"].value, current_balance)
            expected_monthly_loss = calculate_monthly_loss(
                expected_trades,
                self.AS_OF_DATE,
            )
            self.assertAlmostEqual(
                account["B8"].value,
                expected_monthly_loss,
            )
            self.assertLess(account["B8"].value, 0)
            expected_open_risk = tw.calculate_open_theoretical_loss(
                expected_trades,
            )
            self.assertAlmostEqual(account["B9"].value, expected_open_risk)
            monthly_limit = current_balance * 0.05
            self.assertAlmostEqual(account["B7"].value, monthly_limit)
            self.assertAlmostEqual(
                account["B10"].value,
                monthly_limit - expected_open_risk,
            )
            candidate = expected_items[-1]
            candidate_risk = (
                candidate["suggested_shares"]
                * (candidate["buy_price"] - candidate["stop_price"])
            )
            expected_status = tw.calculate_opening_risk_status(
                expected_open_risk,
                candidate_risk,
                monthly_limit,
            )
            self.assertEqual(trade["G101"].value, expected_status)
            self.assertEqual(expected_status, "禁止开仓")
            self.assertAlmostEqual(target["B2"].value, current_balance)
            self.assertAlmostEqual(target["B4"].value, current_balance * 0.10)
            expected_count = calculate_required_trades(
                current_balance * 0.10,
                summary["average_trade_amount"],
                summary["expectancy"],
            )
            self.assertEqual(target["B7"].value, expected_count)

    def test_delivery_writer_creates_clean_and_test_workbooks(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            clean, sample = tw.write_workbooks(
                root,
                as_of_date=self.AS_OF_DATE,
            )
            self.assertEqual(clean.name, "交易管理系统.xlsx")
            self.assertEqual(sample.name, "交易管理系统_测试版.xlsx")
            self.assertTrue(clean.exists())
            self.assertTrue(sample.exists())
            clean_wb = load_workbook(clean, data_only=False)
            sample_wb = load_workbook(sample, data_only=False)
            self.assertIsNone(clean_wb["单次交易"]["A2"].value)
            self.assertEqual(sample_wb["单次交易"]["A2"].value, "T2026001")
            self.assertEqual(sample_wb["买入理由"]["D2"].value, "买入")

    def test_progressive_validator_recalculates_after_each_append(self):
        validator = importlib.import_module("progressive_workbook_validation")
        with tempfile.TemporaryDirectory() as temporary:
            result = validator.run_progressive_validation(
                output_dir=temporary,
                as_of_date=self.AS_OF_DATE,
                steps=3,
                quiet=True,
            )
            self.assertEqual(len(result["audit_rows"]), 3)
            self.assertEqual(
                [row["step"] for row in result["audit_rows"]],
                [1, 2, 3],
            )
            self.assertTrue(
                all(row["result"] == "PASS" for row in result["audit_rows"])
            )
            self.assertTrue(result["workbook_path"].exists())
            self.assertTrue(result["report_path"].exists())
            final = load_workbook(result["workbook_path"], data_only=True)
            self.assertEqual(final["单次交易"]["A4"].value, "T2026003")
            self.assertIsNone(final["单次交易"]["A5"].value)

    def test_main_writes_deliverables_to_requested_directory(self):
        with tempfile.TemporaryDirectory() as temporary:
            paths = tw.main(
                output_dir=temporary,
                as_of_date=self.AS_OF_DATE,
            )
            self.assertEqual(len(paths), 2)
            self.assertTrue(all(path.exists() for path in paths))


if __name__ == "__main__":
    unittest.main()

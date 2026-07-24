import math
import unittest
from datetime import date

from trading_workbook import (
    calculate_compound_loss_ceiling,
    calculate_monthly_loss,
    calculate_position_size,
    calculate_realized_pnl,
    calculate_required_trades,
    calculate_return_rate,
    summarize_trades,
)


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


if __name__ == "__main__":
    unittest.main()

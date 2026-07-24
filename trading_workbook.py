"""Excel trading-management workbook generator."""

from __future__ import annotations

import math
from datetime import date
from typing import Any, Iterable, Mapping


def calculate_position_size(
    account_snapshot: float,
    risk_rate: float,
    buy_price: float,
    stop_price: float,
) -> int | None:
    """Return the risk-sized position rounded down to a 100-share lot."""
    if (
        account_snapshot <= 0
        or risk_rate <= 0
        or buy_price <= 0
        or stop_price <= 0
        or buy_price <= stop_price
    ):
        return None
    raw_shares = account_snapshot * risk_rate / (buy_price - stop_price)
    return math.floor(raw_shares / 100) * 100


def calculate_realized_pnl(
    buy_price: float,
    buy_shares: int,
    sell_price: float,
    sell_shares: int,
    buy_fee: float = 0,
    sell_fee: float = 0,
) -> float:
    """Calculate realized profit after explicit buy and sell fees."""
    return (
        sell_price * sell_shares
        - buy_price * buy_shares
        - buy_fee
        - sell_fee
    )


def calculate_return_rate(
    pnl: float,
    buy_price: float,
    buy_shares: int,
    buy_fee: float = 0,
) -> float | None:
    """Calculate return on the cash committed to the purchase."""
    invested = buy_price * buy_shares + buy_fee
    if invested <= 0:
        return None
    return pnl / invested


def _average(values: Iterable[float]) -> float | None:
    items = list(values)
    if not items:
        return None
    return sum(items) / len(items)


def summarize_trades(
    trades: Iterable[Mapping[str, Any]],
) -> dict[str, float | int | None]:
    """Aggregate closed-trade statistics using the agreed metric definitions."""
    items = list(trades)
    closed = [
        trade
        for trade in items
        if trade.get("sell_date") is not None and trade.get("pnl") is not None
    ]
    wins = [trade for trade in closed if trade["pnl"] > 0]
    losses = [trade for trade in closed if trade["pnl"] < 0]
    flats = [trade for trade in closed if trade["pnl"] == 0]
    decisive_count = len(wins) + len(losses)
    win_rate = len(wins) / decisive_count if decisive_count else None
    loss_rate = len(losses) / decisive_count if decisive_count else None
    average_win = _average(
        trade["return_rate"]
        for trade in wins
        if trade.get("return_rate") is not None
    )
    average_loss = _average(
        abs(trade["return_rate"])
        for trade in losses
        if trade.get("return_rate") is not None
    )
    expectancy = None
    if (
        win_rate is not None
        and loss_rate is not None
        and average_win is not None
        and average_loss is not None
    ):
        expectancy = win_rate * average_win - loss_rate * average_loss

    compound_return = None
    if wins or losses:
        win_factor = (1 + (average_win or 0)) ** len(wins)
        loss_factor = (1 - (average_loss or 0)) ** len(losses)
        compound_return = win_factor * loss_factor - 1

    return {
        "completed_count": len(closed),
        "win_count": len(wins),
        "loss_count": len(losses),
        "flat_count": len(flats),
        "win_rate": win_rate,
        "loss_rate": loss_rate,
        "average_win": average_win,
        "average_loss": average_loss,
        "win_hold_days": sum(trade.get("hold_days") or 0 for trade in wins),
        "loss_hold_days": sum(trade.get("hold_days") or 0 for trade in losses),
        "average_trade_amount": _average(
            trade["trade_amount"]
            for trade in items
            if trade.get("trade_amount") is not None
            and trade["trade_amount"] > 0
        ),
        "expectancy": expectancy,
        "compound_return": compound_return,
    }


def calculate_compound_loss_ceiling(
    average_profit: float | None,
    win_rate: float | None,
) -> float | None:
    """Return the average-loss ceiling implied by geometric break-even."""
    if (
        average_profit is None
        or average_profit <= -1
        or win_rate is None
        or not 0 < win_rate < 1
    ):
        return None
    return 1 - (1 + average_profit) ** (-win_rate / (1 - win_rate))


def calculate_monthly_loss(
    trades: Iterable[Mapping[str, Any]],
    as_of_date: date,
) -> float:
    """Sum current-month realized losses while preserving their negative sign."""
    total = 0.0
    for trade in trades:
        sell_date = trade.get("sell_date")
        pnl = trade.get("pnl")
        if (
            isinstance(sell_date, date)
            and sell_date.year == as_of_date.year
            and sell_date.month == as_of_date.month
            and pnl is not None
            and pnl < 0
        ):
            total += pnl
    return total


def calculate_required_trades(
    target_profit: float,
    average_trade_amount: float,
    expected_return_rate: float,
) -> int | None:
    """Return whole trades required to reach a target, rounded upward."""
    if (
        target_profit <= 0
        or average_trade_amount <= 0
        or expected_return_rate <= 0
    ):
        return None
    return math.ceil(
        target_profit / (average_trade_amount * expected_return_rate)
    )

"""Excel trading-management workbook generator."""

from __future__ import annotations

import math
from datetime import date, timedelta
from pathlib import Path
from typing import Any, Iterable, Mapping

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import (
    Alignment,
    Border,
    Font,
    PatternFill,
    Side,
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

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
    "单笔仓位收益率",
    "账户收益率与平均盈利率差值",
    "持有天数",
    "复利容许平均亏损上限",
    "复利风险判断",
    "交易打分评价",
    "实际账户收益率",
]

REASON_HEADERS = [
    "交易编号",
    "股票代码",
    "所属板块",
    "阶段",
    "记录日期",
    "技术指标1",
    "技术指标2",
    "技术指标3",
    "综述",
]

TECHNICAL_INDICATORS = [
    ("蜡烛图", "价格形态", "观察实体、上下影线及组合形态"),
    ("趋势线", "趋势", "连接关键高点或低点判断趋势方向"),
    ("MACD", "动量", "观察快慢线、信号线和柱状图"),
    ("移动平均线", "趋势", "观察价格与不同周期均线的位置关系"),
    ("成交量", "量价", "确认突破、回调和趋势的参与强度"),
    ("RSI", "动量", "衡量相对强弱及超买超卖状态"),
    ("KDJ", "动量", "观察随机指标交叉和极值区域"),
    ("布林带", "波动", "观察价格相对中轨和上下轨的位置"),
    ("支撑位", "关键价位", "记录预期获得买盘支撑的价格区域"),
    ("压力位", "关键价位", "记录预期遇到卖压的价格区域"),
    ("缺口", "价格形态", "观察跳空缺口是否回补或延续"),
    ("形态突破", "价格形态", "观察平台、箱体或整理形态的突破"),
    ("均线金叉", "趋势", "短周期均线上穿长周期均线"),
    ("均线死叉", "趋势", "短周期均线下穿长周期均线"),
    ("量价背离", "量价", "价格和成交量变化方向不一致"),
    ("趋势背离", "动量", "价格走势与动量指标走势不一致"),
    ("ATR", "波动", "衡量真实波动幅度并辅助设置止损"),
    ("换手率", "量价", "衡量筹码交换活跃程度"),
    ("相对强弱", "比较", "比较个股与板块或指数的强弱"),
    ("板块共振", "市场环境", "确认个股信号与所属板块方向一致"),
]

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
INPUT_FILL = PatternFill("solid", fgColor="DDEBF7")
FORMULA_FILL = PatternFill("solid", fgColor="E7E6E6")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
THIN_BORDER = Border(
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
    top=Side(style="thin", color="D9E2F3"),
    bottom=Side(style="thin", color="D9E2F3"),
)
CURRENCY_FORMAT = '¥#,##0.00;[Red]-¥#,##0.00'
PERCENT_FORMAT = "0.00%"


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


def calculate_account_return(
    pnl: float,
    account_snapshot: float,
) -> float | None:
    """Calculate one trade's realized contribution to total account equity."""
    if account_snapshot <= 0:
        return None
    return pnl / account_snapshot


def calculate_compound_return(
    account_returns: Iterable[float | None],
) -> float | None:
    """Compound actual account returns in transaction order."""
    factor = 1.0
    count = 0
    for account_return in account_returns:
        if account_return is None:
            continue
        factor *= 1 + account_return
        count += 1
    return factor - 1 if count else None


def calculate_open_theoretical_loss(
    trades: Iterable[Mapping[str, Any]],
) -> float:
    """Return stop-loss exposure for positions without a sell date."""
    total = 0.0
    for trade in trades:
        if trade.get("sell_date") is not None:
            continue
        buy_price = trade.get("buy_price")
        stop_price = trade.get("stop_price")
        actual_shares = trade.get("actual_buy_shares")
        if (
            buy_price is None
            or stop_price is None
            or actual_shares is None
            or actual_shares <= 0
        ):
            continue
        loss_per_share = float(buy_price) - float(stop_price)
        if loss_per_share > 0:
            total += loss_per_share * int(actual_shares)
    return total


def calculate_opening_risk_status(
    open_risk: float,
    candidate_risk: float,
    monthly_limit: float | None,
) -> str | None:
    """Return whether a candidate position fits inside the monthly risk limit."""
    if monthly_limit is None or monthly_limit <= 0:
        return None
    if open_risk + candidate_risk >= monthly_limit:
        return "禁止开仓"
    return "允许开仓"


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
        trade["account_return"]
        for trade in wins
        if trade.get("account_return") is not None
    )
    average_loss = _average(
        abs(trade["account_return"])
        for trade in losses
        if trade.get("account_return") is not None
    )
    expectancy = None
    if (
        win_rate is not None
        and loss_rate is not None
        and average_win is not None
        and average_loss is not None
    ):
        expectancy = win_rate * average_win - loss_rate * average_loss

    compound_return = calculate_compound_return(
        trade.get("account_return") for trade in closed
    )

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


def _apply_header(ws, headers: list[str]) -> None:
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = THIN_BORDER
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "A2"


def _add_table(ws, name: str, ref: str) -> None:
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    ws.add_table(table)


def _add_validation(
    ws,
    validation: DataValidation,
    cell_range: str,
) -> None:
    ws.add_data_validation(validation)
    for range_part in cell_range.split():
        validation.add(range_part)


def _style_trade_sheet(ws, end_row: int) -> None:
    input_columns = {
        1,
        2,
        3,
        4,
        5,
        8,
        9,
        10,
        11,
        13,
        14,
        15,
        16,
        17,
        18,
        19,
        20,
        30,
    }
    for row in range(2, end_row + 1):
        for column in range(1, len(TRADE_HEADERS) + 1):
            cell = ws.cell(row, column)
            cell.fill = (
                INPUT_FILL if column in input_columns else FORMULA_FILL
            )
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column in (3, 5, 10, 11, 14, 15, 16, 24):
            ws.cell(row, column).number_format = CURRENCY_FORMAT
        for column in (4, 21, 22, 25, 26, 28, 31):
            ws.cell(row, column).number_format = PERCENT_FORMAT
        for column in (9, 13):
            ws.cell(row, column).number_format = "yyyy-mm-dd"
        ws.cell(row, 23).number_format = "0.00"
        ws.cell(row, 27).number_format = "0"

    width_by_header = {
        "交易编号": 14,
        "股票代码": 14,
        "买入时账户金额": 17,
        "本次允许亏损比例": 18,
        "买入建议股数": 16,
        "开仓风险告警": 16,
        "实际买入股数": 16,
        "买入日期": 13,
        "卖出日期": 13,
        "买入价的由来": 24,
        "止损价的由来": 24,
        "期望卖出价的由来": 24,
        "实际卖出价的由来": 24,
        "复利容许平均亏损上限": 22,
        "复利风险判断": 18,
        "交易打分评价": 18,
        "实际账户收益率": 18,
    }
    for column, header in enumerate(TRADE_HEADERS, start=1):
        ws.column_dimensions[get_column_letter(column)].width = (
            width_by_header.get(header, 14)
        )


def _add_trade_formulas(ws, end_row: int) -> None:
    for row in range(2, end_row + 1):
        ws.cell(
            row,
            6,
            (
                f'=IF(OR(C{row}="",D{row}="",E{row}="",N{row}="",'
                f"C{row}<=0,D{row}<=0,E{row}<=N{row}),"
                f'"",ROUNDDOWN((C{row}*D{row})/(E{row}-N{row})/100,0)*100)'
            ),
        )
        ws.cell(
            row,
            7,
            (
                f'=IF(OR(A{row}="",F{row}="",E{row}="",N{row}="",'
                f"E{row}<=N{row},'账户数据'!$B$7=\"\"),\"\","
                f"IF('账户数据'!$B$9+IF(H{row}=\"\","
                f"F{row}*(E{row}-N{row}),0)>='账户数据'!$B$7,"
                '"禁止开仓","允许开仓"))'
            ),
        )
        ws.cell(
            row,
            12,
            f'=IF(K{row}="","",H{row})',
        )
        ws.cell(
            row,
            21,
            f'=IF(OR(E{row}="",J{row}="",E{row}<=0),"",(J{row}-E{row})/E{row})',
        )
        ws.cell(
            row,
            22,
            f'=IF(OR(E{row}="",N{row}="",E{row}<=0),"",(E{row}-N{row})/E{row})',
        )
        ws.cell(
            row,
            23,
            f'=IF(OR(U{row}="",V{row}="",V{row}<=0),"",U{row}/V{row})',
        )
        ws.cell(
            row,
            24,
            (
                f'=IF(OR(K{row}="",L{row}="",E{row}="",H{row}=""),"",'
                f'K{row}*L{row}-E{row}*H{row}-IF(O{row}="",0,O{row})'
                f'-IF(P{row}="",0,P{row}))'
            ),
        )
        ws.cell(
            row,
            25,
            (
                f'=IF(OR(X{row}="",E{row}="",H{row}=""),"",'
                f'IF(E{row}*H{row}+IF(O{row}="",0,O{row})<=0,"",'
                f'X{row}/(E{row}*H{row}+IF(O{row}="",0,O{row}))))'
            ),
        )
        ws.cell(
            row,
            26,
            (
                f'=IF(OR(AE{row}="",\'多次统计数据\'!$B$8=""),"",'
                f"AE{row}-'多次统计数据'!$B$8)"
            ),
        )
        ws.cell(
            row,
            27,
            f'=IF(OR(I{row}="",M{row}="",M{row}<I{row}),"",M{row}-I{row})',
        )
        ws.cell(
            row,
            28,
            (
                f'=IF(OR(A{row}="",\'多次统计数据\'!$B$8="",'
                "'多次统计数据'!$B$6<=0,"
                "'多次统计数据'!$B$6>=1),\"\","
                "1-POWER(1+'多次统计数据'!$B$8,"
                "-'多次统计数据'!$B$6/"
                "(1-'多次统计数据'!$B$6)))"
            ),
        )
        ws.cell(
            row,
            29,
            (
                f'=IF(OR(AB{row}="",\'多次统计数据\'!$B$9=""),"",'
                f'IF(\'多次统计数据\'!$B$9<AB{row},"低于上限",'
                f'"达到或超过上限"))'
            ),
        )
        ws.cell(
            row,
            31,
            f'=IF(OR(X{row}="",C{row}="",C{row}<=0),"",X{row}/C{row})',
        )


def _add_trade_validations(ws, end_row: int) -> None:
    positive_decimal = DataValidation(
        type="decimal",
        operator="greaterThan",
        formula1="0",
        allow_blank=True,
    )
    _add_validation(
        ws,
        positive_decimal,
        f"C2:C{end_row} E2:E{end_row} J2:K{end_row} N2:N{end_row}",
    )
    risk_rate = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="1",
        allow_blank=True,
    )
    _add_validation(ws, risk_rate, f"D2:D{end_row}")
    nonnegative_fee = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True,
    )
    _add_validation(ws, nonnegative_fee, f"O2:P{end_row}")
    actual_shares = DataValidation(
        type="custom",
        formula1='=OR(H2="",AND(ISNUMBER(H2),H2>0,MOD(H2,100)=0))',
        allow_blank=True,
    )
    actual_shares.error = "实际买入股数必须是大于0的100股整数倍"
    actual_shares.errorTitle = "实际买入股数无效"
    actual_shares.showErrorMessage = True
    _add_validation(ws, actual_shares, f"H2:H{end_row}")
    buy_date = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(2000,1,1)",
        formula2="DATE(2100,12,31)",
        allow_blank=True,
    )
    _add_validation(ws, buy_date, f"I2:I{end_row}")
    sell_date = DataValidation(
        type="custom",
        formula1='=OR(M2="",AND(ISNUMBER(M2),M2>=I2))',
        allow_blank=True,
    )
    _add_validation(ws, sell_date, f"M2:M{end_row}")
    sell_shares = DataValidation(
        type="custom",
        formula1='=OR(L2="",L2=H2)',
        allow_blank=True,
    )
    _add_validation(ws, sell_shares, f"L2:L{end_row}")
    score = DataValidation(
        type="list",
        formula1='"1-差,2-较差,3-一般,4-良好,5-优秀"',
        allow_blank=True,
    )
    _add_validation(ws, score, f"AD2:AD{end_row}")


def _add_trade_conditional_formatting(ws, end_row: int) -> None:
    ws.conditional_formatting.add(
        f"X2:X{end_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"X2:X{end_row}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        f"W2:W{end_row}",
        CellIsRule(operator="lessThan", formula=["2"], fill=YELLOW_FILL),
    )
    ws.conditional_formatting.add(
        f"AC2:AC{end_row}",
        FormulaRule(
            formula=['$AC2="达到或超过上限"'],
            fill=RED_FILL,
        ),
    )
    ws.conditional_formatting.add(
        f"AC2:AC{end_row}",
        FormulaRule(formula=['$AC2="低于上限"'], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"Y2:Y{end_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"Y2:Y{end_row}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        f"AE2:AE{end_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"AE2:AE{end_row}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        f"G2:G{end_row}",
        FormulaRule(formula=['$G2="禁止开仓"'], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        f"G2:G{end_row}",
        FormulaRule(formula=['$G2="允许开仓"'], fill=GREEN_FILL),
    )


def _build_trade_sheet(wb: Workbook, end_row: int = 201):
    ws = wb.create_sheet("单次交易")
    _apply_header(ws, TRADE_HEADERS)
    _add_trade_formulas(ws, end_row)
    _style_trade_sheet(ws, end_row)
    _add_trade_validations(ws, end_row)
    _add_trade_conditional_formatting(ws, end_row)
    _add_table(ws, "TradeRecords", f"A1:AE{end_row}")
    ws["C1"].comment = Comment(
        "开仓时从“账户数据”的当前总金额复制，并粘贴为数值；不要保留公式。",
        "Codex",
    )
    ws["D1"].comment = Comment(
        "开仓时从“账户数据”的单次交易允许亏损比例复制，并粘贴为数值。",
        "Codex",
    )
    ws["F1"].comment = Comment(
        "按账户快照、本次风险比例和止损距离计算；仅作仓位建议。",
        "Codex",
    )
    ws["G1"].comment = Comment(
        "比较全部未平仓理论亏损（含本行拟开仓风险）与当月允许最大亏损金额。",
        "Codex",
    )
    ws["H1"].comment = Comment(
        "手动填写实际成交股数；后续盈亏、收益率和统计均引用此列。",
        "Codex",
    )
    ws["L1"].comment = Comment(
        "默认等于实际买入股数；本系统按一次卖出处理。",
        "Codex",
    )
    ws["AE1"].comment = Comment(
        "实际盈亏金额÷买入时账户金额快照；用于账户口径统计与逐笔复利。",
        "Codex",
    )
    return ws


def _build_reason_sheet(wb: Workbook, end_row: int = 501):
    ws = wb.create_sheet("买入理由")
    _apply_header(ws, REASON_HEADERS)
    for row in range(2, end_row + 1):
        for column in range(1, len(REASON_HEADERS) + 1):
            cell = ws.cell(row, column)
            cell.fill = INPUT_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(row, 5).number_format = "yyyy-mm-dd"
    widths = [14, 14, 16, 14, 13, 24, 24, 24, 48]
    for column, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(column)].width = width
    stage = DataValidation(
        type="list",
        formula1='"买入,持续追踪,卖出"',
        allow_blank=True,
    )
    _add_validation(ws, stage, f"D2:D{end_row}")
    log_date = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(2000,1,1)",
        formula2="DATE(2100,12,31)",
        allow_blank=True,
    )
    _add_validation(ws, log_date, f"E2:E{end_row}")
    indicator = DataValidation(
        type="list",
        formula1="=技术指标列表",
        allow_blank=True,
    )
    indicator.error = "请从“技术指标”工作表维护的列表中选择"
    indicator.errorTitle = "技术指标无效"
    indicator.showErrorMessage = True
    _add_validation(ws, indicator, f"F2:H{end_row}")
    _add_table(ws, "TradeReasonLog", f"A1:I{end_row}")
    return ws


def _build_technical_indicator_sheet(wb: Workbook, end_row: int = 201):
    ws = wb.create_sheet("技术指标")
    _apply_header(ws, ["技术指标", "分类", "说明"])
    for row, values in enumerate(TECHNICAL_INDICATORS, start=2):
        for column, value in enumerate(values, start=1):
            ws.cell(row, column).value = value
    for row in range(2, end_row + 1):
        for column in range(1, 4):
            cell = ws.cell(row, column)
            cell.fill = INPUT_FILL
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 56
    _add_table(ws, "TechnicalIndicatorCatalog", f"A1:C{end_row}")
    wb.defined_names.add(
        DefinedName(
            "技术指标列表",
            attr_text=f"'技术指标'!$A$2:$A${end_row}",
        )
    )
    ws["A1"].comment = Comment(
        "可在本列继续添加或修改指标；买入理由页的三个下拉框会引用A2:A201。",
        "Codex",
    )
    return ws


def _style_panel(ws) -> None:
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 72
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)


def _build_statistics_sheet(wb: Workbook):
    ws = wb.create_sheet("多次统计数据")
    rows = [
        ("指标", "当前统计", "口径说明"),
        ("已完成交易数", "=COUNT('单次交易'!M2:M201)", "存在卖出日期的交易"),
        ("盈利交易数", '=COUNTIF(\'单次交易\'!X2:X201,">0")', "实际盈亏金额大于0"),
        ("亏损交易数", '=COUNTIF(\'单次交易\'!X2:X201,"<0")', "实际盈亏金额小于0"),
        (
            "持平交易数",
            '=COUNTIFS(\'单次交易\'!M2:M201,"<>",\'单次交易\'!X2:X201,"=0")',
            "已卖出且实际盈亏为0；不进入胜负样本",
        ),
        ("盈利概率", '=IFERROR(B3/(B3+B4),"")', "盈利数÷胜负交易数"),
        ("亏损概率", '=IFERROR(B4/(B3+B4),"")', "亏损数÷胜负交易数"),
        (
            "平均盈利百分比",
            '=IFERROR(AVERAGEIF(\'单次交易\'!X2:X201,">0",\'单次交易\'!AE2:AE201),"")',
            "盈利交易实际账户收益率的平均值",
        ),
        (
            "平均亏损百分比",
            '=IFERROR(-AVERAGEIF(\'单次交易\'!X2:X201,"<0",\'单次交易\'!AE2:AE201),"")',
            "亏损交易实际账户收益率绝对值的平均值",
        ),
        (
            "盈利的持有总天数",
            '=SUMIF(\'单次交易\'!X2:X201,">0",\'单次交易\'!AA2:AA201)',
            "自然日",
        ),
        (
            "亏损的持有总天数",
            '=SUMIF(\'单次交易\'!X2:X201,"<0",\'单次交易\'!AA2:AA201)',
            "自然日",
        ),
        (
            "平均交易金额",
            '=IFERROR(SUMPRODUCT(\'单次交易\'!E2:E201,\'单次交易\'!H2:H201)/COUNT(\'单次交易\'!H2:H201),"")',
            "买入价×实际买入股数的平均值，包含未卖出交易",
        ),
        ("期望收益率", '=IF(OR(B6="",B7="",B8="",B9=""),"",B6*B8-B7*B9)', "盈利贡献减亏损贡献"),
        (
            "迄今为止的复利净利润率",
            '=IF(COUNT(\'单次交易\'!AE2:AE201)=0,"",'
            'EXP(SUMPRODUCT(IFERROR(LN(1+\'单次交易\'!AE2:AE201),0)))-1)',
            "按每笔实际账户收益率逐笔连乘",
        ),
    ]
    for row in rows:
        ws.append(row)
    _style_panel(ws)
    for row in range(2, 15):
        ws.cell(row, 2).fill = FORMULA_FILL
    for row in (6, 7, 8, 9, 13, 14):
        ws.cell(row, 2).number_format = PERCENT_FORMAT
    for row in (12,):
        ws.cell(row, 2).number_format = CURRENCY_FORMAT
    return ws


def _build_account_sheet(wb: Workbook):
    ws = wb.create_sheet("账户数据")
    rows = [
        ("指标", "值", "录入/计算说明"),
        ("初始总金额", None, "手动输入；当前金额从此值开始累计"),
        (
            "当前总金额",
            "=IF(B2=\"\",\"\",B2+SUM('单次交易'!X2:X201))",
            "初始总金额＋全部已实现盈亏",
        ),
        ("单次交易允许的亏损比例", 0.01, "默认1%，可调整；开仓时复制为交易快照"),
        ("单次交易允许的亏损金额", '=IF(B3="","",B3*B4)', "当前总金额×单次允许亏损比例"),
        ("当月允许最大亏损比例", 0.05, "默认5%，可调整"),
        ("当月允许最大亏损金额", '=IF(B3="","",B3*B6)', "当前总金额×当月最大亏损比例"),
        (
            "当月亏损金额",
            '=SUMIFS(\'单次交易\'!X2:X201,\'单次交易\'!X2:X201,"<0",'
            '\'单次交易\'!M2:M201,">="&EOMONTH(TODAY(),-1)+1,'
            '\'单次交易\'!M2:M201,"<="&EOMONTH(TODAY(),0))',
            "按卖出日期统计本月已实现亏损，保留负号",
        ),
        (
            "当前未平仓理论亏损",
            '=IFERROR(SUMPRODUCT((\'单次交易\'!M2:M201="")*'
            '(\'单次交易\'!H2:H201>0)*'
            '(\'单次交易\'!E2:E201>\'单次交易\'!N2:N201)*'
            '\'单次交易\'!H2:H201*'
            '(\'单次交易\'!E2:E201-\'单次交易\'!N2:N201)),0)',
            "未填写卖出日期的实际持仓×买入价与止损价之差",
        ),
        (
            "当月剩余可开仓风险额度",
            '=IF(OR(B7="",B9=""),"",B7-B9)',
            "当月允许最大亏损金额－当前未平仓理论亏损",
        ),
        (
            "账户实际累计收益率",
            '=IF(OR(B2="",B3="",B2<=0),"",B3/B2-1)',
            "当前总金额÷初始总金额－1；无入金出金时为账户权威累计收益",
        ),
    ]
    for row in rows:
        ws.append(row)
    _style_panel(ws)
    for cell in ("B2", "B4", "B6"):
        ws[cell].fill = INPUT_FILL
    for cell in ("B3", "B5", "B7", "B8", "B9", "B10", "B11"):
        ws[cell].fill = FORMULA_FILL
    for cell in ("B2", "B3", "B5", "B7", "B8", "B9", "B10"):
        ws[cell].number_format = CURRENCY_FORMAT
    for cell in ("B4", "B6", "B11"):
        ws[cell].number_format = PERCENT_FORMAT
    positive_amount = DataValidation(
        type="decimal",
        operator="greaterThan",
        formula1="0",
        allow_blank=False,
    )
    _add_validation(ws, positive_amount, "B2")
    risk_rates = DataValidation(
        type="decimal",
        operator="between",
        formula1="0",
        formula2="1",
        allow_blank=False,
    )
    _add_validation(ws, risk_rates, "B4 B6")
    ws.conditional_formatting.add(
        "B8",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        "B10",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    return ws


def _build_target_sheet(wb: Workbook):
    ws = wb.create_sheet("目标收益")
    rows = [
        ("指标", "值", "口径说明"),
        ("总金额", "='账户数据'!B3", "来自账户当前总金额"),
        ("目标收益率", 0.10, "手动输入"),
        ("目标收益金额", '=IF(OR(B2="",B3=""),"",B2*B3)', "总金额×目标收益率"),
        ("平均交易金额", "='多次统计数据'!B12", "来自历史交易统计"),
        ("历史期望收益率", "='多次统计数据'!B13", "来自历史胜率与平均盈亏"),
        (
            "所需投资笔数",
            '=IF(OR(B3="",B6="",B3<=0,B6<=0),"暂不可计算",ROUNDUP(LN(1+B3)/LN(1+B6),0))',
            "按目标收益率与历史账户期望收益率复利计算并向上取整",
        ),
    ]
    for row in rows:
        ws.append(row)
    _style_panel(ws)
    ws["B3"].fill = INPUT_FILL
    for cell in ("B2", "B4", "B5", "B6", "B7"):
        ws[cell].fill = FORMULA_FILL
    for cell in ("B2", "B4", "B5"):
        ws[cell].number_format = CURRENCY_FORMAT
    for cell in ("B3", "B6"):
        ws[cell].number_format = PERCENT_FORMAT
    target_rate = DataValidation(
        type="decimal",
        operator="greaterThan",
        formula1="0",
        allow_blank=False,
    )
    _add_validation(ws, target_rate, "B3")
    return ws


def generate_sample_transactions(
    as_of_date: date,
    count: int = 100,
) -> list[dict[str, Any]]:
    """Generate deterministic trades for progressive workbook verification."""
    if count < 0 or count > 200:
        raise ValueError("count must be between 0 and 200")

    stock_codes = [
        "600519",
        "000001",
        "300750",
        "002594",
        "601318",
        "688981",
        "000858",
        "600036",
        "002415",
        "601899",
    ]
    sectors = [
        "消费",
        "银行",
        "新能源",
        "汽车",
        "保险",
        "半导体",
        "食品饮料",
        "金融",
        "电子",
        "资源",
    ]
    indicator_names = [item[0] for item in TECHNICAL_INDICATORS]
    start_date = as_of_date - timedelta(days=count)
    current_balance = 100_000.0
    result: list[dict[str, Any]] = []

    for index in range(1, count + 1):
        cycle = index % 10
        if index % 25 == 0:
            outcome = "candidate"
        elif cycle == 0:
            outcome = "open"
        elif cycle in (1, 2, 3, 4):
            outcome = "win"
        elif cycle in (5, 6, 7, 8):
            outcome = "loss"
        else:
            outcome = "flat"

        buy_price = round(10 + (index % 11) * 1.25, 2)
        stop_distance = 0.75 + (index % 4) * 0.25
        stop_price = round(buy_price - stop_distance, 2)
        risk_rate = 0.01
        suggested_shares = calculate_position_size(
            current_balance,
            risk_rate,
            buy_price,
            stop_price,
        )
        if suggested_shares is None or suggested_shares <= 0:
            raise ValueError("sample inputs must produce a valid position")
        position_fraction = (0.2, 0.4, 0.6, 0.8, 1.0)[(index - 1) % 5]
        actual_buy_shares = max(
            100,
            math.floor(suggested_shares * position_fraction / 100) * 100,
        )
        actual_buy_shares = min(actual_buy_shares, suggested_shares)
        if outcome == "candidate":
            actual_buy_shares = None

        buy_date = start_date + timedelta(days=index)
        sell_date = None
        sell_price = None
        buy_fee = 0 if outcome == "flat" else (5 if index % 3 else 0)
        sell_fee = 0 if outcome == "flat" else (5 if index % 4 else 0)
        if outcome == "win":
            sell_price = round(
                buy_price * (1.12 + (index % 3) * 0.01),
                2,
            )
        elif outcome == "loss":
            sell_price = round(
                min(stop_price, buy_price * (0.95 - (index % 2) * 0.01)),
                2,
            )
        elif outcome == "flat":
            sell_price = buy_price
        if outcome in {"win", "loss", "flat"}:
            sell_date = min(
                buy_date + timedelta(days=1 + index % 7),
                as_of_date,
            )

        indicators = tuple(
            indicator_names[(index - 1 + offset * 5) % len(indicator_names)]
            for offset in range(3)
        )
        score = {
            "win": "4-良好",
            "loss": "2-较差",
            "flat": "3-一般",
            "open": "",
            "candidate": "",
        }[outcome]
        sources = (
            f"{indicators[0]}与{indicators[1]}共同确认",
            f"止损设在{stop_price:.2f}",
            f"目标参考{indicators[2]}",
            "按计划成交" if sell_date is not None else "",
        )
        item = {
            "trade_id": f"T{as_of_date.year}{index:03d}",
            "stock_code": stock_codes[(index - 1) % len(stock_codes)],
            "sector": sectors[(index - 1) % len(sectors)],
            "account_snapshot": round(current_balance, 2),
            "risk_rate": risk_rate,
            "buy_price": buy_price,
            "suggested_shares": suggested_shares,
            "actual_buy_shares": actual_buy_shares,
            "buy_date": buy_date,
            "expected_sell_price": round(buy_price * 1.10, 2),
            "sell_price": sell_price,
            "sell_date": sell_date,
            "stop_price": stop_price,
            "buy_fee": buy_fee,
            "sell_fee": sell_fee,
            "sources": sources,
            "indicators": indicators,
            "summary": f"{outcome}场景，第{index}笔逐步验证数据",
            "score": score,
            "outcome": outcome,
        }
        result.append(item)

        if sell_date is not None and actual_buy_shares is not None:
            current_balance += calculate_realized_pnl(
                buy_price,
                actual_buy_shares,
                sell_price,
                actual_buy_shares,
                buy_fee,
                sell_fee,
            )

    return result


def sample_trade_metrics(as_of_date: date) -> list[dict[str, Any]]:
    """Return the calculation-grain records represented by sample workbook rows."""
    result = []
    for item in generate_sample_transactions(as_of_date):
        shares = item["actual_buy_shares"]
        pnl = None
        return_rate = None
        account_return = None
        hold_days = None
        if (
            shares is not None
            and item["sell_price"] is not None
            and item["sell_date"] is not None
        ):
            pnl = calculate_realized_pnl(
                item["buy_price"],
                shares,
                item["sell_price"],
                shares,
                item["buy_fee"],
                item["sell_fee"],
            )
            return_rate = calculate_return_rate(
                pnl,
                item["buy_price"],
                shares,
                item["buy_fee"],
            )
            account_return = calculate_account_return(
                pnl,
                item["account_snapshot"],
            )
            hold_days = (item["sell_date"] - item["buy_date"]).days
        result.append(
            {
                "pnl": pnl,
                "return_rate": return_rate,
                "account_return": account_return,
                "account_snapshot": item["account_snapshot"],
                "hold_days": hold_days,
                "trade_amount": (
                    item["buy_price"] * shares
                    if shares is not None
                    else None
                ),
                "sell_date": item["sell_date"],
                "buy_price": item["buy_price"],
                "stop_price": item["stop_price"],
                "actual_buy_shares": shares,
            }
        )
    return result


def append_trade_to_workbook(
    wb: Workbook,
    item: Mapping[str, Any],
    row: int,
) -> None:
    """Append one trade input record without touching adjacent rows."""
    trade_ws = wb["单次交易"]
    values = {
        1: item["trade_id"],
        2: item["stock_code"],
        3: item["account_snapshot"],
        4: item["risk_rate"],
        5: item["buy_price"],
        8: item["actual_buy_shares"],
        9: item["buy_date"],
        10: item["expected_sell_price"],
        11: item["sell_price"],
        13: item["sell_date"],
        14: item["stop_price"],
        15: item["buy_fee"],
        16: item["sell_fee"],
        17: item["sources"][0],
        18: item["sources"][1],
        19: item["sources"][2],
        20: item["sources"][3],
        30: item["score"],
    }
    for column, value in values.items():
        trade_ws.cell(row, column).value = value


def append_reason_to_workbook(
    wb: Workbook,
    item: Mapping[str, Any],
    row: int,
) -> None:
    """Append one buy-reason input record without touching adjacent rows."""
    reason_ws = wb["买入理由"]
    values = (
        item["trade_id"],
        item["stock_code"],
        item["sector"],
        "买入",
        item["buy_date"],
        item["indicators"][0],
        item["indicators"][1],
        item["indicators"][2],
        item["summary"],
    )
    for column, value in enumerate(values, start=1):
        reason_ws.cell(row, column).value = value


def _populate_sample_data(wb: Workbook, as_of_date: date) -> None:
    wb["账户数据"]["B2"] = 100_000
    wb["目标收益"]["B3"] = 0.10
    for row, item in enumerate(
        generate_sample_transactions(as_of_date),
        start=2,
    ):
        append_trade_to_workbook(wb, item, row)
        append_reason_to_workbook(wb, item, row)


def build_workbook(
    with_sample_data: bool = False,
    as_of_date: date | None = None,
) -> Workbook:
    """Build the six-sheet trading workbook."""
    wb = Workbook()
    wb.remove(wb.active)
    _build_trade_sheet(wb)
    _build_reason_sheet(wb)
    _build_statistics_sheet(wb)
    _build_account_sheet(wb)
    _build_target_sheet(wb)
    _build_technical_indicator_sheet(wb)
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    if with_sample_data:
        _populate_sample_data(wb, as_of_date or date.today())
    return wb


def write_workbooks(
    output_dir: str | Path,
    as_of_date: date | None = None,
) -> tuple[Path, Path]:
    """Write the clean template and synthetic test workbook."""
    destination = Path(output_dir)
    destination.mkdir(parents=True, exist_ok=True)
    effective_date = as_of_date or date.today()
    clean_path = destination / "交易管理系统.xlsx"
    sample_path = destination / "交易管理系统_测试版.xlsx"
    build_workbook(with_sample_data=False).save(clean_path)
    build_workbook(
        with_sample_data=True,
        as_of_date=effective_date,
    ).save(sample_path)
    return clean_path, sample_path


def main(
    output_dir: str | Path = ".",
    as_of_date: date | None = None,
) -> tuple[Path, Path]:
    """Generate both user-facing workbook deliverables."""
    paths = write_workbooks(output_dir, as_of_date=as_of_date)
    for path in paths:
        print(path)
    return paths


if __name__ == "__main__":
    main()

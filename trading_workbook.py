"""Excel trading-management workbook generator."""

from __future__ import annotations

import math
from datetime import date
from typing import Any, Iterable, Mapping

from openpyxl import Workbook
from openpyxl.comments import Comment
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
    "买入股数",
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
        7,
        8,
        9,
        11,
        12,
        13,
        14,
        15,
        16,
        17,
        18,
        28,
    }
    for row in range(2, end_row + 1):
        for column in range(1, len(TRADE_HEADERS) + 1):
            cell = ws.cell(row, column)
            cell.fill = (
                INPUT_FILL if column in input_columns else FORMULA_FILL
            )
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for column in (3, 5, 8, 9, 12, 13, 14, 22):
            ws.cell(row, column).number_format = CURRENCY_FORMAT
        for column in (4, 19, 20, 23, 24, 26):
            ws.cell(row, column).number_format = PERCENT_FORMAT
        for column in (7, 11):
            ws.cell(row, column).number_format = "yyyy-mm-dd"
        ws.cell(row, 21).number_format = "0.00"
        ws.cell(row, 25).number_format = "0"

    width_by_header = {
        "交易编号": 14,
        "股票代码": 14,
        "买入时账户金额": 17,
        "本次允许亏损比例": 18,
        "买入日期": 13,
        "卖出日期": 13,
        "买入价的由来": 24,
        "止损价的由来": 24,
        "期望卖出价的由来": 24,
        "实际卖出价的由来": 24,
        "复利容许平均亏损上限": 22,
        "复利风险判断": 18,
        "交易打分评价": 18,
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
                f'=IF(OR(C{row}="",D{row}="",E{row}="",L{row}="",'
                f"C{row}<=0,D{row}<=0,E{row}<=L{row}),"
                f'"",ROUNDDOWN((C{row}*D{row})/(E{row}-L{row})/100,0)*100)'
            ),
        )
        ws.cell(row, 10, f'=IF(I{row}="","",F{row})')
        ws.cell(
            row,
            19,
            f'=IF(OR(E{row}="",H{row}="",E{row}<=0),"",(H{row}-E{row})/E{row})',
        )
        ws.cell(
            row,
            20,
            f'=IF(OR(E{row}="",L{row}="",E{row}<=0),"",(E{row}-L{row})/E{row})',
        )
        ws.cell(
            row,
            21,
            f'=IF(OR(S{row}="",T{row}="",T{row}<=0),"",S{row}/T{row})',
        )
        ws.cell(
            row,
            22,
            (
                f'=IF(OR(I{row}="",J{row}="",E{row}="",F{row}=""),"",'
                f"I{row}*J{row}-E{row}*F{row}-IF(M{row}="",0,M{row})"
                f'-IF(N{row}="",0,N{row}))'
            ),
        )
        ws.cell(
            row,
            23,
            (
                f'=IF(OR(V{row}="",E{row}="",F{row}="",'
                f"E{row}*F{row}+IF(M{row}=\"\",0,M{row})<=0),\"\","
                f"V{row}/(E{row}*F{row}+IF(M{row}=\"\",0,M{row})))"
            ),
        )
        ws.cell(
            row,
            24,
            (
                f'=IF(OR(W{row}="",\'多次统计数据\'!$B$8=""),"",'
                f"W{row}-'多次统计数据'!$B$8)"
            ),
        )
        ws.cell(
            row,
            25,
            f'=IF(OR(G{row}="",K{row}="",K{row}<G{row}),"",K{row}-G{row})',
        )
        ws.cell(
            row,
            26,
            (
                "=IF(OR('多次统计数据'!$B$8=\"\","
                "'多次统计数据'!$B$6<=0,"
                "'多次统计数据'!$B$6>=1),\"\","
                "1-POWER(1+'多次统计数据'!$B$8,"
                "-'多次统计数据'!$B$6/"
                "(1-'多次统计数据'!$B$6)))"
            ),
        )
        ws.cell(
            row,
            27,
            (
                f'=IF(OR(Z{row}="",\'多次统计数据\'!$B$9=""),"",'
                f'IF(\'多次统计数据\'!$B$9<Z{row},"低于上限",'
                f'"达到或超过上限"))'
            ),
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
        f"C2:C{end_row} E2:E{end_row} H2:I{end_row} L2:L{end_row}",
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
    _add_validation(ws, nonnegative_fee, f"M2:N{end_row}")
    buy_date = DataValidation(
        type="date",
        operator="between",
        formula1="DATE(2000,1,1)",
        formula2="DATE(2100,12,31)",
        allow_blank=True,
    )
    _add_validation(ws, buy_date, f"G2:G{end_row}")
    sell_date = DataValidation(
        type="custom",
        formula1='=OR(K2="",AND(ISNUMBER(K2),K2>=G2))',
        allow_blank=True,
    )
    _add_validation(ws, sell_date, f"K2:K{end_row}")
    sell_shares = DataValidation(
        type="custom",
        formula1='=OR(J2="",J2=F2)',
        allow_blank=True,
    )
    _add_validation(ws, sell_shares, f"J2:J{end_row}")
    score = DataValidation(
        type="list",
        formula1='"1-差,2-较差,3-一般,4-良好,5-优秀"',
        allow_blank=True,
    )
    _add_validation(ws, score, f"AB2:AB{end_row}")


def _add_trade_conditional_formatting(ws, end_row: int) -> None:
    ws.conditional_formatting.add(
        f"V2:V{end_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"V2:V{end_row}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )
    ws.conditional_formatting.add(
        f"U2:U{end_row}",
        CellIsRule(operator="lessThan", formula=["2"], fill=YELLOW_FILL),
    )
    ws.conditional_formatting.add(
        f"AA2:AA{end_row}",
        FormulaRule(
            formula=['$AA2="达到或超过上限"'],
            fill=RED_FILL,
        ),
    )
    ws.conditional_formatting.add(
        f"AA2:AA{end_row}",
        FormulaRule(formula=['$AA2="低于上限"'], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"X2:X{end_row}",
        CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL),
    )
    ws.conditional_formatting.add(
        f"X2:X{end_row}",
        CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL),
    )


def _build_trade_sheet(wb: Workbook, end_row: int = 201):
    ws = wb.create_sheet("单次交易")
    _apply_header(ws, TRADE_HEADERS)
    _add_trade_formulas(ws, end_row)
    _style_trade_sheet(ws, end_row)
    _add_trade_validations(ws, end_row)
    _add_trade_conditional_formatting(ws, end_row)
    _add_table(ws, "TradeRecords", f"A1:AB{end_row}")
    ws["C1"].comment = Comment(
        "开仓时从“账户数据”的当前总金额复制，并粘贴为数值；不要保留公式。",
        "Codex",
    )
    ws["D1"].comment = Comment(
        "开仓时从“账户数据”的单次交易允许亏损比例复制，并粘贴为数值。",
        "Codex",
    )
    ws["J1"].comment = Comment(
        "默认等于买入股数；本系统按一次卖出处理。",
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
    _add_table(ws, "TradeReasonLog", f"A1:I{end_row}")
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
        ("已完成交易数", "=COUNT('单次交易'!K2:K201)", "存在卖出日期的交易"),
        ("盈利交易数", '=COUNTIF(\'单次交易\'!V2:V201,">0")', "实际盈亏金额大于0"),
        ("亏损交易数", '=COUNTIF(\'单次交易\'!V2:V201,"<0")', "实际盈亏金额小于0"),
        (
            "持平交易数",
            '=COUNTIFS(\'单次交易\'!K2:K201,"<>",\'单次交易\'!V2:V201,"=0")',
            "已卖出且实际盈亏为0；不进入胜负样本",
        ),
        ("盈利概率", '=IFERROR(B3/(B3+B4),"")', "盈利数÷胜负交易数"),
        ("亏损概率", '=IFERROR(B4/(B3+B4),"")', "亏损数÷胜负交易数"),
        (
            "平均盈利百分比",
            '=IFERROR(AVERAGEIF(\'单次交易\'!V2:V201,">0",\'单次交易\'!W2:W201),"")',
            "盈利交易实际收益率的平均值",
        ),
        (
            "平均亏损百分比",
            '=IFERROR(-AVERAGEIF(\'单次交易\'!V2:V201,"<0",\'单次交易\'!W2:W201),"")',
            "以正数显示亏损幅度",
        ),
        (
            "盈利的持有总天数",
            '=SUMIF(\'单次交易\'!V2:V201,">0",\'单次交易\'!Y2:Y201)',
            "自然日",
        ),
        (
            "亏损的持有总天数",
            '=SUMIF(\'单次交易\'!V2:V201,"<0",\'单次交易\'!Y2:Y201)',
            "自然日",
        ),
        (
            "平均交易金额",
            '=IFERROR(SUMPRODUCT(\'单次交易\'!E2:E201,\'单次交易\'!F2:F201)/COUNT(\'单次交易\'!E2:E201),"")',
            "全部已填写买入价交易的平均买入金额，包含未卖出交易",
        ),
        ("期望收益率", '=IF(OR(B6="",B7="",B8="",B9=""),"",B6*B8-B7*B9)', "盈利贡献减亏损贡献"),
        (
            "迄今为止的复利净利润率",
            '=IF(OR(B3+B4=0,B8="",B9=""),"",POWER(1+B8,B3)*POWER(1-B9,B4)-1)',
            "基于历史平均盈亏率的几何复利模拟",
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
            "=IF(B2=\"\",\"\",B2+SUM('单次交易'!V2:V201))",
            "初始总金额＋全部已实现盈亏",
        ),
        ("单次交易允许的亏损比例", 0.01, "默认1%，可调整；开仓时复制为交易快照"),
        ("单次交易允许的亏损金额", '=IF(B3="","",B3*B4)', "当前总金额×单次允许亏损比例"),
        ("累计交易允许的亏损比例", 0.05, "默认5%，可调整"),
        ("累计交易允许的亏损总金额", '=IF(B3="","",B3*B6)', "当前总金额×累计允许亏损比例"),
        (
            "当月亏损金额",
            '=SUMIFS(\'单次交易\'!V2:V201,\'单次交易\'!V2:V201,"<0",'
            '\'单次交易\'!K2:K201,">="&EOMONTH(TODAY(),-1)+1,'
            '\'单次交易\'!K2:K201,"<="&EOMONTH(TODAY(),0))',
            "按卖出日期统计本月已实现亏损，保留负号",
        ),
    ]
    for row in rows:
        ws.append(row)
    _style_panel(ws)
    for cell in ("B2", "B4", "B6"):
        ws[cell].fill = INPUT_FILL
    for cell in ("B3", "B5", "B7", "B8"):
        ws[cell].fill = FORMULA_FILL
    for cell in ("B2", "B3", "B5", "B7", "B8"):
        ws[cell].number_format = CURRENCY_FORMAT
    for cell in ("B4", "B6"):
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
            '=IF(OR(B4="",B5="",B6="",B5<=0,B6<=0),"暂不可计算",ROUNDUP(B4/(B5*B6),0))',
            "向上取整；历史期望收益率不为正时不计算",
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


def build_workbook(
    with_sample_data: bool = False,
    as_of_date: date | None = None,
) -> Workbook:
    """Build the five-sheet trading workbook."""
    del with_sample_data, as_of_date
    wb = Workbook()
    wb.remove(wb.active)
    _build_trade_sheet(wb)
    _build_reason_sheet(wb)
    _build_statistics_sheet(wb)
    _build_account_sheet(wb)
    _build_target_sheet(wb)
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    return wb

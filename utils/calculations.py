"""
Финансовые расчёты: NPV, IRR, срок окупаемости, выручка, дефлятор.
"""
import streamlit as st

def calc_npv(cashflows: list[float], rate: float) -> float:
    """NPV: CF_n / (1+r)^n, n=1,2,..."""
    if rate == 0.0:
        return sum(cashflows)
    return sum(cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cashflows))


def calc_irr(cashflows: list[float]) -> float:
    def npv_r(r):
        return sum(cf / (1 + r) ** (i + 1) for i, cf in enumerate(cashflows))

    has_negative = any(cf < 0 for cf in cashflows)
    has_positive = any(cf > 0 for cf in cashflows)
    if not has_negative or not has_positive:
        return float("nan")

    # Динамический поиск границ
    lo = 1e-6
    hi = None

    # Ищем hi где NPV < 0, пробуем широкий диапазон
    for candidate in [0.5, 1.0, 2.0, 5.0, 10.0, 50.0, 100.0, 500.0, 1000.0, 5000.0]:
        if npv_r(candidate) < 0:
            hi = candidate
            break

    # Если NPV везде >= 0 — IRR не существует
    if hi is None:
        return float("nan")

    # Уточняем lo — ищем где NPV > 0
    for candidate in [1e-6, 0.01, 0.05, 0.1, 0.5, 1.0]:
        if candidate < hi and npv_r(candidate) > 0:
            lo = candidate
            break

    try:
        for _ in range(3000):
            mid = (lo + hi) / 2
            val = npv_r(mid)
            if val > 0:
                lo = mid
            else:
                hi = mid
            if hi - lo < 1e-9:
                break

        result = (lo + hi) / 2 * 100
        return result if result < 1_000_000 else float("nan")

    except Exception:
        return float("nan")

def calc_payback(cashflows: list[float]) -> float:
    """Простой срок окупаемости (лет)."""
    cum = 0.0
    for i, cf in enumerate(cashflows):
        prev = cum
        cum += cf
        if prev < 0 <= cum:
            return i + (-prev / (cum - prev))
    return float("inf")


def calc_dpayback(cashflows: list[float], rate: float) -> float:
    """Дисконтированный срок окупаемости (лет)."""
    cum = 0.0
    for i, cf in enumerate(cashflows):
        prev = cum
        dcf = cf / (1 + rate) ** (i + 1) if rate != 0 else cf
        cum += dcf
        if prev < 0 <= cum:
            return i + (-prev / (cum - prev))
    return float("inf")


def calculate_deflator(year, base_year, price_indices):
    """Расчет дефлятора относительно базового года"""
    if year == base_year:
        return 1.0
    if year > base_year:
        result = 1.0
        for y in range(base_year + 1, year + 1):
            result *= price_indices.get(y, 1.0)
        return result
    else:
        result = 1.0
        for y in range(year + 1, base_year + 1):
            result *= price_indices.get(y, 1.0)
        return 1.0 / result if result != 0 else 1.0
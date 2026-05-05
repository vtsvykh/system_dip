"""
Финансовые расчёты: NPV, IRR, срок окупаемости, выручка, дефлятор.
"""


def calculate_npv(cash_flows, discount_rate):
    """Расчет чистой приведенной стоимости"""
    npv = 0
    years = sorted(cash_flows.keys())
    for i, year in enumerate(years):
        npv += cash_flows[year] / ((1 + discount_rate) ** i)
    return npv


def calculate_irr(cash_flows, max_iterations=100):
    """Расчет внутренней нормы доходности"""
    def npv(rate):
        return calculate_npv(cash_flows, rate)

    low, high = -0.5, 1.0
    for _ in range(max_iterations):
        mid = (low + high) / 2
        if npv(mid) > 0:
            low = mid
        else:
            high = mid
        if abs(high - low) < 0.0001:
            return mid
    return (low + high) / 2


def calculate_payback_period(cash_flows, initial_investment):
    """Расчет срока окупаемости"""
    cumulative = 0
    years = sorted(cash_flows.keys())

    for i, year in enumerate(years):
        cumulative += cash_flows[year]
        if cumulative >= initial_investment:
            if i == 0:
                return 1
            prev_cumulative = cumulative - cash_flows[year]
            return i + (initial_investment - prev_cumulative) / cash_flows[year]
    return float('inf')


def calculate_revenue(volumes_by_product, prices, export_shares, import_shares, years):
    """Расчет выручки"""
    results = {
        'total_revenue': {year: 0 for year in years},
        'export_revenue': {year: 0 for year in years},
        'domestic_revenue': {year: 0 for year in years},
        'export_volume': {year: 0 for year in years},
        'domestic_volume': {year: 0 for year in years},
    }

    for product in volumes_by_product:
        for year in years:
            volume = volumes_by_product[product].get(year, 0)
            price = prices.get(year, 0)
            export_share = export_shares.get(product, {}).get(year, 0)
            import_share = import_shares.get(product, {}).get(year, 0)

            export_vol = volume * export_share
            domestic_vol = volume * import_share

            results['export_volume'][year] += export_vol
            results['domestic_volume'][year] += domestic_vol
            results['export_revenue'][year] += export_vol * price
            results['domestic_revenue'][year] += domestic_vol * price * 1.2
            results['total_revenue'][year] += volume * price

    return results


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
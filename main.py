import json
from datetime import datetime
from io import BytesIO

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from data.okpd_catalog import ALL_UNITS, DEFAULT_UNITS, OKPD_CATALOG
from data.product_types import PRODUCT_TYPES_MAPPING
from models.project_data import ProjectData
from utils.calculations import (
    calculate_deflator,
    calculate_irr,
    calculate_npv,
    calculate_payback_period,
    calculate_revenue,
)
from utils.excel_export import create_excel_report

def read_uploaded_excel(uploaded_file):
    """Чтение Excel-файла и базовая очистка таблицы"""
    if uploaded_file is None:
        return None

    try:
        df = pd.read_excel(uploaded_file, header=None)
        df = df.dropna(how="all").dropna(axis=1, how="all")
        df = df.reset_index(drop=True)
        df.columns = [f"Колонка {i+1}" for i in range(df.shape[1])]
        return df
    except Exception as e:
        st.error(f"Ошибка чтения файла {uploaded_file.name}: {e}")
        return None

def normalize_year(value):
    """Преобразование заголовка года к целому числу"""
    try:
        if pd.isna(value):
            return None
        value_str = str(value).strip()

        if value_str.startswith("31.12."):
            tail = value_str.split(".")[-1]
            year = int(tail)
            return 2000 + year if year < 100 else year

        year_float = float(value_str)
        year_int = int(year_float)
        if 1900 <= year_int <= 2100:
            return year_int
    except:
        return None

    return None


def extract_row_by_code(df, code):
    """Извлекает значения по строке отчетности (например, 2110 или 4112)"""
    if df is None or df.empty:
        return {}

    result = {}

    for i in range(len(df)):
        row_code = df.iloc[i, 1] if df.shape[1] > 1 else None
        try:
            if pd.notna(row_code) and int(float(row_code)) == int(code):
                for j in range(2, df.shape[1]):
                    year = normalize_year(df.iloc[1, j]) if len(df) > 1 else None
                    if year is not None:
                        value = df.iloc[i, j]
                        if pd.isna(value):
                            value = 0.0
                        result[year] = float(value)
                break
        except:
            continue

    return result

def build_sales_matrix(project_data, years, deflators, vat_rates,
                       other_operating_receipts, liquidation_by_year,
                       revenue_without_vat_report, product_base_prices):
    rows = []

    total_other_revenue = {year: 0.0 for year in years}
    total_receipts = {year: other_operating_receipts.get(year, 0.0) for year in years}
    total_liquidation = {year: liquidation_by_year.get(year, 0.0) for year in years}
    total_revenue_with_vat = {year: 0.0 for year in years}
    total_revenue_without_vat = {year: 0.0 for year in years}
    total_accrued_vat = {year: 0.0 for year in years}

    for product_name, product_years in project_data.products.items():
        row_product_name = {"Наименование статьи": product_name}
        row_price_total = {"Наименование статьи": "Цена"}
        row_volume_total = {"Наименование статьи": "Объем"}
        row_revenue_total = {"Наименование статьи": "Выручка от реализации"}

        row_domestic_header = {"Наименование статьи": "В т.ч. продажи на внутренний рынок"}
        row_domestic_price = {"Наименование статьи": "Цена"}
        row_domestic_volume = {"Наименование статьи": "Объем импорта"}
        row_domestic_revenue = {"Наименование статьи": "Выручка с НДС"}

        row_export_header = {"Наименование статьи": "Продажи на экспорт"}
        row_export_price = {"Наименование статьи": "Цена"}
        row_export_volume = {"Наименование статьи": "Объем экспорта"}
        row_export_revenue = {"Наименование статьи": "Выручка без НДС"}


        for year in years:
            volume = project_data.products.get(product_name, {}).get(year, 0.0)
            export_share = project_data.export_shares.get(product_name, {}).get(year, 0.0)
            export_volume = volume * export_share
            import_volume = volume - export_volume


            vat_rate = vat_rates.get(year, 0.20)
            deflator = deflators.get(year, 1.0)
            # Сначала пробуем взять фиксированную цену по году
            fixed = st.session_state.get("fixed_product_prices", {})
            if product_name in fixed and year in fixed[product_name]:
                price_wo_vat = fixed[product_name][year]
            else:
                input_price = product_base_prices.get(product_name, 0.0)
                price_wo_vat = input_price / deflator if deflator not in (0, None) else 0.0

            domestic_revenue_vat = price_wo_vat * import_volume * (1 + vat_rate)
            export_revenue_wo_vat = price_wo_vat * export_volume
            total_product_revenue = domestic_revenue_vat + export_revenue_wo_vat

            row_product_name[year] = None
            row_price_total[year] = price_wo_vat
            row_volume_total[year] = volume
            row_revenue_total[year] = total_product_revenue * 1000000

            row_domestic_header[year] = None
            row_domestic_price[year] = price_wo_vat
            row_domestic_volume[year] = import_volume
            row_domestic_revenue[year] = domestic_revenue_vat * 1000000

            row_export_header[year] = None
            row_export_price[year] = price_wo_vat
            row_export_volume[year] = export_volume
            row_export_revenue[year] = export_revenue_wo_vat * 1000000

            total_revenue_with_vat[year] += total_product_revenue
            total_revenue_without_vat[year] += (
                domestic_revenue_vat / (1 + vat_rate) if (1 + vat_rate) != 0 else 0.0
            ) + export_revenue_wo_vat

        rows.extend([
            row_product_name,
            row_price_total,
            row_volume_total,
            row_revenue_total,
            row_domestic_header,
            row_domestic_price,
            row_domestic_volume,
            row_domestic_revenue,
            row_export_header,
            row_export_price,
            row_export_volume,
            row_export_revenue,
        ])

    row_other_revenue = {"Наименование статьи": "Прочая выручка"}
    row_other_receipts = {"Наименование статьи": "Прочие поступления от операций"}
    row_liquidation = {"Наименование статьи": "Ликвидационная стоимость"}
    row_total_with_vat = {"Наименование статьи": "Выручка с НДС"}
    row_total_without_vat = {"Наименование статьи": "Выручка без НДС"}
    row_vat = {"Наименование статьи": "НДС начисленный"}
    row_base_wo_other = {"Наименование статьи": "Выручка без НДС без прочей"}

    for year in years:
        deflator = deflators.get(year, 1.0)
        receipts = other_operating_receipts.get(year, 0.0)
        liquidation = liquidation_by_year.get(year, 0.0)
        revenue_2110 = revenue_without_vat_report.get(year, 0.0)

        receipts_deflated = receipts / deflator if deflator not in (0, None) else 0.0

        # База: продукты + прочие ОДДС + ликвидация (всё в постоянных ценах)
        base_wo_other = (
            total_revenue_without_vat[year] * 1000000
            + receipts_deflated
            + liquidation
        )

        if year >= 2025:
            other_revenue = 0.0
            receipts_for_calc = 0.0
            revenue_without_vat = base_wo_other
            revenue_with_vat = total_revenue_with_vat[year] * 1000000 + liquidation
        else:
            receipts_for_calc = receipts
            # Прочая выручка — балансирующая разница до revenue_2110
            other_revenue = (revenue_2110 / deflator) - base_wo_other if deflator not in (0, None) else 0.0
            # Итоговая выручка без НДС = revenue_2110 / deflator
            revenue_without_vat = revenue_2110 / deflator if deflator not in (0, None) else 0.0
            revenue_with_vat = revenue_without_vat + (
                total_revenue_with_vat[year] * 1000000
                - total_revenue_without_vat[year] * 1000000
            )

        vat_value = revenue_with_vat - revenue_without_vat

        row_base_wo_other[year] = base_wo_other
        row_other_revenue[year] = other_revenue
        row_other_receipts[year] = receipts_deflated
        row_liquidation[year] = liquidation
        row_total_with_vat[year] = revenue_with_vat
        row_total_without_vat[year] = revenue_without_vat
        row_vat[year] = vat_value

    rows.extend([
        row_base_wo_other,
        row_other_revenue,
        row_other_receipts,
        row_liquidation,
        row_total_with_vat,
        row_total_without_vat,
        row_vat,
    ])

    return pd.DataFrame(rows)

def build_opex_matrix(project_data, years, deflators, vat_rates,
                      cost_prices, cost_volumes,
                      avg_cost_structure, avg_salary,
                      cslyab_data, transport_price, transport_product,
                      cash_flow_df, income_df,
                      equipment_share, construction_share, pir_share,
                      equipment_dep, construction_dep):

    rows = []

    # ── Вспомогательные данные ──────────────────────────────────────
    all_cost_items = list(cost_prices.keys())
    raw_materials_total = {year: 0.0 for year in years}
    fuel_energy_total   = {year: 0.0 for year in years}

    raw_material_names = st.session_state.get('raw_materials', [])
    fuel_energy_names  = st.session_state.get('fuel_energy', [])

    # Доли из усредненной структуры (в долях, не %)
    def get_share(name):
        return avg_cost_structure.get(name, 0.0) / 100.0

    share_rm        = get_share("Сырьё и материалы")
    share_fuel      = get_share("Топливо")
    share_energy    = get_share("Энергия")
    share_labor     = get_share("Затраты на оплату труда")
    share_social    = get_share("Отчисления на социальные нужды")
    share_works     = get_share("Работы и услуги производственного характера, выполненные сторонними организациями, и приобретённые комплектующие изделия")
    share_other     = get_share("Прочие затраты")
    share_depr      = get_share("Амортизация основных средств")

    # Доля отчислений на соц. нужды относительно оплаты труда
    social_rate = share_social / share_labor if share_labor > 1e-9 else 0.0

    # ── По каждому виду издержек ─────────────────────────────────────
    for item in all_cost_items:
        row_name    = {"Наименование статьи": item}
        row_price   = {"Наименование статьи": "Цена"}
        row_volume  = {"Наименование статьи": "Объём"}
        row_expense = {"Наименование статьи": "Затраты"}

        for year in years:
            price = float(cost_prices.get(item, {}).get(year, 0.0))
            vol = float(cost_volumes.get(item, {}).get(year, 0.0))
            expense = price * vol * 1_000_000

            row_name[year]    = None
            row_price[year]   = price
            row_volume[year]  = vol
            row_expense[year] = expense

            if item in raw_material_names:
                raw_materials_total[year] += expense
            if item in fuel_energy_names:
                fuel_energy_total[year] += expense

        rows.extend([row_name, row_price, row_volume, row_expense])

    # ── Зарплата ─────────────────────────────────────────────────────
    row_salary_net = {"Наименование статьи": "Зарплата без начислений"}
    row_schr = {"Наименование статьи": "СЧР"}  # ← только один раз
    row_accruals = {"Наименование статьи": "Начисления"}
    row_salary_total = {"Наименование статьи": "Зарплата сотрудникам"}

    _years_schr = list(range(2016, 2036))
    _schr_values = [
        0, 356, 1237, 2012, 2012, 2012, 2012, 2379,
        2379, 2379, 2379, 2379, 2569, 2759, 2759, 2759,
        2759, 2759, 2759, 2759
    ]
    _schr_dict = dict(zip(_years_schr, _schr_values))

    _years_salary = list(range(2016, 2036))
    _salary_values = [
        0, 900_625, 2_704_492, 2_967_318, 2_807_261, 2_476_037,
        2_043_990, 2_760_321, 3_337_132, 2_755_693, 2_755_693,
        2_755_693, 2_975_778, 3_195_863, 3_195_863, 3_195_863,
        3_195_863, 3_195_863, 3_195_863, 3_195_863
    ]
    _salary_dict = dict(zip(_years_salary, _salary_values))

    for year in years:
        salary_total = float(_salary_dict.get(year, 0.0))
        salary_net = salary_total / (1 + social_rate) if (1 + social_rate) != 0 else 0.0
        schr = float(_schr_dict.get(year, 0.0))
        accruals = salary_net * social_rate

        row_salary_net[year] = salary_net
        row_schr[year] = schr
        row_accruals[year] = accruals
        row_salary_total[year] = salary_total

    rows.extend([row_salary_net, row_schr, row_accruals, row_salary_total])

    # ── Работы и услуги сторонних организаций ───────────────────────
    row_works = {"Наименование статьи": "Работы и услуги, выполненные сторонними организациями"}

    for year in years:
        defl    = deflators.get(year, 1.0)
        vat     = vat_rates.get(year, 0.20)
        rm_sum  = raw_materials_total[year]
        ratio   = (share_works / share_rm) if share_rm > 1e-9 else 0.0
        works   = rm_sum * ratio
        row_works[year] = works

    rows.append(row_works)

    # ── Коммерческие и управленческие расходы ───────────────────────
    # Фиксированные коммерческие и управленческие расходы — вне цикла
    _years_comm = list(range(2016, 2036))
    _comm_values = [
        0, 384_394, 1_026_910, 1_900_390, 2_464_723, 4_251_418,
        3_150_418, 4_123_032, 4_337_506, 5_382_520, 5_398_114,
        5_394_331, 13_863_751, 12_354_453, 12_352_307, 12_350_524,
        12_349_040, 12_347_803, 12_346_767, 12_345_899
    ]
    _comm_dict = dict(zip(_years_comm, _comm_values))

    # ── Коммерческие и управленческие расходы ────────────────────────
    row_comm = {"Наименование статьи": "Коммерческие и управленческие расходы"}

    for year in years:
        row_comm[year] = float(_comm_dict.get(year, 0.0))  # ← фиксированное

    rows.append(row_comm)

    # ── Акциз ────────────────────────────────────────────────────────
    row_cslyab_x_usd = {"Наименование статьи": "ЦСЛЯБ × Курс доллара"}
    row_excise_per_t = {"Наименование статьи": "Акциз на 1 тонну"}
    row_slab_vol = {"Наименование статьи": "Объём сляба"}
    row_excise = {"Наименование статьи": "Акциз"}

    # Фиксированные значения ЦСЛЯБ × Курс доллара
    _years_cs_usd = list(range(2016, 2036))
    _cs_usd_values = [
        0, 0, 0, 0, 0, 0,
        32782, 44285, 41276, 28704, 33665, 33665,
        33665, 33665, 33665, 33665, 33665, 33665,
        33665, 33665
    ]
    _cs_usd_dict = dict(zip(_years_cs_usd, _cs_usd_values))

    excise_rate = st.session_state.get('excise_rate', 2.7) / 100

    # Общий объём продукции из таблицы продаж
    total_product_volumes = {}
    for pname, pvols in project_data.products.items():
        for year, vol in pvols.items():
            total_product_volumes[year] = total_product_volumes.get(year, 0.0) + vol

    for year in years:
        defl = deflators.get(year, 1.0)
        cs_x_usd = float(_cs_usd_dict.get(year, 0.0))
        excise_1t = cs_x_usd * excise_rate if cs_x_usd > 30000 else 0.0
        slab_vol = total_product_volumes.get(year, 0.0) * 1.02
        excise = excise_1t * slab_vol * 1000

        row_cslyab_x_usd[year] = cs_x_usd
        row_excise_per_t[year] = excise_1t
        row_slab_vol[year] = slab_vol
        row_excise[year] = excise

    rows.extend([row_cslyab_x_usd, row_excise_per_t, row_slab_vol, row_excise])

    # ── Транспортные затраты ─────────────────────────────────────────
    row_transport = {"Наименование статьи": "Транспортные затраты"}

    transport_volumes = {}
    if transport_product:
        transport_volumes = cost_volumes.get(transport_product, {})

    for year in years:
        vol = float(transport_volumes.get(year, 0.0))
        row_transport[year] = vol * transport_price * 1_000_000

    rows.append(row_transport)

    # ── Амортизация (скрытые расчёты) ───────────────────────────────
    opc_4221 = extract_row_by_code(cash_flow_df, 4221)

    row_depreciation = {"Наименование статьи": "Амортизация"}

    amort = {}
    for year in years:
        defl    = deflators.get(year, 1.0)
        vat     = vat_rates.get(year, 0.20)
        inv_421 = - opc_4221.get(year, 0.0)

        inv_equip  = inv_421 * equipment_share * (1 + vat) / defl
        va_equip   = inv_equip / (1 + vat) if (1 + vat) != 0 else 0.0
        amo_equip  = equipment_dep * va_equip / defl if defl not in (0, None) else 0.0

        inv_smr   = inv_421 * construction_share * (1 + vat)
        va_smr    = inv_smr / (1 + vat) if (1 + vat) != 0 else 0.0
        amo_smr   = construction_dep * va_smr / defl if defl not in (0, None) else 0.0

        inv_pir   = inv_421 * pir_share * (1 + vat)
        va_pir    = inv_pir / (1 + vat) if (1 + vat) != 0 else 0.0
        amo_pir   = construction_dep * va_pir / defl if defl not in (0, None) else 0.0

        amort[year] = amo_equip + amo_smr + amo_pir
        row_depreciation[year] = amort[year]

    # ── Амортизация с накопительным учётом ВА ───────────────────────
    opc_4221 = extract_row_by_code(cash_flow_df, 4221)

    row_depreciation = {"Наименование статьи": "Амортизация"}
    amort = {}

    # Фиксированные инвестиции в ПИР — вне цикла
    _years_pir = list(range(2016, 2036))
    _pir_values = [
        0, 9716, 90, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0
    ]
    _pir_invest_dict = dict(zip(_years_pir, _pir_values))

    # Фиксированные значения ВА на начало периода для ПИР
    _va_pir_begin_dict = dict(zip(_years_pir, [
        0, 8097, 7970, 7970, 7970, 7970, 7970, 7970, 7970, 7970,
        7970, 7970, 7970, 7970, 7970, 7970, 7970, 7970, 7970, 7970
    ]))

    # Начальные значения ВА на конец предыдущего периода (только для equip и smr)
    va_equip_end = 0.0
    va_smr_end = 0.0

    debug_rows = []

    for year in years:
        defl = deflators.get(year, 1.0)
        vat = vat_rates.get(year, 0.20)
        inv_421 = -opc_4221.get(year, 0.0)

        # Инвестиции по категориям без НДС
        inv_equip_novat = inv_421 * equipment_share
        inv_smr_novat = inv_421 * construction_share
        inv_pir_novat = float(_pir_invest_dict.get(year, 0.0))

        # ВА на начало периода
        va_equip_begin = va_equip_end + inv_equip_novat
        va_smr_begin = va_smr_end + inv_smr_novat
        va_pir_begin = float(_va_pir_begin_dict.get(year, 0.0))  # ← фиксированное

        # Амортизация за период
        amo_equip = equipment_dep * va_equip_begin / defl
        amo_smr = construction_dep * va_smr_begin / defl
        amo_pir = construction_dep * va_pir_begin / defl

        # ВА на конец периода (только для equip и smr — накопительно)
        va_equip_end = max(0.0, va_equip_begin - amo_equip)
        va_smr_end = max(0.0, va_smr_begin - amo_smr)
        # ВА на конец для ПИР = ВА начало - амортизация (для отладки)
        va_pir_end = max(0.0, va_pir_begin - amo_pir)

        amort[year] = amo_equip + amo_smr + amo_pir
        row_depreciation[year] = amort[year]

    rows.append(row_depreciation)

    # ── Прочие затраты ───────────────────────────────────────────────
    fr_2120 = extract_row_by_code(income_df, 2120)
    row_other_costs = {"Наименование статьи": "Прочие затраты"}

    # Фиксированные прочие затраты (в тех же единицах, что и остальные затраты)
    _years_other = list(range(2016, 2036))
    _other_values = [
        0, 987_028, 2_741_531, 3_752_624, 2_711_144, 6_000_939,
        3_662_948, 4_678_828, 3_138_789, 6_962_630, 6_962_630,
        6_962_630, 18_258_401, 16_520_549, 16_520_549, 16_520_549,
        16_520_549, 16_520_549, 16_520_549, 16_520_549
    ]
    _other_costs_dict = dict(zip(_years_other, _other_values))

    cs_without_vat = {}  # С/С без НДС без прочего — скрытый расчёт
    for year in years:
        defl = deflators.get(year, 1.0)
        vat = vat_rates.get(year, 0.20)

        salary_t = row_salary_total.get(year, 0.0)
        works_v = row_works.get(year, 0.0)

        ss = (raw_materials_total[year]
              + fuel_energy_total[year]
              + salary_t
              + works_v
              + amort[year])
        cs_without_vat[year] = ss

        other = float(_other_costs_dict.get(year, 0.0))  # ← заменили расчёт на фиксированное
        row_other_costs[year] = other

    rows.append(row_other_costs)

    # ── Итоговые затраты ─────────────────────────────────────────────
    row_costs_with_vat = {"Наименование статьи": "Себестоимость с НДС"}
    row_net_costs = {"Наименование статьи": "Чистая себестоимость"}
    row_vat_in_costs = {"Наименование статьи": "НДС в затратах"}
    row_opex_with_vat = {"Наименование статьи": "Операционные затраты с НДС"}
    row_opex_without_vat = {"Наименование статьи": "Операционные затраты без НДС"}
    row_opex_nodep_vat = {"Наименование статьи": "Опex без амортизации с НДС"}
    row_opex_nodep_novat = {"Наименование статьи": "Опex без амортизации без НДС"}

    # Фиксированный НДС в затратах — вне цикла
    _years_vat_costs = list(range(2016, 2036))
    _vat_costs_values = [
        0, 893_308, 2_961_744, 4_629_532, 3_459_737, 7_066_643,
        5_116_099, 6_597_577, 5_346_340, 6_962_802, 6_962_802,
        6_962_802, 12_899_843, 16_582_858, 16_582_858, 16_582_858,
        16_582_858, 16_582_858, 16_582_858, 16_582_858
    ]
    _vat_costs_dict = dict(zip(_years_vat_costs, _vat_costs_values))

    for year in years:
        defl = deflators.get(year, 1.0)
        vat = vat_rates.get(year, 0.20)
        rm_fe = raw_materials_total[year] + fuel_energy_total[year]
        sal = row_salary_total.get(year, 0.0)
        works_v = row_works.get(year, 0.0)
        exc = row_excise.get(year, 0.0)
        dep = amort[year]
        other = row_other_costs.get(year, 0.0)
        transp = row_transport.get(year, 0.0) if year >= 2026 else 0.0
        comm = row_comm.get(year, 0.0)

        costs_with_vat = rm_fe + sal + works_v + exc + dep + other + transp
        net_costs = costs_with_vat - dep
        vat_in_costs = float(_vat_costs_dict.get(year, 0.0))  # ← фиксированное
        opex_with_vat = costs_with_vat + comm
        opex_wo_vat = opex_with_vat - vat_in_costs
        opex_nodep_vat = opex_with_vat - dep - exc
        opex_nodep_novat = opex_wo_vat - dep

        row_costs_with_vat[year] = costs_with_vat
        row_net_costs[year] = net_costs
        row_vat_in_costs[year] = vat_in_costs
        row_opex_with_vat[year] = opex_with_vat
        row_opex_without_vat[year] = opex_wo_vat
        row_opex_nodep_vat[year] = opex_nodep_vat
        row_opex_nodep_novat[year] = opex_nodep_novat

    rows.extend([
        row_costs_with_vat, row_net_costs, row_vat_in_costs,
        row_opex_with_vat, row_opex_without_vat,
        row_opex_nodep_vat, row_opex_nodep_novat,
    ])

    return pd.DataFrame(rows)

def build_nwc_matrix(balance_df, years):
    """
    Блок Оборотный капитал.
    Использует данные бухгалтерского баланса (РСБУ) только за периоды,
    за которые есть фактическая отчётность.
    """
    # --- Извлечение строк баланса ---
    cash_1250          = extract_row_by_code(balance_df, 1250)  # Денежные средства
    receivables_1230   = extract_row_by_code(balance_df, 1230)  # Дебиторская задолженность
    investments_1220   = extract_row_by_code(balance_df, 1220)  # Краткосрочные фин. вложения
    inventories_1210   = extract_row_by_code(balance_df, 1210)  # Запасы сырья и материалов
    other_ca_1260      = extract_row_by_code(balance_df, 1260)  # Прочие оборотные активы

    payables_1420      = extract_row_by_code(balance_df, 1420)  # Кредиторская задолженность
    deferred_rev_1430  = extract_row_by_code(balance_df, 1430)  # Доходы будущих периодов
    provisions_1450    = extract_row_by_code(balance_df, 1450)  # Оценочные обязательства

    # Определяем годы, за которые есть хотя бы одна строка баланса
    reported_years = sorted(set(
        list(cash_1250.keys())
        + list(receivables_1230.keys())
        + list(investments_1220.keys())
        + list(inventories_1210.keys())
        + list(other_ca_1260.keys())
        + list(payables_1420.keys())
        + list(deferred_rev_1430.keys())
        + list(provisions_1450.keys())
    ))

    # Фильтруем только те годы, которые входят в общий диапазон модели
    reported_years = [y for y in years if y in reported_years]

    if not reported_years:
        return pd.DataFrame()

    rows = []

    # ── Оборотные активы ─────────────────────────────────────────────
    row_cash = {"Наименование статьи": "Денежные средства"}
    row_recv = {"Наименование статьи": "Дебиторская задолженность"}
    row_inv  = {"Наименование статьи": "Запасы сырья и материалов"}
    row_oca  = {"Наименование статьи": "Прочие ОА"}
    row_ca   = {"Наименование статьи": "Оборотные активы"}

    for year in reported_years:
        cash    = cash_1250.get(year, 0.0)
        recv    = receivables_1230.get(year, 0.0) + investments_1220.get(year, 0.0)
        inv     = inventories_1210.get(year, 0.0)
        oca     = other_ca_1260.get(year, 0.0)
        ca      = cash + recv + inv + oca

        row_cash[year] = cash
        row_recv[year] = recv
        row_inv[year]  = inv
        row_oca[year]  = oca
        row_ca[year]   = ca

    rows.extend([row_cash, row_recv, row_inv, row_oca, row_ca])

    # ── Пустая строка-разделитель ─────────────────────────────────────
    row_empty = {"Наименование статьи": ""}
    for year in reported_years:
        row_empty[year] = None
    rows.append(row_empty)

    # ── Нормируемые краткосрочные обязательства ───────────────────────
    row_pay   = {"Наименование статьи": "Кредиторская задолженность"}
    row_def   = {"Наименование статьи": "Доходы будущих периодов"}
    row_prov  = {"Наименование статьи": "Оценочные обязательства"}
    row_ncl   = {"Наименование статьи": "Нормируемые краткосрочные обязательства"}

    for year in reported_years:
        pay  = payables_1420.get(year, 0.0)
        defd = deferred_rev_1430.get(year, 0.0)
        prov = provisions_1450.get(year, 0.0)
        ncl  = pay + defd + prov

        row_pay[year]  = pay
        row_def[year]  = defd
        row_prov[year] = prov
        row_ncl[year]  = ncl

    rows.extend([row_pay, row_def, row_prov, row_ncl])

    # ── Пустая строка-разделитель ─────────────────────────────────────
    row_empty2 = {"Наименование статьи": ""}
    for year in reported_years:
        row_empty2[year] = None
    rows.append(row_empty2)

    # ── Чистый оборотный капитал ──────────────────────────────────────
    row_nwc = {"Наименование статьи": "Чистый оборотный капитал"}

    for year in reported_years:
        ca  = row_ca[year]
        ncl = row_ncl[year]
        row_nwc[year] = ca - ncl

    rows.append(row_nwc)

    return pd.DataFrame(rows), reported_years

def build_invest_matrix(cash_flow_df, years, deflators, vat_rates,
                        equipment_share, construction_share, pir_share):
    """
    Блок Инвестиции.
    Скрытые расчёты ведутся в текущих ценах (с НДС),
    в таблице отображаются значения, делённые на дефлятор (постоянные цены).
    """

    opc_4221 = extract_row_by_code(cash_flow_df, 4221)

    # ── Диапазоны для расчёта среднего (2018–2024) ───────────────────
    avg_range = [y for y in years if 2018 <= y <= 2024]

    # ── Скрытые расчёты: инвестиции в текущих ценах с НДС ───────────
    inv_equip_raw  = {}   # Инвестиции в оборудование
    inv_smr_raw    = {}   # Инвестиции в СМР
    inv_pir_raw    = {}   # Инвестиции в ПИР

    for year in years:
        vat        = vat_rates.get(year, 0.20)
        inv_421    = -opc_4221.get(year, 0.0)   # знак минус, т.к. ОДДС 4221 отрицательный

        if 2017 <= year <= 2018:
            inv_equip_raw[year] = inv_421 * (equipment_share + pir_share) * (1 + vat)
        elif 2019 <= year <= 2024:
            inv_equip_raw[year] = inv_421 * (equipment_share + pir_share) * (1 + vat)
        else:
            inv_equip_raw[year] = None          # заполним средним ниже

        if 2017 <= year <= 2024:
            inv_smr_raw[year] = inv_421 * construction_share * (1 + vat)
        else:
            inv_smr_raw[year] = None

        if year in (2017, 2018):
            inv_pir_raw[year] = inv_421 * pir_share * (1 + vat)
        else:
            inv_pir_raw[year] = 0.0

    # ── Среднее по 2018–2024 для оборудования и СМР ─────────────────
    avg_equip = (
        sum(inv_equip_raw[y] for y in avg_range if inv_equip_raw.get(y) is not None)
        / len(avg_range) if avg_range else 0.0
    )
    avg_smr = (
        sum(inv_smr_raw[y] for y in avg_range if inv_smr_raw.get(y) is not None)
        / len(avg_range) if avg_range else 0.0
    )

    for year in years:
        if inv_equip_raw.get(year) is None:
            inv_equip_raw[year] = avg_equip if year in (2025, 2026) else 0.0
        if inv_smr_raw.get(year) is None:
            inv_smr_raw[year] = avg_smr if year in (2025, 2026) else 0.0

    # ── Суммарные инвестиции во внеоборотные активы (с НДС) ──────────
    inv_total_raw = {}
    for year in years:
        inv_total_raw[year] = (
            inv_equip_raw[year] + inv_smr_raw[year] + inv_pir_raw[year]
        )

    # ── Строки таблицы (постоянные цены = делим на дефлятор) ─────────
    rows = []

    row_header     = {"Наименование статьи": "Инвестиции во внеоборотные активы"}
    row_equip      = {"Наименование статьи": "Оборудование"}
    row_smr        = {"Наименование статьи": "Строительно-монтажные работы"}
    row_pir        = {"Наименование статьи": "Прочие инвестиции"}
    row_total      = {"Наименование статьи": "Инвестиции во внеоборотные активы (итого)"}
    row_vat_capex  = {"Наименование статьи": "В т.ч. НДС в капитальных затратах"}

    for year in years:
        defl = deflators.get(year, 1.0)

        equip_const = inv_equip_raw[year] / defl if defl else 0.0
        smr_const   = inv_smr_raw[year]   / defl if defl else 0.0
        pir_const   = inv_pir_raw[year]   / defl if defl else 0.0
        total_const = equip_const + smr_const + pir_const
        vat_capex   = total_const * 0.167

        row_header[year]    = None
        row_equip[year]     = equip_const
        row_smr[year]       = smr_const
        row_pir[year]       = pir_const
        row_total[year]     = total_const
        row_vat_capex[year] = vat_capex

    rows.extend([
        row_header,
        row_equip,
        row_smr,
        row_pir,
        row_total,
        row_vat_capex,
    ])

    # Сохраняем скрытые расчёты в session_state для использования в других блоках
    st.session_state.inv_equip_raw  = inv_equip_raw
    st.session_state.inv_smr_raw    = inv_smr_raw
    st.session_state.inv_pir_raw    = inv_pir_raw
    st.session_state.inv_total_raw  = inv_total_raw

    return pd.DataFrame(rows)

def build_fixed_assets_matrix(years, deflators, vat_rates,
                               inv_equip_raw, inv_smr_raw, inv_pir_raw,
                               equipment_dep, construction_dep,
                               property_tax_rate):
    """
    Блок Основной капитал.
    Скрытые расчёты ведутся в текущих ценах (без НДС).
    В таблице значения делятся на дефлятор (постоянные цены).
    """

    # ── Скрытые расчёты (текущие цены, без НДС) ─────────────────────
    va_equip_begin = {}   # ВА на начало периода: оборудование
    va_smr_begin   = {}   # ВА на начало периода: СМР
    va_pir_begin   = {}   # Прочие инвестиции на начало периода

    amo_equip = {}        # Амортизация: оборудование
    amo_smr   = {}        # Амортизация: СМР
    amo_pir   = {}        # Амортизация: ПИР

    va_equip_end = {}     # ВА на конец периода: оборудование
    va_smr_end   = {}     # ВА на конец периода: СМР
    va_pir_end   = {}     # Прочие инвестиции на конец периода

    residual    = {}      # Остаточная стоимость на конец периода
    tax_base    = {}      # Налоговая база налога на имущество
    prop_tax    = {}      # Налог на имущество

    prev_va_equip_end = 0.0
    prev_va_smr_end   = 0.0
    prev_va_pir_end   = 0.0

    for year in years:
        vat  = vat_rates.get(year, 0.20)

        # ВА на начало периода = инвестиции без НДС + ВА конец пред. периода
        equip_inv_novat = inv_equip_raw.get(year, 0.0) / (1 + vat) if (1 + vat) else 0.0
        smr_inv_novat   = inv_smr_raw.get(year, 0.0)   / (1 + vat) if (1 + vat) else 0.0
        pir_inv_novat   = inv_pir_raw.get(year, 0.0)   / (1 + vat) if (1 + vat) else 0.0

        va_equip_begin[year] = equip_inv_novat + prev_va_equip_end
        va_smr_begin[year]   = smr_inv_novat   + prev_va_smr_end
        va_pir_begin[year]   = pir_inv_novat   + prev_va_pir_end

        # Амортизация
        amo_equip[year] = va_equip_begin[year] * equipment_dep
        amo_smr[year]   = va_smr_begin[year]   * construction_dep
        amo_pir[year]   = va_pir_begin[year]   * construction_dep

        # ВА на конец периода
        va_equip_end[year] = max(0.0, va_equip_begin[year] - amo_equip[year])
        va_smr_end[year]   = max(0.0, va_smr_begin[year]   - amo_smr[year])
        va_pir_end[year]   = max(0.0, va_pir_begin[year]   - amo_pir[year])

        # Остаточная стоимость на конец периода
        residual[year] = va_equip_end[year] + va_smr_end[year] + va_pir_end[year]

        # Обновляем предыдущие значения для следующей итерации
        prev_va_equip_end = va_equip_end[year]
        prev_va_smr_end   = va_smr_end[year]
        prev_va_pir_end   = va_pir_end[year]

    # Налоговая база и налог на имущество — считается с 2017 года
    sorted_years = sorted(years)
    for i, year in enumerate(sorted_years):
        if year < 2017:
            tax_base[year] = 0.0
            prop_tax[year] = 0.0
            continue

        next_year = sorted_years[i - 1] if i - 1 < len(sorted_years) else year
        residual_next = residual.get(next_year, residual[year])
        tax_base[year] = (residual[year] + residual_next) / 2
        prop_tax[year] = (property_tax_rate / 100.0) * tax_base[year]

    # ── Сохраняем скрытые расчёты в session_state ────────────────────
    st.session_state.va_equip_begin  = va_equip_begin
    st.session_state.va_smr_begin    = va_smr_begin
    st.session_state.va_pir_begin    = va_pir_begin
    st.session_state.amo_equip       = amo_equip
    st.session_state.amo_smr         = amo_smr
    st.session_state.amo_pir         = amo_pir
    st.session_state.amo_total       = {y: amo_equip[y] + amo_smr[y] + amo_pir[y] for y in years}
    st.session_state.va_equip_end    = va_equip_end
    st.session_state.va_smr_end      = va_smr_end
    st.session_state.va_pir_end      = va_pir_end
    st.session_state.residual_value  = residual
    st.session_state.prop_tax_raw    = prop_tax

    # ── Строки таблицы (постоянные цены = делим на дефлятор) ─────────
    rows = []

    row_va_equip_beg = {"Наименование статьи": "Внеоборотные активы на начало периода в рез. инв. в Оборудование"}
    row_va_smr_beg   = {"Наименование статьи": "Внеоборотные активы на начало периода в рез. инв. в СМР"}
    row_va_pir_beg   = {"Наименование статьи": "Прочие инвестиции"}

    row_amo_equip    = {"Наименование статьи": "Амортизация с ОС в рез.инв. в Оборудование"}
    row_amo_smr      = {"Наименование статьи": "Амортизация с ОС в рез.инв. в СМР"}
    row_amo_pir      = {"Наименование статьи": "Амортизация с ОС в рез.инв. в ПИР"}
    row_amo_total    = {"Наименование статьи": "Амортизационные отчисления"}

    row_va_equip_end = {"Наименование статьи": "Внеоборотные активы на кон. периода в рез. инв. в Оборудование"}
    row_va_smr_end   = {"Наименование статьи": "Внеоборотные активы на кон. периода в рез. инв. в СМР"}
    row_va_pir_end   = {"Наименование статьи": "Прочие инвестиции на конец периода"}

    row_residual     = {"Наименование статьи": "Остаточная стоимость"}
    row_prop_tax     = {"Наименование статьи": "Налог на имущество"}

    for year in years:
        defl = deflators.get(year, 1.0) or 1.0

        row_va_equip_beg[year] = va_equip_begin[year] / defl
        row_va_smr_beg[year]   = va_smr_begin[year]   / defl
        row_va_pir_beg[year]   = va_pir_begin[year]   / defl

        row_amo_equip[year]    = amo_equip[year] / defl
        row_amo_smr[year]      = amo_smr[year]   / defl
        row_amo_pir[year]      = amo_pir[year]   / defl
        row_amo_total[year]    = (amo_equip[year] + amo_smr[year] + amo_pir[year]) / defl

        row_va_equip_end[year] = va_equip_end[year] / defl
        row_va_smr_end[year]   = va_smr_end[year]   / defl
        row_va_pir_end[year]   = va_pir_end[year]   / defl

        row_residual[year]     = residual[year]   / defl
        row_prop_tax[year]     = prop_tax[year]   / defl

    rows.extend([
        row_va_equip_beg,
        row_va_smr_beg,
        row_va_pir_beg,
        row_amo_equip,
        row_amo_smr,
        row_amo_pir,
        row_amo_total,
        row_va_equip_end,
        row_va_smr_end,
        row_va_pir_end,
        row_residual,
        row_prop_tax,
    ])

    return pd.DataFrame(rows)

# ========== ОСНОВНОЕ ПРИЛОЖЕНИЕ ==========
def main():
    st.set_page_config(page_title="Инвестиционный калькулятор", layout="wide")
    st.title("📊 Калькулятор эффективности инвестиционных проектов")

    # Инициализация состояния
    if 'project_data' not in st.session_state:
        st.session_state.project_data = ProjectData()
    if 'revenue_data' not in st.session_state:
        st.session_state.revenue_data = None
    if 'cash_flows' not in st.session_state:
        st.session_state.cash_flows = None
    if 'npv' not in st.session_state:
        st.session_state.npv = 0
    if 'irr' not in st.session_state:
        st.session_state.irr = 0
    if 'payback' not in st.session_state:
        st.session_state.payback = 0
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
    if 'custom_products' not in st.session_state:
        st.session_state.custom_products = []  # Список добавленных пользователем продуктов
    if 'custom_products_by_okpd' not in st.session_state:
        st.session_state.custom_products_by_okpd = {}  # Пользовательские продукты по ОКПД
    if 'product_base_prices' not in st.session_state:
        st.session_state.product_base_prices = {}
    if 'balance_df' not in st.session_state:
        st.session_state.balance_df = None
    if 'income_df' not in st.session_state:
        st.session_state.income_df = None
    if 'cash_flow_df' not in st.session_state:
        st.session_state.cash_flow_df = None
    if 'vat_rates' not in st.session_state:
        st.session_state.vat_rates = {}
    if 'other_operating_receipts' not in st.session_state:
        st.session_state.other_operating_receipts = {}
    if 'revenue_without_vat' not in st.session_state:
        st.session_state.revenue_without_vat = {}
    if 'price_indices_temp' not in st.session_state:
        st.session_state.price_indices_temp = {
            2016: 1.0495,
            2017: 1.0171,
            2018: 1.0456,
            2019: 1.0476,
            2020: 1.0718,
            2021: 1.0627,
            2022: 1.1226,
            2023: 1.0830,
            2024: 1.0817,
            2025: 1.0680,
            2026: 1.0400,
            2027: 1.0400,
            2028: 1.0400,
            2029: 1.0400,
            2030: 1.0400,
            2031: 1.0400,
            2032: 1.0400,
            2033: 1.0400,
            2034: 1.0400,
            2035: 1.0400,
        }
    _OKPD_KEY = "24.10"  # укажи свой реальный код ОКПД

    if 'volumes_temp' not in st.session_state:
        _years_vol = list(range(2016, 2036))
        _vol_zagotovka = [
            0.000, 0.059, 0.1936, 0.2243, 0.2501, 0.2786, 0.3275, 0.4233,
            0.4233, 0.5216, 0.5216, 0.5216, 1.50, 1.745, 1.745, 1.745,
            1.745, 1.745, 1.745, 1.745
        ]
        _vol_sortvoy = [
            0.000, 0.051, 0.166, 0.193, 0.215, 0.239, 0.281, 0.364,
            0.364, 0.364, 0.364, 0.364, 0.364, 0.455, 0.455, 0.455,
            0.455, 0.455, 0.455, 0.455
        ]
        st.session_state.volumes_temp = {
            _OKPD_KEY: {
                "Стальная заготовка": dict(zip(_years_vol, _vol_zagotovka)),
                "Сортовой прокат": dict(zip(_years_vol, _vol_sortvoy)),
            }
        }
    if 'exports_temp' not in st.session_state:
        _years_vol = list(range(2016, 2036))
        _exp_zagotovka = [
            0, 0, 18.0, 18.0, 18.0, 18.0, 18.0, 18.0, 18.0, 15.0,
            15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0
        ]
        _exp_sortvoy = [
            0, 0, 18.0, 18.0, 18.0, 18.0, 18.0, 18.0, 18.0, 15.0,
            15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0, 15.0
        ]
        st.session_state.exports_temp = {
            _OKPD_KEY: {
                "Стальная заготовка": dict(zip(_years_vol, _exp_zagotovka)),
                "Сортовой прокат": dict(zip(_years_vol, _exp_sortvoy)),
            }
        }

    if 'exports_temp' not in st.session_state:
        _years_vol = list(range(2016, 2036))
        st.session_state.exports_temp = {
            _OKPD_KEY: {
                "Стальная заготовка": {year: 0.0 for year in _years_vol},
                "Сортовой прокат": {year: 0.0 for year in _years_vol},
            }
        }

    if 'fixed_product_prices' not in st.session_state:
        _years_p = list(range(2016, 2036))
        _prices_zagotovka = [
            0.000, 44.056, 44.056, 51.911, 37.198, 67.317, 38.663, 44.056,
            33.928, 44.056, 44.056, 44.056, 44.056, 44.056, 44.056, 44.056,
            44.056, 44.056, 44.056, 44.056
        ]
        _prices_sortvoy = [
            0.000, 72.420, 72.420, 95.879, 68.703, 124.334, 71.410, 72.420,
            61.126, 72.420, 72.420, 72.420, 72.420, 72.420, 72.420, 72.420,
            72.420, 72.420, 72.420, 72.420
        ]
        st.session_state.fixed_product_prices = {
            "Стальная заготовка": dict(zip(_years_p, _prices_zagotovka)),
            "Сортовой прокат": dict(zip(_years_p, _prices_sortvoy)),
        }
    if 'residual_value_last_year' not in st.session_state:
        st.session_state.residual_value_last_year = 4_639_210.0

    if 'cost_prices' not in st.session_state:
        _years_cp = list(range(2016, 2036))

        _prices_lom = [
                33.7, 33.2, 31.7, 30.3, 21.7, 39.3, 22.6, 28.5,
                22.3, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5, 28.5,
                28.5, 28.5, 28.5, 28.5
            ]
        _prices_gbzh = [
                26.8, 26.4, 25.2, 24.1, 20.9, 42.9, 20.4, 26.2,
                20.5, 26.2, 26.2, 26.2, 26.2, 26.2, 26.2, 26.2,
                26.2, 26.2, 26.2, 26.2
            ]
        _prices_gaz = [
                0.0072, 0.0071, 0.0068, 0.0065, 0.0061, 0.0057, 0.0051, 0.0047,
                0.0047, 0.0047, 0.0047, 0.0047, 0.0047, 0.0047, 0.0047, 0.0047,
                0.0047, 0.0047, 0.0047, 0.0047
            ]
        _prices_electro = [
                6.2, 6.0, 5.8, 5.8, 5.7, 5.6, 5.4, 5.2,
                5.6, 5.2, 5.2, 5.2, 5.2, 5.2, 5.2, 5.2,
                5.2, 5.2, 5.2, 5.2
            ]

        st.session_state.cost_prices = {
                "Лом и скраб": dict(zip(_years_cp, _prices_lom)),
                "ГБЖ": dict(zip(_years_cp, _prices_gbzh)),
                "Природный газ": dict(zip(_years_cp, _prices_gaz)),
                "Электроэнергия": dict(zip(_years_cp, _prices_electro)),
            }
    if 'raw_materials' not in st.session_state:
        st.session_state.raw_materials = ["Лом и скраб", "ГБЖ"]

    if 'fuel_energy' not in st.session_state:
        st.session_state.fuel_energy = ["Природный газ", "Электроэнергия"]

    if 'cost_volumes' not in st.session_state:
        _years_cv = list(range(2016, 2036))

        _vol_lom = [
                0.00, 0.11, 0.38, 0.62, 0.62, 0.76, 0.79, 0.83,
                0.83, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80, 0.80,
                0.80, 0.80, 0.80, 0.80
            ]
        _vol_gbzh = [
                0.000, 0.000, 0.000, 0.000, 0.000, 0.000, 0.150, 0.150,
                0.150, 0.150, 0.150, 0.150, 0.798, 1.595, 1.595, 1.595,
                1.595, 1.595, 1.595, 1.600
            ]
        _vol_gaz = [
                0.0, 2.5, 8.7, 14.1, 14.1, 17.1, 17.7, 18.7,
                18.7, 18.7, 18.7, 18.7, 20.8, 41.6, 41.6, 41.6,
                41.6, 41.6, 41.6, 41.6
            ]
        _vol_electro = [
                0.0, 0.1, 0.3, 0.6, 0.6, 0.7, 0.7, 0.7,
                0.7, 0.7, 0.7, 0.7, 0.8, 1.6, 1.6, 1.6,
                1.6, 1.6, 1.6, 1.6
            ]

        st.session_state.cost_volumes = {
                "Лом и скраб": dict(zip(_years_cv, _vol_lom)),
                "ГБЖ": dict(zip(_years_cv, _vol_gbzh)),
                "Природный газ": dict(zip(_years_cv, _vol_gaz)),
                "Электроэнергия": dict(zip(_years_cv, _vol_electro)),
            }
    if 'cslyab_data' not in st.session_state:
        _years_cs = list(range(2016, 2036))

        _cslyab = [
                0, 0, 0, 0, 0, 0,
                442, 520, 483, 394, 400, 400,
                400, 400, 400, 400, 400, 400,
                400, 400
            ]
        _dollar = [
                0.0, 58.3, 62.7, 64.7, 72.1, 73.7,
                68.5, 85.2, 92.5, 84.2, 84.2, 84.2,
                84.2, 84.2, 84.2, 84.2, 84.2, 84.2,
                84.2, 84.2
            ]

        st.session_state.cslyab_data = {
                "ЦСЛЯБ": dict(zip(_years_cs, _cslyab)),
                "Курс доллара": dict(zip(_years_cs, _dollar)),
            }

    if 'transport_price' not in st.session_state:
        st.session_state.transport_price = 0.82

    if 'avg_cost_structure' not in st.session_state:
        st.session_state.avg_cost_structure = {
                "Сырьё и материалы": 53.41,
                "Топливо": 9.29,
                "Энергия": 14.85,
                "Затраты на оплату труда": 8.43,
                "Отчисления на социальные нужды": 2.59,
                "Амортизация основных средств": 3.58,
                "Работы и услуги производственного характера, выполненные сторонними организациями, и приобретённые комплектующие изделия": 6.21,
                "Прочие затраты": 1.64,
            }

    if 'excise_rate' not in st.session_state:
        st.session_state.excise_rate = 2.7

    if "fixedproductprices" not in st.session_state:
        years_p = list(range(2016, 2036))

        prices_zagotovka = [
            0.000, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1,  # ← базовая цена 44.1
            44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1, 44.1
        ]
        prices_sortvoy = [
            0.000, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4,  # ← базовая цена 72.4
            72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4, 72.4
        ]
        st.session_state["fixedproductprices"] = {
            "Стальная заготовка": dict(zip(years_p, prices_zagotovka)),
            "Сортовой прокат": dict(zip(years_p, prices_sortvoy)),
        }

    if "property_tax_rate" not in st.session_state:
        st.session_state.property_tax_rate = 2

    # ========== САЙДБАР ДЛЯ ЗАГРУЗКИ ФАЙЛОВ ==========
    with st.sidebar:
        st.header("📁 Загрузка отчетности")

        current_year = datetime.now().year

        # Получаем текущие годы из project_data если они есть
        if st.session_state.project_data.years:
            start_year = st.session_state.project_data.years[0]
            end_year = st.session_state.project_data.years[-1]
            st.info(f"📅 Требуемый период: {start_year}-{current_year}")
        else:
            st.info("📅 Сначала укажите годы во вкладке 'Ввод данных'")
            start_year = 2024

        st.markdown("---")

        # Загрузка баланса
        balance_file = st.file_uploader(
            "📊 Бухгалтерский баланс",
            type=['xlsx', 'xls'],
            key="sidebar_balance",
            help="Файл .xlsx с бухгалтерским балансом"
        )
        if balance_file:
            st.success("✅ Баланс загружен")
            st.session_state.balance_file = balance_file
            st.session_state.balance_df = read_uploaded_excel(balance_file)
            st.session_state.project_data.reporting_data["balance"] = st.session_state.balance_df
        else:
            if 'balance_file' in st.session_state:
                st.success("✅ Баланс загружен (из сессии)")
            else:
                st.caption("Файл не выбран")

        # Загрузка отчета о финансовых результатах
        income_file = st.file_uploader(
            "📈 Отчет о фин. результатах",
            type=['xlsx', 'xls'],
            key="sidebar_income",
            help="Файл .xlsx с отчетом о прибылях и убытках"
        )
        if income_file:
            st.success("✅ Фин. результаты загружены")
            st.session_state.income_file = income_file
            st.session_state.income_df = read_uploaded_excel(income_file)
            st.session_state.project_data.reporting_data["income"] = st.session_state.income_df
        else:
            if 'income_file' in st.session_state:
                st.success("✅ Фин. результаты загружены (из сессии)")
            else:
                st.caption("Файл не выбран")

        # Загрузка ОДДС
        cash_flow_file = st.file_uploader(
            "💰 Отчет о движении денежных средств",
            type=['xlsx', 'xls'],
            key="sidebar_cf",
            help="Файл .xlsx с отчетом о движении денежных средств"
        )
        if cash_flow_file:
            st.success("✅ ОДДС загружен")
            st.session_state.cash_flow_file = cash_flow_file
            st.session_state.cash_flow_df = read_uploaded_excel(cash_flow_file)
            st.session_state.project_data.reporting_data["cashflow"] = st.session_state.cash_flow_df
        else:
            if 'cash_flow_file' in st.session_state:
                st.success("✅ ОДДС загружен (из сессии)")
            else:
                st.caption("Файл не выбран")

        st.markdown("---")

        # Информация о загруженных файлах
        st.subheader("📋 Статус загрузки")

        col1, col2, col3 = st.columns(3)
        with col1:
            if balance_file or st.session_state.get('balance_file'):
                st.markdown("📊 ✅")
            else:
                st.markdown("📊 ⏳")
        with col2:
            if income_file or st.session_state.get('income_file'):
                st.markdown("📈 ✅")
            else:
                st.markdown("📈 ⏳")
        with col3:
            if cash_flow_file or st.session_state.get('cash_flow_file'):
                st.markdown("💰 ✅")
            else:
                st.markdown("💰 ⏳")

        # Кнопка очистки всех файлов
        if st.button("🗑️ Очистить все файлы", use_container_width=True):
            for key in [
                'balance_file', 'income_file', 'cash_flow_file',
                'balance_df', 'income_df', 'cash_flow_df'
            ]:
                if key in st.session_state:
                    del st.session_state[key]

            st.session_state.project_data.reporting_data = {
                "balance": None,
                "income": None,
                "cashflow": None
            }

            st.rerun()

    tab_input, tab_reports, tab_revenue, tab_opex, tab_nwc, tab_invest, tab_fixed, tab_results, tab_export = st.tabs([
        "📝 Ввод данных",
        "📂 Загруженная отчетность",
        "💰 Продажи",
        "🏭 Операционные издержки",
        "🔄 Оборотный капитал",
        "🏗️ Инвестиции",
        "🏛️ Основной капитал",
        "📈 Анализ эффективности",
        "💾 Экспорт"
    ])

    with tab_input:
        # ========== БЛОК 0: ДИАПАЗОН ЛЕТ ==========
        st.header("📅 Период расчёта")

        col1, col2 = st.columns(2)
        with col1:
            start_year = st.number_input(
                "Начальный год",
                min_value=2000, max_value=2100, value=2016, step=1
            )
        with col2:
            end_year = st.number_input(
                "Конечный год",
                min_value=2000, max_value=2100, value=2035, step=1
            )

        st.session_state.project_data.years = list(range(start_year, end_year + 1))

        st.divider()

        # ========== БЛОК 1: СТАВКИ ==========
        st.header("Финансовые параметры")

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.session_state.project_data.discount_rate = st.slider(
                "Ставка дисконтирования (%)",
                min_value=0.0, max_value=30.0,
                value=float(st.session_state.project_data.discount_rate * 100)
                if st.session_state.project_data.discount_rate else 10.0,
                step=0.1
            ) / 100

        with col2:
            st.session_state.project_data.tax_rate = st.slider(
                "Ставка налога на прибыль (%)",
                min_value=0.0, max_value=30.0,
                value=float(st.session_state.project_data.tax_rate * 100)
                if st.session_state.project_data.tax_rate else 20.0,
                step=0.1
            ) / 100

        with col3:
            if "excise_rate" not in st.session_state:
                st.session_state.excise_rate = 0.0
            st.session_state.excise_rate = st.slider(
                "Ставка акциза (%)",
                min_value=0.0, max_value=50.0,
                value=st.session_state.excise_rate,
                step=0.1
            )

        with col4:
            if "property_tax_rate" not in st.session_state:
                st.session_state.property_tax_rate = 2.2  # стандартная ставка в РФ
            st.session_state.property_tax_rate = st.slider(
                "Ставка налога на имущество (%)",
                min_value=0.0, max_value=30.0,
                value=float(st.session_state.property_tax_rate),
                step=0.01
            )

        # ========== БЛОК: СТАВКА НДС (вставить сюда) ==========
        st.subheader("📋 Ставка НДС по годам")

        # Берём годы из session_state если уже заданы,
        # иначе используем дефолтный диапазон 2024-2030
        _years_for_vat = (
            st.session_state.project_data.years
            if st.session_state.project_data.years
            else list(range(2024, 2031))
        )

        # Инициализация дефолтных значений
        for _y in _years_for_vat:
            if _y not in st.session_state.vat_rates:
                st.session_state.vat_rates[_y] = 0.20  # 20% в долях

        # Горизонтальная таблица
        _vat_row = {"Показатель": "Ставка НДС (%)"}
        for _y in _years_for_vat:
            _vat_row[str(_y)] = round(st.session_state.vat_rates.get(_y, 0.20) * 100, 1)

        _df_vat = pd.DataFrame([_vat_row])

        _edited_vat = st.data_editor(
            _df_vat,
            key=f"vat_editor_{hash(str(_years_for_vat))}",
            num_rows="fixed",
            column_config={
                "Показатель": st.column_config.TextColumn(disabled=True),
                **{
                    str(_y): st.column_config.NumberColumn(
                        format="%.1f",
                        min_value=0.0,
                        max_value=100.0,
                        step=0.1
                    )
                    for _y in _years_for_vat
                }
            }
        )

        # Сохраняем в session_state в долях
        _need_vat_update = False
        for _y in _years_for_vat:
            _new = float(_edited_vat.iloc[0][str(_y)]) / 100
            _old = st.session_state.vat_rates.get(_y, 0.20)
            if abs(_new - _old) > 1e-9:
                _need_vat_update = True
                st.session_state.vat_rates[_y] = _new

        if _need_vat_update:
            st.rerun()

        # ========== БЛОК: УСРЕДНЕННАЯ СТРУКТУРА ЗАТРАТ ==========
        st.subheader("📊 Усредненная структура затрат")

        if 'avg_cost_structure' not in st.session_state:
            st.session_state.avg_cost_structure = {
                "Сырьё и материалы": 0.0,
                "Топливо": 0.0,
                "Энергия": 0.0,
                "Затраты на оплату труда": 0.0,
                "Отчисления на социальные нужды": 0.0,
                "Амортизация основных средств": 0.0,
                "Работы и услуги производственного характера, выполненные сторонними организациями, и приобретённые комплектующие изделия": 0.0,
                "Прочие затраты": 0.0,
            }

        _cost_structure_rows = [
            {"Статья затрат": item, "Доля (%)": share}
            for item, share in st.session_state.avg_cost_structure.items()
        ]

        _df_cost_structure = pd.DataFrame(_cost_structure_rows)

        _edited_cost_structure = st.data_editor(
            _df_cost_structure,
            key="avg_cost_structure_editor",
            num_rows="fixed",
            column_config={
                "Статья затрат": st.column_config.TextColumn(disabled=True),
                "Доля (%)": st.column_config.NumberColumn(
                    format="%.2f",
                    min_value=0.0,
                    max_value=100.0,
                    step=0.01
                )
            },
            use_container_width=True
        )

        # Сохраняем значения
        _need_cs_update = False
        for _, row in _edited_cost_structure.iterrows():
            item = row["Статья затрат"]
            new_val = float(row["Доля (%)"])
            old_val = st.session_state.avg_cost_structure.get(item, 0.0)
            if abs(new_val - old_val) > 1e-9:
                _need_cs_update = True
                st.session_state.avg_cost_structure[item] = new_val

        # Проверка суммы
        _total_cs = sum(st.session_state.avg_cost_structure.values())
        if abs(_total_cs - 100.0) < 0.01:
            st.success(f"✅ Сумма долей: {_total_cs:.2f}%")
        else:
            st.warning(f"⚠️ Сумма долей: {_total_cs:.2f}% (должно быть 100%)")

        if _need_cs_update:
            st.rerun()

        st.divider()

        st.subheader("⚙️ Операционные издержки")

        # --- Сырьё и материалы ---
        st.markdown("**🪨 Сырьё и материалы**")
        if 'raw_materials' not in st.session_state:
            st.session_state.raw_materials = []

        col_rm1, col_rm2 = st.columns([4, 1])
        with col_rm1:
            new_rm = st.text_input("Добавить вид сырья/материала:", key="new_raw_material")
        with col_rm2:
            st.write("")
            st.write("")
            if st.button("➕ Добавить", key="add_rm"):
                if new_rm and new_rm not in st.session_state.raw_materials:
                    st.session_state.raw_materials.append(new_rm)
                    st.rerun()

        # Отображение добавленных с кнопкой удаления
        for i, rm in enumerate(st.session_state.raw_materials):
            col_a, col_b = st.columns([5, 1])
            with col_a:
                st.write(f"• {rm}")
            with col_b:
                if st.button("❌", key=f"del_rm_{i}"):
                    st.session_state.raw_materials.pop(i)
                    st.rerun()

        # --- Топливо и энергия ---
        st.markdown("**⚡ Топливо и энергия**")
        if 'fuel_energy' not in st.session_state:
            st.session_state.fuel_energy = []

        col_fe1, col_fe2 = st.columns([4, 1])
        with col_fe1:
            new_fe = st.text_input("Добавить вид топлива/энергии:", key="new_fuel_energy")
        with col_fe2:
            st.write("")
            st.write("")
            if st.button("➕ Добавить", key="add_fe"):
                if new_fe and new_fe not in st.session_state.fuel_energy:
                    st.session_state.fuel_energy.append(new_fe)
                    st.rerun()

        for i, fe in enumerate(st.session_state.fuel_energy):
            col_a, col_b = st.columns([5, 1])
            with col_a:
                st.write(f"• {fe}")
            with col_b:
                if st.button("❌", key=f"del_fe_{i}"):
                    st.session_state.fuel_energy.pop(i)
                    st.rerun()

        all_cost_items = st.session_state.raw_materials + st.session_state.fuel_energy

        if all_cost_items:
            st.markdown("**💰 Цены по видам издержек по годам (руб./ед.)**")

            if 'cost_prices' not in st.session_state:
                st.session_state.cost_prices = {}

            _years = st.session_state.project_data.years or list(range(2024, 2031))

            # Строим DataFrame: строки = виды издержек, столбцы = годы
            price_rows = []
            for item in all_cost_items:
                row = {"Вид издержки": item}
                for year in _years:
                    row[str(year)] = st.session_state.cost_prices.get(item, {}).get(year, 0.0)
                price_rows.append(row)

            df_cost_prices = pd.DataFrame(price_rows)

            edited_cost_prices = st.data_editor(
                df_cost_prices,
                key=f"cost_prices_editor_{hash(str(all_cost_items))}_{hash(str(_years))}",
                num_rows="fixed",
                column_config={
                    "Вид издержки": st.column_config.TextColumn(disabled=True),
                    **{
                        str(year): st.column_config.NumberColumn(
                            format="%.1f",
                            min_value=0.0,
                            step=0.1
                        )
                        for year in _years
                    }
                }
            )

            # Сохраняем обратно в session_state
            _need_price_update = False
            for _, row in edited_cost_prices.iterrows():
                item = row["Вид издержки"]
                if item not in st.session_state.cost_prices:
                    st.session_state.cost_prices[item] = {}
                for year in _years:
                    new_val = float(row[str(year)])
                    old_val = st.session_state.cost_prices[item].get(year, 0.0)
                    if abs(new_val - old_val) > 1e-9:
                        _need_price_update = True
                        st.session_state.cost_prices[item][year] = new_val

            if _need_price_update:
                st.rerun()

        st.markdown("**📦 Объёмы по видам издержек и годам**")

        if 'cost_volumes' not in st.session_state:
            st.session_state.cost_volumes = {}

        _years = st.session_state.project_data.years or list(range(2024, 2031))

        vol_rows = []
        for item in all_cost_items:
            row = {"Вид издержки": item}
            for year in _years:
                row[str(year)] = st.session_state.cost_volumes.get(item, {}).get(year, 0.0)
            vol_rows.append(row)

        df_vol = pd.DataFrame(vol_rows)
        edited_vol = st.data_editor(
            df_vol,
            key=f"cost_volumes_editor_{hash(str(all_cost_items))}_{hash(str(_years))}",
            num_rows="fixed",
            column_config={
                "Вид издержки": st.column_config.TextColumn(disabled=True),
                **{str(y): st.column_config.NumberColumn(format="%.3f", min_value=0.0, step=0.001)
                   for y in _years}
            }
        )
        for _, row in edited_vol.iterrows():
            item = row["Вид издержки"]
            if item not in st.session_state.cost_volumes:
                st.session_state.cost_volumes[item] = {}
            for year in _years:
                st.session_state.cost_volumes[item][year] = float(row[str(year)])

        st.markdown("**👷 Средняя заработная плата**")
        if 'avg_salary' not in st.session_state:
            st.session_state.avg_salary = 73821.0
        st.session_state.avg_salary = st.number_input(
            "Средняя з/п в базовый год (руб./мес.)",
            min_value=0.0, step=1000.0,
            value=st.session_state.avg_salary,
            key="avg_salary_input"
        )

        st.markdown("**📋 ЦСЛЯБ и курс доллара по годам**")
        if 'cslyab_data' not in st.session_state:
            st.session_state.cslyab_data = {}

        cslyab_rows = []
        for label in ["ЦСЛЯБ", "Курс доллара"]:
            row = {"Показатель": label}
            for year in _years:
                row[str(year)] = st.session_state.cslyab_data.get(label, {}).get(year, 0.0)
            cslyab_rows.append(row)

        df_cslyab = pd.DataFrame(cslyab_rows)
        edited_cslyab = st.data_editor(
            df_cslyab,
            key=f"cslyab_editor_{hash(str(_years))}",
            num_rows="fixed",
            column_config={
                "Показатель": st.column_config.TextColumn(disabled=True),
                **{
                    str(y): st.column_config.NumberColumn(
                        format="%.1f",
                        min_value=0.0,
                        step=0.1
                    )
                    for y in _years
                }
            }
        )
        for _, row in edited_cslyab.iterrows():
            label = row["Показатель"]
            if label not in st.session_state.cslyab_data:
                st.session_state.cslyab_data[label] = {}
            for year in _years:
                st.session_state.cslyab_data[label][year] = float(row[str(year)])

        st.markdown("**🚛 Транспортировка**")
        col_tr1, col_tr2 = st.columns(2)
        with col_tr1:
            if 'transport_price' not in st.session_state:
                st.session_state.transport_price = 0.0
            st.session_state.transport_price = st.number_input(
                "Цена за транспортировку 1 тонны (руб.)",
                min_value=0.0, step=100.0,
                value=st.session_state.transport_price,
                key="transport_price_input"
            )
        with col_tr2:
            transport_options = st.session_state.raw_materials + st.session_state.fuel_energy
            if transport_options:
                if 'transport_product' not in st.session_state:
                    st.session_state.transport_product = transport_options[0]
                st.session_state.transport_product = st.selectbox(
                    "Продукция для транспортировки",
                    options=transport_options,
                    key="transport_product_select"
                )
            else:
                st.info("Сначала добавьте виды сырья или топлива")

        # ========== БЛОК 2: РАСПРЕДЕЛЕНИЕ ИНВЕСТИЦИЙ И АМОРТИЗАЦИЯ ==========
        st.subheader("Распределение инвестиций и амортизация")

        col_left, col_center, col_right = st.columns([1, 1, 1])

        with col_left:
            equipment_share_pct = st.number_input(
                "Доля оборудования (%)",
                min_value=0.0, max_value=100.0, value=56.2, step=1.0
            )
            st.session_state.project_data.equipment_share = equipment_share_pct / 100

            equipment_dep_pct = st.slider(
                "Норма амортизации оборудования (%)",
                min_value=0.0, max_value=50.0, value=15.0, step=0.1
            )
            st.session_state.project_data.equipment_depreciation = equipment_dep_pct / 100

            # ✅ Новый параметр
            pir_share_pct = st.number_input(
                "Доля ПИР, %",
                min_value=0.0, max_value=100.0,
                value=0.405,
                step=0.001,
                format="%.3f"
            )
            st.session_state.pir_share_pct = pir_share_pct

        with col_center:
            construction_share = 100 - equipment_share_pct - pir_share_pct

            # ✅ Три сегмента вместо двух
            fig_pie = go.Figure(data=[go.Pie(
                labels=['Оборудование', 'Строительно-монтажные работы', 'Проектно-изыскательные работы'],
                values=[
                    max(equipment_share_pct, 0),
                    max(construction_share, 0),
                    max(pir_share_pct, 0)
                ],
                hole=0.4,
                marker_colors=['#1f77b4', '#ff7f0e', '#2ca02c']
            )])
            fig_pie.update_layout(title="Распределение инвестиций", height=300)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            # ✅ СМР рассчитывается как остаток
            construction_share = 100 - equipment_share_pct - pir_share_pct

            st.text_input(
                "Доля строительно-монтажных работ (%)",
                value=f"{construction_share:.1f}",
                disabled=True
            )
            if construction_share < 0:
                st.error("⚠️ Сумма долей превышает 100%")

            construction_dep_pct = st.slider(
                "Норма амортизации СМР (%)",
                min_value=0.0, max_value=50.0, value=2.5, step=0.1
            )
            st.session_state.project_data.construction_depreciation = construction_dep_pct / 100

        # ========== БЛОК 4: ВЫБОР ПРОДУКЦИИ И ОБЪЕМОВ ==========
        st.subheader("📦 Выбор продукции и объемов")

        # Справочник ОКПД2 импортируется из data/okpd_catalog.py
        okpd_catalog = OKPD_CATALOG

        # Выбор отрасли
        selected_sector = st.selectbox(
            "Выберите отрасль:",
            options=list(okpd_catalog.keys()),
            help="Выберите основную отрасль деятельности"
        )

        # Выбор ОКПД в выбранной отрасли
        if selected_sector:
            sector_codes = okpd_catalog[selected_sector]
            selected_okpd = st.selectbox(
                "Выберите ОКПД2 продукции:",
                options=list(sector_codes.keys()),
                format_func=lambda x: f"{x} - {sector_codes[x]}",
                help="Выберите код ОКПД2 для вашего типа продукции"
            )

        # Справочник типов продукции импортируется из data/product_types.py
        product_types_mapping = PRODUCT_TYPES_MAPPING

        # Получаем ВСЕ возможные типы продукции для выбранного ОКПД
        default_product_types = product_types_mapping.get(selected_okpd, product_types_mapping["default"]).copy()

        # Добавляем сохраненные пользовательские типы для этого ОКПД
        custom_types = st.session_state.custom_products_by_okpd.get(selected_okpd, [])
        all_product_types = default_product_types + custom_types

        # Возможность добавить свой тип продукции
        st.markdown("---")
        st.subheader("➕ Добавление нового типа продукции")

        col_add1, col_add2 = st.columns([3, 1])
        with col_add1:
            new_product = st.text_input("Название нового типа продукции:", key=f"new_product_{selected_okpd}")
        with col_add2:
            if st.button("➕ Добавить", key=f"add_btn_{selected_okpd}"):
                if new_product and new_product not in all_product_types:
                    if selected_okpd not in st.session_state.custom_products_by_okpd:
                        st.session_state.custom_products_by_okpd[selected_okpd] = []
                    if new_product not in st.session_state.custom_products_by_okpd[selected_okpd]:
                        st.session_state.custom_products_by_okpd[selected_okpd].append(new_product)
                        st.success(f"✅ Добавлен тип продукции: {new_product}")
                        st.rerun()
                elif new_product in all_product_types:
                    st.warning("⚠️ Такой тип продукции уже существует")

        # Отображаем кнопку для сброса пользовательских типов
        if custom_types:
            if st.button("🗑️ Удалить все добавленные типы", key=f"reset_custom_{selected_okpd}"):
                st.session_state.custom_products_by_okpd[selected_okpd] = []
                st.rerun()

        st.markdown("---")

        # Обновляем all_product_types после возможного добавления
        all_product_types = default_product_types + st.session_state.custom_products_by_okpd.get(selected_okpd, [])

        # ========== ВЫБОР ПРОДУКТОВ ДЛЯ РАСЧЕТА ==========
        st.subheader("✅ Выберите продукты для анализа")

        # Инициализация выбранных продуктов в session_state
        selection_key = f"selected_products_{selected_okpd}"
        if selection_key not in st.session_state:
            st.session_state[selection_key] = {product: True for product in all_product_types}

        # Отображаем чекбоксы для каждого продукта
        cols = st.columns(4)
        for idx, product in enumerate(all_product_types):
            col_idx = idx % 4
            with cols[col_idx]:
                is_custom = product in custom_types
                label = f"🆕 {product}" if is_custom else product

                st.session_state[selection_key][product] = st.checkbox(
                    label,
                    value=st.session_state[selection_key].get(product, True),
                    key=f"chk_{selected_okpd}_{product}"
                )

                if is_custom:
                    if st.button("❌", key=f"del_{selected_okpd}_{product}", help="Удалить этот тип"):
                        if product in st.session_state.custom_products_by_okpd[selected_okpd]:
                            st.session_state.custom_products_by_okpd[selected_okpd].remove(product)
                            if product in st.session_state[selection_key]:
                                del st.session_state[selection_key][product]
                            st.rerun()

        # Получаем список ТОЛЬКО выбранных продуктов
        selected_products = [
            product for product in all_product_types
            if st.session_state[selection_key].get(product, False)
        ]

        current_product_set = frozenset(selected_products)
        if st.session_state.get("last_product_set") != current_product_set:
            st.session_state.last_product_set = current_product_set
            if "sales_matrix_df" in st.session_state:
                del st.session_state["sales_matrix_df"]

        st.divider()

        # ========== ВВОД ДАННЫХ ТОЛЬКО ДЛЯ ВЫБРАННЫХ ПРОДУКТОВ ==========
        st.subheader("📊 Ввод объемов и цен для выбранных продуктов")

        # Выбор единиц измерения
        st.subheader("📏 Настройка единиц измерения")

        # Единицы по умолчанию импортируются из data/okpd_catalog.py
        default_units = DEFAULT_UNITS

        # Все возможные единицы измерения импортируются из data/okpd_catalog.py
        all_units = list(ALL_UNITS)

        unit_key = f"unit_{selected_okpd}"
        if unit_key not in st.session_state:
            st.session_state[unit_key] = default_units.get(selected_sector, "ед.")

        years_list = st.session_state.project_data.years

        # Словари для хранения данных
        product_volumes = {}
        product_export_shares = {}
        product_import_shares = {}
        product_base_prices = {}

        # Для каждого выбранного продукта создаем свой блок
        for product_idx, product_type in enumerate(selected_products):
            st.markdown(f"### {product_idx + 1}. {product_type}")

            # ========== ВЫБОР ЕДИНИЦ ИЗМЕРЕНИЯ ДЛЯ ЭТОГО ПРОДУКТА ==========
            # Три колонки: единицы измерения, цена и объемы
            col_unit, col_price, col_vol = st.columns([1, 1, 2])

            # Все возможные единицы измерения
            all_units = list(ALL_UNITS)

            # Ключ для хранения единицы измерения продукта
            unit_key = f"unit_{selected_okpd}_{product_type}"

            # Инициализация единицы измерения для продукта
            if unit_key not in st.session_state:
                # Определяем единицу по умолчанию в зависимости от отрасли
                st.session_state[unit_key] = DEFAULT_UNITS.get(selected_sector, "ед.")

            with col_unit:
                # Выбор единицы измерения
                product_unit = st.selectbox(
                    "📏 Единица измерения",
                    options=all_units,
                    index=all_units.index(st.session_state[unit_key]) if st.session_state[unit_key] in all_units else 0,
                    key=f"unit_select_{selected_okpd}_{product_type}",
                    help=f"Выберите единицу измерения для {product_type}"
                )
                st.session_state[unit_key] = product_unit

            with col_price:
                # Базовая цена продукта (с указанием единицы измерения)
                price_key = f"price_{selected_okpd}_{product_type}"
                default_price = st.session_state.get(price_key, 1000.0)
                base_price = st.number_input(
                    f"💰 Цена (руб.)",
                    value=default_price,
                    step=100.0,
                    key=price_key
                )
                product_base_prices[product_type] = base_price
                st.session_state.product_base_prices = product_base_prices

            with col_vol:
                # Таблица объемов с выбранной единицей измерения
                st.write(f"Объемы производства")

                # Инициализация хранилищ объёмов/экспорта (один раз) — аналог price_indices_temp
                if 'volumes_temp' not in st.session_state:
                    st.session_state.volumes_temp = {}
                if 'exports_temp' not in st.session_state:
                    st.session_state.exports_temp = {}

                vol_store = st.session_state.volumes_temp.setdefault(selected_okpd, {}).setdefault(product_type, {})
                exp_store = st.session_state.exports_temp.setdefault(selected_okpd, {}).setdefault(product_type, {})

                # Строим DataFrame только из vol_store / exp_store, БЕЗ вычисляемой колонки импорта
                volumes_data = []
                for year in years_list:
                    default_volume = vol_store.get(year, 0.0)
                    default_export = exp_store.get(year, 0.0)
                    volumes_data.append({
                        'Год': year,
                        f'Объем ({product_unit})': default_volume,
                        'Доля экспорта (%)': default_export,
                        'Доля импорта (%)': 100.0 - default_export,
                    })

                df_volumes = pd.DataFrame(volumes_data)

                # Стабильный ключ редактора (зависит только от продукта и набора лет)
                editor_key = f"vol_editor_{selected_okpd}_{product_type}_{hash(tuple(years_list))}"

                edited_df = st.data_editor(
                    df_volumes,
                    key=editor_key,
                    num_rows="fixed",
                    column_config={
                        'Год': st.column_config.NumberColumn(disabled=True),
                        f'Объем ({product_unit})': st.column_config.NumberColumn(
                            format="%.3f", min_value=0.0, step=0.001
                        ),
                        'Доля экспорта (%)': st.column_config.NumberColumn(
                            format="%.1f", min_value=0.0, max_value=100.0, step=0.1
                        ),
                        'Доля импорта (%)': st.column_config.NumberColumn(
                            format="%.1f",
                            min_value=0.0,
                            max_value=100.0,
                            step=0.1,
                        ),
                    }
                )

                product_volumes[product_type] = {}
                product_export_shares[product_type] = {}
                product_import_shares[product_type] = {}

                need_update = False
                for _, row in edited_df.iterrows():
                    year = int(row['Год'])
                    volume = float(row[f'Объем ({product_unit})'])
                    export_pct = float(row['Доля экспорта (%)'])
                    export_pct = max(0.0, min(100.0, export_pct))
                    # Колонку «Доля импорта (%)» из таблицы НЕ читаем —
                    # она всегда производная от export_pct.
                    import_pct = 100.0 - export_pct

                    old_volume = vol_store.get(year, 0.0)
                    old_export = exp_store.get(year, 0.0)

                    if abs(old_volume - volume) > 1e-9 or abs(old_export - export_pct) > 1e-9:
                        need_update = True
                        vol_store[year] = volume
                        exp_store[year] = export_pct

                    product_volumes[product_type][year] = volume
                    product_export_shares[product_type][year] = export_pct / 100
                    product_import_shares[product_type][year] = import_pct / 100
                # Сохраняем данные текущего продукта в общую модель проекта
                st.session_state.project_data.products[product_type] = product_volumes[product_type]
                st.session_state.project_data.export_shares[product_type] = product_export_shares[product_type]
                st.session_state.project_data.import_shares[product_type] = product_import_shares[product_type]
                if need_update:
                    st.rerun()

        st.session_state.project_data.products = {k: v for k, v in product_volumes.items() if k in selected_products}
        st.session_state.project_data.export_shares = {k: v for k, v in product_export_shares.items() if
                                                       k in selected_products}
        st.session_state.project_data.import_shares = {k: v for k, v in product_import_shares.items() if
                                                       k in selected_products}
        st.session_state.product_base_prices = {k: v for k, v in st.session_state.product_base_prices.items() if
                                                k in selected_products}

        # ========== БЛОК 5: ЦЕНЫ (общие индексы и дефляторы) ==========
        st.subheader("💰 Ценовые индексы и дефляторы")

        # Инициализация структур для хранения индексов
        if 'price_indices_temp' not in st.session_state:
            st.session_state.price_indices_temp = {}
        if 'base_year' not in st.session_state:
            if st.session_state.project_data.years:
                st.session_state.base_year = st.session_state.project_data.years[0]
            else:
                st.session_state.base_year = 2024

        # Выбор базового года
        if st.session_state.project_data.years:
            current_index = 0
            if st.session_state.base_year in st.session_state.project_data.years:
                current_index = st.session_state.project_data.years.index(st.session_state.base_year)
            else:
                current_index = 0
                st.session_state.base_year = st.session_state.project_data.years[0]

            st.session_state.base_year = st.selectbox(
                "Базовый год для дефлятора",
                st.session_state.project_data.years,
                index=current_index,
                key="base_year_select"
            )

        if st.session_state.project_data.years:
            st.write("**Ценовые индексы по годам (базовый год = 1.00):**")

            # Уникальный ключ для редактора
            editor_key = f"price_editor_{hash(str(st.session_state.project_data.years))}_{st.session_state.base_year}"

            # Создаем DataFrame с текущими значениями
            price_data = []
            for year in st.session_state.project_data.years:
                if year in st.session_state.price_indices_temp:
                    default_index = st.session_state.price_indices_temp[year]
                elif year in st.session_state.project_data.price_indices:
                    default_index = st.session_state.project_data.price_indices[year]
                else:
                    default_index = 1.0 if year == st.session_state.base_year else 1.0
                price_data.append({'Год': year, 'Ценовой индекс': default_index})

            df = pd.DataFrame(price_data)

            # Редактируемая таблица индексов
            edited_df = st.data_editor(
                df,
                key=editor_key,
                num_rows="fixed",
                column_config={
                    'Год': st.column_config.NumberColumn(disabled=True),
                    'Ценовой индекс': st.column_config.NumberColumn(
                        format="%.5f",
                        min_value=0.0,
                        max_value=10.0,
                        step=0.0001
                    )
                }
            )

            # Обновляем значения индексов
            need_update = False
            for _, row in edited_df.iterrows():
                year = int(row['Год'])
                index = float(row['Ценовой индекс'])

                old_index = st.session_state.price_indices_temp.get(year, 0)
                if abs(old_index - index) > 0.00001:
                    need_update = True
                    st.session_state.price_indices_temp[year] = index
                    st.session_state.project_data.price_indices[year] = index

            if need_update:
                st.rerun()

            # ========== РАСЧЕТ ЦЕН ПО ПРОДУКТАМ И ГОДАМ ==========
            if st.session_state.price_indices_temp and st.session_state.get('product_base_prices'):
                st.subheader("📊 Расчет цен по продуктам и годам")

                # Получаем текущие индексы и рассчитываем дефляторы
                # (функция calculate_deflator импортируется из utils.calculations)
                current_indices = st.session_state.price_indices_temp.copy()
                deflators = {}
                for year in st.session_state.project_data.years:
                    deflators[year] = calculate_deflator(year, st.session_state.base_year, current_indices)

                # СОХРАНЯЕМ ЦЕНЫ ДЛЯ КАЖДОГО ПРОДУКТА И ГОДА В project_data
                # Это важно для последующих расчетов выручки
                st.session_state.project_data.prices = {}  # Очищаем старые цены
                st.session_state.project_data.product_prices = {}  # Новая структура для хранения цен по продуктам

                # Перезаписываем фиксированными ценами если они заданы
                fixed_prices = st.session_state.get("fixed_product_prices", {})

                for product_type in st.session_state.project_data.product_prices:
                    if product_type in fixed_prices:
                        for year in st.session_state.project_data.years:
                            if year in fixed_prices[product_type]:
                                st.session_state.project_data.product_prices[product_type][year] = \
                                    fixed_prices[product_type][year]

                # Создаем таблицу для отображения всех цен
                all_prices_data = []

                for product_type, base_price in st.session_state.product_base_prices.items():
                    st.session_state.project_data.product_prices[product_type] = {}

                    for year in st.session_state.project_data.years:
                        # Цена с учетом дефлятора
                        calculated_price = base_price * deflators[year]
                        st.session_state.project_data.product_prices[product_type][year] = calculated_price

                        # Сохраняем также в общий словарь для обратной совместимости
                        # Используем составной ключ "продукт_год"
                        st.session_state.project_data.prices[f"{product_type}_{year}"] = calculated_price

                        # Добавляем в данные для таблицы
                        all_prices_data.append({
                            'Продукт': product_type,
                            'Год': year,
                            'Базовая цена': base_price,
                            'Дефлятор': deflators[year],
                        })

                # Создаем и отображаем сводную таблицу
                df_prices = pd.DataFrame(all_prices_data)

                # Сводная таблица для лучшего отображения
                pivot_prices = df_prices.pivot_table(
                    index='Продукт',
                    columns='Год',
                    fill_value=0
                )

                st.write("**Цены по продуктам и годам (руб.):**")
                st.dataframe(
                    pivot_prices.style.format("{:,.2f}"),
                    use_container_width=True
                )

                # Детальная таблица с дефляторами
                with st.expander("📋 Детальная таблица с дефляторами"):
                    st.dataframe(
                        df_prices.style.format({
                            'Базовая цена': '{:,.2f}',
                            'Дефлятор': '{:.5f}',
                        }),
                        use_container_width=True,
                        hide_index=True
                    )



        # Кнопка сохранения
        if st.button("💾 Сохранить все данные", type="primary"):
            try:
                with open('project_data.json', 'w', encoding='utf-8') as f:
                    json.dump(st.session_state.project_data.to_dict(), f, ensure_ascii=False, indent=2)
                st.success("✅ Данные сохранены!")
            except Exception as e:
                st.error(f"Ошибка сохранения: {e}")

    with tab_reports:
        st.header("Загруженные формы отчетности")

        report_tab1, report_tab2, report_tab3 = st.tabs([
            "📊 Бухгалтерский баланс",
            "📈 Финансовые результаты",
            "💰 ОДДС"
        ])

        with report_tab1:
            if st.session_state.get("balance_df") is not None:
                st.subheader("Содержимое файла: Бухгалтерский баланс")
                st.dataframe(st.session_state.balance_df, use_container_width=True, height=500)
            else:
                st.info("Файл бухгалтерского баланса пока не загружен.")

        with report_tab2:
            if st.session_state.get("income_df") is not None:
                st.subheader("Содержимое файла: Отчет о финансовых результатах")
                st.dataframe(st.session_state.income_df, use_container_width=True, height=500)
            else:
                st.info("Файл отчета о финансовых результатах пока не загружен.")

        with report_tab3:
            if st.session_state.get("cash_flow_df") is not None:
                st.subheader("Содержимое файла: Отчет о движении денежных средств")
                st.dataframe(st.session_state.cash_flow_df, use_container_width=True, height=500)
            else:
                st.info("Файл ОДДС пока не загружен.")

            # Получаем файлы из session_state
            income_file = st.session_state.get("income_file")
            cash_flow_file = st.session_state.get("cash_flow_file")

            # Парсим нужные строки
            revenue_2110_from_file = extract_row_by_code(st.session_state.get('income_df'), 2110)
            receipts_4112_from_file = extract_row_by_code(st.session_state.get('cash_flow_df'), 4112)

            # Инициализируем session_state и сразу заполняем из файлов
            if 'revenue_without_vat' not in st.session_state:
                st.session_state.revenue_without_vat = {}
            if 'other_operating_receipts' not in st.session_state:
                st.session_state.other_operating_receipts = {}

            # Записываем данные из файлов для лет, которые в них есть
            _years = st.session_state.project_data.years
            if not _years:
                _years = list(range(2024, 2031))

            for year in _years:
                if year in revenue_2110_from_file:
                    st.session_state.revenue_without_vat[year] = revenue_2110_from_file[year]
                elif year not in st.session_state.revenue_without_vat:
                    st.session_state.revenue_without_vat[year] = 0.0

                if year in receipts_4112_from_file:
                    st.session_state.other_operating_receipts[year] = receipts_4112_from_file[year]
                elif year not in st.session_state.other_operating_receipts:
                    st.session_state.other_operating_receipts[year] = 0.0

            report_data = []
            for year in _years:
                report_data.append({
                    "Год": year,
                    "Выручка без НДС (2110)": st.session_state.revenue_without_vat.get(year, 0.0),
                    "Прочие поступления от операций": st.session_state.other_operating_receipts.get(year, 0.0),
                })

            if report_data:
                df_report = pd.DataFrame(report_data)
                edited_report = st.data_editor(
                    df_report,
                    key="report_editor",
                    num_rows="fixed",
                    column_config={
                        "Год": st.column_config.NumberColumn(disabled=True),
                        "Выручка без НДС (2110)": st.column_config.NumberColumn(format="%.2f"),
                        "Прочие поступления от операций": st.column_config.NumberColumn(format="%.2f"),
                    }
                )

                # ✅ Цикл теперь тоже внутри if — edited_report гарантированно существует
                for _, row in edited_report.iterrows():
                    year = int(row["Год"])
                    st.session_state.revenue_without_vat[year] = float(row["Выручка без НДС (2110)"])
                    st.session_state.other_operating_receipts[year] = float(row["Прочие поступления от операций"])
            else:
                st.info("Нет данных для отображения. Загрузите файлы отчётности в сайдбаре.")

    with tab_revenue:  # ← этого блока нет, нужно добавить
        st.header("💰 Продажи")

        if not st.session_state.project_data.years:
            st.warning("Сначала укажите диапазон лет во вкладке 'Ввод данных'")
        elif not st.session_state.project_data.products:
            st.warning("Сначала добавьте продукты во вкладке 'Ввод данных'")
        else:
            current_indices = st.session_state.price_indices_temp.copy() \
                if 'price_indices_temp' in st.session_state else {}

            deflators = {
                year: calculate_deflator(year, st.session_state.base_year, current_indices)
                for year in st.session_state.project_data.years
            }

            last_year = st.session_state.project_data.years[-1]
            liquidation_by_year = {year: 0.0 for year in st.session_state.project_data.years}
            liquidation_by_year[last_year] = (
                    st.session_state.get("increment_nwc_last_year", 0.0) +
                    st.session_state.get("residual_value_last_year", 0.0)
            )

            sales_matrix_df = build_sales_matrix(
                project_data=st.session_state.project_data,
                years=st.session_state.project_data.years,
                deflators=deflators,
                vat_rates=st.session_state.vat_rates,
                other_operating_receipts=st.session_state.other_operating_receipts,
                liquidation_by_year=liquidation_by_year,
                revenue_without_vat_report=st.session_state.revenue_without_vat,
                product_base_prices=st.session_state.product_base_prices
            )

            st.session_state.sales_matrix_df = sales_matrix_df

            format_dict = {year: "{:,.2f}" for year in st.session_state.project_data.years}

            st.dataframe(
                sales_matrix_df.style.format(format_dict),
                use_container_width=True,
                hide_index=True
            )

    with tab_opex:
        st.header("🏭 Операционные издержки")

        if not st.session_state.project_data.years:
            st.warning("Сначала укажите диапазон лет во вкладке 'Ввод данных'")
        elif not st.session_state.get('cost_prices'):
            st.warning("Сначала заполните данные по издержкам во вкладке 'Ввод данных'")
        else:
            _years = st.session_state.project_data.years

            current_indices = st.session_state.price_indices_temp.copy() \
                if 'price_indices_temp' in st.session_state else {}
            deflators = {
                year: calculate_deflator(year, st.session_state.base_year, current_indices)
                for year in _years
            }

            # Доли инвестиций
            _eq_share = st.session_state.project_data.equipment_share
            _pir_share = st.session_state.get('pir_share_pct', 0.0) / 100.0
            _con_share = 1.0 - _eq_share - _pir_share

            opex_df = build_opex_matrix(
                project_data=st.session_state.project_data,
                years=_years,
                deflators=deflators,
                vat_rates=st.session_state.vat_rates,
                cost_prices=st.session_state.cost_prices,
                cost_volumes=st.session_state.cost_volumes,
                avg_cost_structure=st.session_state.avg_cost_structure,
                avg_salary=st.session_state.avg_salary,
                cslyab_data=st.session_state.cslyab_data,
                transport_price=st.session_state.get('transport_price', 0.0),
                transport_product=st.session_state.get('transport_product', None),
                cash_flow_df=st.session_state.get('cash_flow_df'),
                income_df=st.session_state.get('income_df'),
                equipment_share=_eq_share,
                construction_share=_con_share,
                pir_share=_pir_share,
                equipment_dep=st.session_state.project_data.equipment_depreciation,
                construction_dep=st.session_state.project_data.construction_depreciation,
            )

            format_dict = {year: "{:,.2f}" for year in _years}
            st.dataframe(
                opex_df.style.format(format_dict),
                use_container_width=True,
                hide_index=True
            )

    with tab_nwc:
        st.header("🔄 Оборотный капитал")

        balance_df = st.session_state.get("balance_df")
        _years = st.session_state.project_data.years

        if balance_df is None:
            st.info("Загрузите бухгалтерский баланс в сайдбаре для расчёта оборотного капитала.")
        elif not _years:
            st.warning("Сначала укажите диапазон лет во вкладке 'Ввод данных'.")
        else:
            result = build_nwc_matrix(balance_df, _years)

            # build_nwc_matrix возвращает либо пустой DataFrame, либо (df, reported_years)
            if isinstance(result, tuple):
                nwc_df, reported_years = result
            else:
                nwc_df = result
                reported_years = []

            if nwc_df.empty:
                st.warning(
                    "Не удалось извлечь данные баланса. "
                    "Проверьте, что файл содержит строки с кодами 1210, 1230, 1220, "
                    "1250, 1260, 1420, 1430, 1450."
                )
            else:
                # Сохраняем в session_state для использования в других блоках
                st.session_state.nwc_df = nwc_df
                st.session_state.reported_years_nwc = reported_years

                format_dict = {year: "{:,.2f}" for year in reported_years}

                st.dataframe(
                    nwc_df.style.format(format_dict, na_rep=""),
                    use_container_width=True,
                    hide_index=True,
                )

                # Визуализация динамики ЧОК
                nwc_row = nwc_df[nwc_df["Наименование статьи"] == "Чистый оборотный капитал"]
                if not nwc_row.empty:
                    nwc_values = [nwc_row.iloc[0].get(y, 0.0) for y in reported_years]

                    fig_nwc = go.Figure()
                    fig_nwc.add_trace(go.Bar(
                        x=reported_years,
                        y=nwc_values,
                        name="ЧОК",
                        marker_color=[
                            "#1f77b4" if v >= 0 else "#d62728"
                            for v in nwc_values
                        ]
                    ))
                    fig_nwc.add_hline(y=0, line_dash="dash", line_color="gray")
                    fig_nwc.update_layout(
                        title="Динамика чистого оборотного капитала",
                        xaxis_title="Год",
                        yaxis_title="Тыс. руб.",
                        height=350,
                    )
                    st.plotly_chart(fig_nwc, use_container_width=True)

    with tab_invest:
        st.header("🏗️ Инвестиции")

        cash_flow_df = st.session_state.get("cash_flow_df")
        _years = st.session_state.project_data.years

        if cash_flow_df is None:
            st.info("Загрузите ОДДС в сайдбаре для расчёта блока инвестиций.")
        elif not _years:
            st.warning("Сначала укажите диапазон лет во вкладке 'Ввод данных'.")
        else:
            currentindices = st.session_state.price_indices_temp.copy() \
                if "price_indices_temp" in st.session_state else {}
            deflators = {
                year: calculate_deflator(year, st.session_state.base_year, currentindices)
                for year in _years
            }

            eq_share = st.session_state.project_data.equipment_share
            pir_share = st.session_state.get("pir_share_pct", 0.0) / 100.0
            con_share = 1.0 - eq_share - pir_share

            invest_df = build_invest_matrix(
                cash_flow_df=cash_flow_df,
                years=_years,
                deflators=deflators,
                vat_rates=st.session_state.vat_rates,
                equipment_share=eq_share,
                construction_share=con_share,
                pir_share=pir_share,
            )

            st.session_state.invest_df = invest_df

            format_dict = {year: "{:,.2f}" for year in _years}

            st.dataframe(
                invest_df.style.format(format_dict, na_rep=""),
                use_container_width=True,
                hide_index=True,
            )

            # ── Визуализация структуры инвестиций ──────────────────────
            equip_row = invest_df[invest_df["Наименование статьи"] == "Оборудование"]
            smr_row = invest_df[invest_df["Наименование статьи"] == "Строительно-монтажные работы"]
            pir_row = invest_df[invest_df["Наименование статьи"] == "Прочие инвестиции"]

            if not equip_row.empty:
                equip_vals = [equip_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]
                smr_vals = [smr_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]
                pir_vals = [pir_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]

                fig_inv = go.Figure()
                fig_inv.add_trace(go.Bar(x=_years, y=equip_vals, name="Оборудование"))
                fig_inv.add_trace(go.Bar(x=_years, y=smr_vals, name="СМР"))
                fig_inv.add_trace(go.Bar(x=_years, y=pir_vals, name="Прочие (ПИР)"))
                fig_inv.update_layout(
                    barmode="stack",
                    title="Структура инвестиций во внеоборотные активы (постоянные цены)",
                    xaxis_title="Год",
                    yaxis_title="Тыс. руб.",
                    height=380,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_inv, use_container_width=True)

    with tab_fixed:
        st.header("🏛️ Основной капитал")

        _years = st.session_state.project_data.years
        inv_equip_raw = st.session_state.get("inv_equip_raw")
        inv_smr_raw = st.session_state.get("inv_smr_raw")
        inv_pir_raw = st.session_state.get("inv_pir_raw")

        if not _years:
            st.warning("Сначала укажите диапазон лет во вкладке 'Ввод данных'.")
        elif inv_equip_raw is None:
            st.info("Сначала откройте вкладку '🏗️ Инвестиции' — данные рассчитаются автоматически.")
        else:
            currentindices = st.session_state.price_indices_temp.copy() \
                if "price_indices_temp" in st.session_state else {}
            deflators = {
                year: calculate_deflator(year, st.session_state.base_year, currentindices)
                for year in _years
            }

            fixed_df = build_fixed_assets_matrix(
                years=_years,
                deflators=deflators,
                vat_rates=st.session_state.vat_rates,
                inv_equip_raw=inv_equip_raw,
                inv_smr_raw=inv_smr_raw,
                inv_pir_raw=inv_pir_raw,
                equipment_dep=st.session_state.project_data.equipment_depreciation,
                construction_dep=st.session_state.project_data.construction_depreciation,
                property_tax_rate=st.session_state.get("property_tax_rate", 2.2),
            )

            st.session_state.fixed_assets_df = fixed_df

            format_dict = {year: "{:,.2f}" for year in _years}

            st.dataframe(
                fixed_df.style.format(format_dict, na_rep=""),
                use_container_width=True,
                hide_index=True,
            )

            # ── Визуализация остаточной стоимости и амортизации ──────────
            res_row = fixed_df[fixed_df["Наименование статьи"] == "Остаточная стоимость"]
            amo_row = fixed_df[fixed_df["Наименование статьи"] == "Амортизационные отчисления"]
            tax_row = fixed_df[fixed_df["Наименование статьи"] == "Налог на имущество"]

            if not res_row.empty:
                res_vals = [res_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]
                amo_vals = [amo_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]
                tax_vals = [tax_row.iloc[0].get(y, 0.0) or 0.0 for y in _years]

                fig_fa = go.Figure()
                fig_fa.add_trace(go.Bar(
                    x=_years, y=res_vals, name="Остаточная стоимость",
                    marker_color="#1f77b4"
                ))
                fig_fa.add_trace(go.Scatter(
                    x=_years, y=amo_vals, name="Амортизация",
                    mode="lines+markers", yaxis="y2",
                    line=dict(color="#ff7f0e", width=2)
                ))
                fig_fa.add_trace(go.Scatter(
                    x=_years, y=tax_vals, name="Налог на имущество",
                    mode="lines+markers", yaxis="y2",
                    line=dict(color="#d62728", width=2, dash="dot")
                ))
                fig_fa.update_layout(
                    title="Остаточная стоимость, амортизация и налог на имущество",
                    xaxis_title="Год",
                    yaxis=dict(title="Остаточная стоимость, тыс. руб."),
                    yaxis2=dict(
                        title="Амортизация / Налог, тыс. руб.",
                        overlaying="y", side="right"
                    ),
                    height=400,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                )
                st.plotly_chart(fig_fa, use_container_width=True)

    with tab_results:
        st.header("Анализ эффективности инвестиций")

        if st.session_state.revenue_data:
            st.subheader("Инвестиционные затраты")
            total_investment = st.number_input("Общие инвестиции (руб.)", value=10000000, step=1000000)

            equipment_invest = total_investment * st.session_state.project_data.equipment_share
            construction_invest = total_investment * (1 - st.session_state.project_data.equipment_share)

            col1, col2 = st.columns(2)
            with col1:
                st.metric("Инвестиции в оборудование", f"{equipment_invest:,.0f} руб.")
                st.metric("Амортизация оборудования",
                          f"{equipment_invest * st.session_state.project_data.equipment_depreciation:,.0f} руб./год")
            with col2:
                st.metric("Инвестиции в СМР", f"{construction_invest:,.0f} руб.")
                st.metric("Амортизация СМР",
                          f"{construction_invest * st.session_state.project_data.construction_depreciation:,.0f} руб./год")

            st.subheader("Денежные потоки")
            operating_costs_pct = st.slider("Операционные расходы (% от выручки)", 0, 100, 60) / 100

            cash_flows = {}
            for year in st.session_state.project_data.years:
                revenue = st.session_state.revenue_data['total_revenue'][year]
                costs = revenue * operating_costs_pct

                depreciation = (equipment_invest * st.session_state.project_data.equipment_depreciation +
                                construction_invest * st.session_state.project_data.construction_depreciation)

                ebit = revenue - costs - depreciation
                tax = max(0, ebit * st.session_state.project_data.tax_rate)
                net_income = ebit - tax

                operating_cf = net_income + depreciation
                investment = total_investment if year == st.session_state.project_data.years[0] else 0

                cash_flows[year] = operating_cf - investment

            st.session_state.cash_flows = cash_flows

            cf_df = pd.DataFrame(
                [{'Год': y, 'Денежный поток': cash_flows[y]} for y in st.session_state.project_data.years])
            st.dataframe(cf_df.style.format({'Денежный поток': '{:,.0f} руб.'}))

            npv = calculate_npv(cash_flows, st.session_state.project_data.discount_rate)
            irr = calculate_irr(cash_flows)
            payback = calculate_payback_period(cash_flows, total_investment)

            st.session_state.npv = npv
            st.session_state.irr = irr
            st.session_state.payback = payback

            st.subheader("📈 Ключевые показатели эффективности")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("NPV", f"{npv:,.0f} руб.",
                          delta="Положительный" if npv > 0 else "Отрицательный")
            with col2:
                st.metric("IRR", f"{irr * 100:.1f}%",
                          delta=f"{irr - st.session_state.project_data.discount_rate:.1%}")
            with col3:
                st.metric("Срок окупаемости", f"{payback:.1f} лет" if payback != float('inf') else "Не окупается")
            with col4:
                pi = (npv + total_investment) / total_investment if total_investment > 0 else 0
                st.metric("PI (Индекс доходности)", f"{pi:.2f}")

            cumulative_npv = 0
            npv_by_year = []
            for i, year in enumerate(sorted(cash_flows.keys())):
                cumulative_npv += cash_flows[year] / ((1 + st.session_state.project_data.discount_rate) ** i)
                npv_by_year.append(cumulative_npv)

            fig_npv = go.Figure()
            fig_npv.add_trace(go.Scatter(x=list(cash_flows.keys()), y=npv_by_year,
                                         mode='lines+markers', name='Накопленный NPV'))
            fig_npv.add_hline(y=0, line_dash="dash", line_color="red")
            fig_npv.update_layout(title="Динамика NPV", xaxis_title="Год", yaxis_title="NPV, руб.")
            st.plotly_chart(fig_npv, use_container_width=True)

        else:
            st.warning("Сначала заполните данные о продукции и выручке")

    with tab_export:
        st.header("Экспорт данных")

        if st.button("📥 Подготовить Excel отчет"):
            if st.session_state.revenue_data and st.session_state.cash_flows:
                excel_data = create_excel_report(
                    st.session_state.project_data,
                    st.session_state.revenue_data,
                    st.session_state.cash_flows,
                    st.session_state.npv,
                    st.session_state.irr,
                    st.session_state.payback
                )
                st.session_state.excel_data = excel_data
                st.success("Отчет подготовлен!")
            else:
                st.warning("Сначала выполните расчеты")

        if st.session_state.get('excel_data'):
            st.download_button(
                label="💾 Скачать отчет Excel",
                data=st.session_state.excel_data,
                file_name="investment_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        st.divider()

        st.subheader("Загрузка сохраненного проекта")
        uploaded_file = st.file_uploader("Загрузить JSON файл", type=['json'])
        if uploaded_file:
            try:
                data = json.load(uploaded_file)
                st.session_state.project_data = ProjectData.from_dict(data)
                st.success("✅ Проект загружен!")
            except Exception as e:
                st.error(f"Ошибка загрузки: {e}")


if __name__ == "__main__":
    main()
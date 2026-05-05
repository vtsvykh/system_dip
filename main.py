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
            for key in ['balance_file', 'income_file', 'cash_flow_file']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()

    # Вкладки
    tab_input, tab_revenue, tab_results, tab_export = st.tabs([
        "📝 Ввод данных", "💰 Выручка и продажи", "📈 Анализ эффективности", "💾 Экспорт"
    ])

    with tab_input:
        # ========== БЛОК 1: СТАВКИ ==========
        st.header("Финансовые параметры")

        col1, col2 = st.columns(2)
        with col1:
            st.session_state.project_data.discount_rate = st.slider(
                "Ставка дисконтирования (%)",
                min_value=0.0,
                max_value=30.0,
                value=10.0,
                step=0.1
            ) / 100

        with col2:
            st.session_state.project_data.tax_rate = st.slider(
                "Ставка налога на прибыль (%)",
                min_value=0.0,
                max_value=30.0,
                value=20.0,
                step=0.1
            ) / 100

        # ========== БЛОК 2: РАСПРЕДЕЛЕНИЕ ИНВЕСТИЦИЙ И АМОРТИЗАЦИЯ ==========
        st.subheader("Распределение инвестиций и амортизация")

        col_left, col_center, col_right = st.columns([1, 1, 1])

        with col_left:
            # Доля оборудования (вводит пользователь)
            equipment_share_pct = st.number_input(
                "Доля оборудования (%)",
                min_value=0.0,
                max_value=100.0,
                value=60.0,
                step=1.0
            )
            st.session_state.project_data.equipment_share = equipment_share_pct / 100

            # Норма амортизации оборудования
            equipment_dep_pct = st.slider(
                "Норма амортизации оборудования (%)",
                min_value=0.0,
                max_value=50.0,
                value=15.0,
                step=0.1
            )
            st.session_state.project_data.equipment_depreciation = equipment_dep_pct / 100

        with col_center:
            # Круговая диаграмма
            construction_share = 100 - equipment_share_pct
            fig_pie = go.Figure(data=[go.Pie(
                labels=['Оборудование', 'Строительно-монтажные работы'],
                values=[equipment_share_pct, construction_share],
                hole=0.4,
                marker_colors=['#1f77b4', '#ff7f0e']
            )])
            fig_pie.update_layout(title="Распределение инвестиций", height=300)
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_right:
            # Доля СМР (рассчитывается автоматически)
            st.text_input(
                "Доля строительно-монтажных работ (%)",
                value=f"{construction_share:.1f}",
                disabled=True
            )

            # Норма амортизации СМР
            construction_dep_pct = st.slider(
                "Норма амортизации СМР (%)",
                min_value=0.0,
                max_value=50.0,
                value=5.0,
                step=0.1
            )
            st.session_state.project_data.construction_depreciation = construction_dep_pct / 100

        # ========== БЛОК 3: ДИАПАЗОН ЛЕТ ==========
        col1, col2 = st.columns(2)
        with col1:
            start_year = st.number_input("Начальный год", min_value=2000, max_value=2100, value=2024, step=1)
        with col2:
            end_year = st.number_input("Конечный год", min_value=2000, max_value=2100, value=2030, step=1)

        st.session_state.project_data.years = list(range(start_year, end_year + 1))

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

                if need_update:
                    st.rerun()

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

    with tab_revenue:
        st.header("Расчет выручки")

        if st.session_state.project_data.products and st.session_state.project_data.prices:
            revenue_data = calculate_revenue(
                st.session_state.project_data.products,
                st.session_state.project_data.prices,
                st.session_state.project_data.export_shares,
                st.session_state.project_data.import_shares,
                st.session_state.project_data.years
            )

            st.session_state.revenue_data = revenue_data

            # Таблица выручки
            st.subheader("📊 Выручка по годам")
            revenue_table = []
            for year in st.session_state.project_data.years:
                revenue_table.append({
                    'Год': year,
                    'Экспорт (без НДС)': revenue_data['export_revenue'][year],
                    'Внутренний рынок (с НДС)': revenue_data['domestic_revenue'][year],
                    'Общая выручка': revenue_data['total_revenue'][year]
                })

            df_revenue = pd.DataFrame(revenue_table)
            st.dataframe(df_revenue.style.format({
                'Экспорт (без НДС)': '{:,.0f}',
                'Внутренний рынок (с НДС)': '{:,.0f}',
                'Общая выручка': '{:,.0f}'
            }))

            # График
            fig = go.Figure()
            fig.add_trace(go.Bar(name='Экспорт', x=st.session_state.project_data.years,
                                 y=[revenue_data['export_revenue'][y] for y in st.session_state.project_data.years]))
            fig.add_trace(go.Bar(name='Внутренний рынок', x=st.session_state.project_data.years,
                                 y=[revenue_data['domestic_revenue'][y] for y in st.session_state.project_data.years]))
            fig.update_layout(title="Динамика выручки", barmode='stack', xaxis_title="Год", yaxis_title="Руб.")
            st.plotly_chart(fig, use_container_width=True)

        else:
            st.warning("Сначала введите данные о продукции и ценах во вкладке 'Ввод данных'")

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
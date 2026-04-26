import requests
import pandas as pd


def get_cpi_from_fedstat(region_name, years_list):
    """
    Загружает индексы потребительских цен с fedstat.ru

    Args:
        region_name: название региона (например, "Амурская область")
        years_list: список лет (например, [2021, 2022, 2023, 2024])

    Returns:
        словарь {год: индекс ИПЦ (в % к декабрю предыдущего года)}
    """
    # ID индикатора для ИПЦ (на конец периода, в % к декабрю предыдущего года)
    # Это можно получить, изучив URL на fedstat.ru
    indicator_id = "31074"  # Индексы потребительских цен на товары и услуги

    # Базовый URL для API fedstat (нужно уточнить)
    base_url = "https://www.fedstat.ru/api/v1/data"

    params = {
        "indicator_id": indicator_id,
        "territory": region_name,  # или код ОКАТО
        "filter_years": years_list
    }

    try:
        response = requests.get(base_url, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        # Парсим ответ в словарь (структура зависит от API fedstat)
        cpi_data = {}
        for item in data.get("result", []):
            year = item.get("year")
            value = item.get("value")
            if year in years_list:
                cpi_data[year] = float(value)

        return cpi_data
    except Exception as e:
        st.error(f"Ошибка загрузки данных с fedstat.ru: {e}")
        return {}
# Справочник регионов (ОКАТО коды)
regions_okato = {
    "Амурская область": "64310000000",
    "Москва": "64345000000",
    "Санкт-Петербург": "64340000000",
    "Московская область": "64346000000",
    "Краснодарский край": "64303000000",
    # ... остальные регионы
}
# В блоке "Выбор региона и ИПЦ"
import requests

st.subheader("📊 Автозагрузка индексов потребительских цен (ИПЦ)")

col_reg, col_load = st.columns([3, 1])

with col_reg:
    # Справочник регионов (можно расширить)
    regions = {
        "Амурская область": "64310000000",
        "Москва": "64345000000",
        "Санкт-Петербург": "64340000000",
        # ... добавьте другие регионы
    }
    selected_region = st.selectbox("Выберите регион", list(regions.keys()))

with col_load:
    if st.button("📥 Загрузить ИПЦ с fedstat.ru"):
        with st.spinner("Загружаем данные с fedstat.ru..."):
            # Получаем годы из project_data
            years = st.session_state.project_data.years

            # Загружаем ИПЦ
            cpi_data = get_cpi_from_fedstat(selected_region, years)

            if cpi_data:
                # Автоматически заполняем таблицу ценовых индексов
                for year, cpi_value in cpi_data.items():
                    # CPI обычно в процентах, переводим в индекс (например, 105.5% -> 1.055)
                    index = cpi_value / 100

                    # Сохраняем в session_state
                    if 'saved_indices' not in st.session_state:
                        st.session_state.saved_indices = {}
                    st.session_state.saved_indices[year] = index

                    # Рассчитываем цену
                    if 'base_price' in locals():
                        st.session_state.project_data.prices[year] = base_price * index

                st.success(f"✅ Загружены ИПЦ для {selected_region} за {len(cpi_data)} лет")
                st.rerun()  # Обновляем страницу для отображения данных
            else:
                st.error("Не удалось загрузить данные. Попробуйте позже.")
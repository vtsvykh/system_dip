#
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import json
import os
import numpy as np
from io import BytesIO

# ========== УПРОЩЕННЫЕ МОДЕЛИ ДАННЫХ ==========
class ProjectData:
    """Хранилище всех данных проекта"""

    def __init__(self):
        self.years = []
        self.volumes = {}
        self.prices = {}
        self.discount_rate = 0.10
        self.tax_rate = 0.20
        self.equipment_share = 0.60
        self.equipment_depreciation = 0.15
        self.construction_depreciation = 0.05
        self.products = {}  # {product_type: {year: volume}}
        self.export_shares = {}
        self.import_shares = {}
        self.sector = ""  # Выбранная отрасль
        self.okpd_code = ""  # Выбранный код ОКПД2
        self.price_indices = {}

    def to_dict(self):
        return {
            'years': self.years,
            'volumes': self.volumes,
            'prices': self.prices,
            'discount_rate': self.discount_rate,
            'tax_rate': self.tax_rate,
            'equipment_share': self.equipment_share,
            'equipment_depreciation': self.equipment_depreciation,
            'construction_depreciation': self.construction_depreciation,
            'products': self.products,
            'export_shares': self.export_shares,
            'import_shares': self.import_shares,
            'sector': self.sector,
            'okpd_code': self.okpd_code,
            'price_indices': self.price_indices
        }

    @classmethod
    def from_dict(cls, data):
        obj = cls()
        obj.__dict__.update(data)
        return obj


# ========== ФИНАНСОВЫЕ РАСЧЕТЫ ==========
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
        'domestic_volume': {year: 0 for year in years}
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


def create_excel_report(project_data, revenue_data, cash_flows, npv, irr, payback):
    """Создание Excel отчета"""
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        params = pd.DataFrame([
            ['Ставка дисконтирования', f"{project_data.discount_rate * 100:.1f}%"],
            ['Ставка налога', f"{project_data.tax_rate * 100:.1f}%"],
            ['Доля оборудования', f"{project_data.equipment_share * 100:.1f}%"],
            ['Амортизация оборудования', f"{project_data.equipment_depreciation * 100:.1f}%"],
            ['Амортизация СМР', f"{project_data.construction_depreciation * 100:.1f}%"],
            ['Отрасль', project_data.sector],
            ['Код ОКПД2', project_data.okpd_code]
        ], columns=['Параметр', 'Значение'])
        params.to_excel(writer, sheet_name='Параметры', index=False)

        if revenue_data:
            revenue_df = pd.DataFrame([{
                'Год': year,
                'Экспорт': revenue_data['export_revenue'][year],
                'Внутренний рынок': revenue_data['domestic_revenue'][year],
                'Всего': revenue_data['total_revenue'][year]
            } for year in project_data.years])
            revenue_df.to_excel(writer, sheet_name='Выручка', index=False)

        if cash_flows:
            cf_df = pd.DataFrame([{'Год': year, 'Денежный поток': cash_flows[year]}
                                  for year in sorted(cash_flows.keys())])
            cf_df.to_excel(writer, sheet_name='Денежные потоки', index=False)

        metrics = pd.DataFrame([
            ['NPV', f"{npv:,.0f} руб."],
            ['IRR', f"{irr * 100:.1f}%"],
            ['Срок окупаемости', f"{payback:.1f} лет" if payback != float('inf') else "Не окупается"]
        ], columns=['Показатель', 'Значение'])
        metrics.to_excel(writer, sheet_name='Итоговые показатели', index=False)

    return output.getvalue()


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

    # ========== САЙДБАР ДЛЯ ЗАГРУЗКИ ФАЙЛОВ ==========
    with st.sidebar:
        st.header("📁 Загрузка отчетности")

        from datetime import datetime
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

        # УНИВЕРСАЛЬНЫЙ справочник ОКПД2 для всех отраслей
        okpd_catalog = {
            "Продукция сельского, лесного и рыбного хозяйства": {
                "01.11": "Культуры зерновые (кроме риса), зернобобовые, семена масличных культур",
                "01.12": "Выращивание овощей",
                "01.13": "Выращивание фруктов и ягод",
                "01.21": "Разведение крупного рогатого скота",
                "01.22": "Разведение свиней",
                "01.23": "Разведение птицы",
                "01.24": "Разведение овец и коз"
            },
            "ДОБЫЧА ПОЛЕЗНЫХ ИСКОПАЕМЫХ": {
                "05.10": "Добыча угля",
                "06.10": "Добыча нефти",
                "06.20": "Добыча природного газа",
                "07.10": "Добыча железных руд",
                "07.21": "Добыча руд цветных металлов",
                "08.11": "Добыча строительного камня"
            },
            "ОБРАБАТЫВАЮЩИЕ ПРОИЗВОДСТВА": {
                "10.11": "Производство мяса",
                "10.12": "Производство молочных продуктов",
                "10.13": "Производство продуктов мукомольной промышленности",
                "10.51": "Производство молочных продуктов",
                "10.71": "Производство хлеба и мучных кондитерских изделий",
                "11.01": "Производство алкогольных напитков",
                "13.10": "Производство текстильных изделий",
                "13.20": "Производство одежды",
                "14.11": "Производство кожаной одежды",
                "16.10": "Производство пиломатериалов",
                "17.11": "Производство целлюлозы",
                "17.12": "Производство бумаги и картона",
                "20.11": "Производство промышленных газов",
                "20.12": "Производство красителей и пигментов",
                "20.13": "Производство основных органических химических веществ",
                "20.14": "Производство основных неорганических химических веществ",
                "20.15": "Производство удобрений и азотных соединений",
                "20.16": "Производство пластмасс в первичных формах",
                "20.17": "Производство синтетического каучука",
                "21.10": "Производство фармацевтической продукции",
                "22.11": "Производство шин и покрышек",
                "22.19": "Производство прочих резиновых изделий",
                "23.11": "Производство листового стекла",
                "23.12": "Производство стеклянной тары",
                "23.13": "Производство стекловолокна",
                "23.14": "Производство бытовых стеклянных изделий",
                "23.31": "Производство керамических плиток",
                "23.32": "Производство кирпича и черепицы",
                "23.41": "Производство цемента",
                "23.42": "Производство извести и гипса",
                "23.43": "Производство готовых бетонных изделий",
                "24.10": "Производство чугуна, стали и ферросплавов",
                "24.20": "Производство стальных труб",
                "24.31": "Производство стального проката",
                "24.32": "Производство стальных профилей",
                "24.33": "Производство стальной арматуры",
                "24.34": "Производство стальных проволок и канатов",
                "24.41": "Производство благородных металлов",
                "24.42": "Производство алюминия",
                "24.43": "Производство свинца, цинка и олова",
                "24.44": "Производство меди",
                "24.45": "Производство прочих цветных металлов",
                "25.11": "Производство металлических конструкций",
                "25.12": "Производство строительных металлических изделий",
                "25.21": "Производство радиаторов и котлов",
                "25.29": "Производство прочих металлических емкостей",
                "25.30": "Производство паровых котлов",
                "26.11": "Производство электронных компонентов",
                "26.20": "Производство компьютеров",
                "26.30": "Производство коммуникационного оборудования",
                "26.40": "Производство бытовой электроники",
                "27.11": "Производство электродвигателей",
                "27.12": "Производство трансформаторов",
                "27.13": "Производство распределительных устройств",
                "27.20": "Производство батарей и аккумуляторов",
                "27.31": "Производство кабелей и проводов",
                "27.32": "Производство электроосветительного оборудования",
                "28.11": "Производство двигателей и турбин",
                "28.12": "Производство насосов и компрессоров",
                "28.13": "Производство арматуры и кранов",
                "28.14": "Производство подшипников",
                "28.15": "Производство редукторов и зубчатых передач",
                "28.21": "Производство печей и горелок",
                "28.22": "Производство подъемно-транспортного оборудования",
                "28.23": "Производство офисной техники",
                "28.24": "Производство ручных инструментов",
                "28.25": "Производство промышленного холодильного оборудования",
                "29.10": "Производство автомобилей",
                "29.20": "Производство кузовов для автомобилей",
                "29.31": "Производство электрического и электронного оборудования для автомобилей",
                "30.11": "Производство судов",
                "30.12": "Производство прогулочных и спортивных судов",
                "30.20": "Производство железнодорожного подвижного состава",
                "30.30": "Производство летательных аппаратов",
                "30.40": "Производство военной боевой техники",
                "30.91": "Производство мотоциклов",
                "30.92": "Производство велосипедов",
                "31.01": "Производство мебели для офисов",
                "31.02": "Производство кухонной мебели",
                "31.03": "Производство матрасов",
                "32.11": "Производство монет",
                "32.12": "Производство ювелирных изделий",
                "32.13": "Производство бижутерии"
            },
            "ЭНЕРГЕТИКА": {
                "35.11": "Производство электроэнергии",
                "35.12": "Передача электроэнергии",
                "35.13": "Распределение электроэнергии",
                "35.14": "Торговля электроэнергией",
                "35.21": "Производство газообразного топлива",
                "35.22": "Распределение газообразного топлива",
                "35.23": "Торговля газообразным топливом"
            },
            "СТРОИТЕЛЬСТВО": {
                "41.10": "Разработка строительных проектов",
                "41.20": "Строительство жилых зданий",
                "42.11": "Строительство автомобильных дорог",
                "42.12": "Строительство железных дорог",
                "42.13": "Строительство мостов и тоннелей",
                "43.11": "Разборка и снос зданий",
                "43.12": "Подготовка строительной площадки"
            },
            "ТОРГОВЛЯ": {
                "45.11": "Торговля легковыми автомобилями",
                "45.19": "Торговля прочими автотранспортными средствами",
                "46.11": "Торговля через агентов",
                "47.11": "Торговля в неспециализированных магазинах"
            },
            "ТРАНСПОРТ": {
                "49.10": "Деятельность железнодорожного транспорта",
                "49.20": "Деятельность грузового автомобильного транспорта",
                "49.31": "Деятельность городского пассажирского транспорта",
                "50.10": "Деятельность морского пассажирского транспорта",
                "51.10": "Деятельность воздушного пассажирского транспорта"
            },
            "УСЛУГИ": {
                "55.10": "Деятельность гостиниц",
                "56.10": "Деятельность ресторанов",
                "58.11": "Издание книг",
                "58.13": "Издание газет",
                "59.11": "Производство кинофильмов",
                "60.10": "Деятельность в области радиовещания",
                "61.10": "Деятельность в области связи"
            }
        }

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


        # УНИВЕРСАЛЬНЫЙ справочник типов продукции для всех ОКПД
        product_types_mapping = {
            "01.11": ["Пшеница", "Ячмень, рожь и овес", "Овес", "Кукуруза", "Сорго, просо и прочие зерновые культуры"],
            "01.12": ["Картофель", "Морковь", "Капуста", "Лук", "Томаты", "Огурцы"],
            "01.13": ["Яблоки", "Груши", "Вишня", "Слива", "Смородина", "Малина"],
            "01.21": ["Говядина", "Молоко", "Мясо КРС"],
            "01.22": ["Свинина", "Мясо свиней"],
            "01.23": ["Куриное мясо", "Яйца", "Мясо индейки"],
            "05.10": ["Каменный уголь", "Бурый уголь"],
            "06.10": ["Сырая нефть", "Газовый конденсат"],
            "06.20": ["Природный газ"],
            "07.10": ["Железная руда", "Железорудный концентрат"],
            "24.10": ["Чугун передельный", "Чугун литейный", "Сталь углеродистая", "Сталь легированная"],
            "24.20": ["Трубы бесшовные", "Трубы сварные", "Трубы профильные"],
            "24.31": ["Прокат листовой", "Прокат сортовой", "Прокат фасонный"],
            "24.41": ["Золото", "Серебро", "Платина"],
            "24.42": ["Алюминий первичный", "Алюминиевые сплавы"],
            "20.13": ["Этилен", "Пропилен", "Бензол", "Метанол"],
            "20.15": ["Азотные удобрения", "Фосфорные удобрения", "Калийные удобрения"],
            "20.16": ["Полиэтилен", "Полипропилен", "ПВХ", "Полистирол"],
            "28.11": ["Дизельные двигатели", "Газовые турбины", "Паровые турбины"],
            "29.10": ["Легковые автомобили", "Грузовые автомобили", "Автобусы"],
            "10.11": ["Говядина", "Свинина", "Баранина", "Мясо птицы"],
            "10.12": ["Молоко", "Сыр", "Йогурт", "Сливочное масло", "Творог"],
            "10.71": ["Хлеб", "Булочки", "Пироги", "Печенье"],
            "16.10": ["Пиломатериалы хвойные", "Пиломатериалы лиственные", "Древесные плиты"],
            "35.11": ["Электроэнергия"],
            "35.21": ["Природный газ", "Сжиженный газ"],
            "default": ["Основная продукция", "Побочная продукция", "Полуфабрикаты"]
        }

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

        default_units = {
            "Продукция сельского, лесного и рыбного хозяйства": "тонн",
            "ДОБЫЧА ПОЛЕЗНЫХ ИСКОПАЕМЫХ": "тонн",
            "ОБРАБАТЫВАЮЩИЕ ПРОИЗВОДСТВА": "тонн",
            "ЭНЕРГЕТИКА": "кВт·ч",
            "СТРОИТЕЛЬСТВО": "м²",
            "ТОРГОВЛЯ": "шт",
            "ТРАНСПОРТ": "пассажиро-км",
            "УСЛУГИ": "услуг"
        }

        all_units = ["млн.тонн", "тонн", "кг", "литров", "м³", "м²", "кВт·ч", "шт", "пар", "комплектов",
                     "пассажиро-км", "услуг", "ед.", "тыс. руб.", "млн. руб."]

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
            all_units = ["млн.тонн", "тонн", "кг", "литров", "м³", "м²", "кВт·ч", "шт", "пар", "комплектов",
                         "пассажиро-км", "услуг", "ед.", "тыс. руб.", "млн. руб."]

            # Ключ для хранения единицы измерения продукта
            unit_key = f"unit_{selected_okpd}_{product_type}"

            # Инициализация единицы измерения для продукта
            if unit_key not in st.session_state:
                # Определяем единицу по умолчанию в зависимости от отрасли
                default_units = {
                    "Продукция сельского, лесного и рыбного хозяйства": "тонн",
                    "ДОБЫЧА ПОЛЕЗНЫХ ИСКОПАЕМЫХ": "тонн",
                    "ОБРАБАТЫВАЮЩИЕ ПРОИЗВОДСТВА": "тонн",
                    "ЭНЕРГЕТИКА": "кВт·ч",
                    "СТРОИТЕЛЬСТВО": "м²",
                    "ТОРГОВЛЯ": "шт",
                    "ТРАНСПОРТ": "пассажиро-км",
                    "УСЛУГИ": "услуг"
                }
                st.session_state[unit_key] = default_units.get(selected_sector, "ед.")

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

            with col_vol:
                # Таблица объемов с выбранной единицей измерения
                st.write(f"Объемы производства")

                # Создаем данные для таблицы
                volumes_data = []
                for year in years_list:
                    volume_key = f"volume_{selected_okpd}_{product_type}_{year}"
                    export_key = f"export_{selected_okpd}_{product_type}_{year}"

                    default_volume = st.session_state.get(volume_key, 0.0)
                    default_export = st.session_state.get(export_key, 0.0)

                    volumes_data.append({
                        'Год': year,
                        f'Объем ({product_unit})': default_volume,
                        'Доля экспорта (%)': default_export,
                        'Доля импорта (%)': 100.0 - default_export
                    })

                df_volumes = pd.DataFrame(volumes_data)

                # Редактируемая таблица
                edited_df = st.data_editor(
                    df_volumes,
                    key=f"editor_{selected_okpd}_{product_type}",
                    num_rows="fixed",
                    column_config={
                        'Год': st.column_config.NumberColumn(disabled=True),
                        f'Объем ({product_unit})': st.column_config.NumberColumn(
                            format="%.3f",
                            min_value=0.0,
                            step=0.001
                        ),
                        'Доля экспорта (%)': st.column_config.NumberColumn(
                            format="%.1f",
                            min_value=0.0,
                            max_value=100.0,
                            step=0.1
                        ),
                        'Доля импорта (%)': st.column_config.NumberColumn(
                            disabled=True,
                            format="%.1f"
                        )
                    }
                )

                # Обрабатываем данные
                product_volumes[product_type] = {}
                product_export_shares[product_type] = {}
                product_import_shares[product_type] = {}

                for _, row in edited_df.iterrows():
                    year = int(row['Год'])
                    volume = float(row[f'Объем ({product_unit})'])
                    export_pct = float(row['Доля экспорта (%)'])
                    export_pct = max(0.0, min(100.0, export_pct))
                    import_pct = 100.0 - export_pct

                    product_volumes[product_type][year] = volume
                    product_export_shares[product_type][year] = export_pct / 100
                    product_import_shares[product_type][year] = import_pct / 100

                    st.session_state[f"volume_{selected_okpd}_{product_type}_{year}"] = volume
                    st.session_state[f"export_{selected_okpd}_{product_type}_{year}"] = export_pct

            # Показываем итоги с выбранной единицей измерения
            total_volume = sum(product_volumes[product_type].values())
            avg_export = sum(product_export_shares[product_type].values()) / len(years_list) * 100 if years_list else 0

            col1, col2 = st.columns(2)
            with col1:
                st.metric(f"📦 Общий объем", f"{total_volume:,.3f} {product_unit}")
            with col2:
                st.metric(f"📤 Средний экспорт", f"{avg_export:.1f}%")

            st.divider()

        # Сохраняем в session_state
        st.session_state.product_volumes = product_volumes
        st.session_state.product_export_shares = product_export_shares
        st.session_state.product_import_shares = product_import_shares
        st.session_state.product_base_prices = product_base_prices

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
            if st.session_state.price_indices_temp and hasattr(st.session_state, 'product_base_prices'):
                st.subheader("📊 Расчет цен по продуктам и годам")

                # Функция расчета дефлятора
                def calculate_deflator(year, base_year, price_indices):
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

                # Получаем текущие индексы и рассчитываем дефляторы
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
                            'Цена с учетом дефлятора (руб.)': calculated_price
                        })

                # Создаем и отображаем сводную таблицу
                df_prices = pd.DataFrame(all_prices_data)

                # Сводная таблица для лучшего отображения
                pivot_prices = df_prices.pivot_table(
                    index='Продукт',
                    columns='Год',
                    values='Цена с учетом дефлятора (руб.)',
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
                            'Цена с учетом дефлятора (руб.)': '{:,.2f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                # Информация о расчете
                st.info(f"""
                **Формула расчета:** Цена продукта в году Y = Базовая цена продукта × Дефлятор года Y

                **Дефлятор для {st.session_state.base_year} года = 1.00** (базовый год)
                - Для годов > {st.session_state.base_year}: цены растут (дефлятор > 1)
                - Для годов < {st.session_state.base_year}: цены снижаются (дефлятор < 1)
                """)
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
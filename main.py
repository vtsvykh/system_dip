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
        # Таблица для ввода объемов
        # НОВЫЙ БЛОК: Выбор ОКПД и объемов продукции
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

            # Показываем описание выбранного ОКПД
            if selected_okpd in sector_codes:
                st.info(f"**Выбрано:** {sector_codes[selected_okpd]}")

        # УНИВЕРСАЛЬНЫЙ справочник типов продукции для всех ОКПД
        product_types_mapping = {
            # Сельское хозяйство
            "01.11": ["Пшеница", "Ячмень, рожь и овес", "Овес", "Кукуруза", "Сорго, просо и прочие зерновые культуры"],
            "01.12": ["Картофель", "Морковь", "Капуста", "Лук", "Томаты", "Огурцы"],
            "01.13": ["Яблоки", "Груши", "Вишня", "Слива", "Смородина", "Малина"],
            "01.21": ["Говядина", "Молоко", "Мясо КРС"],
            "01.22": ["Свинина", "Мясо свиней"],
            "01.23": ["Куриное мясо", "Яйца", "Мясо индейки"],

            # Добыча полезных ископаемых
            "05.10": ["Каменный уголь", "Бурый уголь"],
            "06.10": ["Сырая нефть", "Газовый конденсат"],
            "06.20": ["Природный газ"],
            "07.10": ["Железная руда", "Железорудный концентрат"],

            # Металлургия
            "24.10": ["Чугун передельный", "Чугун литейный", "Сталь углеродистая", "Сталь легированная"],
            "24.20": ["Трубы бесшовные", "Трубы сварные", "Трубы профильные"],
            "24.31": ["Прокат листовой", "Прокат сортовой", "Прокат фасонный"],
            "24.41": ["Золото", "Серебро", "Платина"],
            "24.42": ["Алюминий первичный", "Алюминиевые сплавы"],

            # Химическая промышленность
            "20.13": ["Этилен", "Пропилен", "Бензол", "Метанол"],
            "20.15": ["Азотные удобрения", "Фосфорные удобрения", "Калийные удобрения"],
            "20.16": ["Полиэтилен", "Полипропилен", "ПВХ", "Полистирол"],

            # Машиностроение
            "28.11": ["Дизельные двигатели", "Газовые турбины", "Паровые турбины"],
            "29.10": ["Легковые автомобили", "Грузовые автомобили", "Автобусы"],

            # Пищевая промышленность
            "10.11": ["Говядина", "Свинина", "Баранина", "Мясо птицы"],
            "10.12": ["Молоко", "Сыр", "Йогурт", "Сливочное масло", "Творог"],
            "10.71": ["Хлеб", "Булочки", "Пироги", "Печенье"],

            # Деревообработка
            "16.10": ["Пиломатериалы хвойные", "Пиломатериалы лиственные", "Древесные плиты"],

            # Энергетика
            "35.11": ["Электроэнергия"],
            "35.21": ["Природный газ", "Сжиженный газ"],

            # Универсальные типы для кодов без специфики
            "default": ["Основная продукция", "Побочная продукция", "Полуфабрикаты"]
        }

        # Получаем типы продукции для выбранного ОКПД
        current_product_types = product_types_mapping.get(selected_okpd, product_types_mapping["default"])

        # Возможность добавить свой тип продукции
        add_custom_product = st.checkbox("Добавить свой тип продукции")
        if add_custom_product:
            custom_product = st.text_input("Название типа продукции:")
            if custom_product and custom_product not in current_product_types:
                current_product_types.append(custom_product)

        # Создаем таблицы для каждого типа продукции
        st.subheader("Ввод объемов производства по типам продукции")

        # ПОЛУЧАЕМ ГОДЫ ИЗ SESSION_STATE
        years_list = st.session_state.project_data.years
        if not years_list:
            st.warning("Сначала укажите диапазон лет в блоке 'Диапазон лет'")
            years_list = []

        # Определяем единицы измерения в зависимости от отрасли
        measurement_units = {
            "Продукция сельского, лесного и рыбного хозяйства": "тонн",
            "ДОБЫЧА ПОЛЕЗНЫХ ИСКОПАЕМЫХ": "тонн",
            "ОБРАБАТЫВАЮЩИЕ ПРОИЗВОДСТВА": "тонн",
            "ЭНЕРГЕТИКА": "кВт·ч",
            "СТРОИТЕЛЬСТВО": "м²",
            "ТОРГОВЛЯ": "шт",
            "ТРАНСПОРТ": "пассажиро-км",
            "УСЛУГИ": "услуг"
        }

        unit = measurement_units.get(selected_sector, "ед.")

        # Создаем вкладки для каждого типа продукции
        if len(current_product_types) > 1:
            product_tabs = st.tabs(current_product_types)
        else:
            product_tabs = [st.container()]

        # Словарь для хранения объемов по типам продукции
        product_volumes = {}

        product_export_shares = {}
        product_import_shares = {}

        for i, product_type in enumerate(current_product_types):
            with product_tabs[i]:
                st.write(f"**{product_type}**")

                # Создаем DataFrame
                volumes_data = []
                for year in years_list:  # ИСПРАВЛЕНО: years_list вместо years_range
                    volumes_data.append({
                        'Год': year,
                        f'Объем ({unit})': 0.0,
                        'Доля экспорта (%)': 0.0,
                        'Доля импорта (%)': 100.0  # Начальное значение
                    })

                df = pd.DataFrame(volumes_data)

                # Редактируемая таблица
                edited_df = st.data_editor(
                    df,
                    key=f"editor_{selected_okpd}_{product_type}_{i}",
                    num_rows="fixed",
                    column_config={
                        'Год': st.column_config.NumberColumn(disabled=True),
                        f'Объем ({unit})': st.column_config.NumberColumn(
                            format="%.0f",
                            min_value=0.0
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

                # АВТОМАТИЧЕСКИ ПЕРЕСЧИТЫВАЕМ ИМПОРТ ПОСЛЕ РЕДАКТИРОВАНИЯ
                final_df = edited_df.copy()
                for idx in range(len(final_df)):
                    export_value = float(final_df.at[idx, 'Доля экспорта (%)'])
                    # Ограничиваем значение от 0 до 100
                    export_value = max(0.0, min(100.0, export_value))
                    import_value = 100.0 - export_value

                    # Обновляем значение в DataFrame
                    final_df.at[idx, 'Доля импорта (%)'] = import_value
                    final_df.at[idx, 'Доля экспорта (%)'] = export_value

                # Показываем таблицу с уже обновленными значениями
                st.dataframe(
                    final_df.style.format({
                        f'Объем ({unit})': '{:.0f}',
                        'Доля экспорта (%)': '{:.1f}',
                        'Доля импорта (%)': '{:.1f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # Сохраняем данные для расчетов ИЗ ОБНОВЛЕННОГО DataFrame
                product_volumes[product_type] = {}
                product_export_shares[product_type] = {}
                product_import_shares[product_type] = {}

                for _, row in final_df.iterrows():
                    year = int(row['Год'])
                    product_volumes[product_type][year] = float(row[f'Объем ({unit})'])
                    product_export_shares[product_type][year] = float(row['Доля экспорта (%)']) / 100
                    product_import_shares[product_type][year] = float(row['Доля импорта (%)']) / 100

                # Показываем итого по продукту с проверкой
                total_volume = sum(product_volumes[product_type].values())
                avg_export_share = sum(product_export_shares[product_type].values()) / len(
                    years_list) * 100 if years_list else 0  # ИСПРАВЛЕНО
                avg_import_share = sum(product_import_shares[product_type].values()) / len(
                    years_list) * 100 if years_list else 0  # ИСПРАВЛЕНО

                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric(f"Общий объем", f"{total_volume:,.0f} {unit}")
                with col2:
                    st.metric(f"Средняя доля экспорта", f"{avg_export_share:.1f}%")
                with col3:
                    st.metric(f"Средняя доля импорта", f"{avg_import_share:.1f}%")

                # Проверка: сумма долей должна быть 100%
                for year in years_list:  # ИСПРАВЛЕНО
                    export_share = product_export_shares[product_type].get(year, 0)
                    import_share = product_import_shares[product_type].get(year, 0)
                    if abs((export_share + import_share) - 1.0) > 0.001:  # Допуск 0.1%
                        st.error(
                            f"Ошибка в {year}: экспорт ({export_share * 100:.1f}%) + импорт ({import_share * 100:.1f}%) ≠ 100%")

        # Сохраняем в session_state для использования в других вкладках
        st.session_state.product_volumes = product_volumes
        st.session_state.product_export_shares = product_export_shares
        st.session_state.product_import_shares = product_import_shares



        # Общие объемы по всем типам продукции
        if product_volumes:
            with st.expander("📊 Сводка по объемам производства и экспорту", expanded=False):
                st.subheader("📊 Сводка по объемам производства и экспорту")
                # Создаем сводную таблицу
                summary_data = []
                for year in years_list:  # ИСПРАВЛЕНО
                    year_data = {'Год': year}
                    year_total_volume = 0
                    year_total_export_volume = 0
                    year_total_import_volume = 0

                    for product_type in current_product_types:
                        volume = product_volumes.get(product_type, {}).get(year, 0)
                        export_share = product_export_shares.get(product_type, {}).get(year, 0)
                        import_share = product_import_shares.get(product_type, {}).get(year, 0)

                        year_data[f'{product_type} - Объем'] = volume
                        year_data[f'{product_type} - Экспорт'] = volume * export_share
                        year_data[f'{product_type} - Импорт'] = volume * import_share

                        year_total_volume += volume
                        year_total_export_volume += volume * export_share
                        year_total_import_volume += volume * import_share

                    year_data['Общий объем'] = year_total_volume
                    year_data['Экспорт всего'] = year_total_export_volume
                    year_data['Импорт всего'] = year_total_import_volume
                    year_data['Доля экспорта (%)'] = (
                            year_total_export_volume / year_total_volume * 100) if year_total_volume > 0 else 0
                    year_data['Доля импорта (%)'] = (
                            year_total_import_volume / year_total_volume * 100) if year_total_volume > 0 else 0

                    summary_data.append(year_data)

                summary_df = pd.DataFrame(summary_data)

                # Форматируем таблицу
                format_dict = {}
                for col in summary_df.columns:
                    if 'Объем' in col or 'Экспорт' in col or 'Импорт' in col:
                        format_dict[col] = '{:,.0f}'
                    elif 'Доля' in col:
                        format_dict[col] = '{:.1f}%'

                st.dataframe(summary_df.style.format(format_dict))

            # График объемов производства с разбивкой на экспорт/импорт
            with st.expander("📈 Объемы производства с разбивкой на экспорт и импорт", expanded=False):
                st.subheader("📈 Объемы производства с разбивкой на экспорт и импорт")

                fig_volumes_breakdown = go.Figure()

                for product_type in current_product_types:
                    # Экспортные объемы
                    export_volumes = []
                    import_volumes = []
                    for year in years_list:  # ИСПРАВЛЕНО
                        volume = product_volumes.get(product_type, {}).get(year, 0)
                        export_share = product_export_shares.get(product_type, {}).get(year, 0)
                        import_share = product_import_shares.get(product_type, {}).get(year, 0)
                        export_volumes.append(volume * export_share)
                        import_volumes.append(volume * import_share)

                    # Стек для экспорта
                    fig_volumes_breakdown.add_trace(go.Bar(
                        name=f'{product_type} - Экспорт',
                        x=years_list,  # ИСПРАВЛЕНО
                        y=export_volumes,
                        hovertemplate=f"{product_type} - Экспорт<br>Год: %{{x}}<br>Объем: %{{y:,.0f}} {unit}<extra></extra>",
                        marker_color='#1f77b4'
                    ))

                    # Стек для импорта
                    fig_volumes_breakdown.add_trace(go.Bar(
                        name=f'{product_type} - Импорт',
                        x=years_list,  # ИСПРАВЛЕНО
                        y=import_volumes,
                        hovertemplate=f"{product_type} - Импорт<br>Год: %{{x}}<br>Объем: %{{y:,.0f}} {unit}<extra></extra>",
                        marker_color='#ff7f0e'
                    ))

                fig_volumes_breakdown.update_layout(
                    title="Распределение объемов на экспорт и импорт",
                    xaxis_title="Год",
                    yaxis_title=f"Объем, {unit}",
                    barmode='stack',
                    height=500,
                    showlegend=True
                )
                st.plotly_chart(fig_volumes_breakdown, use_container_width=True)

            # График долей экспорта/импорта
            with st.expander("📊 Доли экспорта и импорта по годам", expanded=False):
                st.subheader("📊 Доли экспорта и импорта по годам")

                fig_shares = go.Figure()

                export_shares_total = []
                import_shares_total = []
                for year in years_list:  # ИСПРАВЛЕНО
                    total_volume = sum(
                        product_volumes.get(product_type, {}).get(year, 0) for product_type in current_product_types)
                    total_export = sum(
                        product_volumes.get(product_type, {}).get(year, 0) * product_export_shares.get(product_type,
                                                                                                       {}).get(year, 0)
                        for
                        product_type in current_product_types)
                    total_import = sum(
                        product_volumes.get(product_type, {}).get(year, 0) * product_import_shares.get(product_type,
                                                                                                       {}).get(year, 0)
                        for
                        product_type in current_product_types)

                    export_share = (total_export / total_volume * 100) if total_volume > 0 else 0
                    import_share = (total_import / total_volume * 100) if total_volume > 0 else 0

                    export_shares_total.append(export_share)
                    import_shares_total.append(import_share)

                fig_shares.add_trace(go.Scatter(
                    name='Доля экспорта',
                    x=years_list,  # ИСПРАВЛЕНО
                    y=export_shares_total,
                    mode='lines+markers',
                    line=dict(color='#1f77b4', width=3),
                    marker=dict(size=8)
                ))

                fig_shares.add_trace(go.Scatter(
                    name='Доля импорта',
                    x=years_list,  # ИСПРАВЛЕНО
                    y=import_shares_total,
                    mode='lines+markers',
                    line=dict(color='#ff7f0e', width=3),
                    marker=dict(size=8)
                ))

                fig_shares.update_layout(
                    title="Динамика долей экспорта и импорта",
                    xaxis_title="Год",
                    yaxis_title="Доля, %",
                    yaxis_range=[0, 100],
                    height=400
                )
                st.plotly_chart(fig_shares, use_container_width=True)

            # ========== БЛОК 5: ЦЕНЫ ==========
            st.subheader("💰 Цены")

            col1, col2 = st.columns(2)
            with col1:
                base_price = st.number_input("Базовая цена (руб.)", value=1000.0, step=100.0)

            with col2:
                if st.session_state.project_data.years:
                    base_year = st.selectbox("Базовый год", st.session_state.project_data.years)

            # Индексы цен по годам - ТОЧНО КАК В БЛОКАХ С ОБЪЕМАМИ
            if st.session_state.project_data.years:
                st.write("**Ценовые индексы по годам (базовый год = 1.00):**")

                # Создаем DataFrame (как в объемах)
                price_data = []
                for year in st.session_state.project_data.years:
                    # Пытаемся получить сохраненные данные
                    if st.session_state.project_data.prices and f"index_{year}" in st.session_state.project_data.prices:
                        default_index = st.session_state.project_data.prices[f"index_{year}"]
                    else:
                        default_index = 1.0 if year == base_year else 1.0
                    price_data.append({'Год': year, 'Ценовой индекс': default_index})

                df = pd.DataFrame(price_data)

                # Редактируемая таблица (КАК В ОБЪЕМАХ)
                edited_df = st.data_editor(
                    df,
                    key=f"price_editor_{base_year}",
                    num_rows="fixed",
                    column_config={
                        'Год': st.column_config.NumberColumn(disabled=True),
                        'Ценовой индекс': st.column_config.NumberColumn(
                            format="%.3f",
                            min_value=0.0,
                            max_value=10.0,
                            step=0.001
                        )
                    }
                )

                # СОХРАНЯЕМ ДАННЫЕ (КАК В ОБЪЕМАХ)
                for _, row in edited_df.iterrows():
                    year = int(row['Год'])
                    index = float(row['Ценовой индекс'])
                    st.session_state.project_data.prices[f"index_{year}"] = index
                    st.session_state.project_data.prices[year] = base_price * index

                # НИКАКОЙ ДОПОЛНИТЕЛЬНОЙ ТАБЛИЦЫ!
                # В блоках с объемами НЕТ второго st.dataframe после редактирования!

            # Дефляторы - только после того, как данные уже сохранены
            if st.session_state.project_data.years and st.session_state.project_data.prices:
                # Проверяем, есть ли у нас сохраненные индексы
                has_indices = any(f"index_{year}" in st.session_state.project_data.prices for year in
                                  st.session_state.project_data.years)

                if has_indices:
                    st.subheader("📉 Дефляторы по годам")

                    # Получаем индексы
                    price_indexes = {}
                    for year in st.session_state.project_data.years:
                        price_indexes[year] = st.session_state.project_data.prices.get(f"index_{year}", 1.0)

                    # Функция расчета дефлятора
                    def calculate_deflator(year, base_year, price_indexes):
                        if year == base_year:
                            return 1.0
                        if year > base_year:
                            result = 1.0
                            for k in range(base_year + 1, year + 1):
                                result *= price_indexes.get(k, 1.0)
                            return result
                        else:
                            result = 1.0
                            for k in range(year + 1, base_year + 1):
                                result *= price_indexes.get(k, 1.0)
                            return 1.0 / result if result != 0 else 1.0

                    # Рассчитываем дефляторы
                    deflators = {}
                    for year in st.session_state.project_data.years:
                        deflators[year] = calculate_deflator(year, base_year, price_indexes)

                    # Показываем результаты (ТОЛЬКО ДЛЯ ПРОСМОТРА, без редактирования)
                    results_data = []
                    for year in st.session_state.project_data.years:
                        results_data.append({
                            'Год': year,
                            'Цена (руб.)': st.session_state.project_data.prices.get(year, 0),
                            'Индекс': price_indexes.get(year, 1.0),
                            'Дефлятор': deflators[year]
                        })

                    results_df = pd.DataFrame(results_data)
                    st.dataframe(
                        results_df.style.format({
                            'Цена (руб.)': '{:,.2f}',
                            'Индекс': '{:.6f}',
                            'Дефлятор': '{:.6f}'
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                    # График
                    fig = go.Figure()
                    colors = ['#2E8B57' if deflators[y] >= 1 else '#FF6B6B' for y in
                              st.session_state.project_data.years]
                    fig.add_trace(go.Bar(
                        x=list(st.session_state.project_data.years),
                        y=list(deflators.values()),
                        marker_color=colors,
                        text=[f"{deflators[y]:.4f}" for y in st.session_state.project_data.years],
                        textposition='outside'
                    ))
                    fig.add_hline(y=1, line_dash="dash", line_color="gray")
                    fig.update_layout(title="Динамика дефляторов", xaxis_title="Год", yaxis_title="Дефлятор",
                                      height=400)
                    st.plotly_chart(fig, use_container_width=True)

                with st.expander("ℹ️ Что такое дефлятор?"):
                    st.markdown(f"""
                    **Дефлятор** - коэффициент изменения цен относительно базового года ({base_year}).

                    - Дефлятор > 1 - цены выросли
                    - Дефлятор = 1 - цены не изменились  
                    - Дефлятор < 1 - цены снизились

                    **Расчет:** цепное произведение ценовых индексов
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
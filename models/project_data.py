"""
Модель данных проекта.
"""


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
            'price_indices': self.price_indices,
        }

    @classmethod
    def from_dict(cls, data):
        obj = cls()
        obj.__dict__.update(data)
        return obj
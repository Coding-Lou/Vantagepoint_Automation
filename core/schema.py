from dataclasses import dataclass, field, asdict
from typing import List


@dataclass
class InvoiceItem:
    item_number: str = ""
    description: str = ""
    quantity: float = 0
    shipped: float = 0
    unit_price: float = 0
    amount: float = 0
    uom: str = ""

    def to_dict(self):
        return asdict(self)


@dataclass
class Invoice:
    vendor: str = ""

    invoice_number: str = ""
    invoice_date: str = ""
    due_date: str = ""

    purchase_order: str = ""
    sales_order: str = ""

    sold_to: str = ""
    ship_to: str = ""

    items: List[InvoiceItem] = field(
        default_factory=list
    )

    subtotal: float = 0
    gst: float = 0
    pst: float = 0
    total: float = 0

    def to_dict(self):
        data = asdict(self)
        data["items"] = [
            item.to_dict()
            for item in self.items
        ]
        return data
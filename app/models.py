from dataclasses import dataclass
from decimal import Decimal


@dataclass
class LineItem:
    description: str
    quantity: Decimal
    unit_price: Decimal

    @property
    def line_total(self) -> Decimal:
        return self.quantity * self.unit_price


@dataclass
class InvoiceTotals:
    subtotal: Decimal
    tax_amount: Decimal
    total_amount: Decimal

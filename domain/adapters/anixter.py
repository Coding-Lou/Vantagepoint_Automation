import re
import pdfplumber
from core.schema import Invoice, InvoiceItem

class AnixterParser:
    def parse(self, pdf_file):
        invoice = Invoice()
        invoice.vendor = "Anixter Canada"

        with pdfplumber.open(pdf_file) as pdf:
            header_text = pdf.pages[0].extract_text() or ""
            all_text = "\n".join(
                page.extract_text(layout=True) or ""
                for page in pdf.pages
            )

        header_line = self.extract_field(header_text, "INVOICE #")
        header_tokens = header_line.split()
        invoice.invoice_number = header_tokens[0] if len(header_tokens) >= 1 else ""
        invoice.invoice_date = self.format_date(header_tokens[1]) if len(header_tokens) >= 2 else ""
        invoice.due_date = self.format_date(header_tokens[2]) if len(header_tokens) >= 3 else ""

        po_so_line = self.extract_field(header_text, "PURCHASE ORDER #")
        invoice.purchase_order = self.extract_po(po_so_line)
        invoice.sales_order = self.extract_so(po_so_line)

        invoice.sold_to = self.extract_section(
            header_text,
            "VENDU À - SOLD TO",
            "EXPÉDIÉ À - SHIP TO"
        )
        invoice.ship_to = self.extract_section(
            header_text,
            "EXPÉDIÉ À - SHIP TO",
            "EXPÉDIÉ DE"
        )
        invoice.items = self.parse_items(all_text)
        invoice.subtotal = self.extract_amount(all_text, "SALES TOTAL")
        invoice.total = self.extract_amount(all_text, "TOTAL DUE")
        invoice.gst, invoice.pst = self.extract_tax(all_text)

        return invoice

    def extract_field(self, text, keyword):
        lines = [
            line.strip()
            for line in text.splitlines()
            if line.strip()
        ]

        for index, line in enumerate(lines):
            if keyword in line:
                if index + 1 < len(lines):
                    return lines[index + 1]

        return ""

    def extract_section(self, text, start, end):
        try:
            section = text.split(start, 1)[1]
            section = section.split(end, 1)[0]
            return section.strip()
        except:
            return ""

    def parse_items(self, text):
        lines = [
            x.strip()
            for x in text.splitlines()
            if x.strip()
        ]

        items = []

        for index, line in enumerate(lines):
            match = re.match(
                r"^(\d{5})\s+(?:\d{5}\s+)?([\w-]+)\s+(\d+)\s+(\d+)\s+(\d+).*?(\$[\d,]+\.\d+)/([A-Z]+)\s+(\$[\d,]+\.\d+)",
                line
            )

            if not match:
                continue

            item = InvoiceItem()

            item.po_line = match.group(1)
            item.item_number = match.group(2)
            item.quantity = float(match.group(3))
            item.shipped = float(match.group(4))
            item.unit_price = self.money(match.group(6))
            item.uom = match.group(7)
            item.amount = self.money(match.group(8))

            description = []
            next_index = index + 1

            while next_index < len(lines):
                next_line = lines[next_index]

                if re.match(r"^\d{5}\s+\S", next_line):
                    break
                if "MONTANT" in next_line:
                    break
                if "$" in next_line:
                    break
                if re.match(r"^NO\.$", next_line):
                    break
                if "TOTAL EN DOLLARS" in next_line or next_line.startswith("Date d"):
                    next_index += 1
                    continue

                description.append(next_line)
                next_index += 1

            item.description = " ".join(description)
            items.append(item)

        return items

    def extract_amount(self, text, keyword):
        match = re.search(
            keyword +
            r".*?\$([\d,]+\.\d+)",
            text,
            re.S
        )

        if match:
            return self.money(
                match.group(1)
            )

        return 0

    def format_date(self, date_str):
        match = re.match(r"(\d{2})/(\d{2})/(\d{4})", date_str)
        if match:
            mm, dd, yyyy = match.groups()
            return f"{yyyy}-{mm}-{dd}"
        return date_str

    def extract_po(self, line):
        for token in line.split():
            if re.match(r"^\d{4,6}$", token):
                return token
        return ""

    def extract_so(self, line):
        for token in line.split():
            if re.match(r"^[A-Z0-9]+$", token) and re.search(r"[A-Z]", token) and re.search(r"\d", token):
                return token
        return ""

    def extract_tax(self, text):
        gst = 0
        pst = 0

        gst_match = re.search(
            r"(?:TPS|GST)[^\n]*\$([\d,]+\.\d+)",
            text
        )

        pst_match = re.search(
            r"(?:TVQ|TVP|PST)[/\s][^\n]*\$([\d,]+\.\d+)",
            text
        )

        if gst_match:
            gst = self.money(gst_match.group(1))

        if pst_match:
            pst = self.money(pst_match.group(1))

        return gst, pst

    def money(self, value):
        return float(
            value
            .replace("$", "")
            .replace(",", "")
        )
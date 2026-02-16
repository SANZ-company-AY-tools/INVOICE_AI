"""
Excel and CSV generator module for invoice data export.
Creates clean Excel files and SAP-compatible CSV from extracted invoice data.
"""

import csv
from typing import List, Dict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os


def _expand_invoices(data: List[Dict]) -> List[Dict]:
    """Expand invoices so each combination of order number + tax line gets its own row.

    Examples:
    - Invoice with 2 order numbers and 1 tax rate = 2 rows
    - Invoice with 1 order number and 2 tax rates = 2 rows
    - Invoice with 2 order numbers and 2 tax rates = 4 rows
    - Invoice with 0 order numbers and 1 tax rate = 1 row
    """
    expanded = []
    for invoice in data:
        if invoice.get('status') != 'success':
            continue

        order_nums = invoice.get('order_numbers', [])
        if not isinstance(order_nums, list):
            order_nums = [order_nums] if order_nums else []
        if not order_nums:
            order_nums = ['']

        tax_lines = invoice.get('tax_lines', [])
        if not tax_lines or not isinstance(tax_lines, list):
            # Fallback for legacy format
            tax_lines = [{
                'base_amount': invoice.get('base_amount'),
                'tax_rate': invoice.get('tax_rate'),
                'tax_amount': invoice.get('tax_amount'),
            }]

        # Create one row per order_number × tax_line combination
        for on in order_nums:
            for tl in tax_lines:
                row = dict(invoice)
                row['order_number'] = str(on) if on else ''
                row['base_amount'] = tl.get('base_amount')
                row['tax_rate'] = tl.get('tax_rate')
                row['tax_amount'] = tl.get('tax_amount')
                expanded.append(row)

    return expanded


class ExcelGenerator:
    """Generate Excel files from extracted invoice data."""

    def __init__(self):
        # Styles
        self.header_font = Font(bold=True, color='FFFFFF', name='Calibri', size=11)
        self.header_fill = PatternFill('solid', start_color='0D9488')  # Turquoise
        self.header_alignment = Alignment(horizontal='center', vertical='center')

        self.data_font = Font(name='Calibri', size=10, color='000000')
        self.number_font = Font(name='Calibri', size=10, color='000000')
        self.account_font = Font(name='Calibri', size=10, bold=True, color='000000')

        self.border = Border(
            bottom=Side(style='thin', color='E5E5E5')
        )

        # Column configuration - now uses order_number (singular, one per row)
        self.columns = [
            {'key': 'date', 'header': 'Fecha', 'width': 12},
            {'key': 'company_name', 'header': 'Emisor', 'width': 30},
            {'key': 'tax_id', 'header': 'CIF Emisor', 'width': 14},
            {'key': 'receiver_name', 'header': 'Receptor', 'width': 30},
            {'key': 'receiver_tax_id', 'header': 'CIF Receptor', 'width': 14},
            {'key': 'order_number', 'header': 'Nº Pedido', 'width': 16},
            {'key': 'concept', 'header': 'Concepto', 'width': 35},
            {'key': 'currency', 'header': 'Divisa', 'width': 8},
            {'key': 'base_amount', 'header': 'Base', 'width': 12, 'format': '#,##0.00'},
            {'key': 'tax_rate', 'header': '%IVA', 'width': 8, 'format': '0"%"'},
            {'key': 'tax_amount', 'header': 'IVA', 'width': 12, 'format': '#,##0.00'},
            {'key': 'total', 'header': 'Total', 'width': 14, 'format': '#,##0.00'},
            {'key': 'accounting_account', 'header': 'Cuenta', 'width': 10},
            {'key': 'accounting_description', 'header': 'Descripción Cuenta', 'width': 25},
        ]

    def create_workbook(self, data: List[Dict], output_path: str) -> str:
        """Create a clean Excel workbook from invoice data."""
        wb = Workbook()
        ws = wb.active
        ws.title = "Facturas"

        # Write headers
        for col_idx, col_config in enumerate(self.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=col_config['header'])
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            ws.column_dimensions[get_column_letter(col_idx)].width = col_config['width']

        # Freeze header
        ws.freeze_panes = 'A2'

        # Expand: each order number gets its own row
        rows = _expand_invoices(data)

        # Write data
        current_row = 2
        for invoice in rows:
            for col_idx, col_config in enumerate(self.columns, 1):
                value = invoice.get(col_config['key'])
                cell = ws.cell(row=current_row, column=col_idx)

                if col_config['key'] == 'accounting_account':
                    cell.value = value
                    cell.font = self.account_font
                    cell.alignment = Alignment(horizontal='center')
                elif 'format' in col_config and value is not None:
                    cell.value = value
                    cell.number_format = col_config['format']
                    cell.alignment = Alignment(horizontal='right')
                    cell.font = self.number_font
                else:
                    cell.value = value if value else ''
                    cell.font = self.data_font

                cell.border = self.border

            current_row += 1

        # Add totals row
        last_row = current_row - 1
        if last_row > 1:
            total_row = last_row + 2

            # Total label (column G = 7 = Concepto)
            ws.cell(row=total_row, column=7, value="TOTAL").font = Font(bold=True, name='Calibri', size=11, color='000000')

            # Base total (column I = 9)
            total_base = ws.cell(row=total_row, column=9)
            total_base.value = f'=SUM(I2:I{last_row})'
            total_base.number_format = '#,##0.00'
            total_base.font = Font(bold=True, name='Calibri', size=11, color='000000')

            # IVA total (column K = 11)
            total_iva = ws.cell(row=total_row, column=11)
            total_iva.value = f'=SUM(K2:K{last_row})'
            total_iva.number_format = '#,##0.00'
            total_iva.font = Font(bold=True, name='Calibri', size=11, color='000000')

            # Grand total (column L = 12)
            total_cell = ws.cell(row=total_row, column=12)
            total_cell.value = f'=SUM(L2:L{last_row})'
            total_cell.number_format = '#,##0.00'
            total_cell.font = Font(bold=True, name='Calibri', size=12, color='000000')

        wb.save(output_path)
        return output_path

    def generate_from_invoices(self, invoices: List[Dict], output_dir: str, filename: str = "facturas.xlsx") -> str:
        """Generate Excel file from list of extracted invoices."""
        output_path = os.path.join(output_dir, filename)
        return self.create_workbook(invoices, output_path)


class SAPCSVGenerator:
    """Generate SAP-compatible CSV from extracted invoice data."""

    # CSV columns for SAP import
    SAP_COLUMNS = [
        'Fecha',            # Posting date
        'Nº Factura',       # Invoice number / reference
        'Nº Pedido',        # Purchase order number
        'CIF Emisor',       # Vendor tax ID (to match SAP vendor)
        'Emisor',           # Vendor name
        'CIF Receptor',     # Receiver tax ID
        'Concepto',         # Description / text
        'Divisa',           # Currency
        'Base Imponible',   # Net amount
        '% IVA',            # Tax rate
        'Importe IVA',      # Tax amount
        'Total',            # Gross amount
        'Cuenta Gasto',     # Expense GL account
    ]

    def generate_csv(self, invoices: List[Dict], output_dir: str, filename: str = "facturas_sap.csv") -> str:
        """Generate SAP-compatible CSV. Each order number gets its own row."""
        output_path = os.path.join(output_dir, filename)
        rows = _expand_invoices(invoices)

        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
            writer.writerow(self.SAP_COLUMNS)

            for inv in rows:
                writer.writerow([
                    inv.get('date', ''),
                    inv.get('invoice_number', ''),
                    inv.get('order_number', ''),
                    inv.get('tax_id', ''),
                    inv.get('company_name', ''),
                    inv.get('receiver_tax_id', ''),
                    inv.get('concept', ''),
                    inv.get('currency', 'EUR'),
                    self._format_decimal(inv.get('base_amount')),
                    self._format_decimal(inv.get('tax_rate')),
                    self._format_decimal(inv.get('tax_amount')),
                    self._format_decimal(inv.get('total')),
                    inv.get('accounting_account', '629'),
                ])

        return output_path

    def _format_decimal(self, value) -> str:
        """Format number for CSV (use comma as decimal separator for SAP)."""
        if value is None:
            return ''
        try:
            return f"{float(value):.2f}".replace('.', ',')
        except (ValueError, TypeError):
            return ''

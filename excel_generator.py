"""
Excel generator module for invoice data export.
Creates clean, simple Excel files from extracted invoice data.
"""

from typing import List, Dict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os


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

        # Simple column configuration
        self.columns = [
            {'key': 'date', 'header': 'Fecha', 'width': 12},
            {'key': 'company_name', 'header': 'Emisor', 'width': 30},
            {'key': 'tax_id', 'header': 'CIF', 'width': 14},
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

        # Write data
        for row_idx, invoice in enumerate(data, 2):
            if invoice.get('status') != 'success':
                continue  # Skip failed extractions

            for col_idx, col_config in enumerate(self.columns, 1):
                value = invoice.get(col_config['key'])
                cell = ws.cell(row=row_idx, column=col_idx)

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

        # Add totals row (note: totals only make sense for same-currency invoices)
        last_row = len([d for d in data if d.get('status') == 'success']) + 1
        if last_row > 1:
            total_row = last_row + 2

            # Total label
            ws.cell(row=total_row, column=4, value="TOTAL").font = Font(bold=True, name='Calibri', size=11, color='000000')

            # Base total (column F = 6, after currency column)
            total_base = ws.cell(row=total_row, column=6)
            total_base.value = f'=SUM(F2:F{last_row})'
            total_base.number_format = '#,##0.00'
            total_base.font = Font(bold=True, name='Calibri', size=11, color='000000')

            # IVA total (column H = 8)
            total_iva = ws.cell(row=total_row, column=8)
            total_iva.value = f'=SUM(H2:H{last_row})'
            total_iva.number_format = '#,##0.00'
            total_iva.font = Font(bold=True, name='Calibri', size=11, color='000000')

            # Grand total (column I = 9)
            total_cell = ws.cell(row=total_row, column=9)
            total_cell.value = f'=SUM(I2:I{last_row})'
            total_cell.number_format = '#,##0.00'
            total_cell.font = Font(bold=True, name='Calibri', size=12, color='000000')

        wb.save(output_path)
        return output_path

    def generate_from_invoices(self, invoices: List[Dict], output_dir: str, filename: str = "facturas.xlsx") -> str:
        """Generate Excel file from list of extracted invoices."""
        output_path = os.path.join(output_dir, filename)
        return self.create_workbook(invoices, output_path)

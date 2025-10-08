import os
import shutil
from copy import copy
import openpyxl
from openpyxl.worksheet.pagebreak import Break, PageBreak
from datetime import datetime, timedelta

from config.settings import TEMPLATE_PATH, TEMP_DIR, EXCEL_START_ROW, EXCEL_TABLE_END_ROW, ROWS_PER_ITEM
from core.utils import format_address_for_excel, number_to_ringgit

class ExcelGenerator:
    def __init__(self):
        self.temp_filepath = None

    def generate_po_excel(self, po_data):
        print("Converting JSON to Excel...")
        
        # Create temporary file
        self.temp_filepath = TEMP_DIR / f"temp_po_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        gui_data = po_data.get('gui_data', {})
        
        try:
            workbook = openpyxl.load_workbook(TEMPLATE_PATH)
            sheet = workbook.active

            # Populate header with GUI data
            self._populate_header(sheet, gui_data)
            
            # Populate supplier information from quote
            self._populate_supplier_info(sheet, po_data, gui_data)
            
            # Populate items table
            total_cost = self._populate_items_table(sheet, po_data)
            
            # Add totals and formatting
            self._add_totals_and_formatting(sheet, total_cost, gui_data)
            
            # Save to temporary file
            workbook.save(self.temp_filepath)
            print(f"✅ Successfully created temporary PO: {self.temp_filepath}")
            
            return self.temp_filepath

        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            raise

    def _populate_header(self, sheet, gui_data):
        """Populate header information from GUI data"""
        sheet['E8'] = gui_data.get('po_number', '')
        sheet['H8'] = gui_data.get('po_issue_date', '')
        sheet['H9'] = gui_data.get('project_number', '')
        sheet['H10'] = gui_data.get('project_name', '')
        sheet['G18'] = f"Contact: {gui_data.get('purchaser_name', '')} ({gui_data.get('purchaser_phone', '')})"
        sheet['G20'] = f"PURCHASE ORDER NO.: {gui_data.get('po_number', '')}"

    def _populate_supplier_info(self, sheet, po_data, gui_data):
        """Populate supplier information from extracted PDF data"""
        sheet['B9'] = po_data.get('companyName')
        
        addr_line1, addr_line2 = format_address_for_excel(po_data.get('address', ''))
        sheet['B10'] = addr_line1
        sheet['B11'] = addr_line2
        
        sheet['C13'] = f": {po_data.get('pic', {}).get('name', '')}"
        sheet['C14'] = f": {po_data.get('pic', {}).get('phone', '')}"
        sheet['C15'] = f": {po_data.get('pic', {}).get('fax', '')}"
        sheet['C16'] = f": {po_data.get('pic', {}).get('email', '')}"
        sheet['A18'] = 'N/A'
        sheet['A21'] = po_data.get('terms', {}).get('payment')
        
        # Calculate delivery date
        delivery_weeks_str = str(po_data.get('terms', {}).get('deliveryWeeks', '0'))
        delivery_weeks_int = int(delivery_weeks_str) if delivery_weeks_str.isdigit() else 0
        if delivery_weeks_int > 0:
            po_issue_date = gui_data.get('po_issue_date', '')
            issue_date_obj = datetime.strptime(po_issue_date, "%d/%m/%Y")
            delivery_date = issue_date_obj + timedelta(weeks=delivery_weeks_int)
            sheet['H24'] = delivery_date.strftime("%d/%m/%Y")
        
        sheet['D29'] = f"With reference to your quotation {po_data.get('quotationNumber', '')}:"

    def _populate_items_table(self, sheet, po_data):
        """Populate items table and return total cost"""
        items_list = po_data.get('items', [])
        start_row = EXCEL_START_ROW
        table_end_row = EXCEL_TABLE_END_ROW
        
        num_items_to_add = len(items_list)
        available_item_slots = (table_end_row - start_row + 1) // ROWS_PER_ITEM

        # Expand table if needed
        if num_items_to_add > available_item_slots:
            self._expand_items_table(sheet, num_items_to_add, available_item_slots, table_end_row)

        total_cost_calculated = 0
        
        # Populate items data
        for index, item in enumerate(items_list):
            current_row = start_row + index * 2
            quantity = item.get('quantity', 0)
            unit_price = item.get('unitPrice', 0)
            line_total = quantity * unit_price
            total_cost_calculated += line_total

            sheet[f'A{current_row}'] = index + 1
            sheet[f'B{current_row}'] = quantity
            sheet[f'C{current_row}'] = item.get('unit')
            sheet[f'D{current_row}'] = item.get('description')
            sheet[f'H{current_row}'] = unit_price
            sheet[f'I{current_row}'] = line_total

        return total_cost_calculated

    def _expand_items_table(self, sheet, num_items_to_add, available_item_slots, table_end_row):
        """Expand the items table if there are more items than available slots"""
        num_rows_to_insert = (num_items_to_add - available_item_slots) * ROWS_PER_ITEM
        insertion_point = table_end_row + 1

        merges_to_shift = []
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row >= insertion_point:
                merges_to_shift.append(merged_range)

        # Unmerge cells before inserting rows
        for merged_range in merges_to_shift:
            sheet.unmerge_cells(str(merged_range))

        # Insert blank rows
        sheet.insert_rows(insertion_point, amount=num_rows_to_insert)

        # Re-apply merges at new locations
        for merged_range in merges_to_shift:
            merged_range.shift(0, num_rows_to_insert)
            sheet.merge_cells(str(merged_range))

        # Copy styles to new rows
        for i in range(0, num_rows_to_insert, ROWS_PER_ITEM):
            new_blank_row_num = table_end_row + i + 1
            new_item_row_num = table_end_row + i + 2
            self._copy_row_style(sheet, table_end_row - 1, new_blank_row_num)
            self._copy_row_style(sheet, table_end_row, new_item_row_num)

    def _add_totals_and_formatting(self, sheet, total_cost, gui_data):
        """Add totals, formatting, and signatures"""
        items_list = []
        num_items = len(items_list)
        if num_items > ((EXCEL_TABLE_END_ROW - EXCEL_START_ROW + 1) // ROWS_PER_ITEM):
            final_table_row = EXCEL_START_ROW + (num_items * ROWS_PER_ITEM) - 1
        else:
            final_table_row = EXCEL_TABLE_END_ROW
        
        # Add total cost
        total_cost_row = final_table_row + 1
        sheet[f'I{total_cost_row}'] = f"=SUM(I{EXCEL_START_ROW}:I{final_table_row})"

        # Add total in words
        total_cost_in_words_row = final_table_row + 3
        sheet[f'E{total_cost_in_words_row}'] = number_to_ringgit(total_cost)
        
        # Add signatures
        name_rows = final_table_row + 10
        sheet[f'G{name_rows}'] = gui_data.get('purchaser_name', '')
        sheet[f'H{name_rows}'] = gui_data.get('director_manager', '')

        # Add page break
        sheet.row_breaks = PageBreak()
        page_break_row = final_table_row + 13
        sheet.row_breaks.append(Break(id=page_break_row))

    def _copy_row_style(self, ws, source_row_num, dest_row_num):
        """Copy row style from source to destination"""
        source_row = ws[source_row_num]
        dest_row = ws[dest_row_num]

        ws.row_dimensions[dest_row_num].height = ws.row_dimensions[source_row_num].height

        for cell in source_row:
            new_cell = dest_row[cell.col_idx - 1]
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = cell.number_format
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

    def cleanup_temp_file(self):
        """Clean up temporary file"""
        if self.temp_filepath and os.path.exists(self.temp_filepath):
            try:
                os.remove(self.temp_filepath)
            except:
                pass  # Ignore cleanup errors
        self.temp_filepath = None
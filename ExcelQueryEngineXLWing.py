import xlwings as xw

class ExcelQueryEngine:
    def __init__(self, excel_path):
        self.wb = xw.Book(excel_path)
        self.data = {}
        for sheet in self.wb.sheets:
            # Convert each sheet into a 2D list (like your openpyxl version)
            values = sheet.used_range.value
            if values is None:
                self.data[sheet.name] = []
            elif isinstance(values[0], list):
                self.data[sheet.name] = values
            else:
                # Single row/col case
                self.data[sheet.name] = [values]

    def get_cell(self, sheet, row, col):
        return self.data.get(sheet, [])[row][col]

    def find_by_value(self, sheet, target):
        matches = []
        sheet_data = self.data.get(sheet, [])
        for row_idx, row in enumerate(sheet_data):
            for col_idx, value in enumerate(row):
                if value == target:
                    matches.append((row_idx, col_idx))
        return matches

    def get_adjacent_value(self, sheet, search_value, offset=(0, 1)):
        matches = self.find_by_value(sheet, search_value)
        if not matches:
            return None
        row, col = matches[0]
        adj_row = row + offset[0]
        adj_col = col + offset[1]
        return self.data[sheet][adj_row][adj_col]

    def extract_table_from_header(self, sheet, header_row_idx):
        sheet_data = self.data.get(sheet, [])
        headers = sheet_data[header_row_idx]
        table_rows = sheet_data[header_row_idx + 1:]
        return [dict(zip(headers, row)) for row in table_rows if any(row)]

    def get_range(self, sheet, start, end=None, size=None):
        result = []
        start_row, start_col = start

        if end:
            end_row, end_col = end
        elif size:
            end_row = start_row + size[0] - 1
            end_col = start_col + size[1] - 1
        else:
            raise ValueError("You must provide either 'end' or 'size'.")

        for row_idx in range(start_row, end_row + 1):
            row_values = []
            for col_idx in range(start_col, end_col + 1):
                value = self.get_cell(sheet, row_idx, col_idx)
                row_values.append(value)
            result.append(row_values)

        return result

    @staticmethod
    def add_offset(base, offset):
        base_row, base_col = base
        offset_row, offset_col = offset
        return (base_row + offset_row, base_col + offset_col)

    @staticmethod
    def excel_ref_to_index(cell_ref):
        # xlwings uses Excel's native addressing, so we need to parse manually
        import re
        match = re.match(r"([A-Z]+)([0-9]+)", cell_ref.upper())
        if not match:
            raise ValueError(f"Invalid cell reference: {cell_ref}")
        col_letters, row_str = match.groups()

        # Convert letters to index (A=1 → 0-based index)
        col = 0
        for char in col_letters:
            col = col * 26 + (ord(char) - ord("A") + 1)
        return (int(row_str) - 1, col - 1)

    @staticmethod
    def parse_excel_range(range_ref):
        if ":" in range_ref:
            start_ref, end_ref = range_ref.split(":")
            return (
                ExcelQueryEngine.excel_ref_to_index(start_ref),
                ExcelQueryEngine.excel_ref_to_index(end_ref)
            )
        else:
            idx = ExcelQueryEngine.excel_ref_to_index(range_ref)
            return (idx, idx)

    def get_range_by_ref(self, sheet, ref):
        start, end = self.parse_excel_range(ref)
        return self.get_range(sheet, start=start, end=end)

    def get_columns_from_row(self, sheet, columns, start_row):
        result = []
        sheet_data = self.data.get(sheet, [])
        start_idx = start_row - 1
        col_indices = [self.excel_ref_to_index(col + "1")[1] for col in columns]

        for row in sheet_data[start_idx:]:
            selected = []
            for col_idx in col_indices:
                if col_idx < len(row):
                    selected.append(row[col_idx])
                else:
                    selected.append(None)
            result.append(selected)

        return result

    def get_columns_from_row_to_row(self, sheet, columns, start_row, end_row):
        result = []
        sheet_data = self.data.get(sheet, [])
        start_idx = start_row - 1
        end_idx = end_row - 1
        col_indices = [self.excel_ref_to_index(col + "1")[1] for col in columns]

        for row in sheet_data[start_idx:end_idx]:
            selected = []
            for col_idx in col_indices:
                if col_idx < len(row):
                    selected.append(row[col_idx])
                else:
                    selected.append(None)
            result.append(selected)

        return result

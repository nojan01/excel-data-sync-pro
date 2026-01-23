#!/usr/bin/env python3
import sys
sys.path.insert(0, 'python')
from excel_writer import write_sheet

result = write_sheet(
    '/Users/nojan/Documents/Projects/mvms-productions/2025-001 - Schuler Bau Jubiläum 2025/Material-01 - Vorlagen Messe/MVK-Schuler-Bau.xlsx',
    '/Users/nojan/Documents/GitHub/mvms-tool-electron/test-exports/test-row-col-combined.xlsx',
    'Daten',  # sheet_name
    {
        'headers': ['Status', 'Typ', 'Bezeichnung', 'Zusatz', 'Kommentar', 'P1', 'H', 'P2', 'J', 'S/K', 'Annahme (S)', 'Annahme (K)', 'Rückgabe (S)', 'Rückgabe (K)'],
        'deleted_columns': [3],
        'inserted_columns': {'operations': [{'position': 7, 'count': 1}, {'position': 9, 'count': 1}]},
        'row_mapping': {
            1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10,
            11: 11, 12: 12, 13: 13, 14: 14, 15: 15, 16: 16, 17: 17, 18: 18, 19: 19, 20: 20,
            21: 21, 22: 22, 23: 23, 24: 24, 25: 25, 26: 26, 27: 27, 28: 28, 29: 29, 30: 30,
            31: -1, 32: -1, 33: -1
        }
    },
    original_path='/Users/nojan/Documents/Projects/mvms-productions/2025-001 - Schuler Bau Jubiläum 2025/Material-01 - Vorlagen Messe/MVK-Schuler-Bau.xlsx'
)
print(f'Result: {result}')

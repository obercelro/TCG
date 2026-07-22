"""
Data Loading and Formatting Engine for TCG.
Utilizes the xlwings COM interface to inject styling, conditional formatting, and generate data tables within the final output artifact.
"""

import xlwings as xw
import time
import os
from typing import List

def format_test_case(test_case: str, sim_summaries: List[str], available_sims: List[str], template_name: str) -> None:
    """
    Applies visual styling, QA templates, and metadata aggregations to the raw DataFrame outputs.

    Args:
        test_case (str): The absolute filepath to the unformatted pandas output file.
        sim_summaries (List[str]): List of absolute filepaths to the original simulation summary files.
        available_sims (List[str]): Correlating list of parsed system identifiers.
        template_name (str): Configuration template to utilize for stylistic structure.
    """
    start = time.time()
    excel_idx = 1
    available_idx = 0
    test_case_template = os.path.join(os.path.dirname(__file__), f'{template_name}_Template.xlsx')

    print('Beginning test case formatting and SIM summary transfer.')

    with xw.App(visible=False) as app:
        tc = app.books.open(test_case)
        template = app.books.open(test_case_template)
        
        # Format all target test cases and configure tabs
        for sheet in tc.sheets:
            template.sheets['TEMPLATE'].range('1:5').copy(sheet.range('1:5'))
            sheet.range((4, 1)).api.Font.Bold = True
            sheet.range((4, 1)).api.Font.Underline = True
            sheet.api.Tab.ColorIndex = 6 
            
            # Apply conditional QA formatting
            for cell in sheet.range('B6:B40'):
                if cell.value == 'PENDING':
                    cell.color = '#FFFF00'

        # Construct dynamic aggregate statistics page
        stats_sheet = tc.sheets.add(name='TEST STATISTICS', before=tc.sheets[0])
        num_sims = str(len(available_sims))
        
        template.sheets['TEST STATISTICS'].range(f'1:{num_sims}').copy(stats_sheet.range(f'1:{num_sims}'))
        
        for col in ['A', 'B', 'C', 'M', 'N', 'O', 'P']:
            stats_sheet.api.Columns(col).Hidden = True
            
        if len(stats_sheet.tables) == 0:
            stats_sheet.tables.add(stats_sheet.range(f'D1:L{num_sims}'), table_style_name='TableStyleLight8')

        # Import relevant sub-sheets from legacy source systems
        print('Initial formatting finished, starting SIM summary transfer.')
        for sims in sim_summaries:
            sim = app.books.open(sims)
            identifier = available_sims[available_idx]
            sheet_names = [s.name for s in sim.sheets]

            # Route sheet extraction based on legacy file structures
            if sims.endswith('.xls') or (len(sheet_names) > 1 and 'SIM Modes' in sheet_names):
                copy_sheet = sim.sheets['SIM Modes']
            elif len(sheet_names) > 1 and 'DETAIL MODE DESCRIPTIONS' in sheet_names:
                copy_sheet = sim.sheets['DETAIL MODE DESCRIPTIONS']
            elif identifier in sheet_names:
                copy_sheet = sim.sheets[identifier]
            elif 'Specific Sheet Name' in sheet_names:
                copy_sheet = sim.sheets['Param Sets']
            else:
                copy_sheet = sim.sheets[0]

            # Load extracted metadata into finalized output
            copy_sheet.copy(after=tc.sheets[excel_idx], name=f'SIM {identifier}')
            tc.sheets[excel_idx].range((4, 1)).value = identifier

            print(f'{identifier} summary added; {identifier} test case finalized.')
            excel_idx += 2
            available_idx += 1
            sim.close()

        tc.save()
        template.close()
        tc.close()

    print(f'Generated test case with {len(available_sims)} emitters in {(time.time()-start)/60:.2f} minutes.')

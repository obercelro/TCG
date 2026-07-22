import xlwings as xw
import time
import os

def format_test_case(test_case, sim_summaries, available_sims, template_name):
    start = time.time()
    excel_idx = 1
    available_idx = 0
    test_case_template = os.path.join(os.path.dirname(__file__), f'{template_name}_Template.xlsx')

    print('Beginning test case formatting and SIM summary transfer.')

    # Combine operations into a single Excel instance to vastly improve speed
    with xw.App(visible=False) as app:
        tc = app.books.open(test_case)
        template = app.books.open(test_case_template)
        
        # 1. Format all initial test cases
        for sheet in tc.sheets:
            template.sheets['TEMPLATE'].range('1:5').copy(sheet.range('1:5'))
            sheet.range((4, 1)).api.Font.Bold = True
            sheet.range((4, 1)).api.Font.Underline = True
            sheet.api.Tab.ColorIndex = 6 # Fixed 'fc' NameError here
            
            for cell in sheet.range('B6:B40'):
                if cell.value == 'PENDING':
                    cell.color = '#FFFF00'

        # 2. Add overall statistics page
        stats_sheet = tc.sheets.add(name='TEST STATISTICS', before=tc.sheets[0])
        num_sims = str(len(available_sims))
        
        template.sheets['TEST STATISTICS'].range(f'1:{num_sims}').copy(stats_sheet.range(f'1:{num_sims}'))
        
        for col in ['A', 'B', 'C', 'M', 'N', 'O', 'P']:
            stats_sheet.api.Columns(col).Hidden = True
            
        if len(stats_sheet.tables) == 0:
            stats_sheet.tables.add(stats_sheet.range(f'D1:L{num_sims}'), table_style_name='TableStyleLight8')

        # 3. Pull over summary sheets
        print('Initial formatting finished, starting SIM summary transfer.')
        for sims in sim_summaries:
            sim = app.books.open(sims)
            identifier = available_sims[available_idx]
            sheet_names = [s.name for s in sim.sheets]

            # Simplified logic checking
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

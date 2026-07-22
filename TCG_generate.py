import os
import time
import re
import pandas as pd
import numpy as np
from openpyxl.workbook.child import INVALID_TITLE_REGEX
import xlwings as xw

def get_identifiers(system_list_file, out_list=False):
    system_info = pd.read_excel(system_list_file)
    output_path = os.path.dirname(system_list_file) + '\\'

    identifier_column, name_column = 0, 0
    
    # Locate headers dynamically
    for index, row in system_info.iterrows():
        if 'Identifier' in row.values:
            identifier_column = row.index[row == 'Identifier'][0]
        if row.astype(str).str.contains('Partial Name').any():
            name_column = row.index[row.astype(str).str.find('Partial Name') == 0][0]
            break

    systems = system_info[[identifier_column, name_column]].dropna()
    systems.columns = systems.iloc[0]
    systems.columns.values[1] = 'Full Name'
    systems = systems[1:]

    if out_list: 
        systems.iloc[:, 0].to_csv(output_path + 'System simulations to download.txt', index=False,
                                  header=False, sep=' ', mode='a', lineterminator='')
    return systems, output_path

def make_lists(systems, sim_location, template_name):
    sim_summaries, available_summaries, unsupported_summaries = [], [], []

    system_list = sorted(systems.iloc[:, 0].unique().tolist())
    
    check_names = systems['Full Name'].value_counts()
    duplicated_names = check_names[check_names > 1]

    if template_name == "Template 1":
        for file in sorted(os.scandir(sim_location), key=lambda e: e.name):
            if str(file.name)[11:16] in system_list:
                sim_summaries.append(os.path.join(sim_location, file.name))
                available_summaries.append(str(file.name)[11:16])
    else:
        # Fixed nested loop duplicating entries
        for dirpath, dirnames, filenames in os.walk(sim_location):
            for filename in filenames:
                if filename.endswith('.scn'):
                    unsupported_summaries.append(filename[0:5])

    no_sim = sorted(list(set(system_list) - set(available_summaries) - set(unsupported_summaries)))

    print(f'The following systems are in provided list but no SIM summary exists:\n {no_sim}')
    print(f'The following systems are in provided list but SIM summary doctype is unsupported:\n {unsupported_summaries}')

    return sim_summaries, available_summaries, duplicated_names

def generate_test_case(system_list_file, out_name, sim_location, template_name):
    template_location = os.path.join(os.path.dirname(__file__), f'{template_name}_Template.xlsx')
    test_case_template = pd.read_excel(template_location, sheet_name='TEMPLATE')

    start = time.time()
    duplicate_count = 0
    if not sim_location: 
        sim_location = '/Path/to/default/location'

    systems, out_path = get_identifiers(system_list_file)
    sim_summaries, available_sims, duplicate_names = make_lists(systems, sim_location, template_name)

    test_case_path = os.path.join(out_path, f'{out_name}_test_case.xlsx')
    
    # Open the ExcelWriter ONCE to avoid massive I/O bottlenecks
    with pd.ExcelWriter(test_case_path, engine='openpyxl') as writer:
        
        for available_idx, sim in enumerate(sim_summaries):
            if sim.endswith('.xls'):
                with xw.App(visible=False) as app:
                    simfile = app.books.open(sim)
                    sheet = simfile.sheets['SIM Modes'].used_range.value
                    system = pd.DataFrame(sheet)
                    simfile.close()
            else:
                check_multiple_sheets = pd.ExcelFile(sim)
                sheet_names = check_multiple_sheets.sheet_names # Fixed TypeError
                
                if len(sheet_names) > 1 and 'SIM Modes' in sheet_names:
                    system = pd.read_excel(sim, sheet_name='SIM Modes')
                elif len(sheet_names) > 1 and 'DETAIL MODE DESCRIPTIONS' in sheet_names:
                    system = pd.read_excel(sim, sheet_name='DETAIL MODE DESCRIPTIONS')
                elif available_sims[available_idx] in sheet_names:
                    system = pd.read_excel(sim, sheet_name=available_sims[available_idx])
                elif 'Specific Sheet Name' in sheet_names:
                    system = pd.read_excel(sim, sheet_name='Param Sets').iloc[1:]
                    system.columns.values[0] = 'Mode Name'
                    system.columns.values[0] = 'Parameter B'
                else:
                    system = pd.read_excel(sim)

            while system.filter(like='Mode Name').columns.size < 1 and not system.empty:
                new_header = system.iloc[0]
                system = system[1:]
                system.columns = new_header
                system.index = range(len(system))

            temp_test_case = test_case_template.copy()
            sim_cases_list = []

            for index, row in system.iterrows():
                # Fixed infinite while loops here
                if 'Parameter R' in row and pd.notna(row['Parameter R']):
                    temp = [row.filter(like='Mode Name').iloc[0]]
                    sim_cases_list.append(temp + ['PENDING'] + np.repeat('', len(test_case_template.columns)-2).tolist())
                elif 'Parameter B' in row and pd.notna(row['Parameter B']):
                    temp = [row.filter(like='Mode Name').iloc[0]]
                    sim_cases_list.append(temp + ['PENDING'] + np.repeat('', len(test_case_template.columns)-2).tolist())
            
            # Replaced slow iterative appending with direct concat
            sim_cases = pd.DataFrame(sim_cases_list, columns=temp_test_case.columns, dtype=str)

            temp_test_case.iloc[2, 0] = available_sims[available_idx]
            result = pd.concat([temp_test_case, sim_cases])

            sheet_name = systems.loc[systems['Identifier'] == available_sims[available_idx], 'Full Name'].iloc[0]
            sheet_name = re.sub(INVALID_TITLE_REGEX, '_', sheet_name)
            
            if sheet_name in duplicate_names:
                duplicate_count += 1
                sheet_name = f"{sheet_name}_{duplicate_count}"

            try:
                result.to_excel(writer, sheet_name=sheet_name, index=False)
            except ValueError as e:
                print(f"Failed to write sheet {sheet_name}: {e}")

            print(f'Unformatted test case for system {available_sims[available_idx]} generated successfully.')

    print(f'Generated unformatted test case for {len(available_sims)} systems in {round((time.time() - start)/60, 2)} minutes.')
    return test_case_path, sim_summaries, available_sims

import os, time, re
import pandas as pd
import numpy as np
from openpyxl.workbook.child import INVALID_TITLE_REGEX
import xlwings as xw

def get_identifiers(system_list_file, out_list=False):
  system_info = pd.read_excel(system_list_file)
  output_path = os.path.dirname(system_list_file)+'\\'

  # Get identifier and name columns
  identifier_column = 0
  for index, row in system_info.iterrows():
    if 'Identifier' in row.values:
      identifier_column = row.index[row == 'Identifier'][0]
    if row.str.contains('Partial Name').any():
      name_column = row.index[row.str.find('Partial Name') == 0][0]
      break

  # Extract identifiers and names, remove NAs and return data frame
  systems = system_info[[identifier_column, name_column]].dropna()
  systems.columns = systems.iloc[0]
  systems.columns.values[1] = 'Full Name'
  systems = systems[1:]

  # List of simulated system identifiers formatted for entry into simulation download browser
  if out_list == True: systems.iloc[:,0].to_csv(output_path+'System simulations to download.txt', index = None,
                                                header = None, sep = ' ', mode = 'a', lineterminator = '')
  return systems, output_path

def make_lists(systems, sim_location, template_name):
  sim_summaries, available_summaries, unsupported_summaries = ([] for i in range(3))

  system_list = systems.iloc[:,0]
  system_list = system_list.unique().tolist()
  system_list.sort()

  # Make list of duplicated names for later use
  check_names = systems['Full Name'].value_counts()
  duplicated_names = check_names[check_names > 1]

  # Get SIM summary filepaths in specified SIM directory, and find which summaries are in unsupported formats
  # Unsupported formats are .docx (summaries are unstructured) and .xlsm (macros are disabled in restricted environment)
  if template_name == "Template 1":
    for file in sorted(os.scandir(sim_location), key = lambda e: e.name):
      if str(file)[11:16] in system_list:
        sim_summaries.append(os.path.join(sim_location, file))
        available_summaries.append(str(file)[11:16])
  else:
    for dirpath, dirnames, filenames in os.walk(sim_location):
      for filename in [f for f in filenames if f.endswith('.xlsx') or f.endswith('.xls')]:
        for file in os.listdir(dirpath):
          if file.endswith('.scn'):
            unsupported_summaries.append(file[0:5])

  # Create list of systems which have no SIM available
  no_sim = list(set(system_list) - set(available_summaries) - set(unsupported_summaries))
  no_sim.sort()

  print('The following systems are in provided list but no SIM summary exists:\n {0}'.format(no_sim))
  print('The following systems are in provided list but SIM summary doctype is unsupported:\n {0}'.format(unsupported_summaries))

  return(sim_summaries, available_summaries, duplicated_names)

def generate_test_case(system_list_file, out_name, sim_location, template_name):
  # Read in test case template
  template_location = os.path.join(os.path.dirname(__file__), template_name + '_Template.xlsx')
  test_case_template = pd.read_excel(template_location, sheet_name = 'TEMPLATE')

  # Create empty variables for later use
  start = time.time()
  duplicate_count = 0
  if sim_location == "": sim_location = '/Path/to/default/location'

  # Pull list of identifiers from system list file
  systems, out_path = get_identifiers(system_list_file)
  sim_summaries, available_sims, duplicate_names = make_lists(systems, sim_location, template_name)

  # Generate test case
  available_idx = 0
  test_case = os.path.join(out_path, out_name + '_test_case.xlsx')

  for sim in sim_summaries:
    # Workaround for .xls files due to lack of xlrd package
    if sim.endswith('.xls'):
      with xw.App(visible = False) as app:
        simfile = xw.Book(sim)
        sheet = simfile.sheets['SIM Modes'].used_range.value
        system = pd.DataFrame(sheet)
        simfile.close()
    else:
      check_multiple_sheets = pd.ExcelFile(sim)
      if len(check_multiple_sheets) > 1 and 'SIM Modes' in check_multiple_sheets.sheet_names:
        system = pd.read_excel(sim, sheet_name = 'SIM Modes')
      elif len(check_multiple_sheets) > 1 and 'DETAIL MODE DESCRITPIONS' in check_multiple_sheets.sheet_names:
        system = pd.read_excel(sim, sheet_name = 'DETAIL MODE DESCRIPTIONS')
      elif available_sims[available_idx] in check_multiple_sheets.sheet_names:
        system = pd.read_excel(sim, sheet_name = available_sims[available_idx])
      elif 'Specific Sheet Name' in check_multiple_sheets.sheet_names:
        system = pd.read_excel(sim, sheet_name = 'Param Sets')
        system = system[1:]
        system.columns.values[0] = 'Mode Name'
        system.columns.values[0] = 'Parameter B'
      else:
        system = pd.read_excel(sim)

    while system.filter(like = 'Mode Name').columns.size < 1:
      new_header = system.iloc[0]
      system = system[1:]
      system.columns = new_header
      system.index = range(len(system))

    temp_test_case = test_case_template.copy()
    sim_cases = pd.DataFrame(columns = temp_test_case.columns, dtype = str)

    for index, row in system.iterrows():
      if 'Parameter R' in row:
        while pd.notna(row['Parameter R']):
          temp = [row.filter(like = 'Mode Name').iloc[0]]
          sim_cases.loc[len(sim_cases.index)] = temp + ['PENDING'] + np.repeat('', len(test_case_template.columns)-2).tolist()
        break
      else:
        while pd.notna(row['Parameter B']):
          temp = [row.filter(like = 'Mode Name').iloc[0]]
          sim_cases.loc[len(sim_cases.index)] = temp + ['PENDING'] + np.repeat('', len(test_case_template.columns)-2)
        break

    # Add identifier to top of test case
    temp_test_case.iloc[2,0] = available_sims[available_idx]
    result = pd.concat([temp_test_case, sim_cases])

    sheet_name = systems.loc[systems['Identifier'] == available_sims[available_idx], 'Full Name'].iloc[0]
    sheet_name = re.sub(INVALID_TITLE_REGEX, '_', sheet_name)
    if sheet_name in duplicate_names:
      duplicate_count += 1
      sheet_name = sheet_name + '_' + str(duplicate_count)

    if available_idx == 0:
      with pd.ExcelWriter(test_case, engine = 'openpyxl') as writer:
        try:
          result.to_excel(writer, sheet_name = sheet_name, index = False)
        except:
          ValueError
    else:
      with pd.ExcelWriter(test_case, mode = 'a', engine = 'openpyxl') as writer:
        try:
          result.to_excel(writer, sheet_name = sheet_name, index = False)
        except:
          ValueError

    print('Unformatted test case for system {0} generated successfully.'.format(available_sims[available_idx]))
    available_idx += 1

  print('Generated unformatted test case for {0} systems in {1} minutes.'.format(len(available_sims), round((time.time() - start)/60, 2)))
  return(test_case, sim_summaries, available_sims)
          
        

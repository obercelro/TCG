import xlwings as xw
import time, os

def format_test_case(test_case, sim_summaries, available_sims, template_name):
  # Create starting variables
  start = time.time()
  excel_idx = 1
  available_idx = 0
  test_case_template = os.path.join(os.path.dirname(__file__), template_name + '_Template.xlsx')

  # Intiially format all test cases and add overall statistics page
  with xw.App(visible = False) as app:
    tc = xw.Book(test_case)
    template = xw.Book(test_case_template)
    print('Beginning intial test case formatting.')

    for sheet in tc.sheets:
      template.sheets['TEMPLATE'].range('1:5').copy(tc.sheets[sheet].range('1:5'))
      tc.sheets[sheet].range(4,1).api.Font.Bold = True
      tc.sheets[sheet].range(4,1).api.Font.Underline = True
      fc.sheets[sheet].api.Tab.ColorIndex = 6
      for x in tc.sheets[sheet].range('B6:B40'):
        if x.value == 'PENDING':
          x.color = '#FFFF00'

    tc.sheets.add(name = 'Sheet1', before = tc.sheets[0])
    template.sheets['TEST STATISTICS'].range('1:' + str(len(available_sims))).copy(tc.sheets['Sheet1'].range('1:' + str(len(available_sims))))
    for col in ['A','B','C','M','N','O','P']:
      tc.sheets[0].api.Columns(col).Hidden = True
    if len(tc.sheets[0].tables) == 0: tbl = tc.sheets[0].tables.add(tc.sheets[0].range('D1:L' + str(len(available_sims))), table_style_name = 'TableStyleLight8')
    tc.sheets[0].name = 'TEST STATISTICS'
    tc.save()
    tc.close()

  # Pull over the summary sheet for each simulation, and place them after their corresponding test case
  print('Initial formatting finished, starting SIM summary transfer.')
  with xw.App(visible = False) as app:
      tc = xw.Book(test_case)
      for sims in sim_summaries:
          sim = xw.Book(sims)
          identifier = available_sims[available_idx]

          # Conditionally move simulation summary pages based on their formatting and location in respective workbook
          if sims.endswith('.xls') or len(sim.sheet_names) > 1 and 'SIM Modes' in sim.sheet_names:
            sim.sheets['SIM Modes'].copy(after = tc.sheets[excel_idx], name = ('SIM ' + identifier))
          elif len(sim.sheet_names) > 1 and 'DETAIL MODE DESCRIPTIONS' in sim.sheet_names:
            sim.sheets['DETAIL MODE DESCRIPTIONS'].copy(after = tc.sheets[excel_idx], name = ('SIM ' + identifier))
          elif identifier in sim.sheet_names:
            sim.sheets[identifier].copy(after = tc.sheets[excel_idx], name = ('SIM ' + identifier))
          elif 'Specific Sheet Name' in sim.sheet_names:
            sim.sheets['Param Sets'].copy(after = tc.sheets[excel_idx], name = ('SIM ' + identifier))
          else:
            sim.sheets[0].copy(after = tc.sheets[excel_idx], name = ('SIM ' + identifier))

          # Add the correct identifier to the test case
          tc.sheets[excel_idx].range(4,1).value = identifier

          print('{0} summary added; {0} test case finalized.'.format(identifier))
          excel_idx += 2
          available_idx += 1
          sim.close()

      tc.save()
      tc.close()

  print('Generated test case with {0} emitters in {1} minutes.'.format(len(available_sims), (time.time()-start)/60))

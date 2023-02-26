import pandas as pd
import datetime
import calendar

# make sure excel file is in the same folder with the same name

yelt = pd.read_csv(r'X:\ARCH213\Documents\Python\Projects\venv\CAT_QS\Result_81.csv')
fxtable = pd.read_csv(r'X:\ARCH213\Documents\Python\Projects\venv\CAT_QS\FXTABLE.csv')
cessions = pd.ExcelFile(r'X:\ARCH213\Documents\Python\Projects\venv\CAT_QS\CESSIONS.xlsx')
c_table = pd.read_excel(cessions, 'CESSIONS', header=0)

# del yelt, yelt_fx
# del yelt_fx_c
# del grouped
# del grouped2

yelt_fx = pd.merge(yelt, fxtable, how='left', on=['LAYERID', 'CURRENCY']).fillna(1)
yelt_fx_c = pd.merge(yelt_fx, c_table, how='left', on=['LAYERID']).fillna(1)




yelt_fx_c['grossLoss'] = yelt_fx_c.apply(lambda x: x['PROJ_GROSSLOSS'] *
                                                   x['PERIOD FXRATE USD'] *
                                                   x['SHARE'] *
                                                   x['LIMITFACTOR'], axis=1)


yelt_fx_c['grossRP'] = yelt_fx_c.apply(lambda x: x['PROJ_GROSSRP'] *
                                                   x['PERIOD FXRATE USD'] *
                                                   x['SHARE'] *
                                                   x['PREMIUMFACTOR'], axis=1)

yelt_fx_c['netLoss1'] = yelt_fx_c.apply(lambda x: x['grossLoss'] * (1 - x['CESSION1']), axis=1)
yelt_fx_c['netLoss2'] = yelt_fx_c.apply(lambda x: x['grossLoss'] * (1 - x['CESSION1'] - x['CESSION2']), axis=1)
yelt_fx_c['netLoss3'] = yelt_fx_c.apply(lambda x: x['grossLoss'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3']), axis=1)
yelt_fx_c['netLoss4'] = yelt_fx_c.apply(lambda x: x['grossLoss'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3'] - x['CESSION4']), axis=1)
yelt_fx_c['netLoss5'] = yelt_fx_c.apply(lambda x: x['grossLoss'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3'] - x['CESSION4'] - x['CESSION5']), axis=1)

yelt_fx_c['netRP1'] = yelt_fx_c.apply(lambda x: x['grossRP'] * x['CESSION1'], axis=1)
yelt_fx_c['netRP2'] = yelt_fx_c.apply(lambda x: x['grossRP'] * (1 - x['CESSION1'] - x['CESSION2']), axis=1)
yelt_fx_c['netRP3'] = yelt_fx_c.apply(lambda x: x['grossRP'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3']), axis=1)
yelt_fx_c['netRP4'] = yelt_fx_c.apply(lambda x: x['grossRP'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3'] - x['CESSION4']), axis=1)
yelt_fx_c['netRP5'] = yelt_fx_c.apply(lambda x: x['grossRP'] * (1 - x['CESSION1'] - x['CESSION2'] - x['CESSION3'] - x['CESSION4'] - x['CESSION5']), axis=1)

yelt_fx_c['retainedGross'] = yelt_fx_c.apply(lambda x: x['grossLoss'] - x['grossRP'], axis=1)
yelt_fx_c['retainedL1'] = yelt_fx_c.apply(lambda x: x['netLoss1'] - x['netRP1'], axis=1)
yelt_fx_c['retainedL2'] = yelt_fx_c.apply(lambda x: x['netLoss2'] - x['netRP2'], axis=1)
yelt_fx_c['retainedL3'] = yelt_fx_c.apply(lambda x: x['netLoss3'] - x['netRP3'], axis=1)
yelt_fx_c['retainedL4'] = yelt_fx_c.apply(lambda x: x['netLoss4'] - x['netRP4'], axis=1)
yelt_fx_c['retainedL5'] = yelt_fx_c.apply(lambda x: x['netLoss5'] - x['netRP5'], axis=1)

yelt_final = yelt_fx_c[["EVENTYEAR", "TOPUPZONE",
                        "grossLoss", "netLoss1", "netLoss2", "netLoss3", "netLoss4", "netLoss5",
                        "grossRP", "netRP1", "netRP2", "netRP3", "netRP4", "netRP5",
                        "retainedGross", "retainedL1", "retainedL2", "retainedL3", "retainedL4", "retainedL5"]].copy()


grouped = yelt_final.groupby(['EVENTYEAR', 'TOPUPZONE']).sum().reset_index()

grouped2 = yelt_final.groupby(['EVENTYEAR']).sum().reset_index()

grouped3 = grouped[["EVENTYEAR", "TOPUPZONE", "retainedL5"]].copy()

grouped3_pivot = grouped3.pivot(index='EVENTYEAR', columns='TOPUPZONE', values='retainedL5').fillna(0).reset_index(drop=False)

####################################################################################################################

import xlwings as xw
wb = xw.Book()  # this will create a new workbook
# wb = xw.Book(r'X:\ARCH213\Documents\Python\Projects\venv\LargeLoss\ModellingView\LL_Nicole.xlsm')

##Output
ws1 = wb.sheets['Sheet1']
ws1.range("A1").clear_contents()
ws1.range("A1").options(index=False, header=True).value = grouped2

wb = xw.Book()  # this will create a new workbook
# wb = xw.Book(r'X:\ARCH213\Documents\Python\Projects\venv\LargeLoss\ModellingView\LL_Nicole.xlsm')

##Output
ws1 = wb.sheets['Sheet1']
ws1.range("A1").clear_contents()
ws1.range("A1").options(index=False, header=True).value = grouped3_pivot

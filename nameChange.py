import pandas as pd
import re

def main():
    workbook = pd.ExcelFile(r'C:\siteNameChange.xlsx')
    siteName = pd.read_excel(workbook,sheet_name='SiteName').astype('str')
    hub = pd.read_excel(workbook, sheet_name='HUB').astype('str')
    fg = pd.read_excel(workbook, sheet_name='FG').astype('str')
    asr = pd.read_excel(workbook, sheet_name='ASR').astype('str')
    mux = pd.read_excel(workbook, sheet_name='MUX').astype('str')
    solar = pd.read_excel(workbook, sheet_name='Solar').astype('str')
    asc = pd.read_excel(workbook, sheet_name='ASC-48').astype('str')
    sla = pd.read_excel(workbook, sheet_name='SLA').astype('str')

    # siteName = siteName.astype('str')
    # hub = hub.astype('str')
    # fg = fg.astype('str')
    # asr = asr.astype('str')
    # mux = mux.astype('str')
    # solar = solar.astype('str')
    # asc = asc.astype('str')
    # sla = sla.astype('str')

    siteName['same'] = ""
    siteName['oldHub'] = ""
    siteName['newHub'] = ""
    siteName['fg'] = ""
    siteName['asr'] = ""
    siteName['mux'] = ""
    siteName['solar'] = ""
    siteName['asc'] = ""
    siteName['sla'] = ""

    for index, row in siteName.iterrows():
        val = row['site_old_name']
        if checkToStandartName(val):

            siteName.loc[siteName['site_old_name'] == val, 'site_new_name'], siteName['sla'][index] = checkToSLA_update\
                (checkToASC_update
                 (checkToSolar_update
                  (checkToMUX_update
                   (checkToASR_update
                    (checkToFG_update
                     (checkToHub_update(val, hub), fg), asr), mux), solar), asc), sla)
        else:
            siteName.loc[siteName['site_old_name'] == val, 'site_new_name'] = val


    for i in siteName.index:
        siteName['same'][i] = siteName['site_old_name'][i] == siteName['site_new_name'][i]
        siteName['fg'][i] = 'same' if 'FG' in siteName['site_old_name'][i] and 'FG' in siteName['site_new_name'][i] else 'diff' if 'FG' in siteName['site_old_name'][i] or 'FG' in siteName['site_new_name'][i] else ""
        siteName['asr'][i] = 'same' if ('ASR' in siteName['site_old_name'][i] and 'ASR' in siteName['site_new_name'][i]) or ('ATN' in siteName['site_old_name'][i] and 'ATN' in siteName['site_new_name'][i]) else 'diff' if ('ASR' in siteName['site_old_name'][i] or 'ASR' in siteName['site_new_name'][i]) or ('ATN' in siteName['site_old_name'][i] or 'ATN' in siteName['site_new_name'][i]) else ""
        siteName['mux'][i] = 'same' if 'MUX' in siteName['site_old_name'][i] and 'MUX' in siteName['site_new_name'][i] else 'diff' if 'MUX' in siteName['site_old_name'][i] or 'MUX' in siteName['site_new_name'][i] else ""
        siteName['solar'][i] = 'same' if 'Solar' in siteName['site_old_name'][i] and 'Solar' in siteName['site_new_name'][i] else 'diff' if 'Solar' in siteName['site_old_name'][i] or 'Solar' in siteName['site_new_name'][i] else ""
        siteName['asc'][i] = 'same' if 'ASC' in siteName['site_old_name'][i] and 'ASC' in siteName['site_new_name'][i] else 'diff' if 'ASC' in siteName['site_old_name'][i] or 'ASC' in siteName['site_new_name'][i] else ""
        if '_(' in siteName['site_old_name'][i]:
            siteName['oldHub'][i] = siteName['site_old_name'][i].split('(')[1].split(')')[0]
        if '_(' in siteName['site_new_name'][i]:
            siteName['newHub'][i] = siteName['site_new_name'][i].split('(')[1].split(')')[0]

    siteName.to_excel('siteOldNewName.xlsx', index=False)

sla_lst = ['S10','S20','S30','S40','S00','S11','S21','S31','S41','S12','S22','S32','S42','S02','S23','S13','S33','S43','S03']
sla_lst_p = ['S10 ','S20 ','S30 ','S40 ','S00 ','S11 ','S21 ','S31 ','S41 ','S12 ','S22 ','S32 ','S42 ','S02 ','S23 ','S13 ','S33 ','S43 ','S03 ']

def checkToStandartName(name):
    return name[:2].isalpha() and name[2:6].isdigit()

def checkSiteInSheet(name,sheet):
    return name[:6] in sheet['site'].values

    """
    for index, row in sheet.iterrows():
        if row[0] == name[:6]:
            return True
    return False
    """
def insertLabel(name, label):
    if len(name.split('_')) == 1:
        name = name + '_' + label
        return name
    elif (name[-3:] in sla_lst or name[-4:] in sla_lst_p) and len(name.split('_')) == 2:
        name = name[:name.find('_')] + '_' + label + name[name.find('_'):]
        return name
    elif (name[-3:] not in sla_lst or name[-4:] not in sla_lst_p) and len(name.split('_')) == 2 and '_(' in name:
        name = name + '_' + label
        return name

    if label == 'FG' or label == 'DG':
        if '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    elif label in ['ASR', 'ATN']:
        if '_FG' in name or '_DG' in name:
            name = name.replace('_FG', '_FG_' + label, 1)
            name = name.replace('_DG', '_DG_' + label, 1)
            return name
        elif '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    elif label == 'MUX':
        if '_ASR' in name or '_ATN' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
            name = name.replace('_ATN', '_ATN_' + label, 1)
            return name
        elif '_FG' in name or '_DG' in name:
            name = name.replace('_FG', '_FG_' + label, 1)
            name = name.replace('_DG', '_DG_' + label, 1)
            return name
        elif '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    elif label == 'Solar':
        if '_MUX' in name:
            name = name.replace('_MUX', '_MUX_' + label, 1)
            return name
        elif '_ASR' in name or '_ATN' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
            name = name.replace('_ATN', '_ATN_' + label, 1)
            return name
        elif '_FG' in name or '_DG' in name:
            name = name.replace('_FG', '_FG_' + label, 1)
            name = name.replace('_DG', '_DG_' + label, 1)
            return name
        elif '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    elif label == 'ASC':
        if '_Solar' in name:
            name = name.replace('_Solar', '_Solar_' + label, 1)
            return name
        elif '_MUX' in name:
            name = name.replace('_MUX', '_MUX_' + label, 1)
            return name
        elif '_ASR' in name or '_ATN' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
            name = name.replace('_ATN', '_ATN_' + label, 1)
            return name
        elif '_FG' in name or '_DG' in name:
            name = name.replace('_FG', '_FG_' + label, 1)
            name = name.replace('_DG', '_DG_' + label, 1)
            return name
        elif '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    return name


def checkToHub_update(name, hub_sheet):
    name = name.replace('HUB','',1)
    if checkSiteInSheet(name, hub_sheet) and name[6:7] not in ['U', 'L']:
        if '_(' in name:
            index1 = name.find('(')
            index2 = name.find(')')
            name = name[:index1+1] + str(hub_sheet.loc[hub_sheet['site'] == name[:6]].iloc[0,1]) + name[index2:]
            return name
        elif len(name.split('_')) > 1:
            name = name.split('_',1)[0] + '_(' + str(hub_sheet.loc[hub_sheet['site'] == name[:6]].iloc[0,1]) + ')_' + name.split('_',1)[1]
            return name
        else:
            name = name + '_(' + str(hub_sheet.loc[hub_sheet['site'] == name[:6]].iloc[0,1]) + ')'
            return name
    elif '_(' in name:
        name = name.split('_',2)[0] + '_' + name.split('_',2)[2]
        return name
    else:
        return name

def checkToFG_update(name, fg_sheet):
    if checkSiteInSheet(name, fg_sheet):
        if '_FG' in name or '_DG' in name:
            return name
        else:
            return insertLabel(name, fg_sheet.loc[fg_sheet['site'] == name[:6]].iloc[0,1])
    elif '_FG' in name or '_DG' in name:
        name = name.replace('_FG', '', 1)
        name = name.replace('_DG', '', 1)
        return name
    return name

def checkToASR_update(name, asr_sheet):
    if checkSiteInSheet(name, asr_sheet):
        label = asr_sheet.loc[asr_sheet['site'] == name[:6]].iloc[0,1]
        if '_ASR' in name:
            if 'ASR' == label:
                return name
            else:
                return name.replace('_ASR', '_'+label, 1)
        elif '_ATN' in name:
            if 'ATN' == label:
                return name
            else:
                return name.replace('_ATN', '_'+label, 1)
        else:
            return insertLabel(name, label)
    elif '_ASR' in name or '_ATN' in name:
        name = name.replace('_ASR', '', 1)
        name = name.replace('_ATN', '', 1)
        return name
    return name

def checkToMUX_update(name, mux_sheet):
    if checkSiteInSheet(name, mux_sheet):
        if '_MUX' in name:
            return name
        else:
            return insertLabel(name, mux_sheet.loc[mux_sheet['site'] == name[:6]].iloc[0, 1])
    elif '_MUX' in name:
        name = name.replace('_MUX', '', 1)
        return name
    return name

def checkToSolar_update(name, solar_sheet):
    if checkSiteInSheet(name, solar_sheet):
        if '_Solar' in name:
            return name
        else:
            return insertLabel(name, solar_sheet.loc[solar_sheet['site'] == name[:6]].iloc[0, 1])
    elif '_Solar' in name:
        name = name.replace('_Solar', '', 1)
        return name
    return name

def checkToASC_update(name, asc_sheet):
    if checkSiteInSheet(name, asc_sheet):
        if '_ASC' in name:
            return name
        else:
            return insertLabel(name, asc_sheet.loc[asc_sheet['site'] == name[:6]].iloc[0, 1])
    elif '_ASC' in name:
        name = name.replace('_ASC', '', 1)
        return name
    return name

def checkToSLA_update(name, sla_sheet):
    if checkSiteInSheet(name, sla_sheet):
        sla = str(sla_sheet.loc[sla_sheet['site'] == name[:6]].iloc[0, 1])
        if name[-3:] in sla_lst:
            if name[-3:] == sla:

                return name, 'same'
            else:
                return name[:len(name)-3] + sla, 'diff'
        elif name[-4:] in sla_lst_p:
            if name[-4:][:3] == sla:
                return name[:len(name)-1], 'same'
            else:
                return name[:len(name) - 4] + sla, 'diff'
        elif name[-2:] in ['S0','S1','S2','S3']:
            return name[:len(name) - 2] + sla, 'diff'
        elif name[-3:] in ['S0 ', 'S1 ', 'S2 ', 'S3 ']:
            return name[:len(name) - 3] + sla, 'diff'
        else:
            return name + '_' + sla, 'diff'
    else:
        if name[-4:] in sla_lst_p or name[-3:] in ['S0 ','S1 ','S2 ','S3 ']:
            return name[:len(name)-1], 'same'
        elif name[-3:] in sla_lst or name[-2:] in ['S0','S1','S2','S3']:
            return name, 'same'
        else:
            return name, ""



main()






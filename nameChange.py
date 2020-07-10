import pandas as pd
import re

def main():
    workbook = pd.ExcelFile(r'E:\siteNameChange.xlsx')
    siteName = pd.read_excel(workbook,sheet_name='SiteName')
    hub = pd.read_excel(workbook, sheet_name='HUB')
    fg = pd.read_excel(workbook, sheet_name='FG')
    asr = pd.read_excel(workbook, sheet_name='ASR')
    mux = pd.read_excel(workbook, sheet_name='MUX')
    solar = pd.read_excel(workbook, sheet_name='Solar')
    asc = pd.read_excel(workbook, sheet_name='ASC-48')
    sla = pd.read_excel(workbook, sheet_name='SLA')

    for index, row in siteName.iterrows():
        val = row['site_old_name']
        if checkToStandartName(val):

            siteName.loc[siteName['site_old_name'] == val, 'site_new_name'] = checkToSLA_update\
                (checkToASC_update
                 (checkToSolar_update
                  (checkToMUX_update
                   (checkToASR_update
                    (checkToFG_update
                     (checkToHub_update(val, hub), fg), asr), mux), solar), asc), sla)
        else:
            siteName.loc[siteName['site_old_name'] == val, 'site_new_name'] = val

    siteName.to_excel('siteOldNewName.xlsx', index=False)


def checkToStandartName(name):
    return name[:2].isalpha() and name[2:6].isdigit()

def checkSiteInSheet(name,sheet):
    for index, row in sheet.iterrows():
        if row[0] == name[:6]:
            return True
    return False

def insertLabel(name, label):
    if len(name.split('_')) == 1:
        name = name + '_' + label
        return name
    elif (name[-2:] in ['S0','S1','S2','S3'] or name[-3:] in ['S0 ','S1 ','S2 ','S3 ']) and len(name.split('_')) == 2:
        name = name[:name.find('_')] + '_' + label + name[name.find('_'):]
        return name
    elif (name[-2:] in ['S0','S1','S2','S3'] or name[-3:] in ['S0 ','S1 ','S2 ','S3 ']) and len(name.split('_')) == 2 and '_(' in name:
        name = name + '_' + label
        return name

    if label == 'FG' or label == 'DG':
        if '_(' in name:
            name = name.replace(')', ')_' + label, 1)
            return name
        else:
            name = name.replace('_', '_' + label + '_', 1)
            return name
    elif label == 'ASR':
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
        if '_ASR' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
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
        elif '_ASR' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
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
        elif '_ASR' in name:
            name = name.replace('_ASR', '_ASR_' + label, 1)
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
        if '_ASR' in name:
            return name
        else:
            return insertLabel(name, asr_sheet.loc[asr_sheet['site'] == name[:6]].iloc[0,1])
    elif '_ASR' in name:
        name = name.replace('_ASR', '', 1)
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
    if name[-2:] in ['S0','S1','S2','S3'] or name[-3:] in ['S0 ','S1 ','S2 ','S3 ']:
        return name
    elif checkSiteInSheet(name, sla_sheet):
        name = name + '_' + str(sla_sheet.loc[sla_sheet['site'] == name[:6]].iloc[0,1])
        return name
    return name

main()






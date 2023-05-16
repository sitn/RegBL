from openpyxl import load_workbook
from datetime import datetime
import requests
import os
import shutil
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker


def _findRowIndex(ws, searchTerm, column_id=1, limit=1e3):
    index = None
    
    row_id = 1
    while row_id < limit:
        if ws.cell(row_id,column_id).value == searchTerm:
            index = row_id
            break
        row_id += 1

    return index


def _findColumnIndex(ws, searchTerm, row_id=1, limit=1e3):
    index = None
    
    column_id = 1
    while column_id < limit:
        if ws.cell(row_id,column_id).value == searchTerm:
            index = column_id
            break
        column_id += 1

    return index


def _removeRows(ws, searchTerm, row_id=1, column_id=1, limit=1e3):
    while row_id < limit and ws.cell(row_id,column_id).value is not None:
        if ws.cell(row_id,column_id).value == searchTerm:
            row_id += 1
        else:
            ws.delete_rows(row_id)
    return


def _getIssue22OfCommune(wb, commune_id, issue22_list):
    ws = wb.create_sheet('ISSUE22')
    currentLine_idx = 1

    if issue22_list == []:
        return

    # title line
    for j, title in enumerate(issue22_list[0].keys()):
        ws.cell(currentLine_idx,j+1).value = title

    currentLine_idx += 1

    for i22 in issue22_list:
        if i22['COM_FOSNR'] == commune_id:
            ws.cell(currentLine_idx,1).value = i22['COM_FOSNR']
            ws.cell(currentLine_idx,2).value = i22['AV_SOURCE']
            ws.cell(currentLine_idx,3).value = i22['AV_TYPE']
            ws.cell(currentLine_idx,4).value = i22['ISSUE']
            ws.cell(currentLine_idx,5).value = i22['ISSUE_CATEGORY']
            ws.cell(currentLine_idx,6).value = i22['BDG_E']
            ws.cell(currentLine_idx,7).value = i22['BDG_N']
            ws.cell(currentLine_idx,8).hyperlink = os.environ['FEEDBACK_COMMUNES_URL_CONSULTATION_ISSUE_22_SITN_COORD'].format(i22['BDG_E'],i22['BDG_N'])
            ws.cell(currentLine_idx,8).value = 'sitn.ne.ch'
            ws.cell(currentLine_idx,8).style = 'Hyperlink'
            currentLine_idx += 1

    return wb


def loadCommunesOFS():
    communes_ofs = {}

    wb_source = load_workbook(os.environ['STATIC_PATH_REPERTOIRE_COMMUNES_OFS'])
    ws_source = wb_source['GDE']

    i = 0
    while i < 1e4:
        i += 1
        
        if ws_source.cell(i,1).value is None:
            break

        if ws_source.cell(i,1).value != 'NE':
            continue

        communes_ofs[ws_source.cell(i,3).value] = ws_source.cell(i,5).value

    return communes_ofs


def downloadListeCantonNeuchatel(path=None):
    if path is None or not os.path.exists(path):
        path = os.environ['FEEDBACK_COMMUNES_WORKING_DIR']

    r = requests.get(os.environ['STATIC_URL_FEEBDACK_CANTON'], allow_redirects=True)

    filename = datetime.strftime(datetime.now(), '%Y%M%d_Listes_NE.xlsx')

    filepath = os.path.join(path, filename)
    open(filepath, 'wb').write(r.content)

    return filepath


def downloadIssue22CantonNeuchatel(path=None):
    if path is None or not os.path.exists(path):
        path = os.environ['FEEDBACK_COMMUNES_WORKING_DIR']

    r = requests.get(os.environ['FEEDBACK_COMMUNES_URL_ISSUE_22_CANTON'], allow_redirects=True)

    filename = datetime.strftime(datetime.now(), '%Y%M%d_issue22_NE.xlsx')

    filepath = os.path.join(path, filename)
    open(filepath, 'wb').write(r.content)

    wb = load_workbook(filepath)
    ws = wb['NE']
    
    issue22_list = []
    i = 1
    while i < 1e5 and ws.cell(i,1) is not None:
        if ws.cell(i,5).value == 22:
            issue22_list.append({
                'COM_FOSNR': ws.cell(i,1).value,
                'AV_SOURCE': ws.cell(i,2).value,
                'AV_TYPE': ws.cell(i,3).value,
                'ISSUE': ws.cell(i,4).value,
                'ISSUE_CATEGORY': ws.cell(i,5).value,
                'BDG_E': ws.cell(i,6).value,
                'BDG_N': ws.cell(i,7).value
            })

        i += 1

    return (issue22_list, filepath)


def cleanWorkingDirectory(path=None):
    if path is not None and os.path.exists(path):
        shutil.rmtree(path)
    return


def generateCommuneErrorFile(commune_id, commune_name, feedback_canton_filepath, issue22_list, today=datetime.strftime(datetime.now(), '%Y%m%d')):
    # copy canton_file to commune_file
        feedback_commune_filename = '_'.join([str(commune_id), commune_name.replace(' ', '_'), 'feedback', today]) + '.xlsx'
        feedback_commune_filepath = os.path.join(os.environ['FEEDBACK_COMMUNES_WORKING_DIR'], today, feedback_commune_filename)
        shutil.copy2(feedback_canton_filepath, feedback_commune_filepath)
        
        feedback_commune = {
            'commune_id': commune_id,
            'commune_nom': commune_name
        }

        wb = load_workbook(feedback_commune_filepath)
        
        # remove canton sheet
        wb.remove(wb['Cantons'])

        # remove other communes in communes sheet
        ws = wb['Communes']
        row_i = _findRowIndex(ws, 'Commune', column_id=4)
        _removeRows(ws, commune_id, row_id=row_i+2, column_id=3, limit=1e2)

        # go through lists
        nb_errors = 0
        for i in range(6):
            ws = wb['Liste ' + str(i+1)]

            # find first line of table
            current_line_idx = _findRowIndex(ws, 'KT', column_id=1)

            if current_line_idx is None:
                continue

            # find column "issues" if exists
            issues_col_idx = _findColumnIndex(ws, 'ISSUES', row_id=current_line_idx)

            # update line index to first line in table
            current_line_idx += 1

            delete_row_idx = None
            delete_row_amount = 0
            k = 0
            while k < 1e4 and ws.cell(current_line_idx,1).value is not None:
                if ws.cell(current_line_idx,2).value == commune_id:
                    # delete rows if necessary
                    if delete_row_idx is not None and delete_row_amount > 0:
                        ws.delete_rows(delete_row_idx, amount=delete_row_amount)
                        current_line_idx = delete_row_idx
                        delete_row_idx = None
                        delete_row_amount = 0

                    # if column issues containes "<br>", change it to "\n"
                    if issues_col_idx is not None:
                        ws.cell(current_line_idx,issues_col_idx).value = ws.cell(current_line_idx,issues_col_idx).value.replace('</br>',' || ')

                    # get coordinates to make geoportal link
                    if i>0:
                        coord_e = None
                        coord_n = None
                        if i==1:
                            coord_e = ws.cell(current_line_idx,10).value
                            coord_n = ws.cell(current_line_idx,11).value
                        if i==2:
                            coord_e = ws.cell(current_line_idx,20).value
                            coord_n = ws.cell(current_line_idx,21).value
                        if i==3:
                            coord_e = ws.cell(current_line_idx,16).value
                            coord_n = ws.cell(current_line_idx,17).value
                        if i==4:
                            print(ws.cell(current_line_idx,9).value)
                            (coord_e, coord_n) = ws.cell(current_line_idx,9).value.split(' ')
                        if i==5:
                            (coord_e, coord_n) = ws.cell(current_line_idx,9).value.split(' ')

                        ws.cell(current_line_idx,4).hyperlink = os.environ['FEEDBACK_COMMUNES_URL_CONSULTATION_ISSUE_22_SITN_COORD'].format(coord_e,coord_n)
                        ws.cell(current_line_idx,4).style = 'Hyperlink'

                    nb_errors += 1
                else:
                    # delete rows that are not in current commune
                    if delete_row_idx is None:
                        delete_row_idx = current_line_idx
                    delete_row_amount += 1
                

                current_line_idx += 1
                k += 1

            # remove last errors
            if delete_row_idx is not None and delete_row_amount > 0:
                ws.delete_rows(delete_row_idx, amount=delete_row_amount)
            
            feedback_commune['Nb erreurs liste ' + str(i+1)]: nb_errors


        # get issue 22
        wb = _getIssue22OfCommune(wb, commune_id, issue22_list)


        wb.save(feedback_commune_filepath)

        return (feedback_commune_filepath, feedback_commune)


def createDBSession(env_substring):
    engine = create_engine("postgresql+psycopg2://{}:{}@{}:{}/{}".format(
        os.environ[env_substring + '_USERNAME'],
        os.environ[env_substring + '_PASSWORD'],
        os.environ[env_substring + '_HOST'],
        os.environ[env_substring + '_PORT'],
        os.environ[env_substring + '_DATABASE']
    ))

    # create session
    Session = sessionmaker()
    Session.configure(bind=engine)
    return Session()

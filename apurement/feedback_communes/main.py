from dotenv import load_dotenv
from openpyxl import Workbook,  load_workbook

import sys

sys.path.insert(0, r'..\..\utils')
import utils


# dev imports
import os
from datetime import datetime



load_dotenv(r'..\..\.env')



if __name__ == '__main__':
    # get feedback for canton de Neuchâtel
    communes_ofs = utils.loadCommunesOFS()

    # prepare working directory
    today = datetime.strftime(datetime.now(), '%Y%m%d')
    feedback_commune_path = os.path.join(os.environ['FEEDBACK_COMMUNES_WORKING_DIR'], today)
    utils.cleanWorkingDirectory(path=feedback_commune_path)
    os.makedirs(feedback_commune_path)

    # get lists
    feedback_canton_filepath = utils.downloadListeCantonNeuchatel(path=feedback_commune_path)
    (issue22_list, issue22_canton_filepath) = utils.downloadIssue22CantonNeuchatel(path=feedback_commune_path)
    issue_solution = utils.get_issue_solution(os.environ['STATIC_PATH_ISSUE_SOLUTION_SGRF'])

    # go through each commune and create an excel with errors if any
    for commune_id in communes_ofs.keys():
        print(commune_id, communes_ofs[commune_id])
        (feedback_commune_filepath, feedback_commune) = utils.generateCommuneErrorFile(commune_id, communes_ofs[commune_id], feedback_canton_filepath, issue22_list, issue_solution, today)


    # finally clean temp folder
    # utils.cleanWorkingDirectory(path=feedback_commune_path)




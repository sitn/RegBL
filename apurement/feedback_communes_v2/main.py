from openpyxl import load_workbook, Workbook
from dotenv import load_dotenv

import sys

sys.path.insert(0, r'..\..\utils')
import utils


# dev imports
import os
from datetime import datetime
import shutil
import xlwings as xw


load_dotenv(r'..\..\.env')





if __name__ == '__main__':

    # get feedback for canton de Neuchâtel
    communes_ofs = utils.loadCommunesOFS()

    # prepare working directory
    tmp_path = os.environ['RAPPORT_COMMUNES_TMP_DIR']

    # go through each commune and create an excel with errors if any
    for commune_id in communes_ofs.keys():
        if commune_id != 6404:
            continue

        print(commune_id, communes_ofs[commune_id])

        utils.cleanWorkingDirectory(path=tmp_path)
        os.makedirs(tmp_path)

        #####################
        #  INFOS GENERALES
        #####################


        #####################
        #  LISTE 1
        #####################


        #####################
        #  LISTE 2
        #####################


        #####################
        #  LISTE 3
        #####################


        #####################
        #  LISTE 4
        #####################


        #####################
        #  LISTE 5
        #####################


        #####################
        #  LISTE 6
        #####################



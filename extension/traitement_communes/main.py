import openpyxl
import os
import re
import time
from shutil import copy2

from dotenv import load_dotenv

load_dotenv(r'..\..\.env')


input_path = os.environ['TRAITEMENT_COMMUNES_INPUT_PATH']
tmp_path = os.environ['TRAITEMENT_COMMUNES_TMP_PATH']
output_path = os.environ['TRAITEMENT_COMMUNES_OUTPUT_PATH']
fme_path = os.environ['TRAITEMENT_COMMUNES_FME_PATH']
defaultExcelSheetName = '6434'



def changeExcelSheetName(filepath, oldSheetname, newSheetname):
    ss=openpyxl.load_workbook(filepath)
    ss_sheet = ss[oldSheetname]
    ss_sheet.title = newSheetname
    ss_sheet.protection.sheet = False
    ss.save(filepath)
    return



if __name__ == "__main__":
    for root, dirs, files in os.walk(input_path):
        for filename in files:
            if re.match(r'6[4,5]{1}[\d]{2}.*\.xlsx', filename) is not None:
                no_commune = filename.split('.')[0].split('_')[0]
                print(no_commune)
                input_filepath = os.path.join(root, filename)
                tmp_filepath = os.path.join(tmp_path, filename)
                output_filepath = os.path.join(output_path, filename)

                # copy file to tmp (working) and output (destination) directory
                if os.path.exists(tmp_filepath):
                    os.remove(tmp_filepath)
                copy2(input_filepath, tmp_filepath)

                # rename excel sheetname for fme which is unfortunately not able to pass sheet name as an argument ! ****
                changeExcelSheetName(tmp_filepath, no_commune, defaultExcelSheetName)

                # copy file
                if os.path.exists(output_filepath):
                    os.remove(output_filepath)
                copy2(tmp_filepath, output_filepath)

                time.sleep(0.5)

                # fme script
                print('\ncommand: fme "{}" --GDENR "{}" --SourceDataset_XLSXR_8 "{}" --DestDataset_XLSXW "{}"\n'.format(fme_path, no_commune, tmp_filepath, output_filepath))
                os.system('fme "{}" --GDENR "{}" --SourceDataset_XLSXR_8 "{}" --DestDataset_XLSXW "{}"'.format(fme_path, no_commune, tmp_filepath, output_filepath))

                # rename excel sheetname with correct commune number
                changeExcelSheetName(output_filepath, defaultExcelSheetName, no_commune)                

                # remove file in tmp directory
                os.remove(tmp_filepath)

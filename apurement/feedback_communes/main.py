from dotenv import load_dotenv

import csv
import sys

sys.path.insert(0, r"..\..\utils")
import utils


# dev imports
import os
from datetime import datetime


load_dotenv(r"..\..\.env")


if __name__ == "__main__":
    # get arguments
    environ_input = input("Hyperliens [inter | intra (default)]: ")
    environ = "INTER" if environ_input == "inter" else "INTRA"
    print(environ)

    batprojtreat_input = input("Filtrer les batiments projetés [oui (défaut) | non]: ")
    batprojtreat = False if batprojtreat_input == "non" else True
    print("Oui" if batprojtreat is True else "Non")

    egidextfilter_input = input("Filtrer les EGID > 500000000 [oui (défaut) | non]: ")
    egidextfilter = False if egidextfilter_input == "non" else True
    print("Oui" if egidextfilter is True else "Non")

    # get feedback for canton de Neuchâtel
    communes_ofs = utils.loadCommunesOFS()

    # prepare working directory
    today = datetime.strftime(datetime.now(), "%Y%m%d")
    feedback_commune_path = os.path.join(os.environ["FEEDBACK_COMMUNES_WORKING_DIR"], today)
    utils.cleanWorkingDirectory(path=feedback_commune_path)
    os.makedirs(feedback_commune_path)

    # get lists
    feedback_canton_filepath = utils.downloadListeCantonNeuchatel(path=feedback_commune_path, batprojtreat=batprojtreat)
    (issue22_list, issue22_canton_filepath) = utils.downloadIssue22CantonNeuchatel(path=feedback_commune_path)

    path_issue_solution = os.environ["RAPPORT_ANALYZER_ISSUE_SOLUTION_PATH"]
    issue_solution = utils.get_issue_solution(path_issue_solution)

    # write feedback
    feedback_filepath = os.path.join(feedback_commune_path, f"{today}-feedback.csv")

    with open(feedback_filepath, "w", newline="") as csvfile:
        csv_fieldnames = [f"Liste_{i+1}" for i in range(6)]
        csv_fieldnames.insert(0, "Commune")
        csv_fieldnames.append("Issue_22")
        # csv_fieldnames = ["Commune", [f"Liste_{i+1}" for i in range(6)], "Issue_22"]
        writer = csv.DictWriter(csvfile, fieldnames=csv_fieldnames, delimiter=";")
        writer.writeheader()

        # go through each commune and create an excel with errors if any
        for commune_id in communes_ofs.keys():
            # if commune_id not in [6487]:
            # if commune_id not in [6487, 6417, 6458, 6416]:
            #     continue
            print(commune_id, communes_ofs[commune_id])
            (feedback_commune_filepath, feedback_commune, nb_errors_by_list) = utils.generateCommuneErrorFile(commune_id, communes_ofs[commune_id], feedback_canton_filepath, issue22_list, issue_solution, today, environ, egidextfilter, log=False)

            csvfile = writer.writerow(nb_errors_by_list)

        # do the same for the canton
        print("Canton de Neuchâtel")
        (feedback_commune_filepath, feedback_commune, nb_errors_by_list) = utils.generateCantonErrorFile(feedback_canton_filepath, issue_solution, today=datetime.strftime(datetime.now(), "%Y%m%d"), environ="INTER", log=False)
        csvfile = writer.writerow(nb_errors_by_list)

    # print output path to copy and paste in browser
    print("Les fichiers se trouvent ici: ", feedback_commune_path)

    # finally clean temp folder
    # utils.cleanWorkingDirectory(path=feedback_commune_path)

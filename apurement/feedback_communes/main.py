from dotenv import load_dotenv

import csv
import os
import sys
from datetime import datetime


sys.path.insert(0, r".\utils")
import utils


load_dotenv(override=True)


if __name__ == "__main__":
    if len(sys.argv) == 1:
        # get arguments
        environ_input = input("internet ou intranet [INTER | INTRA (défaut)]: ")
        environ = "INTER" if environ_input == "inter" else "INTRA"
        print(environ)

        batprojtreat_input = input("Filtrer les batiments projetés [oui (défaut) | non]: ")
        batprojtreat = False if batprojtreat_input == "non" else True
        print("Oui" if batprojtreat is True else "Non")

        egidextfilter_input = input("Filtrer les EGID > 500000000 [oui | non (défaut) ]: ")
        egidextfilter = True if egidextfilter_input == "oui" else False
        print("Oui" if egidextfilter is True else "Non")

        egidwhitelistfilter_input = input("Filtrer les EGID avec la whiteliste [oui (défaut) | non ]: ")
        egidwhitelistfilter = False if egidwhitelistfilter_input == "non" else True
        print("Oui" if egidwhitelistfilter is True else "Non")

    elif len(sys.argv) > 1:
        if "--auto" in sys.argv:
            environ = "INTER"
            batprojtreat = True
            egidextfilter = False
            egidwhitelistfilter = True
            print("internet ou intranet [INTER | INTRA (défaut)]: " + environ)
            print("Filtrer les batiments projetés [oui (défaut) | non]: " + ("Oui" if batprojtreat is True else "Non"))
            print("Filtrer les EGID > 500000000 [oui | non (défaut)]: " + ("Oui" if egidextfilter is True else "Non"))
            print("Filtrer les EGID avec la whiteliste [oui (défaut) | non ]: " + ("Oui" if egidwhitelistfilter is True else "Non"))

    else:
        raise ValueError("Invalid parameters ")

    # get feedback for canton de Neuchâtel
    # prepare working directory
    today = datetime.strftime(datetime.now(), "%Y%m%d")
    feedback_commune_path = os.path.join(os.environ["FEEDBACK_COMMUNES_WORKING_DIR"], today)
    utils.cleanWorkingDirectory(path=feedback_commune_path)
    os.makedirs(feedback_commune_path)

    # get lists
    feedback_canton_filepath = utils.downloadListeCantonNeuchatel(path=feedback_commune_path, batprojtreat=batprojtreat)
    communes_ofs = utils.loadCommunesOFS(feedback_canton_filepath)
    (issue22_list, issue22_canton_filepath) = utils.downloadIssue22CantonNeuchatel(path=feedback_commune_path)

    path_issue_solution = os.environ["RAPPORT_ANALYZER_ISSUE_SOLUTION_PATH"]
    issue_solution = utils.get_issue_solution(path_issue_solution)

    # get whiteliste
    whitelist_path = os.environ["FEEDBACK_COMMUNES_EGID_WHITELIST_CONTROLS_PATH"]
    egid_whitelist_sgrf = utils.getEGIDWhitelistSGRF(path=whitelist_path) if egidwhitelistfilter is True else []

    # write feedback
    feedback_filepath = os.path.join(feedback_commune_path, f"{today}-feedback.csv")

    with open(feedback_filepath, "w", newline="") as csvfile:
        csv_fieldnames = [f"Liste_{i+1}" for i in range(6)]
        csv_fieldnames.insert(0, "Commune_id")
        csv_fieldnames.insert(1, "Commune")
        csv_fieldnames.append("Issue_22")
        # csv_fieldnames = ["Commune", [f"Liste_{i+1}" for i in range(6)], "Issue_22"]
        writer = csv.DictWriter(csvfile, fieldnames=csv_fieldnames, delimiter=";")
        writer.writeheader()

        # go through each commune and create an excel with errors if any
        for commune_id in communes_ofs.keys():
            # if commune_id not in [6421]:
            # if commune_id not in [6421, 6417, 6458, 6416]:
                # continue
            print(commune_id, communes_ofs[commune_id])
            result = utils.generateCommuneErrorFile(commune_id, communes_ofs[commune_id], feedback_canton_filepath, issue22_list, issue_solution, today, environ, egidextfilter, log=False, egidwhitelist=egid_whitelist_sgrf, whitelist_path=whitelist_path)

            if result is not None:
                (feedback_commune_filepath, feedback_commune, nb_errors_by_list) = result
                csvfile = writer.writerow(nb_errors_by_list)

        # # do the same for the canton
        # print("Canton de Neuchâtel")
        # (feedback_commune_filepath, feedback_commune, nb_errors_by_list) = utils.generateCantonErrorFile(feedback_canton_filepath, issue_solution, today=datetime.strftime(datetime.now(), "%Y%m%d"), environ="INTRA", log=False, egidwhitelist=egid_whitelist_sgrf)
        # csvfile = writer.writerow(nb_errors_by_list)

    # print output path to copy and paste in browser
    print("Les fichiers se trouvent ici: ", feedback_commune_path)

    # finally clean temp folder
    # utils.cleanWorkingDirectory(path=feedback_commune_path)

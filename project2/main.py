import subprocess
import os
import glob

if __name__ == "__main__":
    file_path = os.path.dirname(__file__)
    os.chdir(file_path)

    folder_path_pdf = "{}/fake_confirmations/pdfs".format(file_path)
    for file in glob.glob(f"{folder_path_pdf}/*.pdf", recursive=True):
        os.remove(file)

    for script in ["create_pdfs.py", "rpa.py", "procesmining.py"]:
        subprocess.run(["python", script])
        f"{script} is finished!"

    folder_path_xlsx = "{}/fake_confirmations/xlsx".format(file_path)
    for file in glob.glob(f"{folder_path_xlsx}/*.xlsx", recursive=True):
        os.remove(file)

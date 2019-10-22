import os

from pptx import *

from utils.pptx_converter import PPTXConverter
from utils.translate_dict import TranslateDict

file_name = "2019 APM 9-39.pptx"
pptx_path = os.getcwd() + os.sep + "resources" + os.sep + file_name
db_path = os.getcwd() + os.sep + "sqlite.db"
csv_path = os.getcwd() + os.sep + "resources" + os.sep + "dict.csv"

pr_dict = {"Агрономическое совещание": "Agronomical meeting",
           "Персонал": "Personnel",
           "АКТИВНАЯ ЧИСЛЕННОСТЬ": "ACTIVE PERSONNEL",
           "ПЕРСОНАЛ": "Personnel"}


def main():
    tr_dict = TranslateDict(db_path)
    tr_dict.create_base_table(csv_path)
    pr_dict = tr_dict.get_dict()
    conv = PPTXConverter(pptx_path)
    conv.translate(pr_dict)
    conv.save(os.getcwd()+os.sep+"new.pptx")

    # prs.save(os.getcwd()+os.sep+"ready.pptx")


if __name__ == "__main__":
    main()

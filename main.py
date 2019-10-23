import os

from flask import Flask, render_template, request, send_file

from utils.pptx_converter import PPTXConverter
from utils.translate_dict import TranslateDict

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'resources'

file_name = "2019 APM 9-39.pptx"
pptx_path = os.getcwd() + os.sep + "resources" + os.sep + file_name
db_path = os.getcwd() + os.sep + "sqlite.db"
csv_path = os.getcwd() + os.sep + "resources" + os.sep + "dict.csv"

pr_dict = {"Агрономическое совещание": "Agronomical meeting",
           "Персонал": "Personnel",
           "АКТИВНАЯ ЧИСЛЕННОСТЬ": "ACTIVE PERSONNEL",
           "ПЕРСОНАЛ": "Personnel"}


def main():
    app.run(host="127.0.0.1")
    pass
    # tr_dict = TranslateDict(db_path)
    # pr_dict = tr_dict.get_dict()
    # conv = PPTXConverter(pptx_path)
    # conv.translate(pr_dict)
    # conv.save(os.getcwd()+os.sep+"new.pptx")

@app.route('/')
def translate_page():
    return render_template('index.html')

@app.route('/translate', methods=["POST"])
def translate():
    if len(request.files) > 0:
        file = request.files["pptx"]
        if file.filename.split('.')[1] == 'pptx':
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))

            tr_dict = TranslateDict(db_path)
            pr_dict = tr_dict.get_dict()
            conv = PPTXConverter(os.path.join(app.config['UPLOAD_FOLDER'], file.filename))
            conv.translate(pr_dict)
            conv.save(os.getcwd()+os.sep+f"EN_{file.filename}")
            return send_file(os.getcwd()+os.sep+"EN_"+file.filename)
    else:
        return None


if __name__ == "__main__":
    main()
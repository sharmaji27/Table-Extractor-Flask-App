import os
from werkzeug.utils import secure_filename
from flask import Flask,request,render_template
import pdfplumber
import pandas as pd


UPLOAD_FOLDER = './static/uploads'
ALLOWED_EXTENSIONS = set(['pdf'])

app = Flask(__name__)
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.secret_key = "secret key"

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                table = pd.DataFrame(table)
                table.columns = table.iloc[0]
                table.drop(0,inplace=True)
                tables.append(table)
    return tables

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/extract_table',methods=['POST'])
def extract_table():
    file = request.files['file']
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        tables = extract(UPLOAD_FOLDER+'/'+filename)

        with pd.ExcelWriter('extracted_tables.xlsx') as writer:
            for i, df in enumerate(tables):
                df.to_excel(writer, sheet_name=str(i+1), index=False)

        tables = [x.to_html(index=False) for x in tables]

        return render_template('home.html',org_img_name=filename,tables=tables,ntables=len(tables))


if __name__ == '__main__':
    app.run(debug=True)
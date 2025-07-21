# hppdauto_web/app.py
from flask import Flask, render_template, request, send_file
import os
from datetime import datetime
from hppdauto import run_hppd_comparison_for_date
from werkzeug.utils import secure_filename
import zipfile

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        template_zip = request.files.get("template_zip")
        report_zip = request.files.get("report_zip")
        date_str = request.form.get("date")
        date = datetime.strptime(date_str, "%Y-%m-%d")

        # Save and extract the templates zip
        template_path = os.path.join(UPLOAD_FOLDER, "templates")
        os.makedirs(template_path, exist_ok=True)
        with zipfile.ZipFile(template_zip, 'r') as zip_ref:
            zip_ref.extractall(template_path)

        # Save and extract the reports zip
        report_path = os.path.join(UPLOAD_FOLDER, "reports")
        os.makedirs(report_path, exist_ok=True)
        with zipfile.ZipFile(report_zip, 'r') as zip_ref:
            zip_ref.extractall(report_path)

        output_path = run_hppd_comparison_for_date(
            template_path,
            report_path,
            date.strftime("%Y-%m-%d"),
            UPLOAD_FOLDER
        )
        return send_file(output_path, as_attachment=True)

    return render_template("index_zip.html")

if __name__ == "__main__":
    app.run(debug=True)

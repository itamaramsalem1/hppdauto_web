# hppdauto_web/app.py
from flask import Flask, render_template, request, send_file
import os
from datetime import datetime
from hppdauto import run_hppd_comparison_for_date

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        template_files = request.files.getlist("templates")
        report_files = request.files.getlist("reports")
        date_str = request.form["date"]
        date = datetime.strptime(date_str, "%Y-%m-%d")

        templates_path = os.path.join(UPLOAD_FOLDER, "templates")
        reports_path = os.path.join(UPLOAD_FOLDER, "reports")
        os.makedirs(templates_path, exist_ok=True)
        os.makedirs(reports_path, exist_ok=True)

        for f in template_files:
            f.save(os.path.join(templates_path, f.filename))
        for f in report_files:
            f.save(os.path.join(reports_path, f.filename))

        try:
            output_path = run_hppd_comparison_for_date(templates_path, reports_path, date, UPLOAD_FOLDER)
            return send_file(output_path, as_attachment=True)
        except Exception as e:
            return f"❌ Error during processing: {e}", 500

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)

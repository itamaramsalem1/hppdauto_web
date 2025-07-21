from flask import Flask, render_template, request, send_file
import os
import zipfile
from datetime import datetime
from hppdauto import run_hppd_comparison_for_date

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        zip_file = request.files.get("data_zip")
        date_str = request.form.get("date")

        if not zip_file or not date_str:
            return "Missing ZIP file or date.", 400

        try:
            date = datetime.strptime(date_str, "%Y-%m-%d")
        except ValueError:
            return "Invalid date format.", 400

        zip_path = os.path.join(UPLOAD_FOLDER, "uploaded_data.zip")
        zip_file.save(zip_path)

        # Clean extract path each time
        extract_path = os.path.join(UPLOAD_FOLDER, "unzipped")
        if os.path.exists(extract_path):
            import shutil
            shutil.rmtree(extract_path)
        os.makedirs(extract_path, exist_ok=True)

        # Unzip into expected folders
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)

        templates_path = os.path.join(extract_path, "templates")
        reports_path = os.path.join(extract_path, "reports")

        # Run the comparison
        output_path = run_hppd_comparison_for_date(
            templates_path,
            reports_path,
            date.strftime("%Y-%m-%d"),
            UPLOAD_FOLDER
        )

        return send_file(output_path, as_attachment=True)

    return render_template("index_zip.html")

if __name__ == "__main__":
    app.run(debug=True)

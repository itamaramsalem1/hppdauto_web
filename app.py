from flask import Flask, render_template, request, send_file, jsonify
import os
import shutil
from datetime import datetime
from hppdauto import run_hppd_comparison_for_date
from werkzeug.utils import secure_filename
import zipfile
import tempfile
from uuid import uuid4

app = Flask(__name__)
progress_store = {}  # In-memory store for progress tracking

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        try:
            template_zip = request.files.get("template_zip")
            report_zip = request.files.get("report_zip")
            date_str = request.form.get("date")
            progress_id = request.form.get("progress_id")

            if not template_zip or not report_zip or not date_str or not progress_id:
                return "Missing required files, date, or progress ID", 400

            date = datetime.strptime(date_str, "%Y-%m-%d")

            # Initialize progress
            progress_store[progress_id] = {"percent": 0, "status": "Initializing..."}

            # Create temporary directories
            with tempfile.TemporaryDirectory() as temp_dir:
                upload_folder = os.path.join(temp_dir, "uploads")
                os.makedirs(upload_folder, exist_ok=True)

                # Save and extract the templates zip
                template_path = os.path.join(upload_folder, "templates")
                os.makedirs(template_path, exist_ok=True)

                try:
                    with zipfile.ZipFile(template_zip, 'r') as zip_ref:
                        zip_ref.extractall(template_path)
                except Exception as e:
                    return f"Error extracting template zip: {str(e)}", 400

                # Save and extract the reports zip
                report_path = os.path.join(upload_folder, "reports")
                os.makedirs(report_path, exist_ok=True)

                try:
                    with zipfile.ZipFile(report_zip, 'r') as zip_ref:
                        zip_ref.extractall(report_path)
                except Exception as e:
                    return f"Error extracting report zip: {str(e)}", 400

                # Run the analysis
                try:
                    def update_progress(pct, msg):
                        progress_store[progress_id] = {"percent": pct, "status": msg}

                    output_path = run_hppd_comparison_for_date(
                        template_path,
                        report_path,
                        date.strftime("%Y-%m-%d"),
                        upload_folder,
                        progress_callback=update_progress
                    )

                    return send_file(output_path, as_attachment=True)
                except Exception as e:
                    return f"Error processing files: {str(e)}", 500

        except Exception as e:
            return f"Unexpected error: {str(e)}", 500

    return render_template("index_zip.html")

@app.route("/progress/<progress_id>")
def get_progress(progress_id):
    data = progress_store.get(progress_id, {"percent": 0, "status": "Not started"})
    return jsonify(data)

@app.errorhandler(413)
def too_large(e):
    return "File is too large", 413

if __name__ == "__main__":
    app.run(debug=True)

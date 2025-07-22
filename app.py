from flask import Flask, render_template, request, send_file, jsonify
import os
import shutil
from datetime import datetime
from hppdauto import run_hppd_comparison_for_date
from werkzeug.utils import secure_filename
import zipfile
import tempfile
from uuid import uuid4
import threading

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

            print("DEBUG: Upload check")
            print("template_zip:", template_zip)
            print("report_zip:", report_zip)
            print("date_str:", date_str)
            print("progress_id:", progress_id)

            if not template_zip or not report_zip or not date_str or not progress_id:
                return jsonify({"error": "Missing required files, date, or progress ID"}), 400

            date = datetime.strptime(date_str, "%Y-%m-%d")

            progress_store[progress_id] = {"percent": 0, "status": "Initializing...", "completed": False, "file_path": None}

            def process_files():
                try:
                    def update_progress(pct, msg):
                        print(f"[{progress_id}] {pct}% - {msg}")
                        progress_store[progress_id] = {"percent": pct, "status": msg, "completed": False, "file_path": None}

                    print(f"[{progress_id}] Preparing temp directories...")
                    with tempfile.TemporaryDirectory() as temp_dir:
                        upload_folder = os.path.join(temp_dir, "uploads")
                        os.makedirs(upload_folder, exist_ok=True)

                        template_path = os.path.join(upload_folder, "templates")
                        os.makedirs(template_path, exist_ok=True)
                        try:
                            with zipfile.ZipFile(template_zip, 'r') as zip_ref:
                                zip_ref.extractall(template_path)
                        except Exception as e:
                            progress_store[progress_id] = {"percent": 0, "status": f"Error extracting template zip: {str(e)}", "completed": True, "file_path": None}
                            print(f"[{progress_id}] Template extraction failed: {e}")
                            return

                        report_path = os.path.join(upload_folder, "reports")
                        os.makedirs(report_path, exist_ok=True)
                        try:
                            with zipfile.ZipFile(report_zip, 'r') as zip_ref:
                                zip_ref.extractall(report_path)
                        except Exception as e:
                            progress_store[progress_id] = {"percent": 0, "status": f"Error extracting report zip: {str(e)}", "completed": True, "file_path": None}
                            print(f"[{progress_id}] Report extraction failed: {e}")
                            return

                        try:
                            print(f"[{progress_id}] Starting run_hppd_comparison_for_date")
                            output_path = run_hppd_comparison_for_date(
                                template_path,
                                report_path,
                                date.strftime("%Y-%m-%d"),
                                upload_folder,
                                progress_callback=update_progress
                            )
                            print(f"[{progress_id}] run_hppd_comparison_for_date finished")

                            permanent_path = os.path.join("/tmp", f"hppd_output_{progress_id}.xlsx")
                            shutil.copy2(output_path, permanent_path)

                            progress_store[progress_id] = {
                                "percent": 100,
                                "status": "✅ Analysis complete! Download ready.",
                                "completed": True,
                                "file_path": permanent_path
                            }
                            print(f"[{progress_id}] File saved to {permanent_path}")

                        except Exception as e:
                            progress_store[progress_id] = {
                                "percent": 0,
                                "status": f"Error processing files: {str(e)}",
                                "completed": True,
                                "file_path": None
                            }
                            print(f"[{progress_id}] ERROR in processing: {e}")

                except Exception as e:
                    progress_store[progress_id] = {
                        "percent": 0,
                        "status": f"Unexpected error: {str(e)}",
                        "completed": True,
                        "file_path": None
                    }
                    print(f"[{progress_id}] UNEXPECTED error: {e}")

            print(f"[{progress_id}] Launching thread...")
            thread = threading.Thread(target=process_files)
            thread.start()
            print(f"[{progress_id}] Thread launched successfully.")

            return jsonify({"status": "started", "progress_id": progress_id})

        except Exception as e:
            print("FATAL ERROR in / route:", str(e))
            return jsonify({"error": f"Unexpected server error: {str(e)}"}), 500

    return render_template("index_zip.html")

@app.route("/progress/<progress_id>")
def get_progress(progress_id):
    data = progress_store.get(progress_id, {"percent": 0, "status": "Not started", "completed": False, "file_path": None})
    return jsonify(data)

@app.route("/download/<progress_id>")
def download_file(progress_id):
    data = progress_store.get(progress_id, {})
    file_path = data.get("file_path")

    if file_path and os.path.exists(file_path):
        return send_file(file_path, as_attachment=True, download_name="HPPD_Comparison_Output.xlsx")
    else:
        return "File not found", 404

@app.errorhandler(413)
def too_large(e):
    return "File is too large", 413

if __name__ == "__main__":
    app.run(debug=True)

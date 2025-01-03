from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def process_excel(file_path):
    # Load data as raw list
    raw_data = pd.read_excel(file_path, header=None).values.tolist()
    
    # Find the header row dynamically
    header_row_index = -1
    for i, row in enumerate(raw_data):
        if any("nomor id jaminan" in str(cell).lower() for cell in row):
            header_row_index = i
            break
    
    if header_row_index == -1:
        raise ValueError("Header yang mengandung 'Nomor ID Jaminan' tidak ditemukan.")
    
    # Create DataFrame using the found header
    headers = raw_data[header_row_index]
    data = pd.DataFrame(raw_data[header_row_index + 1:], columns=headers)
    headers_lower = data.columns.str.lower()
    
    # Find necessary columns
    try:
        idx_nama_rs = headers_lower.get_loc("nama rumah sakit")
        idx_status_pembayaran = headers_lower.get_loc("status pembayaran")
        idx_status_verifikasi = headers_lower.get_loc("status verifikasi")
        idx_gl_status = headers_lower.get_loc("gl status")
    except KeyError as e:
        raise ValueError(f"Kolom tidak ditemukan: {e}")
    
    # Filter unpaid rows and GL Status Active
    filtered_data = data[
        (data.iloc[:, idx_status_pembayaran].str.lower() == "unpaid") & 
        (data.iloc[:, idx_gl_status].str.lower() == "active")
    ]
    
    # Rekap data
    result = {}
    for _, row in filtered_data.iterrows():
        nama_rs = row[idx_nama_rs]
        status_verifikasi = str(row[idx_status_verifikasi]).lower()
        
        if nama_rs not in result:
            result[nama_rs] = {
                "Done": 0,
                "Revision": 0,
                "New": 0,
                "Waiting First Layer Verification": 0,
                "Lain-lain": 0,
                "Total": 0
            }
        
        if "done" in status_verifikasi or "resend" in status_verifikasi:
            result[nama_rs]["Done"] += 1
        elif status_verifikasi == "revision":
            result[nama_rs]["Revision"] += 1
        elif status_verifikasi in ["new", "draft"]:
            result[nama_rs]["New"] += 1
        elif status_verifikasi == "waiting first layer verification":
            result[nama_rs]["Waiting First Layer Verification"] += 1
        else:
            result[nama_rs]["Lain-lain"] += 1
        
        result[nama_rs]["Total"] += 1
    
    # Create DataFrame for summary
    summary_data = [
        [nama_rs, *counts.values()] for nama_rs, counts in result.items()
    ]
    summary_df = pd.DataFrame(summary_data, columns=["Nama Rumah Sakit", "Done", "Revision", "New", 
                                                     "Waiting First Layer Verification", "Lain-lain", "Total"])
    # Sort by 'Nama Rumah Sakit' in ascending order
    summary_df = summary_df.sort_values(by="Nama Rumah Sakit", ascending=True)
    # Add grand total
    grand_total = summary_df.iloc[:, 1:].sum(axis=0)
    grand_total["Nama Rumah Sakit"] = "Total"
    summary_df = pd.concat([summary_df, pd.DataFrame([grand_total])], ignore_index=True)
    
    return summary_df


import pandas as pd
from datetime import datetime

def calculate_days(file_path):
    import pandas as pd
    
    # Load data as raw list
    raw_data = pd.read_excel(file_path, header=None).values.tolist()

    # Find the header row dynamically
    header_row_index = -1
    for i, row in enumerate(raw_data):
        if any("nomor id jaminan" in str(cell).lower() for cell in row):
            header_row_index = i
            break

    if header_row_index == -1:
        raise ValueError("Header yang mengandung 'Nomor ID Jaminan' tidak ditemukan.")

    # Create DataFrame using the found header
    headers = raw_data[header_row_index]
    data = pd.DataFrame(raw_data[header_row_index + 1:], columns=headers)
    headers_lower = data.columns.str.lower()

    # Find necessary columns
    try:
        idx_nama_rs = headers_lower.get_loc("nama rumah sakit")
        idx_status_pembayaran = headers_lower.get_loc("status pembayaran")
        idx_status_verifikasi = headers_lower.get_loc("status verifikasi")
        idx_gl_status = headers_lower.get_loc("gl status")
        idx_tanggal_verifikasi = headers_lower.get_loc("tanggal verifikasi")
    except KeyError as e:
        raise ValueError(f"Kolom tidak ditemukan: {e}")

    # Filter unpaid rows
    data = data[data.iloc[:, idx_status_pembayaran].str.lower() == "unpaid"]

    # Filter GL Status as "Active"
    data = data[data.iloc[:, idx_gl_status].str.lower() == "active"]

    # Filter Status Verifikasi as "done" or contains "resend"
    data = data[
        (data.iloc[:, idx_status_verifikasi].str.lower() == "done") |
        (data.iloc[:, idx_status_verifikasi].str.contains("resend", case=False, na=False))
    ]

    # Filter rows with valid "tanggal verifikasi"
    data = data[~data.iloc[:, idx_tanggal_verifikasi].isna()]
    data = data[data.iloc[:, idx_tanggal_verifikasi] != "-"]

    # Convert "tanggal verifikasi" to datetime with specific format
    data["Tanggal_Verifikasi"] = pd.to_datetime(
        data.iloc[:, idx_tanggal_verifikasi],
        format="%d-%m-%Y",  # Format yang sesuai dengan data Anda
        errors='coerce'  # Tetap menangani nilai yang tidak valid
    )

    # Filter rows with valid dates
    data = data[data["Tanggal_Verifikasi"].notna()]  # Hapus tanggal yang tidak valid

    # Normalize to date only
    data["Tanggal_Verifikasi"] = data["Tanggal_Verifikasi"].dt.normalize()
    today = pd.Timestamp.now().normalize()  # Normalize to today's date

    # Calculate "Lama_Hari"
    data["Lama_Hari"] = (today - data["Tanggal_Verifikasi"]).dt.days

    # Grouping by Rumah Sakit and Binning Lama_Hari
    bins = [-1, 9, 14, float('inf')]
    labels = ["<10", "10-14", ">14"]

    # Initialize dictionary for grouped data
    grouped_data = {}

    # Bin data for each Rumah Sakit
    for _, row in data.iterrows():
        nama_rs = row[idx_nama_rs]
        days = row["Lama_Hari"]

        # Initialize the dictionary for new rumah sakit
        if nama_rs not in grouped_data:
            grouped_data[nama_rs] = {label: 0 for label in labels}

        # Assign bin based on Lama_Hari
        if bins[0] < days <= bins[1]:
            grouped_data[nama_rs]["<10"] += 1
        elif bins[1] < days <= bins[2]:
            grouped_data[nama_rs]["10-14"] += 1
        else:
            grouped_data[nama_rs][">14"] += 1

    # Generate Grouped Summary
    summary_data = []
    for nama_rs, counts in grouped_data.items():
        summary_row = [nama_rs] + list(counts.values())
        summary_row.append(sum(counts.values()))  # Grand Total
        summary_data.append(summary_row)

    # Calculate Grand Total for each column (sum across all Rumah Sakit)
    grand_totals = [sum([row[i] for row in summary_data]) for i in range(1, len(labels) + 2)]
    grand_totals.insert(0, "Total")

    # Add grand totals to the summary data
    summary_data.append(grand_totals)

    grouped_summary_df = pd.DataFrame(summary_data, columns=["Nama Rumah Sakit"] + labels + ["Grand Total"])
    
    # Sort by Nama Rumah Sakit in ascending order
    grouped_summary_df = grouped_summary_df.sort_values("Nama Rumah Sakit", ascending=True)

    # Create Detailed Data with calculated "Lama_Hari"
    detailed_data = data[[headers[idx_nama_rs], headers[idx_tanggal_verifikasi], "Lama_Hari"]]
    detailed_data = detailed_data.rename(columns={
        headers[idx_nama_rs]: "Nama Rumah Sakit",
        headers[idx_tanggal_verifikasi]: "Tanggal Verifikasi",
        "Lama_Hari": "Lama Hari"
    })
    # Sort Detailed Data by "Nama Rumah Sakit" in ascending order
    detailed_data = detailed_data.sort_values("Nama Rumah Sakit", ascending=True)

    return grouped_summary_df, detailed_data








def process_and_save(file_path, result_file):
    summary_df = process_excel(file_path)
    grouped_summary, detailed_data = calculate_days(file_path)
    
    with pd.ExcelWriter(result_file, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        grouped_summary.to_excel(writer, index=False, sheet_name="Grouped Summary")
        detailed_data.to_excel(writer, index=False, sheet_name="Detailed Data")


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file:
            return render_template("index.html", error="Harap unggah file.")
        
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        try:
            result_file = os.path.join(app.config['RESULT_FOLDER'], "summary.xlsx")
            process_and_save(file_path, result_file)
            os.remove(file_path)
            return send_file(result_file, as_attachment=True)
        except Exception as e:
            if os.path.exists(file_path):
                os.remove(file_path)
            return render_template("index.html", error=str(e))
    
    return render_template("index.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5500)

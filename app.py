from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
from collections import Counter

app = Flask(__name__)
CORS(app)

@app.route("/api/documents/upload", methods=["POST"])
def upload_documents_excel():
    if "file" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo."}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "El archivo no tiene nombre."}), 400

    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xlsm")):
        return jsonify({"error": "Solo se permiten archivos Excel."}), 400

    workbook = openpyxl.load_workbook(file, data_only=True)
    sheet = workbook.active

    headers = [cell.value for cell in sheet[1]]
    rows = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(value is None for value in row):
            continue

        row_dict = {}
        for i in range(len(headers)):
            column_name = headers[i] if headers[i] is not None else f"col_{i}"
            row_dict[column_name] = row[i]

        rows.append(row_dict)

    status_counter = Counter()

    for row in rows:
        status = row.get("Estatus", "Sin estatus")
        status_counter[str(status)] += 1

    distribution = [
        {"name": key, "value": value}
        for key, value in status_counter.items()
    ]

    return jsonify({
        "totalDocuments": len(rows),
        "distribution": distribution,
        "rows": rows
    })

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
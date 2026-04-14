from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl

app = Flask(__name__)
CORS(app)


def normalize_text(value):
    if value is None:
        return ""
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def normalize_key(value):
    return normalize_text(value).lower()


def find_header(headers, expected_name):
    expected = normalize_key(expected_name)

    for header in headers:
        if normalize_key(header) == expected:
            return str(header)

    return None


@app.route("/api/documents/upload", methods=["POST"])
def upload_rci_excel():
    if "file" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo."}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "El archivo no tiene nombre."}), 400

    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xlsm")):
        return jsonify({"error": "Solo se permiten archivos Excel."}), 400

    workbook = openpyxl.load_workbook(file, data_only=True)
    sheet = workbook.active

    HEADER_ROW = 7
    DATA_START_ROW = 8

    headers = [cell.value for cell in sheet[HEADER_ROW]]

    print("HEADERS RAW:", headers)
    print("HEADERS LIMPIOS:")
    for h in headers:
        print(f"[{normalize_text(h)}]")

    if not any(headers):
        return jsonify({"error": f"La fila {HEADER_ROW} no contiene encabezados válidos."}), 400

    criticidad_header = None
    for h in headers:
        if "criticidad" in normalize_key(h):
            criticidad_header = str(h)
            break

    if criticidad_header is None:
        return jsonify({"error": "No se encontró la columna 'CRITICIDAD' en el archivo."}), 400

    codigo_header = find_header(headers, "CÓDIGO INTERFAZ")
    tramo_header = find_header(headers, "Tramo")
    estado_header = find_header(headers, "ESTADO")
    sistema1_header = find_header(headers, "SISTEMA 1")
    subsistema_lider_header = find_header(headers, "SUBSISTEMA LIDER")
    sistema2_header = find_header(headers, "SISTEMA 2")
    subsistema_participante_header = find_header(headers, "SUBSISTEMA PARTICIPANTE")

    rows = []
    criticidad_alta = 0
    criticidad_media = 0
    criticidad_baja = 0

    for row in sheet.iter_rows(min_row=DATA_START_ROW, values_only=True):
        if all(value is None for value in row):
            continue

        row_dict = {}
        for i in range(len(headers)):
            header_name = headers[i] if headers[i] is not None else f"col_{i}"
            row_dict[str(header_name)] = row[i]

        criticidad_value = normalize_key(row_dict.get(criticidad_header))

        if criticidad_value == "alta":
            criticidad_alta += 1
        elif criticidad_value == "media":
            criticidad_media += 1
        elif criticidad_value == "baja":
            criticidad_baja += 1

        simplified_row = {
            "No.": row_dict.get("No."),
            "Tramo": row_dict.get(tramo_header) if tramo_header else None,
            "CÓDIGO INTERFAZ": row_dict.get(codigo_header) if codigo_header else None,
            "SISTEMA 1": row_dict.get(sistema1_header) if sistema1_header else None,
            "SUBSISTEMA LIDER": row_dict.get(subsistema_lider_header) if subsistema_lider_header else None,
            "SISTEMA 2": row_dict.get(sistema2_header) if sistema2_header else None,
            "SUBSISTEMA PARTICIPANTE": row_dict.get(subsistema_participante_header) if subsistema_participante_header else None,
            "CRITICIDAD": row_dict.get(criticidad_header),
            "ESTADO": row_dict.get(estado_header) if estado_header else None,
        }

        rows.append(simplified_row)

    total_interfaces = len(rows)

    return jsonify({
        "totalInterfaces": total_interfaces,
        "criticidadAlta": criticidad_alta,
        "criticidadMedia": criticidad_media,
        "criticidadBaja": criticidad_baja,
        "distribution": [
            {"name": "Alta", "value": criticidad_alta},
            {"name": "Media", "value": criticidad_media},
            {"name": "Baja", "value": criticidad_baja},
        ],
        "rows": rows
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
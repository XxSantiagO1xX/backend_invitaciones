from flask import Flask, request, jsonify
import openpyxl

app = Flask(__name__)

# Ruta para guardar confirmaciones
@app.route("/api/save-rsvp", methods=["POST"])
def save_rsvp():
    data = request.json  # Obtén los datos enviados por el cliente
    name = data.get("name")
    attendance = data.get("attendance")

    # Nombre del archivo Excel
    file_name = "confirmaciones.xlsx"

    # Intenta cargar el archivo o créalo si no existe
    try:
        workbook = openpyxl.load_workbook(file_name)
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Nombre", "Asistencia"])  # Encabezados

    # Agrega los datos al Excel
    sheet.append([name, attendance])
    workbook.save(file_name)

    return jsonify({"message": "Confirmación guardada exitosamente"}), 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

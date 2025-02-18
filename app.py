from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import os

app = Flask(__name__)
CORS(app)

# Configura las rutas relativas para plantillas y archivos estáticos
app.template_folder = os.path.join(os.path.dirname(__file__), "templates")
app.static_folder = os.path.join(os.path.dirname(__file__), "static")

PPTX_TEMPLATE = os.path.join(os.path.dirname(__file__), "Certificado Diploma Premio.pptx")
IMAGEN_ESTRELLA = os.path.join(os.path.dirname(__file__), "static", "estrella.png")

# Ruta principal
@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generar-diploma", methods=["POST"])
def generar_diploma():
    try:
        data = request.json
        print("Recibido JSON:", data)

        nombre = data.get("nombre", "Nombre")
        estrellas = data.get("estrellas", 0)  # Evita que sea None

        if estrellas is None or estrellas == "":
            estrellas = 0  # Asegurar que siempre sea un número

        estrellas = int(estrellas)  # Convertir a entero
        print("Nombre:", nombre)
        print("Estrellas:", estrellas)

        # Verificar que la plantilla existe
        if not os.path.exists(PPTX_TEMPLATE):
            print("❌ ERROR: El archivo PPTX no se encontró")
            return jsonify({"error": "El archivo PPTX no se encontró"}), 500

        # Verificar que la imagen de la estrella existe
        if not os.path.exists(IMAGEN_ESTRELLA):
            print("❌ ERROR: La imagen de la estrella no se encontró")
            return jsonify({"error": "La imagen de la estrella no se encontró"}), 500

        prs = Presentation(PPTX_TEMPLATE)

        # Reemplazar el nombre en el PPTX y aplicar estilo
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if "[Nombre]" in shape.text:
                        # Reemplazar el marcador con el nombre
                        shape.text = nombre

                        # Aplicar estilo solo al nombre
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if nombre in run.text:  # Aplicar estilo solo al nombre
                                    run.font.name = "TeXGyreChorus"  # Tipografía Allura
                                    run.font.size = Pt(40)    # Tamaño 60
                            paragraph.alignment = PP_ALIGN.CENTER  # Centrar texto

        # Reemplazar las estrellas en el PPTX con imágenes
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    if "[ESTRELLAS]" in shape.text:
                        # Eliminar el marcador de estrellas
                        shape.text = ""

                        # Obtener la posición y dimensiones de la forma
                        left = shape.left
                        top = shape.top
                        width = shape.width

                        # Ajustar el tamaño de las estrellas
                        estrella_width = Inches(0.3)  # Ancho más pequeño
                        estrella_height = Inches(0.3)  # Alto más pequeño

                        # Calcular el espacio total que ocuparán las estrellas
                        espacio_entre_estrellas = Inches(0.2)  # Espacio entre estrellas
                        espacio_total = (estrella_width * estrellas) + (espacio_entre_estrellas * (estrellas - 1))

                        # Calcular la posición inicial para centrar las estrellas
                        left_inicio = left + (width - espacio_total) / 2

                        # Insertar las imágenes de las estrellas
                        for i in range(estrellas):
                            estrella_left = left_inicio + (i * (estrella_width + espacio_entre_estrellas))
                            slide.shapes.add_picture(
                                IMAGEN_ESTRELLA,
                                estrella_left,
                                top,
                                width=estrella_width,
                                height=estrella_height
                            )

        output_filename = f"diploma_{nombre}.pptx"
        prs.save(output_filename)

        print("✅ Diploma generado correctamente:", output_filename)
        return send_file(output_filename, as_attachment=True)

    except Exception as e:
        print("❌ ERROR en Flask:", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
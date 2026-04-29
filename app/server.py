"""
Serveur Flask — Générateur de Comptes Rendus de Visite Agence
"""

import io
import os
import zipfile
from datetime import date
from flask import Flask, render_template, request, send_file, jsonify

from core import generate_reports, generate_consolidated, list_agencies
from gpt import get_gpt_fn

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

# Clé API OpenAI (depuis variable d'environnement ou saisie utilisateur)
_OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/agencies", methods=["POST"])
def get_agencies():
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Format non supporté. Utilisez un fichier .xlsx"}), 400
    try:
        return jsonify({"agencies": list_agencies(f.read())})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/generate", methods=["POST"])
def generate():
    """Génère les rapports individuels + consolidé selon les options."""
    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier reçu"}), 400

    f              = request.files["file"]
    rz_name        = request.form.get("rz_name", "")
    date_vis       = request.form.get("date_visite", date.today().strftime("%d/%m/%Y"))
    selected       = request.form.getlist("agencies")
    use_gpt        = request.form.get("use_gpt", "false").lower() == "true"
    api_key        = request.form.get("api_key", "").strip() or _OPENAI_API_KEY
    include_consol = request.form.get("consolidated", "false").lower() == "true"

    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Format non supporté"}), 400

    file_bytes = f.read()
    agencies   = selected if selected else None
    gpt_fn     = get_gpt_fn(api_key) if use_gpt else None

    # Modèle Word à utiliser si disponible
    _here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    _tpl  = os.path.join(_here, "template_synthese.docx")
    template_path = _tpl if os.path.exists(_tpl) else None

    try:
        reports = generate_reports(file_bytes, rz_name, date_vis, agencies, gpt_fn,
                                   template_path=template_path)
    except Exception as e:
        return jsonify({"error": f"Erreur génération : {str(e)}"}), 500

    if not reports:
        return jsonify({"error": "Aucune agence trouvée dans le fichier"}), 400

    # Rapport consolidé
    consolidated_bytes = None
    if include_consol and len(reports) > 1:
        try:
            consolidated_bytes = generate_consolidated(file_bytes, rz_name, date_vis, agencies, gpt_fn)
        except Exception as e:
            consolidated_bytes = None  # non bloquant

    # Un seul rapport individuel, pas de consolidé → .docx direct
    if len(reports) == 1 and not consolidated_bytes:
        name, docx_bytes = reports[0]
        return send_file(
            io.BytesIO(docx_bytes),
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=f"Synthese_CR_{name.replace(' ','_')}.docx",
        )

    # Plusieurs rapports ou consolidé → ZIP
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, docx_bytes in reports:
            zf.writestr(f"Rapports_individuels/Synthese_CR_{name.replace(' ','_')}.docx", docx_bytes)
        if consolidated_bytes:
            zf.writestr("Rapport_CONSOLIDE_Toutes_Agences.docx", consolidated_bytes)
    zip_buf.seek(0)

    return send_file(zip_buf, mimetype="application/zip", as_attachment=True,
                     download_name="Rapports_Visite_Agences.zip")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    debug = os.environ.get("FLASK_ENV", "production") == "development"
    app.run(host="0.0.0.0", port=port, debug=debug)

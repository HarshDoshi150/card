import pandas as pd
import os
import sys
from datetime import datetime
from PIL import Image, ImageDraw, ImageFont
from flask import Flask, render_template, request, send_file, session
import zipfile
import io
from flask_session import Session
import logging
import re
from flask_cors import CORS  # optional if frontend AJAX/cookies used

# ----------------- App Setup ----------------- #
app = Flask(__name__)
CORS(app, origins="https://card-stec.onrender.com", supports_credentials=True)  # adjust your frontend URL
app.secret_key = os.environ.get("SECRET_KEY", "fallback-key")
app.config["SESSION_TYPE"] = "filesystem"
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024  # 10MB limit
Session(app)

# ----------------- Fixed Positions ----------------- #
NAME_Y = 640
MONTH_X, MONTH_Y = 465, 1020
MONTH_MAX_WIDTH = 350
PARA_Y = 820
LEFT_MARGIN = 300
RIGHT_MARGIN = 230
PARA_LINE_SPACING = 25
NAME_COLOR = "#ab7d3e"
BLACK = "#000000"

# ----------------- Resource Path ----------------- #
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ----------------- Helper Functions ----------------- #
def fit_font(draw, text, font_path, start_size, max_width=None):
    size = start_size
    while size > 20:
        try:
            font = ImageFont.truetype(resource_path(font_path), size)
            if max_width:
                bbox = draw.textbbox((0, 0), text, font=font)
                width = bbox[2] - bbox[0]
                if width <= max_width:
                    return font
            else:
                return font
        except:
            pass
        size -= 1
    return ImageFont.load_default()

def draw_paragraph(draw, text, font, x, y, max_width, line_spacing):
    words = text.split()
    lines = []
    current_line = ""

    for word in words:
        test_line = current_line + (" " + word if current_line else word)
        bbox = draw.textbbox((0, 0), test_line, font=font)
        width = bbox[2] - bbox[0]

        if width <= max_width:
            current_line = test_line
        else:
            if current_line:
                lines.append(current_line)
            current_line = word

    if current_line:
        lines.append(current_line)

    current_y = y
    for line in lines:
        draw.text((x, current_y), line, font=font, fill=BLACK)
        current_y += font.size + line_spacing

# ----------------- Routes ----------------- #
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":

        if "cert_zip" in session:
            session.pop("cert_zip")

        excel_file = request.files.get("excel_file")
        paragraph_template = request.form.get("paragraph")

        if not excel_file or not paragraph_template:
            return render_template("index.html", error="Please upload Excel and type paragraph.")

        # --------------------- Read Excel --------------------- #
        try:
            data = pd.read_excel(excel_file, dtype=str, engine="openpyxl")
            data.columns = data.columns.str.strip().str.title()
        except Exception as e:
            logging.exception("Failed to read Excel")
            return render_template("index.html", error=f"Invalid Excel file: {str(e)}")

        # --------------------- Required Column Check --------------------- #
        required_columns = ["Name", "Month"]
        missing_columns = [col for col in required_columns if col not in data.columns]

        if missing_columns:
            return render_template(
                "index.html",
                error=f"Missing required column(s): {', '.join(missing_columns)}"
            )

        # --------------------- Placeholder Validation --------------------- #
        placeholders_in_template = re.findall(r"{(.*?)}", paragraph_template)

        missing_placeholders = [
            ph for ph in placeholders_in_template
            if ph.title() not in data.columns
        ]

        if missing_placeholders:
            return render_template(
                "index.html",
                error=f"The following placeholders are missing in Excel: {', '.join(missing_placeholders)}"
            )

        # --------------------- Prepare ZIP --------------------- #
        output_buffer = io.BytesIO()
        zipf = zipfile.ZipFile(output_buffer, "w", zipfile.ZIP_DEFLATED)
        generated_count = 0

        for idx, row in data.iterrows():
            try:
                name = str(row.get("Name", "")).strip()
                if not name:
                    continue

                paragraph = paragraph_template
                for ph in placeholders_in_template:
                    column_name = ph.title()
                    raw_value = str(row.get(column_name, "N/A")).strip()
                    if column_name == "Conversion":
                        formatted_value = f"{raw_value}%"
                    else:
                        formatted_value = raw_value
                    paragraph = paragraph.replace(f"{{{ph}}}", formatted_value)

                # Format Month
                month_val = row.get("Month", "")
                try:
                    month = pd.to_datetime(str(month_val)).strftime("%B %Y")
                except:
                    month = str(month_val).strip()

                # Create Image
                img = Image.open(resource_path("static/certificate.png")).convert("RGB")
                draw = ImageDraw.Draw(img)
                img_width, _ = img.size

                # Draw Name
                name_font = fit_font(draw, name, "static/OleoScript-Regular.ttf", 110)
                bbox = draw.textbbox((0, 0), name, font=name_font)
                name_width = bbox[2] - bbox[0]
                name_x = (img_width - name_width) // 2
                draw.text((name_x, NAME_Y), name, font=name_font, fill=NAME_COLOR)

                # Draw Month
                month_font = fit_font(draw, month, "static/Montserrat-Bold.ttf", 42, MONTH_MAX_WIDTH)
                draw.text((MONTH_X, MONTH_Y), month, font=month_font, fill=BLACK)

                # Draw Paragraph
                para_x = LEFT_MARGIN
                para_max_width = img_width - LEFT_MARGIN - RIGHT_MARGIN
                para_font = ImageFont.truetype(resource_path("static/Montserrat-Bold.ttf"), 40)
                draw_paragraph(draw, paragraph, para_font, para_x, PARA_Y, para_max_width, PARA_LINE_SPACING)

                # Save PDF to ZIP
                safe_name = "".join(c for c in name if c.isalnum() or c in (" ", "_")).rstrip().replace(" ", "_")
                img_bytes = io.BytesIO()
                img.save(img_bytes, "PDF", resolution=100.0)
                img_bytes.seek(0)
                zipf.writestr(f"{safe_name}.pdf", img_bytes.read())
                generated_count += 1

            except Exception:
                logging.exception(f"Error processing row {idx+1}")
                continue

        zipf.close()
        output_buffer.seek(0)
        session["cert_zip"] = output_buffer.getvalue()
        session["total"] = generated_count

        return render_template("index.html", generated=True, total=generated_count)

    return render_template("index.html")

# ----------------- Download Route ----------------- #
@app.route("/download")
def download_zip():
    if "cert_zip" not in session:
        return "No certificates generated yet!", 400

    output_buffer = io.BytesIO(session["cert_zip"])
    output_buffer.seek(0)

    return send_file(
        output_buffer,
        as_attachment=True,
        download_name="All_Certificates.zip",
        mimetype="application/zip"
    )

# ----------------- Run App ----------------- #
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # 5000 for local dev
    app.run(host="0.0.0.0", port=port, debug=True)

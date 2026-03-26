"""
VERIDEM — Serveur Python local
Pipeline : Florence-2 (OCR + Caption) + ViT (détection IA) + OpenCV (analyse zones)

Usage :
  pip install -r requirements.txt
  python server.py

Le serveur démarre sur http://localhost:7860
Votre site React l'appelle via fetch("http://localhost:7860/analyze")
"""

import io
import os
import math
import datetime
import numpy as np
import cv2
from PIL import Image
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
from transformers import pipeline, AutoProcessor, AutoModelForCausalLM
import torch

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# CONFIGURATION
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

DEVICE = "cuda" if torch.cuda.is_available() else "cpu"

REPORTS_DIR = os.path.join(os.path.dirname(__file__), "reports")
os.makedirs(REPORTS_DIR, exist_ok=True)
DTYPE = torch.float16 if DEVICE == "cuda" else torch.float32

print(f"[INFO] Device : {DEVICE}")
print(f"[INFO] Chargement des modèles... (première fois = téléchargement ~3 GB)")

# ── Modèle 1 : ViT — Classification IA vs Réel ──
print("[1/3] Chargement du classifieur ViT (umm-maybe/AI-image-detector)...")
vit_classifier = pipeline(
    "image-classification",
    model="umm-maybe/AI-image-detector",
    device=0 if DEVICE == "cuda" else -1,
)
print("      ✓ ViT prêt")

# ── Modèle 2 : Florence-2 — OCR + Description ──
print("[2/3] Chargement de Florence-2 (microsoft/Florence-2-base)...")
florence_processor = AutoProcessor.from_pretrained(
    "microsoft/Florence-2-base",
    trust_remote_code=True,
)
florence_model = AutoModelForCausalLM.from_pretrained(
    "microsoft/Florence-2-base",
    trust_remote_code=True,
    torch_dtype=DTYPE,
).to(DEVICE)
florence_model.eval()
print("      ✓ Florence-2 prêt")

print("[3/3] OpenCV prêt (intégré)")
print()
print("=" * 50)
print("  SERVEUR PRÊT — http://localhost:7860")
print("=" * 50)
print()

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# FONCTIONS DU PIPELINE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def run_florence(image, task, text_input=""):
    """Exécute une tâche Florence-2."""
    prompt = task if not text_input else f"{task}{text_input}"
    inputs = florence_processor(
        text=prompt, images=image, return_tensors="pt"
    ).to(DEVICE, DTYPE)

    with torch.no_grad():
        generated = florence_model.generate(
            input_ids=inputs["input_ids"],
            pixel_values=inputs["pixel_values"],
            max_new_tokens=1024,
            num_beams=3,
        )

    text = florence_processor.batch_decode(generated, skip_special_tokens=False)[0]
    result = florence_processor.post_process_generation(
        text, task=task, image_size=(image.width, image.height)
    )
    return result


def extract_ocr(image):
    """Extrait le texte via Florence-2 <OCR>."""
    result = run_florence(image, "<OCR>")
    return result.get("<OCR>", "(aucun texte détecté)")


def generate_description(image):
    """Description détaillée via Florence-2."""
    result = run_florence(image, "<MORE_DETAILED_CAPTION>")
    return result.get("<MORE_DETAILED_CAPTION>", "")


def generate_caption(image):
    """Caption court via Florence-2."""
    result = run_florence(image, "<CAPTION>")
    return result.get("<CAPTION>", "")


def classify_image(image):
    """Classifie IA vs Réel via ViT."""
    results = vit_classifier(image)
    output = {"probAI": 0.0, "probHuman": 0.0}
    for r in results:
        label = r["label"].lower()
        if "artificial" in label or "ai" in label:
            output["probAI"] = round(r["score"], 4)
        elif "human" in label or "real" in label:
            output["probHuman"] = round(r["score"], 4)
    return output


def detect_photos_opencv(image, min_area_ratio=0.02):
    """Détecte les photos contenues dans l'image via OpenCV."""
    img_np = np.array(image)
    if len(img_np.shape) == 2:
        gray = img_np
    else:
        gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 50, 150)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    dilated = cv2.dilate(edges, kernel, iterations=2)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    img_area = image.width * image.height
    min_area = img_area * min_area_ratio
    max_area = img_area * 0.95

    boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = w * h
        aspect = w / max(h, 1)
        if min_area < area < max_area and 0.2 < aspect < 5.0:
            boxes.append((x, y, w, h))

    if not boxes:
        return []

    boxes_arr = np.array(boxes)
    x1 = boxes_arr[:, 0]
    y1 = boxes_arr[:, 1]
    x2 = boxes_arr[:, 0] + boxes_arr[:, 2]
    y2 = boxes_arr[:, 1] + boxes_arr[:, 3]
    areas = boxes_arr[:, 2] * boxes_arr[:, 3]
    order = areas.argsort()[::-1]
    keep = []

    while order.size > 0:
        i = order[0]
        keep.append(i)
        xx1 = np.maximum(x1[i], x1[order[1:]])
        yy1 = np.maximum(y1[i], y1[order[1:]])
        xx2 = np.minimum(x2[i], x2[order[1:]])
        yy2 = np.minimum(y2[i], y2[order[1:]])
        w = np.maximum(0.0, xx2 - xx1)
        h = np.maximum(0.0, yy2 - yy1)
        inter = w * h
        iou = inter / (areas[i] + areas[order[1:]] - inter)
        remaining = np.where(iou <= 0.3)[0]
        order = order[remaining + 1]

    return [boxes[i] for i in keep]


def analyze_zones(image):
    """Analyse par zone via OpenCV."""
    img_np = np.array(image)
    if len(img_np.shape) == 2:
        gray = img_np
    else:
        gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)

    h, w = gray.shape
    if max(h, w) > 512:
        scale = 512 / max(h, w)
        gray = cv2.resize(gray, (int(w * scale), int(h * scale)))

    lap_var = cv2.Laplacian(gray, cv2.CV_64F).var()
    artifact_score = max(0, min(100, int(100 - min(lap_var / 50, 100))))

    sobelx = cv2.Sobel(gray, cv2.CV_64F, 1, 0, ksize=3)
    sobely = cv2.Sobel(gray, cv2.CV_64F, 0, 1, ksize=3)
    grad_mag = np.sqrt(sobelx**2 + sobely**2)
    grad_std = grad_mag.std()
    texture_score = max(0, min(100, int(100 - min(grad_std / 2, 100))))

    f_transform = np.fft.fft2(gray.astype(np.float64))
    f_shift = np.fft.fftshift(f_transform)
    magnitude = np.log(np.abs(f_shift) + 1)
    spectral_std = magnitude.std()
    spectral_score = max(0, min(100, int(100 - min(spectral_std * 10, 100))))

    edges = cv2.Canny(gray, 100, 200)
    edge_density = edges.sum() / (gray.shape[0] * gray.shape[1] * 255)
    contour_score = max(0, min(100, int(100 - min(edge_density * 500, 100))))

    return [
        {"label": "Artefacts de génération", "score": artifact_score},
        {"label": "Cohérence des textures", "score": texture_score},
        {"label": "Patterns spectraux", "score": spectral_score},
        {"label": "Contours et transitions", "score": contour_score},
    ]


def save_report(base_name: str, filename: str, verdict: str, confidence: int,
                ocr_text: str, description: str) -> str:
    """Sauvegarde OCR + description dans un .txt et retourne son nom de fichier."""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"{base_name}_{timestamp}.txt"
    report_path = os.path.join(REPORTS_DIR, report_filename)

    content = (
        f"=== RAPPORT VERIDEM ===\n"
        f"Fichier  : {filename}\n"
        f"Date     : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        f"Verdict  : {verdict} ({confidence}%)\n"
        f"\n--- OCR ---\n{ocr_text}\n"
        f"\n--- DESCRIPTION ---\n{description}\n"
    )

    with open(report_path, "w", encoding="utf-8") as f:
        f.write(content)

    print(f"  ✓ Rapport sauvegardé : reports/{report_filename}")
    return report_filename


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# SERVEUR FLASK
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

app = Flask(__name__)
CORS(app)


@app.route("/", methods=["GET"])
def index():
    return jsonify({
        "service": "VERIDEM — Pipeline DIGIDEM",
        "status": "ready",
        "device": DEVICE,
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "healthy",
        "models_loaded": True,
        "device": DEVICE,
    })


@app.route("/analyze", methods=["POST"])
def analyze():
    """Pipeline complet d'analyse d'image."""

    if "file" not in request.files:
        return jsonify({"error": "Aucun fichier envoyé. Utilisez le champ 'file'."}), 400

    file = request.files["file"]
    if not file.content_type or not file.content_type.startswith("image/"):
        return jsonify({"error": "Le fichier doit être une image (PNG, JPEG, WebP)"}), 400

    try:
        image = Image.open(file.stream).convert("RGB")
    except Exception as e:
        return jsonify({"error": f"Impossible de lire l'image : {str(e)}"}), 400

    filename = file.filename or "image"
    base_name = os.path.splitext(filename)[0]

    print(f"[ANALYSE] Image reçue : {filename} ({image.width}x{image.height})")

    print("  [1/5] Détection de photos (OpenCV)...")
    photo_boxes = detect_photos_opencv(image)

    print("  [2/5] Extraction OCR (Florence-2)...")
    ocr_text = extract_ocr(image)
    if not ocr_text or not ocr_text.strip():
        ocr_text = "(aucun texte détecté)"

    print("  [3/5] Description (Florence-2)...")
    description = generate_description(image)

    print("  [4/5] Classification IA vs Réel (ViT)...")
    global_class = classify_image(image)
    prob_ai = global_class["probAI"]
    prob_human = global_class["probHuman"]
    is_ai = prob_ai > prob_human
    confidence = round(max(prob_ai, prob_human) * 100)

    photos = []
    for idx, (x, y, w, h) in enumerate(photo_boxes):
        cropped = image.crop((x, y, x + w, y + h))
        photo_caption = generate_caption(cropped)
        photo_class = classify_image(cropped)
        p_ai = photo_class["probAI"]
        p_human = photo_class["probHuman"]
        photos.append({
            "filename": f"{base_name}_photo_{str(idx + 1).zfill(2)}.png",
            "size": f"{w}×{h}",
            "caption": photo_caption,
            "verdict": "IA" if p_ai > p_human else "RÉEL",
            "confidence": round(max(p_ai, p_human), 4),
            "probAI": round(p_ai, 4),
            "probHuman": round(p_human, 4),
        })

    print("  [5/5] Analyse par zone (OpenCV)...")
    zones = analyze_zones(image)

    print("  [6/6] Sauvegarde rapport...")
    verdict_str = "IA" if is_ai else "RÉEL"
    report_filename = save_report(
        base_name, filename, verdict_str, confidence,
        ocr_text, description
    )

    has_exif = hasattr(image, "_getexif") and image._getexif() is not None
    metadata = [
        {"key": "Résolution", "val": f"{image.width} × {image.height} px"},
        {"key": "Format", "val": file.content_type.split("/")[-1].upper()},
        {"key": "Données EXIF", "val": "Présentes — appareil détecté" if has_exif else "Absentes — suspect"},
        {"key": "Bruit capteur", "val": "Uniformité anormale — typique IA" if is_ai else "Bruit naturel — capteur cohérent"},
        {"key": "Photos détectées", "val": f"{len(photos)} photo(s) extraite(s)"},
        {"key": "Texte OCR", "val": "Aucun texte" if ocr_text == "(aucun texte détecté)" else f"{len(ocr_text)} caractères extraits"},
    ]

    total_ai = sum(1 for p in photos if p["verdict"] == "IA")

    result = {
        "isAI": is_ai,
        "confidence": confidence,
        "probAI": round(prob_ai * 100, 1),
        "probReal": round(prob_human * 100, 1),
        "ocrText": ocr_text,
        "ocrCharCount": len(ocr_text),
        "description": description,
        "photos": photos,
        "totalPhotos": len(photos),
        "totalAI": total_ai,
        "totalReal": len(photos) - total_ai,
        "metadata": metadata,
        "zones": zones,
        "reportFile": report_filename,
    }

    print(f"  ✓ Terminé — {'IA détectée' if is_ai else 'Image réelle'} ({confidence}%)")
    return jsonify(result)


@app.route("/reports", methods=["GET"])
def list_reports():
    """Liste tous les rapports .txt disponibles."""
    files = sorted(
        [f for f in os.listdir(REPORTS_DIR) if f.endswith(".txt")],
        reverse=True,
    )
    return jsonify({"reports": files, "count": len(files)})


@app.route("/reports/<path:filename>", methods=["GET"])
def get_report(filename):
    """Retourne le contenu brut d'un rapport .txt."""
    return send_from_directory(REPORTS_DIR, filename, mimetype="text/plain; charset=utf-8")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# LANCEMENT
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    app.run(host="0.0.0.0", port=port, debug=False)
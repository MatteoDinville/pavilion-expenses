import json
import logging
import os
from pathlib import Path
import time
import sentry_sdk

from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from google import genai
from google.genai import types
from google.genai.errors import ServerError
from pydantic import BaseModel, Field

from excel_writer import generate_monthly_excel
from mock import mock_gemini_extraction

# 1. Charger les variables d'environnement avant d'initialiser les services
load_dotenv()

# --- Initialisation de Sentry ---
sentry_sdk.init(
    dsn=os.getenv("SENTRY_DSN"),
    send_default_pii=True,
    enable_logs=True,
    environment=os.getenv("ENVIRONMENT", "development"),
)

# --- Configuration du logging local (Terminal) ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S"
)
logger = logging.getLogger("frais_app")

api_key = os.getenv("GEMINI_API_KEY")

if not api_key:
    logger.error("La variable GEMINI_API_KEY est introuvable.")
    raise ValueError("Erreur : La variable GEMINI_API_KEY est introuvable.")

# Initialisation des clients API
client = genai.Client(api_key=api_key)
app = FastAPI(title="Traitement de Frais Manuscrit")

# --- Modèles Pydantic ---
class DonneeFiche(BaseModel):
    numero_page: str | None = Field(
        default=None, description="Numéro de la page"
    )
    date: str | None = Field(default=None, description="Date inscrite")
    facture_nom: str | None = Field(
        default=None, description="Nom ou intitulé de la facture"
    )
    montant: str | None = Field(default=None, description="Montant exact écrit")
    categorie: str | None = Field(
        default=None, description="Catégorie du document"
    )


class DocumentMultiPages(BaseModel):
    pages: list[DonneeFiche] = Field(
        description="Liste des fiches extraites par page"
    )


class ValidationPayload(BaseModel):
    fiches: list[DonneeFiche]


# --- Interface Web ---
@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    return Path("index.html").read_text(encoding="utf-8")

@app.get("/favicon.svg")
async def serve_favicon_svg():
    return FileResponse("favicon.svg", media_type="image/svg+xml")

# --- 1. Route d'extraction pure via Gemini ---
@app.post("/api/extract-only")
async def extract_only(file: UploadFile = File(...)):
    logger.info(f"--- Nouvelle requête d'extraction pour le fichier : {file.filename} ---")

    # MODE MOCK : aucun appel à Gemini
    if os.getenv("MOCK_GEMINI", "false").lower() == "true":
        logger.info("🧪 MODE MOCK GEMINI ACTIVÉ")

        time.sleep(2)  # simule le temps de traitement

        extractions = mock_gemini_extraction()

        logger.info(
            f"Simulation terminée : {len(extractions)} fiche(s) générée(s)."
        )

        return {
            "status": "success",
            "extractions": extractions,
        }

    if not file.filename.endswith(".pdf"):
        logger.warning(f"Fichier rejeté (format non supporté) : {file.filename}")
        raise HTTPException(
            status_code=400, detail="Veuillez envoyer un fichier PDF."
        )

    temp_path = f"temp_{file.filename}"
    with open(temp_path, "wb") as buffer:
        buffer.write(await file.read())
    logger.info(f"Fichier temporaire créé : {temp_path}")

    try:
        logger.info("Envoi du fichier à l'API Gemini File...")
        pdf_file = client.files.upload(file=temp_path)
        logger.info(f"Fichier téléversé sur Gemini avec l'ID : {pdf_file.name}")

        prompt = (
            "Analyse ce document PDF. Pour CHAQUE page, extrais avec précision"
            " les 5 champs : Numéro de page, Date, Facture (Nom), Montant et"
            " Catégorie."
        )

        response = None
        for attempt in range(1, 4):
            try:
                logger.info(f"Tentative {attempt}/3 de génération de contenu via Gemini...")
                response = client.models.generate_content(
                    model="gemini-3.6-flash",
                    contents=[pdf_file, prompt],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=DocumentMultiPages,
                    ),
                )
                logger.info("Réponse reçue de Gemini avec succès.")
                break
            except ServerError as se:
                logger.warning(f"Erreur serveur Gemini à la tentative {attempt}: {se}")
                time.sleep(2)

        # Nettoyage fichier distant & local
        client.files.delete(name=pdf_file.name)
        logger.info(f"Fichier Gemini nettoyé distant : {pdf_file.name}")
        if os.path.exists(temp_path):
            os.remove(temp_path)
            logger.info(f"Fichier temporaire supprimé : {temp_path}")

        if not response:
            logger.error("Échec de la réponse Gemini après 3 tentatives.")
            raise HTTPException(
                status_code=503,
                detail="Le service de traitement est temporairement indisponible.",
            )

        data = json.loads(response.text)
        extractions = data.get("pages", [])
        logger.info(f"Extraction terminée : {len(extractions)} fiche(s) trouvée(s).")
        for idx, item in enumerate(extractions, 1):
            logger.info(f"  Fiche #{idx} -> Page: {item.get('numero_page')}, Date: {item.get('date')}, Facture: {item.get('facture_nom')}, Montant: {item.get('montant')}, Catégorie: {item.get('categorie')}")

        return {"status": "success", "extractions": extractions}

    except Exception as e:
        error_text = str(e)
        logger.error(f"Erreur lors de l'extraction : {error_text}")
        sentry_sdk.capture_exception(e)  # Transmet la stacktrace à Sentry

        if "RESOURCE_EXHAUSTED" in error_text or "429" in error_text:
            logger.warning("Quota Gemini atteint : limite d'extraction dépassée.")
            if os.path.exists(temp_path):
                os.remove(temp_path)
            raise HTTPException(
                status_code=429,
                detail="Vous avez atteint la limite d'extraction, merci de réessayer demain.",
            )

        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise HTTPException(status_code=500, detail="Une erreur est survenue lors de l'extraction.")


# --- 2. Route de confirmation et d'enregistrement dans Supabase ---
@app.post("/api/confirm-and-fill")
async def confirm_and_fill(payload: ValidationPayload):
    logger.info(
        f"--- Génération Excel demandée avec {len(payload.fiches)} fiche(s) ---"
    )

    try:
        records = [
            {
                "numero_page": f.numero_page,
                "date": f.date,
                "facture_nom": f.facture_nom,
                "montant": f.montant,
                "categorie": f.categorie,
            }
            for f in payload.fiches
        ]

        for idx, record in enumerate(records, 1):
            logger.info(
                f"Fiche #{idx} -> "
                f"Page: {record.get('numero_page')} | "
                f"Date: {record.get('date')} | "
                f"Facture: {record.get('facture_nom')} | "
                f"Montant: {record.get('montant')} | "
                f"Catégorie: {record.get('categorie')}"
            )

        # Génération d'un nouveau fichier Excel
        now = datetime.now()
        month_year_str = now.strftime("%m_%Y")

        output_filename = generate_monthly_excel(
            all_records=records,
            month_year_str=month_year_str,
        )

        logger.info(
            f"Fichier Excel généré avec succès : {output_filename}"
        )

        return FileResponse(
            output_filename,
            filename=output_filename,
            media_type=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            ),
        )

    except Exception as e:
        logger.error(
            f"Erreur lors de la génération de l'Excel : {str(e)}"
        )

        sentry_sdk.capture_exception(e)

        raise HTTPException(
            status_code=500,
            detail="Erreur lors de la génération du fichier Excel.",
        )

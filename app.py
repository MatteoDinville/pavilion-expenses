import calendar
from datetime import datetime, timezone
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
from supabase import create_client

from excel_writer import generate_monthly_excel

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
supabase_url = os.getenv("SUPABASE_URL")
supabase_key = os.getenv("SUPABASE_KEY")

if not api_key:
    logger.error("La variable GEMINI_API_KEY est introuvable.")
    raise ValueError("Erreur : La variable GEMINI_API_KEY est introuvable.")
if not supabase_url or not supabase_key:
    logger.error("SUPABASE_URL ou SUPABASE_KEY est introuvable.")
    raise ValueError(
        "Erreur : SUPABASE_URL ou SUPABASE_KEY est introuvable dans le .env."
    )

# Initialisation des clients API
client = genai.Client(api_key=api_key)
supabase = create_client(supabase_url, supabase_key)

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


# --- 1. Route d'extraction pure via Gemini ---
@app.post("/api/extract-only")
async def extract_only(file: UploadFile = File(...)):
    logger.info(f"--- Nouvelle requête d'extraction pour le fichier : {file.filename} ---")
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
        logger.error(f"Erreur lors de l'extraction : {str(e)}")
        sentry_sdk.capture_exception(e)  # Transmet la stacktrace à Sentry
        if os.path.exists(temp_path):
            os.remove(temp_path)
        raise HTTPException(status_code=500, detail="Une erreur est survenue lors de l'extraction.")


# --- 2. Route de confirmation et d'enregistrement dans Supabase ---
@app.post("/api/confirm-and-fill")
async def confirm_and_fill(payload: ValidationPayload):
    logger.info(f"--- Requête d'insertion Supabase reçue avec {len(payload.fiches)} fiche(s) ---")
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

        for idx, rec in enumerate(records, 1):
            logger.info(f"  Enregistrement Supabase #{idx} -> {rec}")

        # Insertion des données dans Supabase
        res = supabase.table("depense").insert(records).execute()
        logger.info(f"Insertion Supabase réussie. Données insérées : {res.data}")

        return {
            "status": "success",
            "message": f"{len(records)} ligne(s) enregistrée(s) avec succès dans Supabase !",
        }
    except Exception as e:
        logger.error(f"Erreur lors de l'insertion dans Supabase : {str(e)}")
        sentry_sdk.capture_exception(e)  # Transmet la stacktrace à Sentry
        raise HTTPException(status_code=500, detail="Erreur lors de l'enregistrement en base de données.")


# --- 3. Route pour générer le fichier Excel mensuel depuis Supabase ---
@app.get("/api/download-excel")
async def download_excel():
    logger.info("--- Requête de génération et téléchargement de la matrice Excel ---")
    try:
        now = datetime.now(timezone.utc)
        month_year_str = now.strftime("%m_%Y")

        # Récupérer les bornes temporelles du mois en cours
        first_day = datetime(now.year, now.month, 1, 0, 0, 0, tzinfo=timezone.utc)
        _, last_day_num = calendar.monthrange(now.year, now.month)
        last_day = datetime(
            now.year, now.month, last_day_num, 23, 59, 59, tzinfo=timezone.utc
        )

        logger.info(f"Recherche Supabase filtrée sur 'created_at' entre {first_day.isoformat()} et {last_day.isoformat()}")

        # Récupération de l'ensemble des dépenses insérées ce mois-ci
        res = (
            supabase.table("depense")
            .select("*")
            .gte("created_at", first_day.isoformat())
            .lte("created_at", last_day.isoformat())
            .order("created_at", desc=False)
            .execute()
        )
        current_frais = res.data

        logger.info(f"{len(current_frais)} enregistrement(s) trouvé(s) dans Supabase pour le mois en cours ({month_year_str}).")
        for idx, item in enumerate(current_frais, 1):
            logger.info(f"  [Supabase ID: {item.get('id')}] CreatedAt: {item.get('created_at')} | Facture: {item.get('facture_nom')} | Montant: {item.get('montant')} | Cat: {item.get('categorie')}")

        # Générer l'Excel ventilé à partir du modèle vierge MATRICE FRAIS.xlsx
        output_filename = generate_monthly_excel(
            all_records=current_frais, month_year_str=month_year_str
        )
        logger.info(f"Fichier Excel généré avec succès : {output_filename}")

        return FileResponse(
            output_filename,
            filename=output_filename,
            media_type=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
    except Exception as e:
        logger.error(f"Erreur lors du téléchargement de l'Excel : {str(e)}")
        sentry_sdk.capture_exception(e)  # Transmet la stacktrace à Sentry
        raise HTTPException(status_code=500, detail="Erreur lors de la génération du fichier Excel.")
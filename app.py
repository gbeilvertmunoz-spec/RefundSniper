# ============================================================
# REFUNDSNIPER — V1
# Micro-SaaS B2B de détection des colis en retard / perdus
# ============================================================
#
# Lancement :
#   pip install -r requirements.txt
#   streamlit run app.py
#
# Architecture :
#   - Application Streamlit monofichier
#   - Base locale simulée via st.session_state
#   - Import CSV / XLSX
#   - Détection des retards
#   - Estimation des montants potentiellement réclamables
#   - Export CSV / XLSX
#   - Génération de modèles de réclamation
#
# IMPORTANT :
# Les règles de remboursement utilisées dans cette V1 sont des
# règles de démonstration/configuration. Elles ne constituent pas
# une garantie contractuelle de remboursement.
#
# Les données de démonstration sont entièrement synthétiques.
# ============================================================

from __future__ import annotations

import io
import hashlib
import re
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURATION GLOBALE
# ============================================================

APP_NAME = "RefundSniper"
APP_VERSION = "1.0.0"

# Limites de sécurité
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_ROWS = 5000

# Couleurs Cyberpunk Clean
BG = "#0D1117"
CARD = "#111820"
CARD_ALT = "#151D27"
TEXT = "#F4F7FB"
MUTED = "#8D9AA8"
GREEN = "#00FFCC"
PURPLE = "#8A2BE2"
RED = "#FF4D67"
ORANGE = "#FFCC66"
BLUE = "#4DA6FF"

# ------------------------------------------------------------
# Règles de démonstration
# ------------------------------------------------------------
# Ces valeurs sont volontairement configurables depuis la sidebar.
#
# Elles ne représentent PAS automatiquement les CGV de Colissimo,
# DHL ou UPS.
# ------------------------------------------------------------

DEFAULT_CARRIER_RULES = {
    "Colissimo": {
        "late_days": 3,
        "lost_days": 10,
        "cap": 200.0,
    },
    "DHL": {
        "late_days": 4,
        "lost_days": 10,
        "cap": 250.0,
    },
    "UPS": {
        "late_days": 4,
        "lost_days": 10,
        "cap": 250.0,
    },
    "Autre": {
        "late_days": 5,
        "lost_days": 15,
        "cap": 150.0,
    },
}

STATUS_LABELS = {
    "eligible": "ÉLIGIBLE",
    "review": "À VÉRIFIER",
    "not_eligible": "NON ÉLIGIBLE",
    "claimed": "RÉCLAMATION PRÉPARÉE",
}

STATUS_EMOJI = {
    "eligible": "🎯",
    "review": "🔎",
    "not_eligible": "✓",
    "claimed": "📑",
}


# ============================================================
# ALIASES DE COLONNES
# ============================================================

COLUMN_ALIASES = {
    "tracking": [
        "tracking",
        "tracking_number",
        "trackingnumber",
        "numero_suivi",
        "numéro_suivi",
        "numero_de_suivi",
        "n°_suivi",
        "no_suivi",
        "suivi",
        "tracking_id",
        "parcel_id",
    ],
    "carrier": [
        "carrier",
        "transporteur",
        "transporteur_name",
        "shipping_carrier",
        "livreur",
        "delivery_carrier",
    ],
    "order": [
        "order",
        "order_id",
        "order_number",
        "numero_commande",
        "numéro_commande",
        "commande",
        "reference_commande",
        "référence_commande",
    ],
    "shipped": [
        "shipped",
        "shipped_at",
        "shipping_date",
        "date_expedition",
        "date_expédition",
        "date_envoi",
        "sent_at",
    ],
    "promised": [
        "promised",
        "promised_date",
        "delivery_date",
        "expected_delivery",
        "expected_delivery_date",
        "date_livraison_prevue",
        "date_livraison_attendue",
        "livraison_prevue",
        "livraison_attendue",
    ],
    "delivered": [
        "delivered",
        "delivered_at",
        "actual_delivery",
        "actual_delivery_date",
        "date_livraison",
        "date_livraison_reelle",
        "date_livraison_réelle",
    ],
    "shipping_cost": [
        "shipping_cost",
        "shipping_price",
        "transport_cost",
        "frais_port",
        "frais_livraison",
        "cout_transport",
        "coût_transport",
        "prix_livraison",
    ],
    "value": [
        "value",
        "order_value",
        "order_amount",
        "amount",
        "montant",
        "montant_commande",
        "valeur_commande",
        "total",
    ],
    "customer": [
        "customer",
        "customer_name",
        "client",
        "nom_client",
        "nom",
        "destinataire",
    ],
    "city": [
        "city",
        "ville",
        "destination_city",
        "destination",
    ],
    "service": [
        "service",
        "shipping_service",
        "delivery_service",
        "offre",
        "service_livraison",
    ],
    "status": [
        "status",
        "statut",
        "shipping_status",
        "delivery_status",
        "etat",
        "état",
    ],
}


# ============================================================
# UTILITAIRES
# ============================================================

def safe_text(value: Any, default: str = "") -> str:
    """Convertit une valeur en texte sans faire planter l'application."""
    try:
        if value is None:
            return default

        if pd.isna(value):
            return default

        return str(value).strip()

    except Exception:
        return default


def safe_float(value: Any, default: float = 0.0) -> float:
    """Conversion robuste vers float."""
    try:
        if value is None:
            return default

        if isinstance(value, str):
            cleaned = (
                value.replace("€", "")
                .replace(" ", "")
                .replace(",", ".")
                .strip()
            )

            # Gestion basique des nombres du type 1.299,90
            if cleaned.count(".") > 1:
                cleaned = cleaned.replace(".", "")

            return float(cleaned)

        return float(value)

    except Exception:
        return default


def safe_date(value: Any) -> Optional[date]:
    """Conversion robuste vers date."""
    try:
        if value is None:
            return None

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        parsed = pd.to_datetime(value, errors="coerce")

        if pd.isna(parsed):
            return None

        return parsed.date()

    except Exception:
        return None


def money(value: Any) -> str:
    """Format monétaire français."""
    try:
        amount = safe_float(value)
        formatted = f"{amount:,.2f}".replace(",", " ").replace(".", ",")
        return f"{formatted} €"
    except Exception:
        return "0,00 €"


def normalize_column(value: str) -> str:
    """Normalise un nom de colonne pour le mapping automatique."""
    text = safe_text(value).lower()

    replacements = {
        "é": "e",
        "è": "e",
        "ê": "e",
        "ë": "e",
        "à": "a",
        "â": "a",
        "ä": "a",
        "î": "i",
        "ï": "i",
        "ô": "o",
        "ö": "o",
        "ù": "u",
        "û": "u",
        "ü": "u",
        "ç": "c",
    }

    for source, target in replacements.items():
        text = text.replace(source, target)

    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text)

    return text.strip("_")


def make_id(*values: Any) -> str:
    """Crée un identifiant stable à partir de plusieurs valeurs."""
    raw = "|".join(safe_text(v) for v in values)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()[:16]


def carrier_key(carrier: str) -> str:
    """Normalise le nom du transporteur."""
    value = safe_text(carrier).lower()

    if "colissimo" in value:
        return "Colissimo"

    if "dhl" in value:
        return "DHL"

    if "ups" in value:
        return "UPS"

    return "Autre"


# ============================================================
# STRUCTURE D'UN COLIS
# ============================================================

def build_package(
    tracking: str,
    carrier: str,
    order_id: str,
    shipped_date: Optional[date],
    promised_date: Optional[date],
    delivered_date: Optional[date],
    shipping_cost: float,
    order_value: float,
    customer: str,
    city: str,
    service: str,
    source: str = "demo",
) -> Dict[str, Any]:

    package_id = make_id(
        tracking,
        carrier,
        order_id,
    )

    return {
        "id": package_id,
        "tracking": tracking,
        "carrier": carrier_key(carrier),
        "order_id": order_id,
        "shipped_date": shipped_date,
        "promised_date": promised_date,
        "delivered_date": delivered_date,
        "shipping_cost": round(shipping_cost, 2),
        "order_value": round(order_value, 2),
        "customer": customer,
        "city": city,
        "service": service,
        "source": source,

        # Champs calculés par le moteur
        "delay_days": 0,
        "is_lost": False,
        "status": "not_eligible",
        "status_label": STATUS_LABELS["not_eligible"],
        "estimated_refund": 0.0,
        "reason": "",
        "analyzed_at": None,
        "claim_prepared": False,
    }


# ============================================================
# DONNÉES DE DÉMONSTRATION
# ============================================================

def demo_packages() -> List[Dict[str, Any]]:
    """
    Génère 10 colis synthétiques.

    Ces données ne correspondent à aucun client réel.
    """

    today = date.today()

    data = [
        {
            "tracking": "DEMO-COL-0001",
            "carrier": "Colissimo",
            "order": "DEMO-ORD-1001",
            "shipped": today - timedelta(days=9),
            "promised": today - timedelta(days=5),
            "delivered": None,
            "shipping_cost": 8.90,
            "order_value": 129.90,
            "customer": "Client Démo 01",
            "city": "Paris",
            "service": "Colissimo Domicile",
        },
        {
            "tracking": "DEMO-DHL-0002",
            "carrier": "DHL",
            "order": "DEMO-ORD-1002",
            "shipped": today - timedelta(days=8),
            "promised": today - timedelta(days=4),
            "delivered": today - timedelta(days=1),
            "shipping_cost": 18.50,
            "order_value": 249.00,
            "customer": "Client Démo 02",
            "city": "Lyon",
            "service": "DHL Express",
        },
        {
            "tracking": "DEMO-UPS-0003",
            "carrier": "UPS",
            "order": "DEMO-ORD-1003",
            "shipped": today - timedelta(days=6),
            "promised": today - timedelta(days=2),
            "delivered": today - timedelta(days=1),
            "shipping_cost": 14.90,
            "order_value": 89.00,
            "customer": "Client Démo 03",
            "city": "Marseille",
            "service": "UPS Standard",
        },
        {
            "tracking": "DEMO-COL-0004",
            "carrier": "Colissimo",
            "order": "DEMO-ORD-1004",
            "shipped": today - timedelta(days=5),
            "promised": today - timedelta(days=2),
            "delivered": today - timedelta(days=2),
            "shipping_cost": 7.90,
            "order_value": 59.90,
            "customer": "Client Démo 04",
            "city": "Bordeaux",
            "service": "Colissimo Domicile",
        },
        {
            "tracking": "DEMO-DHL-0005",
            "carrier": "DHL",
            "order": "DEMO-ORD-1005",
            "shipped": today - timedelta(days=11),
            "promised": today - timedelta(days=7),
            "delivered": None,
            "shipping_cost": 21.00,
            "order_value": 399.00,
            "customer": "Client Démo 05",
            "city": "Nantes",
            "service": "DHL Express",
        },
        {
            "tracking": "DEMO-UPS-0006",
            "carrier": "UPS",
            "order": "DEMO-ORD-1006",
            "shipped": today - timedelta(days=7),
            "promised": today - timedelta(days=3),
            "delivered": today,
            "shipping_cost": 12.90,
            "order_value": 159.90,
            "customer": "Client Démo 06",
            "city": "Toulouse",
            "service": "UPS Standard",
        },
        {
            "tracking": "DEMO-COL-0007",
            "carrier": "Colissimo",
            "order": "DEMO-ORD-1007",
            "shipped": today - timedelta(days=13),
            "promised": today - timedelta(days=10),
            "delivered": None,
            "shipping_cost": 9.90,
            "order_value": 199.90,
            "customer": "Client Démo 07",
            "city": "Lille",
            "service": "Colissimo Domicile",
        },
        {
            "tracking": "DEMO-DHL-0008",
            "carrier": "DHL",
            "order": "DEMO-ORD-1008",
            "shipped": today - timedelta(days=5),
            "promised": today - timedelta(days=1),
            "delivered": today,
            "shipping_cost": 19.90,
            "order_value": 299.00,
            "customer": "Client Démo 08",
            "city": "Strasbourg",
            "service": "DHL Express",
        },
        {
            "tracking": "DEMO-UPS-0009",
            "carrier": "UPS",
            "order": "DEMO-ORD-1009",
            "shipped": today - timedelta(days=4),
            "promised": today - timedelta(days=1),
            "delivered": None,
            "shipping_cost": 16.50,
            "order_value": 179.00,
            "customer": "Client Démo 09",
            "city": "Nice",
            "service": "UPS Standard",
        },
        {
            "tracking": "DEMO-COL-0010",
            "carrier": "Colissimo",
            "order": "DEMO-ORD-1010",
            "shipped": today - timedelta(days=2),
            "promised": today,
            "delivered": today,
            "shipping_cost": 6.90,
            "order_value": 49.90,
            "customer": "Client Démo 10",
            "city": "Rennes",
            "service": "Colissimo Domicile",
        },
    ]

    packages = []

    for item in data:
        packages.append(
            build_package(
                tracking=item["tracking"],
                carrier=item["carrier"],
                order_id=item["order"],
                shipped_date=item["shipped"],
                promised_date=item["promised"],
                delivered_date=item["delivered"],
                shipping_cost=item["shipping_cost"],
                order_value=item["order_value"],
                customer=item["customer"],
                city=item["city"],
                service=item["service"],
                source="demo",
            )
        )

    return packages


# ============================================================
# MOTEUR DE DÉTECTION
# ============================================================

def analyze_package(
    package: Dict[str, Any],
    rules: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """
    Analyse un colis.

    Logique :
    - Si livré après la date promise : calcul du retard.
    - Si non livré et que le nombre de jours depuis la date promise
      dépasse le seuil de perte : considéré comme "perte potentielle".
    - Si le retard atteint le seuil configuré : éligible potentiel.
    - Si retard intermédiaire : à vérifier.
    """

    result = dict(package)

    now = date.today()
    carrier = carrier_key(result.get("carrier", "Autre"))

    carrier_rule = rules.get(
        carrier,
        rules.get("Autre", DEFAULT_CARRIER_RULES["Autre"]),
    )

    late_threshold = int(
        safe_float(
            carrier_rule.get("late_days", 5),
            5,
        )
    )

    lost_threshold = int(
        safe_float(
            carrier_rule.get("lost_days", 15),
            15,
        )
    )

    refund_cap = safe_float(
        carrier_rule.get("cap", 150.0),
        150.0,
    )

    promised = result.get("promised_date")
    delivered = result.get("delivered_date")

    delay_days = 0
    is_lost = False

    if promised is not None:

        if delivered is not None:
            delta = (delivered - promised).days
            delay_days = max(0, delta)

        else:
            delta = (now - promised).days
            delay_days = max(0, delta)

            if delay_days >= lost_threshold:
                is_lost = True

    result["delay_days"] = delay_days
    result["is_lost"] = is_lost
    result["analyzed_at"] = datetime.now().isoformat(timespec="seconds")

    shipping_cost = safe_float(
        result.get("shipping_cost", 0),
        0,
    )

    estimated_refund = min(
        shipping_cost,
        refund_cap,
    )

    # --------------------------------------------------------
    # Classification
    # --------------------------------------------------------

    if is_lost:
        status = "eligible"
        reason = (
            f"Colis non livré depuis {delay_days} jours après "
            f"la date promise. Seuil de perte configuré : "
            f"{lost_threshold} jours."
        )

    elif delay_days >= late_threshold:
        status = "eligible"
        reason = (
            f"Retard détecté de {delay_days} jours. "
            f"Seuil configuré : {late_threshold} jours."
        )

    elif delay_days > 0:
        status = "review"
        reason = (
            f"Retard de {delay_days} jour(s), inférieur "
            f"au seuil d'éligibilité configuré."
        )

    else:
        status = "not_eligible"
        reason = "Aucun retard détecté selon les règles configurées."

    # Une réclamation déjà préparée conserve son statut.
    if result.get("claim_prepared"):
        status = "claimed"

    result["status"] = status
    result["status_label"] = STATUS_LABELS.get(
        status,
        status,
    )
    result["estimated_refund"] = (
        estimated_refund
        if status in {"eligible", "claimed"}
        else 0.0
    )
    result["reason"] = reason

    return result


def analyze_packages(
    packages: List[Dict[str, Any]],
    rules: Dict[str, Dict[str, Any]],
    show_progress: bool = True,
) -> List[Dict[str, Any]]:
    """
    Analyse tous les colis avec barre de progression.
    """

    results = []

    progress = None
    status_box = None

    if show_progress:
        status_box = st.empty()
        progress = st.progress(0)

    total = max(len(packages), 1)

    for index, package in enumerate(packages):

        try:
            result = analyze_package(
                package,
                rules,
            )

            results.append(result)

        except Exception as exc:
            # Sécurité : un colis défectueux ne doit pas faire
            # planter toute l'analyse.
            fallback = dict(package)
            fallback["status"] = "review"
            fallback["status_label"] = STATUS_LABELS["review"]
            fallback["reason"] = (
                f"Analyse impossible pour ce colis : {exc}"
            )
            fallback["estimated_refund"] = 0.0

            results.append(fallback)

        if progress is not None:
            progress.progress(
                int(((index + 1) / total) * 100)
            )

        if status_box is not None:
            status_box.info(
                f"Analyse du colis {index + 1}/{len(packages)}..."
            )

    if status_box is not None:
        status_box.success(
            f"Analyse terminée — {len(results)} colis traités."
        )

    if progress is not None:
        progress.empty()

    return results


# ============================================================
# IMPORT CSV / XLSX
# ============================================================

def find_column_mapping(
    dataframe: pd.DataFrame,
) -> Dict[str, str]:
    """
    Détecte automatiquement les colonnes connues.
    """

    normalized_columns = {
        normalize_column(column): column
        for column in dataframe.columns
    }

    mapping = {}

    for internal_name, aliases in COLUMN_ALIASES.items():

        for alias in aliases:

            normalized_alias = normalize_column(alias)

            if normalized_alias in normalized_columns:
                mapping[internal_name] = normalized_columns[
                    normalized_alias
                ]
                break

    return mapping


def read_uploaded_file(
    uploaded_file: Any,
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Lit CSV ou XLSX de manière contrôlée.
    """

    try:
        if uploaded_file is None:
            return None, "Aucun fichier fourni."

        file_size = getattr(
            uploaded_file,
            "size",
            0,
        )

        if file_size and file_size > MAX_UPLOAD_BYTES:
            return (
                None,
                (
                    f"Fichier trop volumineux. "
                    f"Limite : {MAX_UPLOAD_BYTES / 1024 / 1024:.0f} Mo."
                ),
            )

        filename = safe_text(
            getattr(uploaded_file, "name", "")
        ).lower()

        raw = uploaded_file.getvalue()

        if filename.endswith(".csv"):
            try:
                dataframe = pd.read_csv(
                    io.BytesIO(raw),
                    sep=None,
                    engine="python",
                )
            except Exception:
                dataframe = pd.read_csv(
                    io.BytesIO(raw),
                    sep=";",
                )

        elif filename.endswith(".xlsx"):
            dataframe = pd.read_excel(
                io.BytesIO(raw),
                engine="openpyxl",
            )

        else:
            return (
                None,
                "Format non supporté. Utilise CSV ou XLSX.",
            )

        if len(dataframe) > MAX_ROWS:
            return (
                None,
                (
                    f"Le fichier contient {len(dataframe)} lignes. "
                    f"Limite de sécurité : {MAX_ROWS}."
                ),
            )

        if dataframe.empty:
            return None, "Le fichier ne contient aucune ligne."

        return dataframe, None

    except Exception as exc:
        return (
            None,
            f"Impossible de lire le fichier : {exc}",
        )


def dataframe_to_packages(
    dataframe: pd.DataFrame,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """
    Transforme le DataFrame importé en objets colis.
    """

    mapping = find_column_mapping(dataframe)

    warnings = []
    packages = []

    if "tracking" not in mapping:
        warnings.append(
            "Colonne de numéro de suivi non détectée. "
            "Un identifiant automatique sera généré."
        )

    if "promised" not in mapping:
        warnings.append(
            "Date de livraison prévue non détectée. "
            "Une date estimée sera utilisée."
        )

    for index, row in dataframe.iterrows():

        try:
            def value_of(field: str, default: Any = "") -> Any:
                column = mapping.get(field)

                if not column:
                    return default

                try:
                    return row[column]
                except Exception:
                    return default

            tracking = safe_text(
                value_of("tracking"),
                f"IMPORT-{index + 1:05d}",
            )

            carrier = carrier_key(
                safe_text(
                    value_of("carrier"),
                    "Autre",
                )
            )

            order_id = safe_text(
                value_of("order"),
                f"ORDER-{index + 1:05d}",
            )

            shipped_date = safe_date(
                value_of("shipped")
            )

            if shipped_date is None:
                shipped_date = date.today() - timedelta(days=5)

            promised_date = safe_date(
                value_of("promised")
            )

            if promised_date is None:
                promised_date = shipped_date + timedelta(days=3)

            delivered_date = safe_date(
                value_of("delivered")
            )

            shipping_cost = safe_float(
                value_of("shipping_cost"),
                0.0,
            )

            order_value = safe_float(
                value_of("value"),
                0.0,
            )

            customer = safe_text(
                value_of("customer"),
                "Client",
            )

            city = safe_text(
                value_of("city"),
                "",
            )

            service = safe_text(
                value_of("service"),
                "Service standard",
            )

            package = build_package(
                tracking=tracking,
                carrier=carrier,
                order_id=order_id,
                shipped_date=shipped_date,
                promised_date=promised_date,
                delivered_date=delivered_date,
                shipping_cost=shipping_cost,
                order_value=order_value,
                customer=customer,
                city=city,
                service=service,
                source="import",
            )

            packages.append(package)

        except Exception as exc:
            warnings.append(
                f"Ligne {index + 2} ignorée : {exc}"
            )

    return packages, warnings


# ============================================================
# SESSION STATE
# ============================================================

def initialize_session() -> None:

    if "initialized" not in st.session_state:
        st.session_state.initialized = True

        st.session_state.packages = demo_packages()

        st.session_state.page = "Dashboard"

        st.session_state.scan_history = []

        st.session_state.company_name = (
            "[NOM DE LA SOCIÉTÉ]"
        )

        st.session_state.company_email = (
            "[EMAIL]"
        )

        st.session_state.company_siret = (
            "[SIRET]"
        )

        st.session_state.host = (
            "[HÉBERGEUR]"
        )

        st.session_state.carrier_rules = {
            key: dict(value)
            for key, value in DEFAULT_CARRIER_RULES.items()
        }


def reset_demo_data() -> None:

    st.session_state.packages = demo_packages()

    st.session_state.scan_history = []

    st.success(
        "Données de démonstration réinitialisées."
    )


# ============================================================
# CSS
# ============================================================

def inject_css() -> None:

    st.markdown(
        f"""
        <style>

        @import url(
            'https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap'
        );

        html,
        body,
        [class*="css"] {{
            font-family: Inter, sans-serif;
        }}

        .stApp {{
            background:
                radial-gradient(
                    circle at 20% 0%,
                    rgba(138,43,226,0.08),
                    transparent 30%
                ),
                radial-gradient(
                    circle at 90% 10%,
                    rgba(0,255,204,0.06),
                    transparent 30%
                ),
                {BG};
            color: {TEXT};
        }}

        section[data-testid="stSidebar"] {{
            background: #0A0E13;
            border-right: 1px solid rgba(138,43,226,0.35);
        }}

        section[data-testid="stSidebar"] * {{
            color: {TEXT};
        }}

        .brand {{
            font-size: 25px;
            font-weight: 800;
            letter-spacing: -0.8px;
            margin-bottom: 3px;
        }}

        .brand span {{
            color: {GREEN};
        }}

        .version {{
            color: {MUTED};
            font-size: 11px;
            margin-bottom: 22px;
        }}

        .page-title {{
            font-size: 32px;
            font-weight: 800;
            letter-spacing: -1px;
            margin-bottom: 2px;
        }}

        .page-subtitle {{
            color: {MUTED};
            font-size: 14px;
            margin-bottom: 25px;
        }}

        .card {{
            background: linear-gradient(
                145deg,
                rgba(17,24,32,0.98),
                rgba(21,29,39,0.92)
            );
            border: 1px solid rgba(138,43,226,0.28);
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.16);
        }}

        .kpi-label {{
            color: {MUTED};
            font-size: 12px;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            font-weight: 700;
        }}

        .kpi-value {{
            font-size: 29px;
            font-weight: 800;
            margin-top: 7px;
        }}

        .kpi-value.green {{
            color: {GREEN};
        }}

        .kpi-value.purple {{
            color: #B987FF;
        }}

        .kpi-value.orange {{
            color: {ORANGE};
        }}

        .section-title {{
            font-size: 18px;
            font-weight: 750;
            margin-top: 18px;
            margin-bottom: 10px;
        }}

        .muted {{
            color: {MUTED};
        }}

        .alert-box {{
            border-radius: 10px;
            padding: 12px 15px;
            margin: 10px 0;
            background: rgba(255,204,102,0.07);
            border: 1px solid rgba(255,204,102,0.25);
            color: #F8E5B2;
        }}

        .success-box {{
            border-radius: 10px;
            padding: 12px 15px;
            margin: 10px 0;
            background: rgba(0,255,204,0.06);
            border: 1px solid rgba(0,255,204,0.25);
            color: {GREEN};
        }}

        .footer {{
            margin-top: 50px;
            padding-top: 20px;
            border-top: 1px solid rgba(255,255,255,0.08);
            color: {MUTED};
            font-size: 11px;
            text-align: center;
        }}

        div[data-testid="stMetric"] {{
            background: rgba(17,24,32,0.95);
            border: 1px solid rgba(138,43,226,0.25);
            padding: 14px;
            border-radius: 12px;
        }}

        div[data-testid="stFileUploader"] {{
            border: 1px dashed rgba(0,255,204,0.35);
            border-radius: 12px;
            padding: 8px;
            background: rgba(0,255,204,0.02);
        }}

        .stButton > button {{
            border-radius: 9px;
            border: 1px solid rgba(138,43,226,0.45);
            background: #141B24;
            color: {TEXT};
            font-weight: 650;
        }}

        .stButton > button:hover {{
            border-color: {GREEN};
            color: {GREEN};
        }}

        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# HEADER / FOOTER
# ============================================================

def page_header(
    title: str,
    subtitle: str,
) -> None:

    st.markdown(
        f"""
        <div class="page-title">{title}</div>
        <div class="page-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True,
    )


def footer() -> None:

    st.markdown(
        f"""
        <div class="footer">
            {APP_NAME} v{APP_VERSION}
            · Outil de démonstration B2B
            · Les règles de remboursement doivent être vérifiées
              selon les contrats applicables.
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# DASHBOARD
# ============================================================

def dashboard_page() -> None:

    page_header(
        "🖥️ Dashboard",
        "Vue d'ensemble de la récupération potentielle de frais de transport.",
    )

    packages = st.session_state.packages

    total = len(packages)

    eligible = [
        p
        for p in packages
        if p.get("status") in {"eligible", "claimed"}
    ]

    delayed = [
        p
        for p in packages
        if safe_float(p.get("delay_days")) > 0
    ]

    potential_gain = sum(
        safe_float(p.get("estimated_refund"))
        for p in eligible
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            f"""
            <div class="card">
                <div class="kpi-label">Total Colis Analysés</div>
                <div class="kpi-value">{total:,}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            f"""
            <div class="card">
                <div class="kpi-label">Colis en Retard Détectés</div>
                <div class="kpi-value orange">{len(delayed):,}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            f"""
            <div class="card">
                <div class="kpi-label">
                    Montant Éligible Potentiel
                </div>
                <div class="kpi-value green">
                    {money(potential_gain)}
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="section-title">💰 Gains potentiels par transporteur</div>',
        unsafe_allow_html=True,
    )

    carrier_totals = {}

    for package in eligible:
        carrier = carrier_key(
            package.get("carrier", "Autre")
        )

        carrier_totals[carrier] = (
            carrier_totals.get(carrier, 0)
            + safe_float(package.get("estimated_refund"))
        )

    if carrier_totals:

        chart_data = pd.DataFrame(
            {
                "Transporteur": list(carrier_totals.keys()),
                "Montant potentiel (€)": list(
                    carrier_totals.values()
                ),
            }
        ).set_index("Transporteur")

        st.bar_chart(
            chart_data,
            use_container_width=True,
        )

    else:

        st.info(
            "Aucun gain potentiel détecté pour le moment."
        )

    st.markdown(
        '<div class="section-title">📊 Répartition des dossiers</div>',
        unsafe_allow_html=True,
    )

    stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)

    counts = {
        "eligible": sum(
            1
            for p in packages
            if p.get("status") == "eligible"
        ),
        "review": sum(
            1
            for p in packages
            if p.get("status") == "review"
        ),
        "not_eligible": sum(
            1
            for p in packages
            if p.get("status") == "not_eligible"
        ),
        "claimed": sum(
            1
            for p in packages
            if p.get("status") == "claimed"
        ),
    }

    with stats_col1:
        st.metric(
            "🎯 Éligibles",
            counts["eligible"],
        )

    with stats_col2:
        st.metric(
            "🔎 À vérifier",
            counts["review"],
        )

    with stats_col3:
        st.metric(
            "✓ Non éligibles",
            counts["not_eligible"],
        )

    with stats_col4:
        st.metric(
            "📑 Réclamations préparées",
            counts["claimed"],
        )

    if st.session_state.scan_history:

        st.markdown(
            '<div class="section-title">🕘 Historique des analyses</div>',
            unsafe_allow_html=True,
        )

        history_df = pd.DataFrame(
            st.session_state.scan_history
        )

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True,
        )


# ============================================================
# SCANNER
# ============================================================

def scanner_page() -> None:

    page_header(
        "🎯 Scanner de Colis",
        "Importe un export transporteur ou utilise le jeu de démonstration.",
    )

    st.markdown(
        """
        <div class="alert-box">
            <strong>Mode sécurisé :</strong>
            les fichiers sont traités uniquement pendant la session.
            Évite d'importer des données inutiles ou sensibles.
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Déposer un fichier CSV ou XLSX",
        type=["csv", "xlsx"],
        help=(
            "Taille maximale : 15 Mo · "
            "5000 lignes maximum."
        ),
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button(
            "🚀 Lancer le scan de démonstration",
            use_container_width=True,
        ):

            try:

                packages = demo_packages()

                results = analyze_packages(
                    packages,
                    st.session_state.carrier_rules,
                    show_progress=True,
                )

                st.session_state.packages = results

                st.session_state.scan_history.append(
                    {
                        "Date": datetime.now().strftime(
                            "%d/%m/%Y %H:%M"
                        ),
                        "Source": "Démonstration",
                        "Colis analysés": len(results),
                        "Dossiers potentiels": sum(
                            1
                            for p in results
                            if p.get("status") == "eligible"
                        ),
                        "Montant potentiel": money(
                            sum(
                                safe_float(
                                    p.get("estimated_refund")
                                )
                                for p in results
                            )
                        ),
                    }
                )

                st.rerun()

            except Exception as exc:

                st.error(
                    f"Le scan de démonstration a échoué : {exc}"
                )

    with col2:

        if uploaded_file is not None:

            if st.button(
                "🔍 Analyser le fichier importé",
                use_container_width=True,
            ):

                dataframe, error = read_uploaded_file(
                    uploaded_file
                )

                if error:

                    st.error(error)

                elif dataframe is not None:

                    with st.expander(
                        "Aperçu du fichier",
                        expanded=True,
                    ):

                        st.dataframe(
                            dataframe.head(20),
                            use_container_width=True,
                            hide_index=True,
                        )

                    packages, warnings = (
                        dataframe_to_packages(
                            dataframe
                        )
                    )

                    for warning in warnings:
                        st.warning(warning)

                    if not packages:

                        st.error(
                            "Aucun colis exploitable détecté."
                        )

                    else:

                        results = analyze_packages(
                            packages,
                            st.session_state.carrier_rules,
                            show_progress=True,
                        )

                        st.session_state.packages = results

                        st.session_state.scan_history.append(
                            {
                                "Date": datetime.now().strftime(
                                    "%d/%m/%Y %H:%M"
                                ),
                                "Source": uploaded_file.name,
                                "Colis analysés": len(results),
                                "Dossiers potentiels": sum(
                                    1
                                    for p in results
                                    if p.get("status") == "eligible"
                                ),
                                "Montant potentiel": money(
                                    sum(
                                        safe_float(
                                            p.get(
                                                "estimated_refund"
                                            )
                                        )
                                        for p in results
                                    )
                                ),
                            }
                        )

                        st.success(
                            f"{len(results)} colis analysés."
                        )

    # --------------------------------------------------------
    # Résultats
    # --------------------------------------------------------

    st.markdown(
        '<div class="section-title">📦 Résultats du scan</div>',
        unsafe_allow_html=True,
    )

    packages = st.session_state.packages

    if not packages:
        st.info("Aucun colis à afficher.")
        return

    filter_col1, filter_col2, filter_col3 = st.columns(3)

    carriers = sorted(
        {
            carrier_key(
                p.get("carrier", "Autre")
            )
            for p in packages
        }
    )

    with filter_col1:

        selected_carriers = st.multiselect(
            "Transporteur",
            options=carriers,
            default=carriers,
        )

    with filter_col2:

        selected_statuses = st.multiselect(
            "Statut",
            options=list(STATUS_LABELS.values()),
            default=list(STATUS_LABELS.values()),
        )

    with filter_col3:

        search = st.text_input(
            "Recherche",
            placeholder="Tracking, commande, client...",
        )

    filtered = []

    for package in packages:

        carrier = carrier_key(
            package.get("carrier", "Autre")
        )

        status_label = package.get(
            "status_label",
            STATUS_LABELS["not_eligible"],
        )

        if carrier not in selected_carriers:
            continue

        if status_label not in selected_statuses:
            continue

        if search:

            haystack = " ".join(
                [
                    safe_text(package.get("tracking")),
                    safe_text(package.get("order_id")),
                    safe_text(package.get("customer")),
                    safe_text(package.get("city")),
                ]
            ).lower()

            if search.lower() not in haystack:
                continue

        filtered.append(package)

    rows = []

    for p in filtered:

        promised = p.get("promised_date")
        delivered = p.get("delivered_date")

        rows.append(
            {
                "Statut": (
                    STATUS_EMOJI.get(
                        p.get("status"),
                        "•",
                    )
                    + " "
                    + safe_text(
                        p.get("status_label")
                    )
                ),
                "Transporteur": carrier_key(
                    p.get("carrier", "Autre")
                ),
                "Tracking": p.get("tracking"),
                "Commande": p.get("order_id"),
                "Date prévue": (
                    promised.strftime("%d/%m/%Y")
                    if isinstance(promised, date)
                    else "-"
                ),
                "Date livraison": (
                    delivered.strftime("%d/%m/%Y")
                    if isinstance(delivered, date)
                    else "Non livré"
                ),
                "Retard": (
                    f"{safe_float(p.get('delay_days')):.0f} j"
                ),
                "Frais transport": money(
                    p.get("shipping_cost")
                ),
                "Remboursement potentiel": money(
                    p.get("estimated_refund")
                ),
            }
        )

    if rows:

        results_df = pd.DataFrame(rows)

        st.dataframe(
            results_df,
            use_container_width=True,
            hide_index=True,
        )

        st.caption(
            f"{len(rows)} résultat(s) affiché(s) sur "
            f"{len(packages)} colis."
        )

        # Détails des anomalies
        eligible_packages = [
            p
            for p in filtered
            if p.get("status") in {"eligible", "claimed"}
        ]

        if eligible_packages:

            st.markdown(
                '<div class="section-title">🎯 Détails des anomalies détectées</div>',
                unsafe_allow_html=True,
            )

            for package in eligible_packages:

                with st.expander(
                    f"{package.get('tracking')} — "
                    f"{package.get('status_label')}"
                ):

                    detail_col1, detail_col2 = st.columns(2)

                    with detail_col1:

                        st.write(
                            f"**Transporteur :** "
                            f"{carrier_key(package.get('carrier'))}"
                        )

                        st.write(
                            f"**Commande :** "
                            f"{package.get('order_id')}"
                        )

                        st.write(
                            f"**Client :** "
                            f"{package.get('customer')}"
                        )

                        st.write(
                            f"**Destination :** "
                            f"{package.get('city') or '-'}"
                        )

                    with detail_col2:

                        st.write(
                            f"**Retard :** "
                            f"{package.get('delay_days', 0)} jour(s)"
                        )

                        st.write(
                            f"**Frais transport :** "
                            f"{money(package.get('shipping_cost'))}"
                        )

                        st.write(
                            f"**Montant potentiel :** "
                            f"{money(package.get('estimated_refund'))}"
                        )

                        st.write(
                            f"**Motif :** "
                            f"{package.get('reason', '-')}"
                        )

            st.warning(
                "⚠️ Une détection « éligible » est une "
                "indication à vérifier. Elle ne signifie ni "
                "fraude, ni remboursement garanti."
            )

    else:

        st.info(
            "Aucun colis ne correspond aux filtres."
        )


# ============================================================
# EXPORTS
# ============================================================

def packages_to_dataframe(
    packages: List[Dict[str, Any]],
) -> pd.DataFrame:

    rows = []

    for p in packages:

        rows.append(
            {
                "tracking": p.get("tracking"),
                "carrier": carrier_key(
                    p.get("carrier", "Autre")
                ),
                "order_id": p.get("order_id"),
                "customer": p.get("customer"),
                "city": p.get("city"),
                "service": p.get("service"),
                "shipped_date": p.get("shipped_date"),
                "promised_date": p.get("promised_date"),
                "delivered_date": p.get("delivered_date"),
                "delay_days": p.get("delay_days"),
                "is_lost": p.get("is_lost"),
                "status": p.get("status_label"),
                "reason": p.get("reason"),
                "shipping_cost": p.get("shipping_cost"),
                "estimated_refund": p.get("estimated_refund"),
                "order_value": p.get("order_value"),
                "source": p.get("source"),
                "claim_prepared": p.get(
                    "claim_prepared",
                    False,
                ),
            }
        )

    return pd.DataFrame(rows)


def dataframe_to_xlsx_bytes(
    dataframe: pd.DataFrame,
) -> bytes:

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="RefundSniper",
        )

    output.seek(0)

    return output.getvalue()


# ============================================================
# LETTRE DE RÉCLAMATION
# ============================================================

def generate_claim_letter(
    package: Dict[str, Any],
) -> str:

    company = st.session_state.get(
        "company_name",
        "[NOM DE LA SOCIÉTÉ]",
    )

    email = st.session_state.get(
        "company_email",
        "[EMAIL]",
    )

    tracking = safe_text(
        package.get("tracking"),
        "[NUMÉRO DE SUIVI]",
    )

    carrier = carrier_key(
        package.get("carrier", "Autre")
    )

    order_id = safe_text(
        package.get("order_id"),
        "[RÉFÉRENCE COMMANDE]",
    )

    promised = package.get("promised_date")

    delivered = package.get("delivered_date")

    promised_text = (
        promised.strftime("%d/%m/%Y")
        if isinstance(promised, date)
        else "[DATE PRÉVUE]"
    )

    delivered_text = (
        delivered.strftime("%d/%m/%Y")
        if isinstance(delivered, date)
        else "Non livré à la date de rédaction"
    )

    delay_days = int(
        safe_float(
            package.get("delay_days"),
            0,
        )
    )

    shipping_cost = money(
        package.get("shipping_cost")
    )

    estimated_refund = money(
        package.get("estimated_refund")
    )

    return f"""Objet : Demande de prise en charge — envoi {tracking}

Madame, Monsieur,

Nous vous contactons au nom de {company} concernant l'expédition
référencée {tracking}, associée à la commande {order_id}.

Transporteur : {carrier}
Date de livraison prévue : {promised_text}
Date de livraison constatée : {delivered_text}
Retard constaté : {delay_days} jour(s)
Frais de transport facturés : {shipping_cost}

Au regard des informations actuellement disponibles, cet envoi
a été identifié par notre système comme pouvant nécessiter
une vérification au regard des conditions applicables à notre
contrat de transport.

Nous vous remercions de bien vouloir examiner ce dossier et de
nous confirmer les éventuelles modalités de prise en charge
applicables.

Montant potentiellement concerné selon notre configuration
interne : {estimated_refund}

Vous trouverez les informations de suivi nécessaires à l'étude
du dossier dans notre export de réclamation.

Nous restons disponibles pour transmettre tout justificatif
complémentaire nécessaire.

Cordialement,

{company}
{email}

---
Modèle généré par {APP_NAME}.
Ce document est un modèle et doit être vérifié avant envoi.
"""


# ============================================================
# PAGE EXPORT & RÉCLAMATIONS
# ============================================================

def claims_page() -> None:

    page_header(
        "📑 Export & Réclamations",
        "Prépare les dossiers détectés et génère les modèles de courrier.",
    )

    packages = st.session_state.packages

    eligible = [
        p
        for p in packages
        if p.get("status") in {"eligible", "claimed"}
    ]

    st.metric(
        "Dossiers potentiellement réclamables",
        len(eligible),
    )

    total = sum(
        safe_float(
            p.get("estimated_refund")
        )
        for p in eligible
    )

    st.metric(
        "Montant potentiel cumulé",
        money(total),
    )

    st.markdown(
        '<div class="section-title">📤 Export des dossiers</div>',
        unsafe_allow_html=True,
    )

    export_df = packages_to_dataframe(
        eligible
    )

    if not export_df.empty:

        csv_bytes = export_df.to_csv(
            index=False
        ).encode("utf-8-sig")

        st.download_button(
            "⬇️ Télécharger CSV",
            data=csv_bytes,
            file_name=(
                "refundsniper_reclamations.csv"
            ),
            mime="text/csv",
            use_container_width=True,
        )

        xlsx_bytes = dataframe_to_xlsx_bytes(
            export_df
        )

        st.download_button(
            "⬇️ Télécharger Excel",
            data=xlsx_bytes,
            file_name=(
                "refundsniper_reclamations.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

    else:

        st.info(
            "Aucun dossier potentiellement réclamable."
        )

    st.markdown(
        '<div class="section-title">✉️ Modèles de réclamation</div>',
        unsafe_allow_html=True,
    )

    if not eligible:

        st.info(
            "Lance d'abord une analyse pour générer des dossiers."
        )

        return

    for package in eligible:

        tracking = safe_text(
            package.get("tracking")
        )

        with st.expander(
            f"📑 {tracking} — "
            f"{money(package.get('estimated_refund'))}"
        ):

            letter = generate_claim_letter(
                package
            )

            st.text_area(
                "Courrier généré",
                value=letter,
                height=350,
                key=f"letter_{package.get('id')}",
            )

            st.download_button(
                "⬇️ Télécharger le courrier .txt",
                data=letter.encode("utf-8"),
                file_name=(
                    f"reclamation_{tracking}.txt"
                ),
                mime="text/plain",
                key=f"download_letter_{package.get('id')}",
            )

            if not package.get(
                "claim_prepared",
                False,
            ):

                if st.button(
                    "📌 Marquer comme préparée",
                    key=f"claim_{package.get('id')}",
                ):

                    for item in st.session_state.packages:

                        if item.get("id") == package.get("id"):

                            item["claim_prepared"] = True
                            item["status"] = "claimed"
                            item["status_label"] = (
                                STATUS_LABELS["claimed"]
                            )

                    st.success(
                        "Dossier marqué comme préparé."
                    )

                    st.rerun()

            else:

                st.success(
                    "Réclamation déjà préparée."
                )


# ============================================================
# MENTIONS LÉGALES / CGU
# ============================================================

def legal_page() -> None:

    page_header(
        "⚖️ Mentions légales & CGU",
        "Configuration juridique et informations de conformité.",
    )

    st.markdown(
        """
        <div class="alert-box">
            <strong>Important :</strong>
            les textes ci-dessous constituent une base de travail
            pour le produit. Ils doivent être adaptés à la société
            qui exploite RefundSniper et faire l'objet d'une
            vérification juridique professionnelle avant publication.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="section-title">🏢 Informations de l'éditeur</div>',
        unsafe_allow_html=True,
    )

    company_col1, company_col2 = st.columns(2)

    with company_col1:

        st.session_state.company_name = st.text_input(
            "Nom de la société",
            value=st.session_state.company_name,
        )

        st.session_state.company_email = st.text_input(
            "Email de contact",
            value=st.session_state.company_email,
        )

    with company_col2:

        st.session_state.company_siret = st.text_input(
            "SIRET",
            value=st.session_state.company_siret,
        )

        st.session_state.host = st.text_input(
            "Hébergeur",
            value=st.session_state.host,
        )

    tabs = st.tabs(
        [
            "Mentions légales",
            "Confidentialité",
            "CGU",
        ]
    )

    with tabs[0]:

        st.markdown(
            f"""
            ## Mentions légales

            **Éditeur**

            {st.session_state.company_name}

            SIRET : {st.session_state.company_siret}

            Email : {st.session_state.company_email}

            **Hébergement**

            {st.session_state.host}

            Les coordonnées exactes de l'hébergeur doivent être
            complétées avant publication.

            **Propriété intellectuelle**

            L'ensemble des éléments graphiques, logiciels, textes,
            logos et interfaces de l'application doivent être
            identifiés et attribués conformément aux droits
            applicables.

            **Responsabilité**

            RefundSniper fournit des outils d'analyse et de
            préparation de dossiers. Une détection automatique
            ne constitue pas une décision contractuelle du
            transporteur et ne garantit pas l'obtention d'un
            remboursement.
            """
        )

    with tabs[1]:

        st.markdown(
            """
            ## Politique de confidentialité

            RefundSniper doit appliquer un principe de minimisation
            des données.

            Les données traitées peuvent notamment comprendre :

            - numéro de suivi ;
            - référence de commande ;
            - informations de livraison ;
            - dates d'expédition et de livraison ;
            - frais de transport ;
            - informations nécessaires au traitement du dossier.

            Les données qui ne sont pas nécessaires au fonctionnement
            du service ne doivent pas être collectées.

            ### Sécurité

            L'application doit mettre en œuvre des mesures adaptées
            afin de protéger les données traitées.

            ### Durée de conservation

            La durée exacte doit être définie selon le fonctionnement
            réel du service et les obligations applicables.

            ### Droits

            Selon le contexte et les données concernées, les personnes
            disposent des droits prévus par la réglementation
            applicable en matière de protection des données.

            Le responsable de traitement doit être identifié dans la
            version finale de cette politique.
            """
        )

    with tabs[2]:

        st.markdown(
            """
            ## Conditions Générales d'Utilisation

            RefundSniper est un outil logiciel destiné à assister
            les professionnels dans l'identification de retards
            potentiels et la préparation de dossiers de réclamation
            auprès des transporteurs.

            ### Utilisation du service

            L'utilisateur est responsable des données qu'il importe
            et doit disposer des droits nécessaires pour les traiter.

            ### Résultats

            Les résultats sont générés à partir des informations
            disponibles et des règles configurées dans l'application.

            Ils doivent être vérifiés par l'utilisateur avant toute
            démarche auprès d'un transporteur.

            ### Remboursements

            L'affichage d'un montant potentiel ne constitue pas une
            promesse de remboursement.

            Les conditions réellement applicables dépendent notamment
            du contrat de transport, du service utilisé, des conditions
            commerciales et des justificatifs disponibles.

            ### Disponibilité

            La disponibilité et les performances du service peuvent
            dépendre de l'environnement technique utilisé.

            ### Responsabilité de l'utilisateur

            L'utilisateur doit vérifier les informations avant
            transmission à un tiers.
            """
        )


# ============================================================
# SIDEBAR
# ============================================================

def sidebar() -> str:

    with st.sidebar:

        st.markdown(
            f"""
            <div class="brand">
                Refund<span>Sniper</span>
            </div>

            <div class="version">
                V{APP_VERSION} · Carrier Recovery OS
            </div>
            """,
            unsafe_allow_html=True,
        )

        page = st.radio(
            "Navigation",
            [
                "🖥️ Dashboard",
                "🎯 Scanner de Colis",
                "📑 Export & Réclamations",
                "⚖️ Mentions Légales & CGU",
            ],
            index=0,
        )

        st.markdown("---")

        st.markdown(
            "### ⚙️ Règles de détection"
        )

        st.caption(
            "Valeurs de démonstration configurables."
        )

        rules = st.session_state.carrier_rules

        for carrier in [
            "Colissimo",
            "DHL",
            "UPS",
            "Autre",
        ]:

            st.markdown(
                f"**{carrier}**"
            )

            late = st.number_input(
                f"Seuil retard — {carrier}",
                min_value=0,
                max_value=90,
                value=int(
                    rules[carrier]["late_days"]
                ),
                step=1,
                key=f"late_{carrier}",
            )

            lost = st.number_input(
                f"Seuil perte — {carrier}",
                min_value=1,
                max_value=180,
                value=int(
                    rules[carrier]["lost_days"]
                ),
                step=1,
                key=f"lost_{carrier}",
            )

            cap = st.number_input(
                f"Plafond indicatif (€) — {carrier}",
                min_value=0.0,
                max_value=10000.0,
                value=float(
                    rules[carrier]["cap"]
                ),
                step=5.0,
                key=f"cap_{carrier}",
            )

            rules[carrier]["late_days"] = late
            rules[carrier]["lost_days"] = lost
            rules[carrier]["cap"] = cap

        st.markdown("---")

        if st.button(
            "♻️ Réinitialiser la démo",
            use_container_width=True,
        ):

            reset_demo_data()
            st.rerun()

        st.caption(
            "Les règles affichées ici sont des paramètres "
            "de démonstration et doivent être alignées avec "
            "les contrats réellement applicables."
        )

    return page


# ============================================================
# APPLICATION PRINCIPALE
# ============================================================

def main() -> None:

    st.set_page_config(
        page_title="RefundSniper",
        page_icon="🎯",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    initialize_session()

    inject_css()

    selected_page = sidebar()

    try:

        if selected_page == "🖥️ Dashboard":
            dashboard_page()

        elif selected_page == "🎯 Scanner de Colis":
            scanner_page()

        elif selected_page == "📑 Export & Réclamations":
            claims_page()

        elif selected_page == "⚖️ Mentions Légales & CGU":
            legal_page()

        else:
            dashboard_page()

    except Exception as exc:

        st.error(
            "Une erreur inattendue est survenue dans l'application."
        )

        with st.expander(
            "Détails techniques"
        ):
            st.exception(exc)

    footer()


# ============================================================
# ENTRYPOINT
# ============================================================

if __name__ == "__main__":
    main()

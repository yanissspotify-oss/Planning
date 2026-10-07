import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta
from html import escape
import random
import uuid
import re
import hmac

# =============================================================================
# APPLICATION : PLANNING DES RENDEZ-VOUS
# =============================================================================
# Cette application Streamlit permet :
#   1. d'ajouter des rendez-vous (écrits dans un Google Sheet) ;
#   2. d'afficher le planning (vue Mois façon agenda Excel, ou Semaine / Jour) ;
#   3. de repérer automatiquement les rendez-vous qui se chevauchent ;
#   4. de consulter et supprimer les rendez-vous.
#
# Les zones faciles à modifier sont marquées :   >>> PERSONNALISATION
# =============================================================================


# =============================================================================
# 1. CATÉGORIES DE RENDEZ-VOUS
# =============================================================================
# >>> PERSONNALISATION — AJOUTER / RENOMMER / RECOLORER UNE CATÉGORIE
#
#   "emoji"   : petit symbole affiché devant le rdv
#   "couleur" : couleur du bloc dans le calendrier (format "#RRGGBB")
#
# Pour les rdv « Client », la couleur vient plutôt du client (voir plus bas).
# =============================================================================

CATEGORIES = {
    "Client":   {"emoji": "🏢", "couleur": "#8697C4"},
    "Médical":  {"emoji": "🩺", "couleur": "#C9B6D9"},
    "Perso":    {"emoji": "👤", "couleur": "#F2C6A0"},
    "Libre":    {"emoji": "🌿", "couleur": "#E4E1DA"},
    "Vacances": {"emoji": "🏖️", "couleur": "#B5473F"},
    "Autre":    {"emoji": "📌", "couleur": "#F5EB3B"},
}

# Catégories qui peuvent durer plusieurs jours (date de début → date de fin)
CATEGORIES_PLUSIEURS_JOURS = ["Vacances"]


# =============================================================================
# 1b. COULEURS DES CLIENTS (reprises de l'agenda Excel)
# =============================================================================
# >>> PERSONNALISATION — COULEUR D'UN CLIENT
#
# Clé = nom du client en minuscules (ou un morceau du nom).
# Un client absent de cette liste reçoit automatiquement une couleur
# de la palette PALETTE_AUTO, toujours la même pour un même nom.
# =============================================================================

COULEURS_CLIENTS = {
    "holcim":      "#F5B800",   # jaune
    "morges":      "#E0701E",   # orange
    "santy":       "#5B8FD0",   # bleu
    "sebastien":   "#7A9A45",   # vert olive
    "sébastien":   "#7A9A45",
    "takeda":      "#7B4BB0",   # violet
    "sushi":       "#EBB7B7",   # rose
    "neuchâtel":   "#22A84F",   # vert
    "comadur":     "#8E3FB5",   # violet foncé
    "infratunnel": "#8E3FB5",
    "unine":       "#F2C14E",   # jaune doux
}

PALETTE_AUTO = [
    "#4DB6AC", "#F06292", "#9575CD", "#4FC3F7", "#AED581",
    "#FFB74D", "#A1887F", "#90A4AE", "#BA68C8", "#81C784",
]


# =============================================================================
# 1c. MOMENTS DE LA JOURNÉE
# =============================================================================
# >>> PERSONNALISATION — HORAIRES DU MATIN / DE L'APRÈS-MIDI
#
# Ces horaires servent à placer le rdv dans l'agenda et à repérer
# les chevauchements (deux rdv le même matin = alerte).
# =============================================================================

MOMENTS = {
    "Matin":         {"emoji": "☀️", "debut": "08:00", "fin": "12:00"},
    "Après-midi":    {"emoji": "🌇", "debut": "13:00", "fin": "17:00"},
    "Journée":       {"emoji": "🗓️", "debut": "08:00", "fin": "17:00"},
    "Heure précise": {"emoji": "🕘", "debut": None,    "fin": None},
}

# Un rdv qui commence avant MIDI est dans la demi-case du matin,
# un rdv qui finit après MIDI est dans la demi-case de l'après-midi.
MIDI = 12 * 60 + 30


# =============================================================================
# 2. HORAIRES PROPOSÉS
# =============================================================================
# >>> PERSONNALISATION — PLAGE HORAIRE ET PAS DES MENUS
#
# Les menus proposent les heures de HEURE_MIN à HEURE_MAX,
# par tranches de PAS_MINUTES minutes.
# =============================================================================

HEURE_MIN = 6
HEURE_MAX = 22
PAS_MINUTES = 15

HORAIRES = [
    f"{h:02d}:{m:02d}"
    for h in range(HEURE_MIN, HEURE_MAX + 1)
    for m in range(0, 60, PAS_MINUTES)
]

# Durées proposées : 0h15, 0h30, ... 8h00
DUREES = [f"{h}h{m:02d}" for h in range(0, 9) for m in (0, 15, 30, 45)][1:33]


# =============================================================================
# 3. COLONNES DU GOOGLE SHEET
# =============================================================================
# L'app écrit ces en-têtes toute seule si la 1re ligne du Sheet est vide.
# =============================================================================

COLONNES = [
    "ID",
    "Date",
    "Début",
    "Fin",
    "Titre",
    "Catégorie",
    "Client",            # Clé client (même nom que dans la feuille "client")
    "Type prestation",   # Visite / OPROMA / Prestation à l'heure / Prestation journalière
    "Lieu",
    "Notes",
    "Créé le",
    "Moment",            # Matin / Après-midi / Journée / Heure précise
    "Date fin",          # pour les vacances sur plusieurs jours (sinon vide)
]
COL_ID = COLONNES.index("ID") + 1

# >>> PERSONNALISATION — NOM DES ONGLETS DANS LE GOOGLE SHEET DES PRESTATIONS
ONGLET_PLANNING = "Planning"   # l'app le crée toute seule s'il n'existe pas
ONGLET_CLIENTS = "client"      # la feuille des fiches clients (déjà existante)

TYPES_PRESTATION = ["Visite", "OPROMA", "Prestation à l'heure", "Prestation journalière"]
EMOJIS_PRESTATION = {
    "Visite": "🏥",
    "OPROMA": "🤰",
    "Prestation à l'heure": "⌚",
    "Prestation journalière": "🗓️",
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
MOIS = ["janvier", "février", "mars", "avril", "mai", "juin",
        "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


# =============================================================================
# 4. CONNEXION AU GOOGLE SHEET
# =============================================================================

@st.cache_resource
def _client():
    """Connexion à Google (gardée en mémoire par Streamlit)."""
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES,
    )
    return gspread.authorize(creds)


def _ws():
    """
    Ouvre l'onglet "Planning" du Google Sheet des prestations.
    S'il n'existe pas encore, il est créé. Les en-têtes sont écrits si besoin.

    IMPORTANT : on n'utilise surtout pas .sheet1, qui est la feuille
    des prestations ! On vise l'onglet par son nom.
    """
    classeur = _client().open_by_url(st.secrets["sheet_url"])
    try:
        ws = classeur.worksheet(ONGLET_PLANNING)
    except gspread.WorksheetNotFound:
        ws = classeur.add_worksheet(title=ONGLET_PLANNING, rows=1000, cols=len(COLONNES))

    # Assez de colonnes dans la feuille pour les nouvelles colonnes
    if ws.col_count < len(COLONNES):
        ws.add_cols(len(COLONNES) - ws.col_count)

    entetes = ws.row_values(1)
    if not entetes:
        ws.append_row(COLONNES)
    elif entetes != COLONNES:
        # Colonnes ajoutées dans une nouvelle version : on réécrit l'en-tête
        ws.update("A1", [COLONNES])
    return ws


@st.cache_data(ttl=60)
def charger_clients():
    """
    Lit la feuille "client" (fiches clients de l'app des prestations).
    Renvoie un dictionnaire { "Holcim": {"Type prestation": "...", "Adresse": ..., ...}, ... }
    Lecture seule : l'app planning ne modifie jamais les fiches.
    """
    try:
        valeurs = _client().open_by_url(st.secrets["sheet_url"]).worksheet(ONGLET_CLIENTS).get_all_values()
    except Exception:
        return {}
    if len(valeurs) < 2:
        return {}
    entetes = valeurs[0]
    fiches = {}
    for ligne in valeurs[1:]:
        fiche = dict(zip(entetes, ligne))
        cle = fiche.get("Clé client", "").strip()
        if cle:
            fiches[cle] = fiche
    return fiches


def types_du_client(fiche):
    """Types de prestation d'une fiche client (stockés séparés par ';')."""
    return [
        t.strip() for t in str(fiche.get("Type prestation", "")).split(";")
        if t.strip() in TYPES_PRESTATION
    ]


def creer_fiche_client(fiche):
    """
    Ajoute une NOUVELLE fiche dans la feuille "client" (jamais de modification
    d'une fiche existante). Les colonnes sont remplies d'après l'en-tête
    réel de la feuille, donc l'ordre des colonnes n'a pas d'importance.
    Renvoie False si la clé client existe déjà.
    """
    ws = _client().open_by_url(st.secrets["sheet_url"]).worksheet(ONGLET_CLIENTS)
    entetes = ws.row_values(1)
    cles = [c.strip().lower() for c in ws.col_values(entetes.index("Clé client") + 1)[1:]]
    if fiche["Clé client"].strip().lower() in cles:
        return False
    ws.append_row([str(fiche.get(col, "")) for col in entetes], value_input_option="RAW")
    charger_clients.clear()
    return True


def lieu_du_client(fiche):
    """Adresse lisible d'une fiche client, pour pré-remplir le lieu."""
    morceaux = [fiche.get("Adresse", "").strip(),
                f"{fiche.get('NPA', '').strip()} {fiche.get('Ville', '').strip()}".strip()]
    return ", ".join(m for m in morceaux if m)


def en_minutes(texte):
    """
    Convertit une heure texte en minutes depuis minuit.
    "09:30" -> 570   |   "9h30" -> 570   |   texte illisible -> None
    """
    m = re.search(r"(\d{1,2})\s*[:hH]\s*(\d{2})", str(texte))
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def en_heure(minutes):
    """570 -> "09:30" """
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


@st.cache_data(ttl=60)
def charger():
    """
    Lit tous les rdv du Sheet et renvoie un DataFrame,
    avec des colonnes techniques (_date, _debut, _fin) pour les calculs.
    """
    try:
        valeurs = _ws().get_all_values()
    except Exception:
        st.cache_resource.clear()
        valeurs = _ws().get_all_values()

    if not valeurs or len(valeurs) < 2:
        df = pd.DataFrame(columns=COLONNES)
    else:
        df = pd.DataFrame(valeurs[1:], columns=valeurs[0])

    # On s'assure que toutes les colonnes existent (même vides)
    for col in COLONNES:
        if col not in df.columns:
            df[col] = ""

    df["_date"] = pd.to_datetime(df["Date"], format="%d/%m/%Y", errors="coerce").dt.date
    df["_debut"] = df["Début"].map(en_minutes)
    df["_fin"] = df["Fin"].map(en_minutes)

    # On garde uniquement les lignes exploitables
    df = df[df["_date"].notna() & df["_debut"].notna() & df["_fin"].notna()].copy()

    # Date de fin (vacances) : vide → même jour
    fin = pd.to_datetime(df["Date fin"], format="%d/%m/%Y", errors="coerce").dt.date
    df["_date_fin"] = [f if pd.notna(f) and f >= d else d for f, d in zip(fin, df["_date"])]

    # Anciennes lignes sans "Moment" : on considère une heure précise
    df["Moment"] = df["Moment"].replace("", "Heure précise")
    return df


def etaler(df):
    """
    Un rdv sur plusieurs jours (vacances du 1er au 10) devient une ligne
    par jour, pour l'affichage et la détection des chevauchements.
    Ajoute "_premier" : True le premier jour (on y écrit le titre).
    """
    if df.empty:
        out = df.copy()
        out["_premier"] = pd.Series(dtype=bool)
        return out
    lignes = []
    for r in df.to_dict("records"):
        jour, dernier = r["_date"], r["_date_fin"]
        while jour <= dernier:
            copie = dict(r)
            copie["_date"] = jour
            copie["_premier"] = (jour == r["_date"])
            lignes.append(copie)
            jour += timedelta(days=1)
    return pd.DataFrame(lignes)


def enregistrer(ligne):
    """Ajoute un rdv à la fin du Sheet."""
    ligne["ID"] = uuid.uuid4().hex[:6]
    ligne["Créé le"] = datetime.now().strftime("%d/%m/%Y %H:%M")
    valeurs = [str(ligne.get(c, "")) for c in COLONNES]

    # RAW = le texte est écrit tel quel ("09:30" reste "09:30")
    try:
        _ws().append_row(valeurs, value_input_option="RAW")
    except Exception:
        st.cache_resource.clear()
        _ws().append_row(valeurs, value_input_option="RAW")
    charger.clear()


def supprimer(id_rdv):
    """Supprime le rdv dont l'ID est donné."""
    ws = _ws()
    ids = ws.col_values(COL_ID)
    for i, val in enumerate(ids):
        if i > 0 and val.strip() == str(id_rdv).strip():
            ws.delete_rows(i + 1)
            charger.clear()
            return True
    return False


# =============================================================================
# 5. DÉTECTION DES CHEVAUCHEMENTS
# =============================================================================
# Deux rdv se chevauchent s'ils sont le même jour ET que
#   le début de l'un est avant la fin de l'autre, et inversement.
# Exemple : 9h-10h et 9h30-11h se chevauchent ; 9h-10h et 10h-11h non.
# =============================================================================

def rdv_en_conflit(df, jour, debut, fin):
    """Renvoie les rdv existants qui chevauchent le créneau proposé."""
    meme_jour = df[df["_date"] == jour]
    return meme_jour[(meme_jour["_debut"] < fin) & (meme_jour["_fin"] > debut)]


def ids_en_conflit(df):
    """Renvoie l'ensemble des ID de rdv qui chevauchent au moins un autre rdv."""
    conflits = set()
    for _, groupe in df.groupby("_date"):
        lignes = groupe.sort_values("_debut").to_dict("records")
        for i, a in enumerate(lignes):
            for b in lignes[i + 1:]:
                if b["_debut"] >= a["_fin"]:
                    break  # triés par début : plus aucun chevauchement possible
                conflits.add(a["ID"])
                conflits.add(b["ID"])
    return conflits


# =============================================================================
# 6. PETITES FONCTIONS UTILITAIRES
# =============================================================================

def date_en_lettres(d):
    jour = "1er" if d.day == 1 else d.day
    return f"{JOURS[d.weekday()]} {jour} {MOIS[d.month - 1]} {d.year}"


def duree_en_minutes(txt):
    h, m = txt.split("h")
    return int(h) * 60 + int(m)


def couleur_rdv(r):
    """
    Couleur de fond d'un rdv : celle du client pour un rdv Client
    (comme dans l'agenda Excel), sinon celle de la catégorie.
    """
    if r.get("Catégorie") == "Client":
        nom = str(r.get("Client", "") or r.get("Titre", "")).strip().lower()
        if nom:
            if nom in COULEURS_CLIENTS:
                return COULEURS_CLIENTS[nom]
            for cle, couleur in COULEURS_CLIENTS.items():
                if cle in nom:                       # "Holcim Eclépens" → holcim
                    return couleur
            # Client inconnu : couleur fixe tirée de la palette d'après son nom
            return PALETTE_AUTO[sum(map(ord, nom)) % len(PALETTE_AUTO)]
    return CATEGORIES.get(r.get("Catégorie"), CATEGORIES["Autre"])["couleur"]


def texte_lisible(fond):
    """Texte blanc sur fond foncé, texte foncé sur fond clair."""
    fond = fond.lstrip("#")
    r, g, b = (int(fond[i:i + 2], 16) for i in (0, 2, 4))
    return "#FFFFFF" if (0.299 * r + 0.587 * g + 0.114 * b) < 140 else "#1B1F3B"


def assombrir(couleur, facteur=0.75):
    """Version plus foncée d'une couleur (pour les bordures)."""
    c = couleur.lstrip("#")
    r, g, b = (int(int(c[i:i + 2], 16) * facteur) for i in (0, 2, 4))
    return f"#{r:02X}{g:02X}{b:02X}"


def infobulle(e):
    """Texte affiché au survol d'un rdv (sur ordinateur)."""
    if e.get("Moment") in ("Matin", "Après-midi", "Journée"):
        quand = e["Moment"]
    else:
        quand = f"{e['Début']}–{e['Fin']}"
    date_fin = str(e.get("Date fin", "")).strip()
    if date_fin and date_fin != e["Date"]:
        quand = f"du {e['Date']} au {date_fin}"
    infos = f"{e['Titre']} · {quand}"
    if str(e.get("Type prestation", "")).strip():
        infos += f" · {e['Type prestation']}"
    if str(e.get("Lieu", "")).strip():
        infos += f" · 📍 {e['Lieu']}"
    if str(e.get("Notes", "")).strip():
        infos += f" · {e['Notes']}"
    return escape(infos.replace("\n", " "), quote=True)


# =============================================================================
# 6b. CALENDRIER DE LA SEMAINE (grille jours × heures)
# =============================================================================
# >>> PERSONNALISATION — TAILLE ET PLAGE DU CALENDRIER
#
#   PX_PAR_HEURE         : hauteur d'une heure en pixels sur ordinateur
#   PX_PAR_HEURE_MOBILE  : hauteur d'une heure sur téléphone
#   CAL_HEURE_MIN  : première heure affichée (si aucun rdv plus tôt)
#   CAL_HEURE_MAX  : dernière heure affichée (si aucun rdv plus tard)
# =============================================================================

PX_PAR_HEURE = 60          # sur ordinateur
PX_PAR_HEURE_MOBILE = 40   # sur téléphone (plus compact)
CAL_HEURE_MIN = 8
CAL_HEURE_MAX = 18


def _placer_cote_a_cote(evenements):
    """
    Quand des rdv se chevauchent, on les met côte à côte dans la colonne du jour.
    Ajoute à chaque rdv : "piste" (sa position) et "nb_pistes" (combien côte à côte).
    """
    evenements = sorted(evenements, key=lambda e: (e["_debut"], e["_fin"]))
    resultat, groupe, fin_groupe = [], [], -1

    def fermer(groupe):
        fins_pistes = []  # heure de fin du dernier rdv de chaque piste
        for e in groupe:
            for i, f in enumerate(fins_pistes):
                if f <= e["_debut"]:
                    e["piste"] = i
                    fins_pistes[i] = e["_fin"]
                    break
            else:
                e["piste"] = len(fins_pistes)
                fins_pistes.append(e["_fin"])
        for e in groupe:
            e["nb_pistes"] = len(fins_pistes)
        resultat.extend(groupe)

    for e in evenements:
        if groupe and e["_debut"] >= fin_groupe:
            fermer(groupe)
            groupe, fin_groupe = [], -1
        groupe.append(e)
        fin_groupe = max(fin_groupe, e["_fin"])
    if groupe:
        fermer(groupe)
    return resultat


def _h(heures):
    """Position verticale : un nombre d'heures × la hauteur d'une heure (variable CSS --h)."""
    return f"calc(var(--h) * {heures:.4f})"


def calendrier(rdv, premier_jour, nb_jours, conflits):
    """
    Construit le HTML du calendrier : nb_jours colonnes à partir de premier_jour.
    La hauteur d'une heure vient de la variable CSS --h (plus petite sur téléphone).
    """
    # Plage horaire : 8h-18h, élargie si un rdv dépasse
    h0, h1 = CAL_HEURE_MIN, CAL_HEURE_MAX
    if not rdv.empty:
        h0 = min(h0, int(rdv["_debut"].min()) // 60)
        h1 = max(h1, -(-int(rdv["_fin"].max()) // 60))  # arrondi à l'heure du dessus
    hauteur = _h(h1 - h0)
    aujourdhui = date.today()
    jours = [premier_jour + timedelta(days=i) for i in range(nb_jours)]

    # Colonne des heures
    heures = "".join(
        f"<div class='cal-heure' style='top:{_h(h - h0)}'>{h:02d}:00</div>"
        for h in range(h0, h1)
    )
    classe_cal = "cal cal-7" if nb_jours == 7 else "cal"
    html = [
        f"<div class='cal-scroll'><div class='{classe_cal}'>",
        "<div class='cal-ligne-entete'><div class='cal-coin'></div>",
    ]

    # En-têtes des jours
    for jour in jours:
        classe = "cal-entete aujourdhui" if jour == aujourdhui else "cal-entete"
        nom = JOURS[jour.weekday()] if nb_jours == 1 else JOURS[jour.weekday()][:3]
        html.append(
            f"<div class='{classe}'><div class='cal-entete-jour'>{nom.capitalize()}</div>"
            f"<div class='cal-entete-num'>{jour.day}</div></div>"
        )
    html.append("</div>")

    # Corps : colonne des heures + une colonne par jour
    html.append(f"<div class='cal-corps'><div class='cal-heures' style='height:{hauteur}'>{heures}</div>")

    for jour in jours:
        classes = "cal-jour"
        if jour == aujourdhui:
            classes += " aujourdhui"
        elif jour.weekday() >= 5:
            classes += " weekend"
        html.append(f"<div class='{classes}' style='height:{hauteur}'>")

        du_jour = rdv[rdv["_date"] == jour].to_dict("records") if not rdv.empty else []
        for e in _placer_cote_a_cote(du_jour):
            cat = CATEGORIES.get(e["Catégorie"], CATEGORIES["Autre"])
            en_conflit = e["ID"] in conflits
            fond = couleur_rdv(e)
            texte = texte_lisible(fond)
            bordure = "#C0392B" if en_conflit else assombrir(fond)

            debut_h = (e["_debut"] - h0 * 60) / 60
            duree_h = (e["_fin"] - e["_debut"]) / 60
            largeur = 100 / e["nb_pistes"]
            gauche = e["piste"] * largeur
            infos = infobulle(e)

            # Peu de place (rdv côte à côte, ou 7 jours sur un petit écran) :
            # la classe "etroit" réduit le texte
            etroit = e["nb_pistes"] > 1
            if etroit:
                horaire = e["Début"]
                titre_html = escape(e["Titre"])
                lieu = ""
            else:
                alerte = "⚠️ " if en_conflit else ""
                horaire = f"{alerte}{e['Début']}–{e['Fin']}"
                titre_html = f"{cat['emoji']} {escape(e['Titre'])}"
                lieu = (
                    f"<div class='cal-evt-l'>📍 {escape(e['Lieu'])}</div>"
                    if str(e["Lieu"]).strip() and duree_h >= 1 else ""
                )

            classe_evt = "cal-evt etroit" if etroit else "cal-evt"
            html.append(
                f"<div class='{classe_evt}' title=\"{infos}\" style='top:{_h(debut_h)};"
                f"height:calc(var(--h) * {duree_h:.4f} - 2px);"
                f"left:calc({gauche:.2f}% + 2px);width:calc({largeur:.2f}% - 4px);"
                f"background:{fond};color:{texte};border-left:4px solid {bordure};"
                f"{'box-shadow:inset 0 0 0 2px #C0392B;' if en_conflit else ''}'>"
                f"<div class='cal-evt-h'>{horaire}</div>"
                f"<div class='cal-evt-t'>{titre_html}</div>{lieu}</div>"
            )

        # Trait rouge "maintenant"
        if jour == aujourdhui:
            minutes = datetime.now().hour * 60 + datetime.now().minute
            if h0 * 60 <= minutes < h1 * 60:
                html.append(
                    f"<div class='cal-maintenant' style='top:{_h((minutes - h0 * 60) / 60)}'></div>"
                )

        html.append("</div>")

    html.append("</div></div></div>")
    return "".join(html)


# =============================================================================
# 6c. VUE MOIS (façon agenda Excel)
# =============================================================================
# Une colonne par mois, une ligne par jour, deux demi-cases :
#   en haut le matin, en bas l'après-midi.
# Dimanches grisés, aujourd'hui en vert, numéro de semaine à côté des lundis.
# =============================================================================

JOURS_COURTS = ["LUN", "MAR", "MER", "JEU", "VEN", "SAM", "D"]


def _blocs_du_jour(evenements, conflits):
    """
    Place les rdv d'un jour dans les deux demi-cases (grille CSS) :
      - un rdv qui couvre matin ET après-midi occupe les deux lignes ;
      - les autres sont rangés côte à côte dans leur demi-case.
    """
    complets, matins, aprems = [], [], []
    for e in sorted(evenements, key=lambda x: (x["_debut"], x["_fin"])):
        au_matin = e["_debut"] < MIDI
        l_aprem = e["_fin"] > MIDI
        if au_matin and l_aprem:
            complets.append(e)
        elif au_matin:
            matins.append(e)
        else:
            aprems.append(e)

    places = []  # (rdv, colonne, lignes CSS)
    col = 1
    for e in complets:
        places.append((e, col, "1 / 3"))
        col += 1
    for k in range(max(len(matins), len(aprems))):
        if k < len(matins):
            places.append((matins[k], col, "1"))
        if k < len(aprems):
            places.append((aprems[k], col, "2"))
        col += 1
    nb_colonnes = col - 1

    html = []
    for e, c, lignes in places:
        fond = couleur_rdv(e)
        texte = texte_lisible(fond)
        en_conflit = e["ID"] in conflits
        classe = "blk conflit" if en_conflit else "blk"
        # Rdv sur plusieurs jours : le titre seulement le 1er jour (bande continue ensuite)
        multi = str(e.get("Date fin", "")).strip() not in ("", e["Date"])
        if multi and not e.get("_premier", True):
            libelle = ""
            classe += " suite"
        else:
            libelle = ("⚠️ " if en_conflit else "") + escape(e["Titre"])
            if e.get("Moment") == "Heure précise":
                libelle = f"<b>{e['Début']}</b>&nbsp;{libelle}"
        html.append(
            f"<div class='{classe}' title=\"{infobulle(e)}\" "
            f"style='grid-column:{c};grid-row:{lignes};background:{fond};color:{texte};'>"
            f"<span>{libelle}</span></div>"
        )
    return nb_colonnes, "".join(html)


def vue_mois(rdv, annee, mois, conflits, classe_sup=""):
    """HTML d'un mois façon agenda Excel."""
    aujourdhui = date.today()
    premier = date(annee, mois, 1)
    nb_jours = (date(annee + (mois == 12), mois % 12 + 1, 1) - premier).days

    # Rdv du mois rangés par jour
    par_jour = {}
    if not rdv.empty:
        dans_mois = rdv[(rdv["_date"] >= premier) & (rdv["_date"] < premier + timedelta(days=nb_jours))]
        for e in dans_mois.to_dict("records"):
            par_jour.setdefault(e["_date"], []).append(e)

    html = [f"<div class='mois {classe_sup}'><div class='mois-titre'>{MOIS[mois - 1]} {annee}</div>"]
    for i in range(nb_jours):
        jour = premier + timedelta(days=i)
        classes = "jr"
        if jour.weekday() == 6:
            classes += " dim"
        if jour == aujourdhui:
            classes += " auj"
        nb_col, blocs = _blocs_du_jour(par_jour.get(jour, []), conflits)
        grille = f"grid-template-columns:repeat({nb_col}, minmax(0, 1fr));" if nb_col else ""
        semaine = jour.isocalendar()[1] if jour.weekday() == 0 else ""
        html.append(
            f"<div class='{classes}'>"
            f"<div class='jr-nom'>{JOURS_COURTS[jour.weekday()]}</div>"
            f"<div class='jr-num'>{jour.day}</div>"
            f"<div class='jr-slots' style='{grille}'>{blocs}</div>"
            f"<div class='jr-sem'>{semaine}</div>"
            f"</div>"
        )
    html.append("</div>")
    return "".join(html)


def vue_plusieurs_mois(rdv, annee, mois, conflits, nb=3):
    """
    Plusieurs mois côte à côte (3 sur ordinateur).
    Sur téléphone, le CSS ne garde que le premier.
    """
    blocs = []
    for k in range(nb):
        m = (mois - 1 + k) % 12 + 1
        a = annee + (mois - 1 + k) // 12
        classe = "" if k == 0 else f"mois-sup mois-{k + 1}"
        blocs.append(vue_mois(rdv, a, m, conflits, classe))
    return f"<div class='mois-grille'>{''.join(blocs)}</div>"


# =============================================================================
# 7. CONFIGURATION DE LA PAGE
# =============================================================================

st.set_page_config(
    page_title="Planning",
    page_icon="📅",
    layout="centered",
)

if "gcp_service_account" not in st.secrets:
    st.error(
        "La clé d'accès n'est pas configurée.\n\n"
        "Vérifie le fichier .streamlit/secrets.toml, puis relance l'application."
    )
    st.stop()


# =============================================================================
# 7b. PROTECTION PAR MOT DE PASSE
# =============================================================================
# L'app est publique (n'importe qui ayant le lien peut l'ouvrir), donc on la
# protège. Deux façons d'entrer :
#   1. taper le mot de passe ;
#   2. ouvrir le lien contenant la clé :  https://.../?cle=LE_MOT_DE_PASSE
#      → pratique pour l'icône sur le téléphone : on ne tape jamais rien.
#
# Le mot de passe se règle dans les Secrets (jamais dans le code) :
#   app_password = "..."
# S'il n'est pas défini (en local par exemple), l'app reste ouverte.
# =============================================================================

def acces_autorise():
    mot_de_passe = str(st.secrets.get("app_password", "")).strip()
    if not mot_de_passe:
        return True  # pas de mot de passe configuré → accès libre
    if st.session_state.get("acces_ok"):
        return True

    # Entrée par le lien (?cle=...)
    if hmac.compare_digest(str(st.query_params.get("cle", "")), mot_de_passe):
        st.session_state["acces_ok"] = True
        return True

    # Sinon : petit écran de connexion
    st.markdown(
        "<div style='text-align:center;margin-top:60px;'>"
        "<div style='font-size:3rem;'>📅🔒</div>"
        "<h2 style='color:#3D52A0;'>Mon planning</h2></div>",
        unsafe_allow_html=True,
    )
    saisie = st.text_input("Mot de passe", type="password", key="mdp_saisie")
    if saisie:
        if hmac.compare_digest(saisie.strip(), mot_de_passe):
            st.session_state["acces_ok"] = True
            st.rerun()
        else:
            st.error("Mot de passe incorrect 🙂")
    return False


if not acces_autorise():
    st.stop()


# =============================================================================
# 8. BARRE LATÉRALE
# =============================================================================

with st.sidebar:
    sombre = st.toggle("🌙 Mode sombre", value=False)
    st.divider()
    if st.button("🔄 Actualiser les données"):
        charger.clear()
        charger_clients.clear()
        st.rerun()


# =============================================================================
# 9. CSS (même style que l'app des prestations)
# =============================================================================

st.markdown(
    """
    <style>
      /* Fond de la page */
      [data-testid="stAppViewContainer"] { background-color: #F6F2FC !important; }

      /* Cartes de rdv */
      .rdv-card {
          border-radius: 12px;
          padding: 10px 14px;
          margin-bottom: 8px;
          color: #1B1F3B;
          animation: glisse .4s ease-out both;
      }
      @keyframes glisse {
          from { transform: translateY(-8px); opacity: 0; }
          to   { transform: translateY(0);    opacity: 1; }
      }

      /* ---------------- CALENDRIER DE LA SEMAINE ---------------- */
      .cal-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; margin-top: 10px; }
      .cal {
          background: #FFFFFF; border-radius: 16px;
          border: 2px solid #D8E6F5; padding: 8px 8px 12px 0;
          box-shadow: 0 4px 14px rgba(61, 82, 160, 0.08);
      }
      .cal-7 { min-width: 640px; }   /* la semaine entière garde une largeur minimale */
      .cal-ligne-entete { display: flex; margin-bottom: 6px; }
      .cal-coin { width: 52px; flex: none; }
      .cal-entete { flex: 1; text-align: center; padding: 6px 0; border-radius: 10px; color: #3D52A0; }
      .cal-entete-jour { font-size: 0.85rem; text-transform: uppercase; opacity: 0.8; }
      .cal-entete-num { font-size: 1.35rem; font-weight: 700; }
      .cal-entete.aujourdhui { background: #CFE5D6; color: #084E00; }
      .cal-corps { display: flex; }
      .cal-heures { width: 52px; flex: none; position: relative; }
      .cal-heure {
          position: absolute; right: 8px; transform: translateY(-50%);
          font-size: 0.75rem; color: #8697C4;
      }
      .cal-heure:first-child { transform: none; }
      .cal-jour {
          flex: 1; position: relative; border-left: 1px solid #E6E0F2;
          background-image: linear-gradient(to bottom, #E6E0F2 1px, transparent 1px);
          background-size: 100% var(--h);
      }
      .cal-jour.weekend { background-color: #FAF8FD; }
      .cal-jour.aujourdhui { background-color: #F2FAF3; }
      .cal-evt {
          position: absolute; box-sizing: border-box; border-radius: 8px; min-height: 20px;
          padding: 3px 6px; overflow: hidden; color: #1B1F3B;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08); cursor: default;
          animation: glisse .4s ease-out both;
      }
      .cal-evt-h { font-size: 0.72rem; font-weight: 700; white-space: nowrap; }
      .cal-evt-t {
          font-size: 0.85rem; font-weight: 700; line-height: 1.15;
          overflow: hidden; text-overflow: ellipsis;
      }
      .cal-evt.etroit { padding: 3px 4px; }
      .cal-evt.etroit .cal-evt-h { white-space: normal; font-size: 0.66rem; line-height: 1.1; }
      .cal-evt.etroit .cal-evt-t { font-size: 0.74rem; overflow-wrap: anywhere; }
      .cal-evt-l { font-size: 0.72rem; opacity: 0.75; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .cal-maintenant {
          position: absolute; left: 0; right: 0; height: 2px; background: #C0392B; z-index: 5;
      }
      .cal-maintenant::before {
          content: ""; position: absolute; left: -5px; top: -4px;
          width: 10px; height: 10px; border-radius: 50%; background: #C0392B;
      }
      .cal-legende { display: flex; flex-wrap: wrap; gap: 14px; margin-top: 12px; font-size: 0.9rem; }
      .cal-legende span { display: inline-flex; align-items: center; gap: 6px; }
      .cal-legende i { width: 14px; height: 14px; border-radius: 4px; display: inline-block; }

      /* ---------------- VUE MOIS (façon agenda Excel) ---------------- */
      .mois-grille {
          display: grid; grid-template-columns: repeat(3, minmax(0, 1fr));
          gap: 12px; margin-top: 10px;
      }
      .mois {
          background: #FFFFFF; border: 2px solid #D8E6F5; border-radius: 14px;
          overflow: hidden; box-shadow: 0 4px 14px rgba(61, 82, 160, 0.08);
      }
      .mois-titre {
          text-align: center; color: #2E8BD6; font-weight: 800; font-size: 1.2rem;
          letter-spacing: .05em; text-transform: uppercase; padding: 8px 0;
          border-bottom: 2px solid #D8E6F5;
      }
      .jr {
          display: grid; grid-template-columns: 34px 28px minmax(0, 1fr) 24px;
          min-height: 38px; border-bottom: 1px solid #EFEFEF;
      }
      .jr-nom, .jr-num, .jr-sem { display: flex; align-items: center; justify-content: center; }
      .jr-nom { font-size: .66rem; font-weight: 600; color: #444; }
      .jr-num { font-size: .9rem; font-weight: 700; color: #1B1F3B; background: #E9EEF5; }
      .jr-sem { font-size: .7rem; color: #555; }
      .jr-slots {
          display: grid; grid-template-rows: 1fr 1fr; column-gap: 2px; row-gap: 1px;
          padding: 1px 3px; min-width: 0;
      }
      .jr.dim { background: #D3D3D3; }
      .jr.dim .jr-nom, .jr.dim .jr-num { background: #8DB8F0; color: #FFFFFF; }
      .jr.auj .jr-num { background: #8EE68E; color: #084E00; }
      .jr.auj { background: #F2FAF3; }
      .blk {
          display: flex; align-items: center; min-width: 0; border-radius: 3px;
          padding: 0 4px; font-size: .7rem; font-weight: 600; cursor: default;
      }
      .blk span { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
      .blk.conflit { box-shadow: inset 0 0 0 2px #C0392B; }
      .blk.suite { border-radius: 0; }

      /* En-tête de chaque jour */
      .jour-titre {
          font-weight: 700; font-size: 1.2rem; color: #3D52A0;
          margin: 22px 0 8px 0; padding-bottom: 4px;
          border-bottom: 2px solid #D8E6F5;
      }
      .jour-titre.aujourdhui { color: #1B7A3D; border-bottom-color: #0F9600; }

      /* Bandeau d'alerte chevauchement */
      .alerte-conflit {
          background: #F8D7DA; border-left: 6px solid #C0392B;
          border-radius: 10px; padding: 12px 16px; color: #842029;
          font-weight: 600; margin-bottom: 12px;
      }
      .ok-creneau {
          background: #E8F5E9; border-left: 6px solid #2E7D32;
          border-radius: 10px; padding: 10px 16px; color: #1B5E20;
      }

      /* Boutons radio (catégories) */
      div[role="radiogroup"] label p {
          font-size: 1.2rem !important; font-weight: 600 !important; color: #A26E5B !important;
      }
      div[role="radiogroup"] label {
          background-color: #F3DDD4 !important; border-radius: 12px !important;
          padding: 10px 20px !important; border: 2px solid #D9C6BE !important;
      }
      div[role="radiogroup"] label:hover { background-color: #EBCFC3 !important; }
      div[role="radiogroup"] label:has(input:checked) {
          background-color: #CFE5D6 !important; border-color: #0F9600 !important;
      }
      div[role="radiogroup"] label:has(input:checked) p {
          color: #084E00 !important; font-weight: 700 !important;
      }

      /* Champs texte, date, menus */
      .stTextInput input, .stDateInput input, .stTextArea textarea {
          min-height: 2.9rem; font-size: 1.05rem;
      }
      [data-testid="stTextInput"] input,
      [data-testid="stTextArea"] textarea { background-color: #FFF6E8 !important; }
      [data-testid="stDateInput"] [role="combobox"],
      [data-testid="stDateInput"] input,
      [role="combobox"] {
          background-color: #FFF6E8 !important;
          border: 2px solid #FFE7C4 !important;
          border-radius: 12px !important;
      }
      div[data-baseweb="select"] > div { min-height: 2.9rem; font-size: 1.05rem; }
      [role="option"]:hover, [role="option"][data-focused="true"] { background-color: #FFE7C4 !important; }
      .react-aria-ComboBox button svg { fill: #FFA51E !important; }
      .react-aria-ComboBox button { background-color: #FFE7C4 !important; }
      .react-aria-ComboBox button:hover { background-color: #FFA51E !important; }
      .react-aria-ComboBox button:hover svg { fill: white !important; }

      /* Titres des champs */
      [data-testid="stWidgetLabel"] p { font-size: 1.5rem !important; font-weight: 600 !important; }

      /* Gros boutons */
      .stButton > button p {
          font-weight: 700; font-size: 1.7rem !important;
          padding: 2rem 6rem; min-height: 70px; width: 100%;
      }
      [data-testid="stBaseButton-primary"] {
          background-color: #A8E6A1 !important; border: 2px solid #2E7D32 !important; color: #1B4332 !important;
      }
      [data-testid="stBaseButton-primary"]:hover {
          background-color: #7DD87D !important; border: 2px solid #1B5E20 !important;
      }

      /* Boutons de la sidebar et de navigation : taille normale */
      [data-testid="stSidebar"] .stButton > button p,
      .st-key-nav_prec .stButton > button p,
      .st-key-nav_auj .stButton > button p,
      .st-key-nav_suiv .stButton > button p,
      .st-key-valider_dialog .stButton > button p,
      .st-key-annuler_dialog .stButton > button p,
      .st-key-ouvrir_suppression .stButton > button p {
          font-size: 1rem !important; padding: 0.4rem 0.8rem !important;
          min-height: auto !important; white-space: nowrap !important;
      }
      .st-key-nav_prec .stButton > button,
      .st-key-nav_auj .stButton > button,
      .st-key-nav_suiv .stButton > button,
      .st-key-annuler_dialog .stButton > button {
          background-color: #EDE8F5 !important; border: 2px solid #8697C4 !important; border-radius: 10px !important;
      }
      .st-key-nav_prec .stButton > button p,
      .st-key-nav_auj .stButton > button p,
      .st-key-nav_suiv .stButton > button p,
      .st-key-annuler_dialog .stButton > button p { color: #3D52A0 !important; }
      .st-key-ouvrir_suppression .stButton > button {
          background-color: #F8D7DA !important; border: 2px solid #C0392B !important; border-radius: 10px !important;
      }
      .st-key-ouvrir_suppression .stButton > button p { color: #842029 !important; }

      /* Onglets */
      .stTabs [data-baseweb="tab-list"] { gap: 15px; }
      [data-testid="stTabs"] { margin-top: 35px !important; }
      [data-testid="stTabs"] [role="tabpanel"] { padding-top: 50px !important; }
      [data-testid="stTab"] {
          background-color: #D8E6F5 !important; border-radius: 13px !important;
          padding: 17px 20px !important; transition: all 0.3s ease !important;
      }
      [data-testid="stTab"]:hover { transform: scale(1.05) !important; }
      [data-testid="stTab"] p { font-size: 1.2rem !important; font-weight: 600 !important; }
      [data-testid="stTab"][aria-selected="true"] {
          background-color: #CFE5D6 !important; box-shadow: 0 4px 12px rgba(15, 255, 0, 0.3) !important;
      }
      [data-testid="stTab"][aria-selected="true"] p { color: #2F2A44 !important; }
      .react-aria-SelectionIndicator { display: none !important; }

      /* Tablette : 2 mois côte à côte */
      @media (max-width: 900px) {
          .mois-grille { grid-template-columns: repeat(2, minmax(0, 1fr)); }
          .mois-3 { display: none; }
      }

      /* Mobile */
      @media (max-width: 640px) {
          .mois-grille { grid-template-columns: 1fr; }
          .mois-sup { display: none; }
          .jr { min-height: 40px; }
          .blk { font-size: .78rem; }
          .cal { padding: 6px 4px 10px 0; border-radius: 12px; }
          .cal-coin, .cal-heures { width: 40px; }
          .cal-heure { font-size: 0.65rem; right: 5px; }
          .cal-entete-jour { font-size: 0.7rem; }
          .cal-entete-num { font-size: 1.05rem; }
          .cal-evt { padding: 2px 4px; border-radius: 6px; }
          .cal-evt-h { font-size: 0.62rem; }
          .cal-evt-t { font-size: 0.75rem; }
          .cal-evt-l { display: none; }
          [data-testid="stAppViewContainer"] h1 { font-size: 26px !important; }
          .stButton > button p { font-size: 1.1rem !important; padding: 0.8rem 1.5rem !important; }
          [data-testid="stTab"] { padding: 10px 12px !important; }
          [data-testid="stTab"] p { font-size: 1rem !important; }
          [data-testid="stWidgetLabel"] p { font-size: 1.1rem !important; }
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# Hauteur d'une heure dans le calendrier (variable CSS --h)
st.markdown(
    f"""<style>
      .cal {{ --h: {PX_PAR_HEURE}px; }}
      @media (max-width: 640px) {{ .cal {{ --h: {PX_PAR_HEURE_MOBILE}px; }} }}
    </style>""",
    unsafe_allow_html=True,
)

if sombre:
    st.markdown(
        """
        <style>
          [data-testid="stAppViewContainer"], [data-testid="stHeader"] { background-color: #12152e; }
          [data-testid="stSidebar"] { background-color: #1a1f42; }
          [data-testid="stTab"] { background-color: #232a52 !important; }
          [data-testid="stTab"] p { color: #EDE8F5 !important; }
          [data-testid="stTab"][aria-selected="true"] { background-color: #3D52A0 !important; }
          [data-testid="stTab"][aria-selected="true"] p { color: white !important; }
          [data-testid="stAppViewContainer"] p, [data-testid="stAppViewContainer"] label,
          [data-testid="stAppViewContainer"] span, [data-testid="stAppViewContainer"] h1,
          [data-testid="stAppViewContainer"] h2, [data-testid="stAppViewContainer"] h3,
          [data-testid="stMarkdownContainer"] { color: #EDE8F5 !important; }
          .rdv-card span, .rdv-card div { color: #1B1F3B !important; }
          input, textarea { background-color: #1c2143 !important; color: #EDE8F5 !important; }
          [data-baseweb="select"] > div { background-color: #1c2143 !important; color: #EDE8F5 !important; }
        </style>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# 10. BANDEAU + SALUTATION
# =============================================================================

st.markdown(
    """
    <div style="background: linear-gradient(135deg, #C6B7E9, #7091E6);
                padding: 22px 24px; border-radius: 16px; margin-bottom: 40px; text-align:center;">
      <h1 style="color:#ffffff !important; margin:0; font-size:40px;">📅 Mon planning</h1>
      <p style="color:#EDE8F5; margin:6px 0 0; font-size:18px;">
          Agenda • Clients • Vacances
      </p>
    </div>
    """,
    unsafe_allow_html=True,
)

maintenant = datetime.now()
if maintenant.hour < 12:
    salutation = "Bonjour Maman ☀️"
elif maintenant.hour < 18:
    salutation = "Bon après-midi Maman 🌤️"
else:
    salutation = "Bonsoir Maman 🌙"

petits_mots = [
    "une journée bien rangée, c'est déjà la moitié du travail ✨",
    "pense à faire une pause ☕",
    "tu gères, comme toujours ✨",
    "ton rythme est le bon rythme 🤎",
    "une respiration et on continue 🌿",
    "tu es la meilleure 💫",
]
petit_mot = random.Random(maintenant.strftime("%Y-%m-%d")).choice(petits_mots)

st.markdown(
    f"""
    <style>
    @keyframes scintille {{
      0%, 100% {{ opacity: .3; transform: scale(.85) rotate(-8deg); }}
      50%      {{ opacity: 1;  transform: scale(1.15) rotate(8deg); }}
    }}
    .etincelle {{ display:inline-block; animation: scintille 7s ease-in-out infinite; }}
    .etincelle.droite {{ animation-delay: 1.3s; }}
    </style>
    <div style="text-align:center; font-size:18px; color:#7091E6; margin-bottom:20px;">
        <span class="etincelle">✨</span> {salutation} • {petit_mot} <span class="etincelle droite">✨</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# =============================================================================
# 11. CHARGEMENT DES DONNÉES
# =============================================================================

try:
    df = charger()
except Exception as e:
    st.error("Impossible de lire le Google Sheet (partage en Éditeur ?).")
    st.caption(f"Détail technique : {e}")
    df = pd.DataFrame(columns=COLONNES + ["_date", "_debut", "_fin", "_date_fin"])

# Une ligne par jour (les vacances sur 10 jours deviennent 10 lignes)
df_jours = etaler(df)
conflits_globaux = ids_en_conflit(df_jours) if not df_jours.empty else set()

tab_planning, tab_ajout, tab_liste = st.tabs([
    "📅 Calendrier",
    "➕ Nouveau rdv",
    "🔎 Tous les rdv",
])


# =============================================================================
# 12. ONGLET : CALENDRIER
# =============================================================================

with tab_planning:

    # ── Choix de l'affichage ───────────────────────────────────────────
    # Le lien peut imposer l'affichage par défaut : ?vue=semaine ou ?vue=jour
    VUES = {"mois": "🗓️ Mois", "semaine": "📅 Semaine", "jour": "📆 Jour"}

    if "vue" not in st.session_state:
        vue_lien = str(st.query_params.get("vue", "mois"))
        st.session_state["vue"] = vue_lien if vue_lien in VUES else "mois"

    vue = st.radio(
        "Affichage",
        list(VUES.keys()),
        format_func=lambda v: VUES[v],
        horizontal=True,
        key="vue",
        label_visibility="collapsed",
    )

    # Date de référence (mémorisée entre deux clics)
    if "date_ref" not in st.session_state:
        st.session_state["date_ref"] = date.today()

    def decaler(sens):
        """Avance (+1) ou recule (-1) d'un mois, d'une semaine ou d'un jour."""
        ref = st.session_state["date_ref"]
        if vue == "mois":
            m = ref.month - 1 + sens
            st.session_state["date_ref"] = date(ref.year + m // 12, m % 12 + 1, 1)
        else:
            st.session_state["date_ref"] = ref + timedelta(days=7 * sens if vue == "semaine" else sens)

    c1, c2, c3 = st.columns(3)
    if c1.button("◀ Précédent", key="nav_prec", use_container_width=True):
        decaler(-1)
        st.rerun()
    if c2.button("Aujourd'hui", key="nav_auj", use_container_width=True):
        st.session_state["date_ref"] = date.today()
        st.rerun()
    if c3.button("Suivant ▶", key="nav_suiv", use_container_width=True):
        decaler(+1)
        st.rerun()

    ref = st.session_state["date_ref"]

    # Période affichée
    if vue == "mois":
        premier = date(ref.year, ref.month, 1)
        m = ref.month - 1 + 3   # 3 mois affichés sur ordinateur
        dernier = date(ref.year + m // 12, m % 12 + 1, 1) - timedelta(days=1)
        titre_periode = None   # chaque mois a son propre titre
    elif vue == "semaine":
        premier = ref - timedelta(days=ref.weekday())   # lundi
        dernier = premier + timedelta(days=6)
        titre_periode = (f"Semaine du {premier.day} {MOIS[premier.month - 1]} "
                         f"au {dernier.day} {MOIS[dernier.month - 1]} {dernier.year}")
    else:
        premier = dernier = ref
        titre_periode = date_en_lettres(ref).capitalize()

    if titre_periode:
        st.markdown(f"### {titre_periode}")

    periode = (df_jours[(df_jours["_date"] >= premier) & (df_jours["_date"] <= dernier)]
               if not df_jours.empty else df_jours)
    nb_conflits = len(set(periode["ID"]) & conflits_globaux) if not periode.empty else 0

    if nb_conflits:
        st.markdown(
            f"<div class='alerte-conflit'>⚠️ Attention : {nb_conflits} rdv se chevauchent "
            f"sur cette période (encadrés en rouge).</div>",
            unsafe_allow_html=True,
        )

    # ── Le calendrier ──────────────────────────────────────────────────
    if vue == "mois":
        st.markdown(vue_plusieurs_mois(periode, ref.year, ref.month, conflits_globaux),
                    unsafe_allow_html=True)
    else:
        st.markdown(calendrier(periode, premier, 7 if vue == "semaine" else 1, conflits_globaux),
                    unsafe_allow_html=True)

    # ── Légende : les clients visibles + les catégories ───────────────
    vus = {}
    if not periode.empty:
        for e in periode.drop_duplicates("ID").to_dict("records"):
            if e["Catégorie"] == "Client":
                nom = str(e.get("Client") or e["Titre"]).strip()
                vus.setdefault(nom, couleur_rdv(e))
    legende = "".join(
        f"<span><i style='background:{coul}'></i>{escape(nom)}</span>" for nom, coul in sorted(vus.items())
    )
    legende += "".join(
        f"<span><i style='background:{c['couleur']}'></i>{c['emoji']} {nom}</span>"
        for nom, c in CATEGORIES.items() if nom != "Client"
    )
    legende += "<span><i style='background:#FFFFFF;box-shadow:inset 0 0 0 2px #C0392B'></i>⚠️ Chevauchement</span>"
    st.markdown(f"<div class='cal-legende'>{legende}</div>", unsafe_allow_html=True)

    if periode.empty:
        st.caption("Rien de prévu sur cette période 🌿")
    if vue == "semaine":
        st.caption("📱 Sur téléphone, la vue « Mois » ou « Jour » est plus confortable.")


# =============================================================================
# 13. ONGLET : NOUVEAU RDV
# =============================================================================

with tab_ajout:

    categorie = st.radio(
        "🏷️ **Étape 1 — Quel type de rdv ?**",
        list(CATEGORIES.keys()),
        format_func=lambda c: f"{CATEGORIES[c]['emoji']} {c}",
        horizontal=True,
        key="saisie_categorie",
    )

    client = ""
    type_prestation = ""
    lieu_defaut = ""
    fiche_a_creer = None   # rempli si on crée une nouvelle fiche client

    st.divider()
    if categorie == "Client":
        # ── Rdv client : on s'appuie sur les fiches clients des prestations
        fiches = charger_clients()
        choix_client = st.selectbox(
            "🏢 **Étape 2 — Pour quel client ?**",
            sorted(fiches.keys()) + ["➕ Autre (pas encore de fiche)"],
            key="saisie_client",
            help="La liste vient des fiches clients de l'app des prestations.",
        )
        if choix_client == "➕ Autre (pas encore de fiche)":
            client = st.text_input(
                "Nom du client *",
                placeholder="Ex : Holcim",
                key="saisie_client_autre",
                help="C'est le nom court qui apparaîtra dans les menus (la « clé client »).",
            ).strip()
            fiche = {}

            # Le client existe peut-être déjà (autre majuscule, espace…)
            doublon = next((c for c in fiches if c.lower() == client.lower()), None)
            if client and doublon:
                st.info(f"💡 « {doublon} » a déjà une fiche : choisis-le directement dans la liste au-dessus.")

            elif client:
                creer_fiche = st.toggle(
                    "🗂️ Créer aussi sa fiche client",
                    value=True,
                    key="saisie_creer_fiche",
                    help="Le client apparaîtra ensuite dans les menus des deux apps.",
                )
                if creer_fiche:
                    with st.container(border=True):
                        st.markdown(f"**🗂️ Nouvelle fiche : {escape(client)}**")
                        nom_officiel = st.text_input(
                            "Nom officiel", placeholder="Ex : Holcim (Suisse) SA", key="saisie_fiche_nom")
                        adresse = st.text_input("Adresse", key="saisie_fiche_adresse")
                        c_npa, c_ville = st.columns([1, 2])
                        npa = c_npa.text_input("NPA", key="saisie_fiche_npa")
                        ville = c_ville.text_input("Ville", key="saisie_fiche_ville")
                        types_fiche = st.multiselect(
                            "🩺 Type(s) de prestation",
                            TYPES_PRESTATION,
                            key="saisie_fiche_types",
                        )
                        st.caption(
                            "L'IBAN, l'e-mail et le délai de paiement se complètent plus tard "
                            "dans l'app des prestations (⚙️ Gérer les fiches clients)."
                        )
                    fiche = {
                        "Clé client": client,
                        "Nom officiel": nom_officiel.strip() or client,
                        "Adresse": adresse.strip(),
                        "NPA": npa.strip(),
                        "Ville": ville.strip(),
                        "Pays": "Suisse",
                        "Délai de paiement": "30",
                        "Référence client": "".join(client.split())[:3].upper() + "-001",
                        "Type prestation": " ; ".join(types_fiche),
                    }
                    fiche_a_creer = fiche
        else:
            client = choix_client
            fiche = fiches.get(client, {})

        # Types proposés : ceux de la fiche du client, sinon tous
        types_proposes = types_du_client(fiche) or TYPES_PRESTATION
        type_prestation = st.selectbox(
            "🩺 Quel type de prestation ?",
            types_proposes,
            format_func=lambda t: f"{EMOJIS_PRESTATION.get(t, '')} {t}",
            key=f"saisie_type_{client}",
        )

        precision = st.text_input(
            "📝 Précision (facultatif)",
            placeholder="Ex : visite de suivi, réunion RH…",
            key="saisie_precision",
        )
        titre = f"{client} — {precision.strip()}" if precision.strip() else client
        lieu_defaut = lieu_du_client(fiche)

    elif categorie in ("Libre", "Vacances"):
        # Titre facultatif : "Libre" / "Vacances" par défaut
        precision = st.text_input(
            "📝 **Étape 2 — Une précision ?** (facultatif)",
            placeholder="Ex : Maroc, Paris, temps pour moi…",
            key="saisie_titre",
        ).strip()
        titre = f"{categorie} — {precision}" if precision else categorie

    else:
        titre = st.text_input(
            "📝 **Étape 2 — C'est quoi ?**",
            placeholder="Ex : Dr Eich, RDZ urologue, permis…",
            key="saisie_titre",
        )

    st.divider()
    plusieurs_jours = categorie in CATEGORIES_PLUSIEURS_JOURS

    if plusieurs_jours:
        # ── Vacances : du … au …
        st.markdown("📅 **Étape 3 — Quelles dates ?**")
        c_du, c_au = st.columns(2)
        jour = c_du.date_input("Du", value=date.today(), format="DD/MM/YYYY", key="saisie_date")
        jour_fin = c_au.date_input("Au", value=jour, min_value=jour, format="DD/MM/YYYY",
                                   key=f"saisie_date_fin_{jour}")
        nb = (jour_fin - jour).days + 1
        st.caption(f"👉 Du **{date_en_lettres(jour)}** au **{date_en_lettres(jour_fin)}** "
                   f"({nb} jour{'s' if nb > 1 else ''})")
        moment = "Journée"
    else:
        jour = st.date_input(
            "📅 **Étape 3 — Quel jour ?**",
            value=date.today(),
            format="DD/MM/YYYY",
            key="saisie_date",
        )
        jour_fin = jour
        st.caption(f"👉 Tu as choisi : **{date_en_lettres(jour)}**")

        st.divider()
        moment = st.radio(
            "🕘 **Étape 4 — À quel moment ?**",
            list(MOMENTS.keys()),
            format_func=lambda m: f"{MOMENTS[m]['emoji']} {m}",
            horizontal=True,
            key="saisie_moment",
        )

    if moment == "Heure précise":
        col_h, col_d = st.columns(2)
        debut_txt = col_h.selectbox(
            "🕘 À quelle heure ?",
            HORAIRES,
            index=HORAIRES.index("09:00"),
            key="saisie_debut",
        )
        duree_txt = col_d.selectbox(
            "⏱️ Combien de temps ?",
            DUREES,
            index=DUREES.index("1h00"),
            key="saisie_duree",
        )
        debut = en_minutes(debut_txt)
        fin = debut + duree_en_minutes(duree_txt)
        fin_txt = en_heure(fin) if fin < 24 * 60 else "23:59"
        st.caption(f"👉 De **{debut_txt}** à **{fin_txt}**")
    else:
        debut_txt, fin_txt = MOMENTS[moment]["debut"], MOMENTS[moment]["fin"]
        debut, fin = en_minutes(debut_txt), en_minutes(fin_txt)

    st.divider()
    lieu = st.text_input(
        "📍 Lieu (facultatif)",
        value=lieu_defaut,
        key=f"saisie_lieu_{client}_{lieu_defaut}",   # se pré-remplit avec l'adresse du client
    )
    notes = st.text_area("🗒️ Notes (facultatif)", key="saisie_notes", height=90)

    # ── Vérification en direct du créneau ──────────────────────────────
    st.divider()
    jours_concernes = [jour + timedelta(days=i) for i in range((jour_fin - jour).days + 1)]
    if df_jours.empty:
        chevauchements = df_jours
    else:
        chevauchements = pd.concat(
            [rdv_en_conflit(df_jours, j, debut, fin) for j in jours_concernes]
        ).drop_duplicates("ID")

    if chevauchements.empty:
        st.markdown("<div class='ok-creneau'>✅ Créneau libre, aucun autre rdv à ce moment-là.</div>",
                    unsafe_allow_html=True)
    else:
        liste = "<br>".join(
            f"• {r['_date'].strftime('%d/%m')} — {escape(r['Titre'])} ({r['Début']} – {r['Fin']})"
            for _, r in chevauchements.sort_values(["_date", "_debut"]).iterrows()
        )
        st.markdown(
            f"<div class='alerte-conflit'>⚠️ Ce créneau chevauche :<br>{liste}</div>",
            unsafe_allow_html=True,
        )

    ligne = {
        "Date": jour.strftime("%d/%m/%Y"),
        "Date fin": jour_fin.strftime("%d/%m/%Y") if jour_fin != jour else "",
        "Moment": moment,
        "Début": debut_txt,
        "Fin": fin_txt,
        "Titre": titre.strip(),
        "Catégorie": categorie,
        "Client": client,
        "Type prestation": type_prestation,
        "Lieu": lieu.strip(),
        "Notes": notes.strip(),
    }

    @st.dialog("📅 Récapitulatif")
    def confirmer(ligne, nb_conflits, fiche_a_creer=None):
        cat = CATEGORIES[ligne["Catégorie"]]
        st.markdown(f"### {cat['emoji']} {ligne['Titre']}")
        d1 = datetime.strptime(ligne["Date"], "%d/%m/%Y").date()
        if ligne["Date fin"]:
            d2 = datetime.strptime(ligne["Date fin"], "%d/%m/%Y").date()
            st.write(f"📅 Du {date_en_lettres(d1)} au {date_en_lettres(d2)}")
        else:
            st.write(f"📅 {date_en_lettres(d1)}")
            if ligne["Moment"] == "Heure précise":
                st.write(f"🕘 {ligne['Début']} – {ligne['Fin']}")
            else:
                st.write(f"{MOMENTS[ligne['Moment']]['emoji']} {ligne['Moment']}")
        if ligne["Type prestation"]:
            st.write(f"{EMOJIS_PRESTATION.get(ligne['Type prestation'], '')} {ligne['Type prestation']}")
        if ligne["Lieu"]:
            st.write(f"📍 {ligne['Lieu']}")
        if ligne["Notes"]:
            st.write(f"🗒️ {ligne['Notes']}")

        if fiche_a_creer:
            st.info(f"🗂️ Une fiche client « {fiche_a_creer['Clé client']} » sera aussi créée.")

        if nb_conflits:
            st.warning(f"⚠️ Ce rdv chevauche {nb_conflits} autre(s) rdv. Tu peux l'enregistrer quand même.")

        st.divider()
        col_ok, col_non = st.columns(2)
        if col_ok.button("✅ Valider", type="primary", key="valider_dialog"):
            with st.spinner("J'enregistre… ⏳"):
                try:
                    message = "Rendez-vous ajouté au planning 🎉"
                    if fiche_a_creer:
                        if creer_fiche_client(fiche_a_creer):
                            message += f" — fiche « {fiche_a_creer['Clé client']} » créée 🗂️"
                        else:
                            message += f" — la fiche « {fiche_a_creer['Clé client']} » existait déjà"
                    enregistrer(ligne)
                    st.session_state["message_ok"] = message
                    for k in list(st.session_state.keys()):
                        if k.startswith("saisie_"):
                            del st.session_state[k]
                except Exception as e:
                    st.session_state["message_erreur"] = str(e)
            st.rerun()
        if col_non.button("✏️ Modifier", key="annuler_dialog"):
            st.rerun()

    if st.button("💾 Enregistrer", type="primary"):
        if not titre.strip():
            st.error("Il manque juste le client 🙂" if categorie == "Client"
                     else "Il manque juste le nom du rendez-vous 🙂")
        else:
            confirmer(ligne, len(chevauchements), fiche_a_creer)

    if "message_ok" in st.session_state:
        st.balloons()
        st.success(st.session_state["message_ok"])
        del st.session_state["message_ok"]

    if "message_erreur" in st.session_state:
        st.error("Impossible d'écrire dans le Google Sheet. Vérifie qu'il est partagé en Éditeur.")
        st.caption(f"Détail technique : {st.session_state['message_erreur']}")
        del st.session_state["message_erreur"]


# =============================================================================
# 14. ONGLET : TOUS LES RDV
# =============================================================================

with tab_liste:

    if df.empty:
        st.info("Aucun rendez-vous enregistré pour le moment.")
    else:
        f1, f2 = st.columns([1, 2])
        periode = f1.selectbox("Afficher", ["À venir", "Passés", "Tous"])
        recherche = f2.text_input("🔎 Rechercher", placeholder="Nom, lieu…")

        vue = df.copy()
        if periode == "À venir":
            vue = vue[vue["_date_fin"] >= date.today()]
        elif periode == "Passés":
            vue = vue[vue["_date_fin"] < date.today()]
        if recherche.strip():
            masque = vue[["Titre", "Client", "Lieu", "Notes"]].apply(
                lambda col: col.str.contains(recherche, case=False, na=False)
            ).any(axis=1)
            vue = vue[masque]

        vue = vue.sort_values(["_date", "_debut"], ascending=(periode != "Passés"))
        vue["⚠️"] = vue["ID"].map(lambda i: "⚠️" if i in conflits_globaux else "")

        st.write(f"**{len(vue)} rendez-vous**")
        st.dataframe(
            vue[["⚠️", "Date", "Date fin", "Moment", "Début", "Fin", "Titre", "Catégorie",
                 "Client", "Type prestation", "Lieu", "Notes"]],
            use_container_width=True,
            hide_index=True,
        )

        # ── Suppression ───────────────────────────────────────────────
        st.divider()
        st.markdown("### 🗑️ Supprimer un rendez-vous")

        options = vue["ID"].tolist()
        if options:
            infos = vue.set_index("ID")
            choix = st.selectbox(
                "Choisir le rdv",
                options,
                format_func=lambda i: f"{infos.loc[i, 'Date']} · {infos.loc[i, 'Début']} — {infos.loc[i, 'Titre']}",
            )

            @st.dialog("🗑️ Supprimer le rendez-vous")
            def confirmer_suppression(id_rdv, libelle):
                st.write(f"**{libelle}**")
                st.warning("Ce rendez-vous sera définitivement supprimé du Sheet.")
                col_ok, col_non = st.columns(2)
                if col_ok.button("🗑️ Confirmer", type="primary", key="valider_dialog"):
                    try:
                        supprimer(id_rdv)
                        st.session_state["msg_suppr"] = "Rendez-vous supprimé ✅"
                    except Exception as e:
                        st.session_state["msg_suppr_err"] = str(e)
                    st.rerun()
                if col_non.button("Annuler", key="annuler_dialog"):
                    st.rerun()

            if st.button("🗑️ Supprimer", key="ouvrir_suppression"):
                confirmer_suppression(
                    choix,
                    f"{infos.loc[choix, 'Date']} · {infos.loc[choix, 'Début']} — {infos.loc[choix, 'Titre']}",
                )

        if "msg_suppr" in st.session_state:
            st.success(st.session_state["msg_suppr"])
            del st.session_state["msg_suppr"]
        if "msg_suppr_err" in st.session_state:
            st.error("Impossible de supprimer.")
            st.caption(f"Détail : {st.session_state['msg_suppr_err']}")
            del st.session_state["msg_suppr_err"]

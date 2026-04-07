"""
📮 LPB — Courrier, Email & Fiche PDP (v3)
Import ZIP Notion → Recherche PDP → Fiche complète → Envoi lettre / email
Application interne La Première Brique pour centraliser les fiches PDP,
préparer les courriers / emails, et préremplir les informations légales LPB.
"""

import streamlit as st
import pandas as pd
import requests
import smtplib
import base64
import json
import re
import io
import zipfile
import tempfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from collections import Counter

# ═══════════════════════════════════════════════
# CONSTANTS
# ═══════════════════════════════════════════════
API_BASE = "https://www.merci-facteur.com/api/1.2/prod/service"

LPB_LEGAL_INFO = {
    "denomination": "LA PREMIERE BRIQUE",
    "siren": "848713442",
    "siret_siege": "84871344200053",
    "tva_intracom": "FR45848713442",
    "eori": "Pas de n° EORI valide",
    "activite_principale": "Conseil pour les affaires et autres conseils de gestion",
    "naf_ape": "70.22Z",
    "activite_principale_naf_2025": "Activités de conseil pour les affaires et autre conseil de gestion (70.20Y)",
    "adresse_postale": "91 COURS CHARLEMAGNE",
    "cp": "69002",
    "ville": "LYON",
    "pays": "France",
    "forme_juridique": "SAS, société par actions simplifiée",
    "effectif": "10 à 19 salariés, en 2023",
    "categorie_entreprise": "Petite ou Moyenne Entreprise (PME), en 2023",
    "date_creation": "25/02/2019",
    "date_inscription_insee": "25/02/2019",
    "date_rne_inpi": "11/01/2022",
    "idcc": "0478",
    "insee_statut": "Inscrite (Insee)",
    "inpi_statut": "Immatriculée au RNE (INPI)",
    "sources": ["INSEE", "VIES", "Douanes", "INPI", "État des inscriptions"],
}

MODES_ENVOI = {
    "Lettre verte (normal)": "normal",
    "Lettre suivie": "suivi",
    "Recommandé avec AR": "lrar",
    "Recommandé AR numérisé": "lrare",
    "Recommandé élec. (email OTP)": "ERE_OTP_MAIL",
    "Recommandé élec. (SMS OTP)": "ERE_OTP_SMS",
}
TYPES_COURRIER = {
    "Lettre": "lettre",
    "Carte postale": "carte_postale",
    "Carte postale enveloppe": "carte_postale_enveloppe",
    "Carte pliée": "carte_pliee",
    "Photo": "photo",
}

# Mapping couleurs/icônes pour niveau CONTENTIEUX
CTX_COLORS = {"Contentieux": "#dc3545", "Précontentieux": "#e67e22", "": "#94a3b8"}
CTX_ICONS = {"Contentieux": "🔴", "Précontentieux": "🟠", "": "⚪"}

# Mapping couleurs pour STADE
STADE_CONF = {
    "Financé":       ("🟢", "#16a34a"),
    "Remboursé":     ("✅", "#6b7280"),
    "Abandonné":     ("⚫", "#374151"),
    "Négociation":   ("💬", "#8b5cf6"),
    "Structuration": ("🟣", "#7c3aed"),
    "Audit":         ("🔵", "#2563eb"),
    "A archiver":    ("📦", "#9ca3af"),
    "Collecte à venir": ("📥", "#0ea5e9"),
    "Clôturé":       ("🔒", "#4b5563"),
}

# Colonnes pour chaque section de la fiche
SEC_IDENTITE = [
    "Nom du porteur", "Société (SIREN)", "🔑 SIREN (clean)", "Téléphone",
    "Email PDP", "Notaire du PDP", "Chargé de projet", "Chargé de suivi",
]
SEC_CTX = [
    "niveau CONTENTIEUX", "📝  Date niveau de CTX", "ℹ️ Motif niveau de CTX",
    "Date début (pré)CTX", "Date fin (pré)CTX",
    "Procédure collective détectée ?", "Type de dernière procédure (LJ/RJ/Sauvegarde)",
    "Dernière date de LJ/RJ/Sauvegarde",
    "🎯 Situation juridique actuelle (BODACC)", "Date du dernier jugement publié",
    "Lien vers le jugement BODACC (PDF officiel)",
    "Hypothèque judiciaire",
]
SEC_RISQUE = [
    "Niveau de risque temporel", "📝 Date remboursement final (estimée)",
    "📝 Potentiel de remboursement (€)",
    "Montant restant total", "Capital restant",
    "🎯 Montant couvert", "🎯 Montant à risque",
    "🎯 Montant capital non couvert", "🎯 Montant intérêts non couverts",
    "🎯 Montant pénalité non couvert", "🎯 Montant bonus non couvert",
    "Taux pénalité (%)",
    "🎯 Segment", "🎯 Tranche",
]
SEC_PROJET = [
    "🔑 ID Back-office", "Nom du projet", "STADE", "BO - Statut",
    "Type d'opération", "Description du projet", "Adresse et lots",
    "Type de financement ",
]
SEC_FINANCIER = [
    "Montant à financer", "BO - Montant cible", "Montant décaissé",
    "Montant reçu, à rembourser", "Taux", "Fees (ht)", "Taux de TVA applicable",
    "CA net AAP net Bonus", "Montant AAP", "Montant du bonus",
    "Type d'intérêts", "Type de décaissement", "Ticket max ",
]
SEC_DATES = [
    "Date début prêt", "Date de collecte", "Date de succès", "Décaissement",
    "Date de réitération", "Date fin échéance min", "Date fin échéance cible",
    "Date fin échéance max", "Date fin échéance prolongation",
    "Durée mini", "Durée cible", "Durée maxi initiale", "Durée prolongée ( en mois)",
    "Dates des points 90.60.30j",
]
SEC_SUIVI = [
    "Contrats", "KYB", "Etat urbanisme", "Garantie(s)",
    "CS de la préqual", "Conditions suspensives après comités",
    "Dernière actualité + Statut", "Date dernière NL",
    "Warning dernière NL", "Warning Point date maximale", "Warning Date cible",
    "Signé chez le notaire?", "Validation Comité?",
]
SEC_HYPOTHEQUE = [
    "Date de fin hypothèque initiale", "Date fin réinscription",
    "Durée avant fin hypothèque (j)", "Nom notaire réinscription hypothèque",
    "Date de demande de réinscription", "Commentaire renouvellement hypothèque",
]


# ═══════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════
st.set_page_config(page_title="LPB — Courrier & Fiche PDP", page_icon="📮",
                   layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
.stApp{font-family:'DM Sans',sans-serif}
.hdr{background:linear-gradient(135deg,#0f172a,#1e293b 50%,#334155);padding:1.6rem 2rem;border-radius:14px;margin-bottom:1.5rem;color:#fff}
.hdr h1{margin:0;font-size:1.7rem;font-weight:700}.hdr p{margin:.3rem 0 0;opacity:.7;font-size:.92rem}
.fc{background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:1rem 1.3rem;margin-bottom:.8rem}
.fc h4{margin:0 0 .6rem;color:#1e293b;font-weight:600;font-size:.95rem}
.fc-alert{border-left:4px solid #dc3545;background:#fef2f2}
.fc-warn{border-left:4px solid #f59e0b;background:#fffbeb}
.badge{display:inline-block;padding:3px 10px;border-radius:16px;font-size:.78rem;font-weight:600;color:#fff;margin:2px 3px 2px 0}
.sc{text-align:center;background:#fff;border:1px solid #e2e8f0;border-radius:10px;padding:.8rem .5rem}
.sc .v{font-size:1.5rem;font-weight:700;color:#1e293b}.sc .l{font-size:.75rem;color:#64748b;margin-top:2px}
.kv{display:flex;padding:3px 0;border-bottom:1px solid #f1f5f9}
.kv .k{flex:0 0 220px;font-weight:600;color:#475569;font-size:.84rem}
.kv .val{flex:1;color:#1e293b;font-size:.84rem;word-break:break-word}
.sb{padding:.7rem 1rem;border-radius:10px;margin:.5rem 0;font-weight:500;font-size:.88rem}
.sb-ok{background:#d1fae5;border-left:4px solid #10b981;color:#065f46}
.sb-err{background:#fee2e2;border-left:4px solid #ef4444;color:#991b1b}
.sb-info{background:#dbeafe;border-left:4px solid #3b82f6;color:#1e40af}
.sb-warn{background:#fef3c7;border-left:4px solid #f59e0b;color:#92400e}
</style>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════
for k, v in {"df": None, "pdp_email": None, "pdp_rows": None,
             "mf_token": None, "mf_token_exp": None, "mf_uid": None,
             "mf_creds": None, "smtp": {}, "history": []}.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ═══════════════════════════════════════════════
# HELPERS — ZIP & CSV LOADING
# ═══════════════════════════════════════════════
def extract_csv_from_zip(uploaded_file) -> pd.DataFrame:
    """
    Gère le format d'export Notion :
    outer.zip → inner.zip (Part-1) → CSV files
    On prend le CSV principal (pas le _all).
    """
    raw = uploaded_file.read()
    outer = zipfile.ZipFile(io.BytesIO(raw))

    # Chercher le zip interne
    inner_name = None
    for name in outer.namelist():
        if name.endswith(".zip"):
            inner_name = name
            break

    if inner_name:
        inner_bytes = outer.read(inner_name)
        inner = zipfile.ZipFile(io.BytesIO(inner_bytes))
        csv_names = [n for n in inner.namelist() if n.endswith(".csv")]
    else:
        # Pas de zip interne → CSVs directement dans le zip
        inner = outer
        csv_names = [n for n in inner.namelist() if n.endswith(".csv")]

    if not csv_names:
        st.error("Aucun CSV trouvé dans le ZIP.")
        return None

    # Prendre le CSV principal (pas _all)
    main_csv = None
    for n in csv_names:
        if "_all" not in n:
            main_csv = n
            break
    if not main_csv:
        main_csv = csv_names[0]

    csv_bytes = inner.read(main_csv)
    try:
        df = pd.read_csv(io.BytesIO(csv_bytes), encoding="utf-8-sig")
    except Exception:
        df = pd.read_csv(io.BytesIO(csv_bytes), encoding="latin-1")

    return df


def normalize_email(raw) -> str:
    if not raw or (isinstance(raw, float) and pd.isna(raw)) or not isinstance(raw, str):
        return ""
    e = raw.strip()
    e = re.sub(r"^mailto:", "", e, flags=re.IGNORECASE)
    e = re.sub(r"\s+", "", e).lower().strip()
    return e


def prep_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Ajoute colonnes normalisées."""
    # Trouver colonne email
    ecol = None
    for c in df.columns:
        if "email" in c.lower() and "pdp" in c.lower():
            ecol = c
            break
    if not ecol:
        for c in df.columns:
            if "email" in c.lower():
                ecol = c
                break
    df["_email"] = df[ecol].apply(normalize_email) if ecol else ""

    # Normaliser niveau CTX
    ctx_col = None
    for c in df.columns:
        if "niveau" in c.lower() and "contentieux" in c.lower():
            ctx_col = c
            break
    df["_ctx"] = df[ctx_col].fillna("").str.strip() if ctx_col else ""

    # Normaliser STADE
    stade_col = None
    for c in df.columns:
        if c.strip().upper() == "STADE":
            stade_col = c
            break
    df["_stade"] = df[stade_col].fillna("").str.strip() if stade_col else ""

    return df


# ═══════════════════════════════════════════════
# HELPERS — RENDERING
# ═══════════════════════════════════════════════
def kv(key, val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    v = str(val).strip()
    if not v or v == "nan":
        return ""
    v = re.sub(r"\(https://www\.notion\.so/[^\)]+\)", "", v).strip()
    v = v.replace("\n", "<br>")
    return f'<div class="kv"><div class="k">{key}</div><div class="val">{v}</div></div>'


def badge(text, color="#94a3b8"):
    return f'<span class="badge" style="background:{color}">{text}</span>'


def section(title, row, cols, extra_class=""):
    cls = f"fc {extra_class}".strip()
    h = f'<div class="{cls}"><h4>{title}</h4>'
    for c in cols:
        if c in row.index:
            h += kv(c, row[c])
    h += "</div>"
    return h


def stade_badge(s):
    icon, color = STADE_CONF.get(s, ("📌", "#94a3b8"))
    return badge(f"{icon} {s}", color)


def ctx_badge(c):
    if not c:
        return ""
    icon = CTX_ICONS.get(c, "⚪")
    color = CTX_COLORS.get(c, "#94a3b8")
    return badge(f"{icon} {c}", color)


# ═══════════════════════════════════════════════
# HELPERS — MERCI FACTEUR API
# ═══════════════════════════════════════════════
def mf_token_get(pub, sec):
    r = requests.get(f"{API_BASE}/getToken", headers={
        "Authorization": f"Basic {base64.b64encode(f'{pub}:{sec}'.encode()).decode()}"
    }, timeout=15)
    return r.json()

def mf_token_ensure():
    c = st.session_state.mf_creds
    if not c:
        return None
    now = datetime.now()
    if st.session_state.mf_token and st.session_state.mf_token_exp and now < st.session_state.mf_token_exp:
        return st.session_state.mf_token
    r = mf_token_get(c["pub"], c["sec"])
    if "result" in r and "token" in r["result"]:
        st.session_state.mf_token = r["result"]["token"]
        st.session_state.mf_token_exp = now + timedelta(hours=23)
        return st.session_state.mf_token
    return None

def mf(method, ep, data=None):
    t = mf_token_ensure()
    if not t:
        return {"error": "Token MF indisponible"}
    h = {"Authorization": f"Bearer {t}", "Content-Type": "application/json"}
    u = f"{API_BASE}/{ep}"
    try:
        if method == "GET":
            r = requests.get(u, headers=h, params=data, timeout=30)
        elif method == "POST":
            r = requests.post(u, headers=h, json=data, timeout=30)
        elif method == "DELETE":
            r = requests.delete(u, headers=h, params=data, timeout=30)
        else:
            return {"error": "?"}
        return r.json()
    except Exception as e:
        return {"error": str(e)}

def mf_send(uid, dest, exp, files, mode="normal", typ="lettre",
            rv=False, coul=True, date=None, ref=None, ref_u=None):
    p = {"userId": uid, "typeCourrier": typ, "modeEnvoi": mode,
         "rectoVerso": rv, "couleur": coul,
         "destinataires": dest, "jsonExp": exp, "base64Files": files}
    if date:
        p["dateEnvoi"] = date
    if ref:
        p["refInterne"] = ref
    if ref_u:
        p["refUnique"] = ref_u
    return mf("POST", "sendCourrier", p)


# ═══════════════════════════════════════════════
# HELPERS — SMTP
# ═══════════════════════════════════════════════
def smtp_send(host, port, user, pw, frm, to, subj, body, atts=None, tls=True):
    try:
        msg = MIMEMultipart("mixed")
        msg["From"], msg["To"], msg["Subject"] = frm, to, subj
        msg.attach(MIMEText(body, "html", "utf-8"))
        for fn, fb in (atts or []):
            a = MIMEApplication(fb, Name=fn)
            a["Content-Disposition"] = f'attachment; filename="{fn}"'
            msg.attach(a)
        srv = smtplib.SMTP(host, port, timeout=15) if tls else smtplib.SMTP_SSL(host, port, timeout=15)
        if tls:
            srv.starttls()
        srv.login(user, pw)
        srv.sendmail(frm, to, msg.as_string())
        srv.quit()
        return {"ok": True, "msg": f"Envoyé à {to}"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ═══════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════
st.markdown("""<div class="hdr">
<h1>📮 LPB — Courrier, Email & Fiche PDP</h1>
<p>Import ZIP Notion → Recherche par email → Fiche complète (CTX, risque, BODACC) → Envoi lettre ou email avec PJ</p>
</div>""", unsafe_allow_html=True)

with st.expander("ℹ️ À propos de l’application", expanded=False):
    st.markdown(f"""
### Objectif
Cette application interne **La Première Brique** permet de :
- charger un export Notion des projets ;
- rechercher un porteur de projet par email ;
- afficher une fiche PDP complète ;
- préparer et envoyer un courrier papier via **Merci Facteur** ;
- envoyer un email via **SMTP** ;
- conserver un historique local des envois.

### Fonctionnement
1. Importer un **ZIP / CSV Notion**
2. Rechercher un **PDP par email**
3. Consulter la **fiche projet / risque / contentieux / hypothèque**
4. Préparer un **courrier** ou un **email**
5. Envoyer puis suivre les envois

### Données légales préremplies — La Première Brique
- **Dénomination** : {LPB_LEGAL_INFO["denomination"]}
- **SIREN** : {LPB_LEGAL_INFO["siren"]}
- **SIRET siège social** : {LPB_LEGAL_INFO["siret_siege"]}
- **TVA intracommunautaire** : {LPB_LEGAL_INFO["tva_intracom"]}
- **N° EORI** : {LPB_LEGAL_INFO["eori"]}
- **Activité principale** : {LPB_LEGAL_INFO["activite_principale"]}
- **Code NAF / APE** : {LPB_LEGAL_INFO["naf_ape"]}
- **Activité principale (NAF 2025)** : {LPB_LEGAL_INFO["activite_principale_naf_2025"]}
- **Adresse** : {LPB_LEGAL_INFO["adresse_postale"]}, {LPB_LEGAL_INFO["cp"]} {LPB_LEGAL_INFO["ville"]}
- **Forme juridique** : {LPB_LEGAL_INFO["forme_juridique"]}
- **Effectif** : {LPB_LEGAL_INFO["effectif"]}
- **Catégorie d'entreprise** : {LPB_LEGAL_INFO["categorie_entreprise"]}
- **Date de création** : {LPB_LEGAL_INFO["date_creation"]}
- **Convention collective** : IDCC {LPB_LEGAL_INFO["idcc"]}

### Référentiels
- INSEE
- VIES
- Douanes
- INPI
- État des inscriptions

### État des inscriptions
- **{LPB_LEGAL_INFO["insee_statut"]}** le {LPB_LEGAL_INFO["date_inscription_insee"]}
- **{LPB_LEGAL_INFO["inpi_statut"]}** le {LPB_LEGAL_INFO["date_rne_inpi"]}
""")
# ═══════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 📂 Import Notion")
    up = st.file_uploader("Glisse ton **export ZIP** Notion", type=["zip", "csv"],
                          help="ZIP d'export Notion (Projets Immobiliers)", key="up")
    if up:
        with st.spinner("Extraction..."):
            if up.name.endswith(".zip"):
                df = extract_csv_from_zip(up)
            else:
                try:
                    df = pd.read_csv(up, encoding="utf-8-sig")
                except Exception:
                    up.seek(0)
                    df = pd.read_csv(up, encoding="latin-1")
            if df is not None:
                df = prep_dataframe(df)
                st.session_state.df = df
                ne = (df["_email"] != "").sum()
                nc = (df["_ctx"] != "").sum()
                st.markdown(f'<div class="sb sb-ok">✅ <b>{len(df)}</b> projets — <b>{ne}</b> emails — <b>{nc}</b> CTX/préCTX</div>', unsafe_allow_html=True)
    elif st.session_state.df is not None:
        st.markdown(f'<div class="sb sb-info">📋 {len(st.session_state.df)} projets en mémoire</div>', unsafe_allow_html=True)

    st.markdown("---")
    with st.expander("🏢 Fiche légale LPB", expanded=False):
        st.markdown(f"""
    **Dénomination**  
    {LPB_LEGAL_INFO["denomination"]}
    
    **SIREN**  
    {LPB_LEGAL_INFO["siren"]}
    
    **SIRET du siège social**  
    {LPB_LEGAL_INFO["siret_siege"]}
    
    **N° TVA intracommunautaire**  
    {LPB_LEGAL_INFO["tva_intracom"]}
    
    **N° EORI**  
    {LPB_LEGAL_INFO["eori"]}
    
    **Activité principale**  
    {LPB_LEGAL_INFO["activite_principale"]}
    
    **Code NAF / APE**  
    {LPB_LEGAL_INFO["naf_ape"]}
    
    **Activité principale (NAF 2025)**  
    {LPB_LEGAL_INFO["activite_principale_naf_2025"]}
    
    **Adresse postale**  
    {LPB_LEGAL_INFO["adresse_postale"]}  
    {LPB_LEGAL_INFO["cp"]} {LPB_LEGAL_INFO["ville"]}
    
    **Forme juridique**  
    {LPB_LEGAL_INFO["forme_juridique"]}
    
    **Effectif salarié**  
    {LPB_LEGAL_INFO["effectif"]}
    
    **Catégorie d'entreprise**  
    {LPB_LEGAL_INFO["categorie_entreprise"]}
    
    **Date de création**  
    {LPB_LEGAL_INFO["date_creation"]}
    
    **Convention collective**  
    IDCC {LPB_LEGAL_INFO["idcc"]}
    
    **État des inscriptions**  
    - {LPB_LEGAL_INFO["insee_statut"]} le {LPB_LEGAL_INFO["date_inscription_insee"]}  
    - {LPB_LEGAL_INFO["inpi_statut"]} le {LPB_LEGAL_INFO["date_rne_inpi"]}
    """)
    st.markdown("## ⚙️ Connexions")
    with st.expander("🏠 Merci Facteur", expanded=False):
        mp = st.text_input("Clé publique", type="password", key="s_mp")
        ms = st.text_input("Clé secrète", type="password", key="s_ms")
        if st.button("🔗 Connecter", key="s_mc", use_container_width=True):
            if mp and ms:
                st.session_state.mf_creds = {"pub": mp, "sec": ms}
                if mf_token_ensure():
                    st.success("✅")
                else:
                    st.error("❌")
        mu = st.text_input("User ID MF", value=st.session_state.mf_uid or "", key="s_mu")
        if mu:
            st.session_state.mf_uid = mu
        if st.session_state.mf_token:
            st.markdown('<div class="sb sb-ok">🟢 Token actif</div>', unsafe_allow_html=True)

    with st.expander("📧 SMTP", expanded=False):
                st.markdown("""
### Guide SMTP Gmail
1. Active la validation en 2 étapes sur ton compte Google.
2. Génère ensuite un mot de passe d’application.
3. Utilise ensuite :
   - **Serveur** : `smtp.gmail.com`
   - **Port** : `587`
   - **TLS** : activé
   - **User** : ton adresse Gmail complète
   - **Pass** : le mot de passe d’application généré
   - **From** : en général la même adresse email

**Liens utiles**
- Mot de passe d’application Google : https://myaccount.google.com/apppasswords
- Aide officielle Google : https://support.google.com/accounts/answer/185833
""")
        st.caption("Pour Gmail : utiliser smtp.gmail.com, port 587, TLS activé, et un mot de passe d’application Google.")
        sh = st.text_input("Serveur", value="smtp.gmail.com", key="s_sh")
        sp = st.number_input("Port", value=587, key="s_sp")
        su = st.text_input("User", key="s_su")
        sw = st.text_input("Pass", type="password", key="s_sw")
        sf = st.text_input("From", value=su, key="s_sf")
        stl = st.checkbox("TLS", value=True, key="s_tl")
        st.session_state.smtp = {"h": sh, "p": sp, "u": su, "w": sw, "f": sf or su, "t": stl}
        if su and sw:
            st.markdown('<div class="sb sb-ok">🟢 SMTP OK</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════
# TABS
# ═══════════════════════════════════════════════
t1, t2, t3, t4 = st.tabs(["🔍 Recherche & Fiche", "📬 Envoyer", "🚚 Suivi MF", "📋 Historique"])

# ───────────────────────────────────────────────
# TAB 1 — RECHERCHE & FICHE PDP
# ───────────────────────────────────────────────
with t1:
    df = st.session_state.df
    if df is None:
        st.info("👈 Importe ton export ZIP Notion dans la sidebar.")
    else:
        # ── Dashboard stats ──
        st.markdown("### 📊 Vue d'ensemble")
        ctx_counts = df["_ctx"].value_counts()
        n_ctx = ctx_counts.get("Contentieux", 0)
        n_pctx = ctx_counts.get("Précontentieux", 0)
        stade_counts = df["_stade"].value_counts()

        cols = st.columns(7)
        stats = [
            ("Total", len(df), "#1e293b"),
            ("Financés", stade_counts.get("Financé", 0), "#16a34a"),
            ("Remboursés", stade_counts.get("Remboursé", 0), "#6b7280"),
            ("Négociation", stade_counts.get("Négociation", 0), "#8b5cf6"),
            ("Abandonné", stade_counts.get("Abandonné", 0), "#374151"),
            ("PréCTX", n_pctx, "#e67e22"),
            ("CTX", n_ctx, "#dc3545"),
        ]
        for col, (lbl, val, clr) in zip(cols, stats):
            col.markdown(f'<div class="sc"><div class="v" style="color:{clr}">{val}</div><div class="l">{lbl}</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        # ── Recherche ──
        st.markdown("### 🔎 Recherche PDP")
        c_s, c_b = st.columns([4, 1])
        with c_s:
            srch = st.text_input("Email du porteur de projet", placeholder="email@domaine.com",
                                 key="srch", label_visibility="collapsed")
        with c_b:
            do_search = st.button("🔍 Rechercher", use_container_width=True)

        snorm = normalize_email(srch)

        if do_search and snorm:
            m = df[df["_email"] == snorm]
            if m.empty:
                m = df[df["_email"].str.contains(snorm, case=False, na=False)]
            if m.empty:
                st.markdown(f'<div class="sb sb-warn">⚠️ Aucun projet pour <b>{snorm}</b></div>', unsafe_allow_html=True)
                st.session_state.pdp_rows = None
            else:
                st.session_state.pdp_email = snorm
                st.session_state.pdp_rows = m

        # ── Fiche PDP ──
        pr = st.session_state.pdp_rows
        if pr is not None and len(pr) > 0:
            f0 = pr.iloc[0]
            nom = str(f0.get("Nom du porteur", "")).strip()
            soc = str(f0.get("Société (SIREN)", "")).strip()
            tel = str(f0.get("Téléphone", "")).strip()
            em = st.session_state.pdp_email
            ctx_list = pr["_ctx"].value_counts()

            st.markdown("---")
            st.markdown(f"### 👤 Fiche PDP — {nom}")

            # 3 blocs résumé
            r1, r2, r3 = st.columns(3)
            with r1:
                h = '<div class="fc"><h4>🪪 Identité</h4>'
                h += kv("Nom", nom)
                h += kv("Société (SIREN)", soc)
                siren_clean = str(f0.get("🔑 SIREN (clean)", "")).strip()
                if siren_clean and siren_clean != "nan":
                    h += kv("SIREN", siren_clean)
                h += kv("Téléphone", tel)
                h += kv("Email", em)
                # Notaire
                notaire = str(f0.get("Notaire du PDP", "")).strip()
                if notaire and notaire != "nan":
                    n_clean = re.sub(r"mailto:\S+", "", notaire).strip()
                    n_emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", notaire)
                    if n_clean:
                        h += kv("Notaire", n_clean)
                    if n_emails:
                        h += kv("Email notaire", ", ".join(e.lower() for e in n_emails))
                h += '</div>'
                st.markdown(h, unsafe_allow_html=True)

            with r2:
                h = '<div class="fc"><h4>📊 Projets</h4>'
                h += kv("Total", str(len(pr)))
                for s, cnt in pr["_stade"].value_counts().items():
                    if s:
                        h += f'<div class="kv"><div class="k">{stade_badge(s)}</div><div class="val">{cnt}</div></div>'
                h += '</div>'
                st.markdown(h, unsafe_allow_html=True)

            with r3:
                has_ctx = (pr["_ctx"] == "Contentieux").any()
                has_pctx = (pr["_ctx"] == "Précontentieux").any()
                cls = "fc fc-alert" if has_ctx else ("fc fc-warn" if has_pctx else "fc")
                h = f'<div class="{cls}"><h4>🚨 Alertes Contentieux</h4>'
                if has_ctx:
                    n = (pr["_ctx"] == "Contentieux").sum()
                    h += f'<div class="sb sb-err">🔴 <b>{n} projet(s) en CONTENTIEUX</b></div>'
                if has_pctx:
                    n = (pr["_ctx"] == "Précontentieux").sum()
                    h += f'<div class="sb sb-warn">🟠 <b>{n} projet(s) en PRÉCONTENTIEUX</b></div>'
                if not has_ctx and not has_pctx:
                    h += '<div class="sb sb-ok">✅ Aucun contentieux</div>'

                # Résumé risque
                for _, row in pr.iterrows():
                    risk = str(row.get("Niveau de risque temporel", "")).strip()
                    if risk and risk != "nan" and "risque" in risk.lower():
                        pname = str(row.get("Nom du projet", "")).strip()[:30]
                        h += f'<div class="kv"><div class="k" style="font-size:.78rem">{pname}</div><div class="val" style="font-size:.78rem">{risk}</div></div>'
                h += '</div>'
                st.markdown(h, unsafe_allow_html=True)

            # ── Détail de chaque projet ──
            st.markdown("---")
            st.markdown(f"### 📁 {len(pr)} projet(s)")

            # Trier : CTX en premier
            sort_order = {"Contentieux": 0, "Précontentieux": 1, "": 2}
            pr_sorted = pr.copy()
            pr_sorted["_sort"] = pr_sorted["_ctx"].map(sort_order).fillna(2)
            pr_sorted = pr_sorted.sort_values("_sort")

            for idx, (_, row) in enumerate(pr_sorted.iterrows()):
                ctx = str(row.get("_ctx", "")).strip()
                stade = str(row.get("_stade", "")).strip()
                nom_p = str(row.get("Nom du projet", "")).strip()
                id_bo = str(row.get("🔑 ID Back-office", "")).strip()
                is_alert = ctx in ("Contentieux", "Précontentieux")

                label_parts = []
                if ctx:
                    label_parts.append(CTX_ICONS.get(ctx, "") + " " + ctx)
                si, _ = STADE_CONF.get(stade, ("📌", ""))
                label_parts.append(f"{si} {stade}" if stade else "")
                label = " | ".join(p for p in label_parts if p)

                with st.expander(f"**{nom_p}** — ID {id_bo} — {label}", expanded=is_alert):
                    # Si CTX → section CTX en premier
                    if is_alert:
                        st.markdown(section("🚨 Contentieux / Juridique", row, SEC_CTX,
                                            "fc-alert" if ctx == "Contentieux" else "fc-warn"),
                                    unsafe_allow_html=True)
                        st.markdown(section("📉 Risque & Montants", row, SEC_RISQUE), unsafe_allow_html=True)

                    p1, p2 = st.columns(2)
                    with p1:
                        st.markdown(section("📋 Projet", row, SEC_PROJET), unsafe_allow_html=True)
                        st.markdown(section("📅 Dates & Durées", row, SEC_DATES), unsafe_allow_html=True)
                        if not is_alert:
                            st.markdown(section("📉 Risque & Montants", row, SEC_RISQUE), unsafe_allow_html=True)
                    with p2:
                        st.markdown(section("💰 Financier", row, SEC_FINANCIER), unsafe_allow_html=True)
                        st.markdown(section("📌 Suivi & Garanties", row, SEC_SUIVI), unsafe_allow_html=True)
                        st.markdown(section("🏠 Hypothèque", row, SEC_HYPOTHEQUE), unsafe_allow_html=True)

                    # Actions
                    a1, a2, a3 = st.columns(3)
                    with a1:
                        if st.button("📮 Lettre", key=f"l{idx}"):
                            st.session_state["pf"] = {
                                "nom": nom, "soc": soc.split(" - ")[0] if " - " in soc else soc,
                                "adr": str(row.get("Adresse et lots", "")).strip(),
                                "ref": f"Projet {id_bo} - {nom_p}", "mode": "lettre",
                            }
                            st.info("→ Onglet 'Envoyer'")
                    with a2:
                        if st.button("📧 Email", key=f"e{idx}"):
                            st.session_state["pf"] = {
                                "email": em, "ref": f"Projet {id_bo} - {nom_p}", "mode": "email",
                            }
                            st.info("→ Onglet 'Envoyer'")
                    with a3:
                        st.download_button("📥 CSV", row.to_frame().T.to_csv(index=False),
                                           f"projet_{id_bo}.csv", "text/csv", key=f"x{idx}")

            st.markdown("---")
            st.download_button("📥 Export complet PDP (CSV)", pr.to_csv(index=False),
                               f"pdp_{em.replace('@','_')}.csv", "text/csv", key="xall")


# ───────────────────────────────────────────────
# TAB 2 — ENVOI
# ───────────────────────────────────────────────
with t2:
    pf = st.session_state.pop("pf", {})
    mode = st.radio(
        "**Mode**",
        ["📮 Lettre (Merci Facteur)", "📧 Email (SMTP)"],
        index=1 if pf.get("mode") == "email" else 0,
        horizontal=True
    )
    st.markdown("---")
    st.info(
        f"Les informations expéditeur courrier sont préremplies avec les données légales de "
        f"{LPB_LEGAL_INFO['denomination']} (SIREN {LPB_LEGAL_INFO['siren']})."
    )

    st.markdown("### 🎯 Destinataire")
    d1, d2 = st.columns(2)
    with d1:
        dc = st.selectbox("Civilité", ["M", "Mme", ""], key="dc")
        dn = st.text_input("Nom *", value=pf.get("nom", ""), key="dn")
        dp = st.text_input("Prénom", key="dp")
        ds = st.text_input("Société", value=pf.get("soc", ""), key="ds")
    with d2:
        da1 = st.text_input("Adresse 1 *", value=pf.get("adr", ""), key="da1")
        da2 = st.text_input("Adresse 2", key="da2")
        dcp = st.text_input("CP *", key="dcp")
        dv = st.text_input("Ville *", key="dv")
    d3, d4 = st.columns(2)
    with d3:
        dpays = st.text_input("Pays", value="France", key="dpy")
    with d4:
        dem = st.text_input(
            "Email" + (" *" if "Email" in mode else ""),
            value=pf.get("email", ""),
            key="dem"
        )

    st.markdown("---")

    if "Lettre" in mode:
        st.markdown("### 🏢 Expéditeur")
        e1, e2 = st.columns(2)
        with e1:
            en = st.text_input("Nom exp. *", value="", key="en")
            epr = st.text_input("Prénom exp.", key="epr")
            eso = st.text_input("Société exp.", value=LPB_LEGAL_INFO["denomination"], key="eso")
        with e2:
            ea = st.text_input("Adresse exp. *", value=LPB_LEGAL_INFO["adresse_postale"], key="ea")
            ec = st.text_input("CP exp. *", value=LPB_LEGAL_INFO["cp"], key="ec")
            evi = st.text_input("Ville exp. *", value=LPB_LEGAL_INFO["ville"], key="evi")
            epy = st.text_input("Pays exp.", value=LPB_LEGAL_INFO["pays"], key="epy")

        st.markdown("---")
        st.markdown("### 📝 Contenu")
        o1, o2, o3 = st.columns(3)
        with o1:
            tc = st.selectbox("Type", list(TYPES_COURRIER.keys()), key="tc")
        with o2:
            me = st.selectbox("Mode postal", list(MODES_ENVOI.keys()), key="me")
        with o3:
            rv = st.checkbox("Recto-verso", key="rv")
            cl = st.checkbox("Couleur", value=True, key="cl")

        pdfs = st.file_uploader("📎 PDF(s) *", type=["pdf"], accept_multiple_files=True, key="pdfs")

        r1, r2 = st.columns(2)
        with r1:
            dt = st.date_input(
                "Date programmée",
                value=datetime.now().date(),
                min_value=datetime.now().date(),
                key="dt"
            )
        with r2:
            ri = st.text_input("Réf. interne", value=pf.get("ref", ""), key="ri")
            ru = st.text_input("Réf. anti-doublon", key="ru", max_chars=200)

    else:
        st.markdown("### 📝 Contenu email")
        sj = st.text_input(
            "Objet *",
            value=f"Concernant : {pf.get('ref','')}" if pf.get("ref") else "",
            key="sj"
        )
        bd = st.text_area(
            "Corps (HTML) *",
            height=220,
            placeholder="<p>Bonjour,</p>\n<p>...</p>\n<p>Cordialement,<br>LPB</p>",
            key="bd"
        )
        at = st.file_uploader("📎 Pièces jointes", accept_multiple_files=True, key="at")

    st.markdown("---")
    cs, cr = st.columns([4, 1])

    with cs:
        if st.button("🚀 ENVOYER", type="primary", use_container_width=True, key="go"):
            if "Lettre" in mode:
                errs = []
                if not dn and not ds:
                    errs.append("Nom/société dest.")
                if not dcp:
                    errs.append("CP")
                if not dv:
                    errs.append("Ville")
                if not pdfs:
                    errs.append("PDF(s)")
                if not st.session_state.mf_token:
                    errs.append("Non connecté MF")
                if not st.session_state.mf_uid:
                    errs.append("User ID MF")

                if errs:
                    st.error("❌ Manquant : " + ", ".join(errs))
                else:
                    with st.spinner("📮 Envoi..."):
                        fb = [
                            {"fileName": f.name, "fileBase64": base64.b64encode(f.read()).decode()}
                            for f in pdfs
                        ]
                        res = mf_send(
                            st.session_state.mf_uid,
                            [{
                                "civilite": dc,
                                "nom": dn,
                                "prenom": dp,
                                "societe": ds,
                                "adresse1": da1,
                                "adresse2": da2,
                                "adresse3": "",
                                "cp": dcp,
                                "ville": dv,
                                "pays": dpays,
                                "email": dem
                            }],
                            {
                                "nom": en,
                                "prenom": epr,
                                "societe": eso,
                                "adresse1": ea,
                                "adresse2": "",
                                "adresse3": "",
                                "cp": ec,
                                "ville": evi,
                                "pays": epy
                            },
                            fb,
                            MODES_ENVOI[me],
                            TYPES_COURRIER[tc],
                            rv,
                            cl,
                            str(dt) if dt else None,
                            ri or None,
                            ru or None,
                        )

                        if "error" not in res and res.get("result"):
                            st.markdown(
                                '<div class="sb sb-ok">✅ <b>Courrier envoyé !</b></div>',
                                unsafe_allow_html=True
                            )
                            st.json(res["result"])
                            st.session_state.history.append({
                                "type": "📮",
                                "dest": f"{dn} {dp}".strip(),
                                "mode": me,
                                "ref": ri,
                                "date": datetime.now().strftime("%Y-%m-%d %H:%M")
                            })
                        else:
                            st.markdown(
                                f'<div class="sb sb-err">❌ {json.dumps(res, ensure_ascii=False)}</div>',
                                unsafe_allow_html=True
                            )

            else:
                errs = []
                if not dem:
                    errs.append("Email dest.")
                if not sj:
                    errs.append("Objet")
                if not bd:
                    errs.append("Corps")

                sm = st.session_state.smtp
                if not sm.get("u") or not sm.get("w"):
                    errs.append("SMTP")

                if errs:
                    st.error("❌ Manquant : " + ", ".join(errs))
                else:
                    with st.spinner("📧 Envoi..."):
                        al = [(f.name, f.read()) for f in at] if at else None
                        res = smtp_send(
                            sm["h"], sm["p"], sm["u"], sm["w"], sm["f"],
                            dem, sj, bd, al, sm.get("t", True)
                        )

                        if res["ok"]:
                            st.markdown(
                                '<div class="sb sb-ok">✅ <b>Email envoyé !</b></div>',
                                unsafe_allow_html=True
                            )
                            st.session_state.history.append({
                                "type": "📧",
                                "dest": dem,
                                "sujet": sj,
                                "pj": len(al) if al else 0,
                                "date": datetime.now().strftime("%Y-%m-%d %H:%M")
                            })
                        else:
                            st.markdown(
                                f'<div class="sb sb-err">❌ {res["msg"]}</div>',
                                unsafe_allow_html=True
                            )

    with cr:
        if st.button("🗑️", use_container_width=True, key="rst"):
            st.rerun()


# ───────────────────────────────────────────────
# TAB 3 — SUIVI MF
# ───────────────────────────────────────────────
with t3:
    st.markdown("### 🚚 Suivi Merci Facteur")
    if not st.session_state.mf_token:
        st.info("Connecte-toi à MF dans la sidebar.")
    else:
        s1, s2, s3 = st.tabs(["📋 Envois", "🔎 Détail", "📊 Quotas"])
        with s1:
            if st.button("🔄 Charger", key="sl"):
                st.json(mf("GET", "listEnvois", {"userId": st.session_state.mf_uid}).get("result", {}))
        with s2:
            eid = st.text_input("ID envoi", key="sei")
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("📄", key="sd") and eid:
                    st.json(mf("GET", "getEnvoi", {"idEnvoi": eid}).get("result", {}))
            with c2:
                if st.button("🚚", key="ss") and eid:
                    st.json(mf("GET", "getSuiviEnvoi", {"idEnvoi": eid}).get("result", {}))
            with c3:
                if st.button("📜", key="sp") and eid:
                    st.json(mf("GET", "getProof", {"idEnvoi": eid}).get("result", {}))
        with s3:
            if st.button("📊", key="sq"):
                st.json(mf("GET", "getQuotaCompte").get("result", {}))


# ───────────────────────────────────────────────
# TAB 4 — HISTORIQUE
# ───────────────────────────────────────────────
with t4:
    st.markdown("### 📋 Historique")
    if not st.session_state.history:
        st.info("Aucun envoi.")
    else:
        for i, e in enumerate(reversed(st.session_state.history)):
            with st.expander(f"{e['type']} → {e.get('dest','?')} — {e['date']}", expanded=(i == 0)):
                st.json(e)
        if st.button("🗑️ Vider", key="hc"):
            st.session_state.history = []
            st.rerun()

st.markdown("---")
st.markdown(
    f'<div style="text-align:center;opacity:.35;font-size:.75rem">'
    f'📮 LPB v3 — {LPB_LEGAL_INFO["denomination"]} — SIREN {LPB_LEGAL_INFO["siren"]} — '
    f'Merci Facteur API + SMTP — Données locales'
    f'</div>',
    unsafe_allow_html=True
)

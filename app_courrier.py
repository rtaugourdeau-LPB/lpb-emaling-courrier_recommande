"""
📮 LPB — Interface Courrier & Email avec Fiche PDP
Streamlit app — Notion CSV import + Merci Facteur API + SMTP Email
"""

import streamlit as st
import pandas as pd
import requests
import smtplib
import base64
import json
import re
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
API_BASE = "https://www.merci-facteur.com/api/1.2/prod/service"

MODES_ENVOI = {
    "Lettre verte (normal)": "normal",
    "Lettre suivie": "suivi",
    "Recommandé avec AR": "lrar",
    "Recommandé avec AR numérisé": "lrare",
    "Recommandé électronique (email OTP)": "ERE_OTP_MAIL",
    "Recommandé électronique (SMS OTP)": "ERE_OTP_SMS",
}

TYPES_COURRIER = {
    "Lettre": "lettre",
    "Carte postale": "carte_postale",
    "Carte postale sous enveloppe": "carte_postale_enveloppe",
    "Carte pliée": "carte_pliee",
    "Photo": "photo",
}

STADE_COLORS = {
    "Contentieux": "#dc3545",
    "Précontentieux": "#e67e22",
    "Précontentieux, Retard": "#e67e22",
    "Retard": "#f39c12",
    "Financé, Précontentieux": "#e67e22",
    "Financé": "#28a745",
    "Remboursé": "#6c757d",
    "Audit": "#17a2b8",
    "Structuration": "#6610f2",
    "Commercial": "#20c997",
    "Prise de contact": "#adb5bd",
    "Dead": "#343a40",
    "A archiver, Dead": "#343a40",
}

STADE_ICONS = {
    "Contentieux": "🔴",
    "Précontentieux": "🟠",
    "Précontentieux, Retard": "🟠",
    "Retard": "🟡",
    "Financé, Précontentieux": "🟠",
    "Financé": "🟢",
    "Remboursé": "⚪",
    "Audit": "🔵",
    "Structuration": "🟣",
    "Commercial": "💚",
    "Prise de contact": "⬜",
    "Dead": "⚫",
    "A archiver, Dead": "⚫",
}

FICHE_COLS_IDENTITE = [
    "Nom du porteur", "Société (SIREN)", "Téléphone", "Email PDP", "Notaire du PDP",
]
FICHE_COLS_PROJET = [
    "ID Back-office", "Nom du projet", "Stade", "Type d'opération",
    "Description du projet", "Adresse et lots",
]
FICHE_COLS_FINANCIER = [
    "Montant à financer", "Taux", "Fees (ht)", "CA net",
    "Ticket max ", "Type d'intérêts", "Type de décaissement", "Type de financement ",
]
FICHE_COLS_DATES = [
    "Début du prêt", "Date de collecte", "Décaissement", "Date de réitération",
    "Date mini", "Date cible", "Date maximale intiale", "Date maximale prolongée",
    "Durée mini", "Durée cible", "Durée maxi initiale", "Durée prolongée ( en mois)",
]
FICHE_COLS_SUIVI = [
    "Chargé de projet", "Chargé de suivi", "Contrats", "KYB",
    "Etat urbanisme", "Garantie(s)", "CS de la préqual",
    "Dernière actualité + Statut", "Date dernière NL",
    "Warning dernière NL", "Warning Point date maximale", "Warning Date cible",
    "Dates des points 90.60.30j",
]


# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="📮 LPB — Courrier & Fiche PDP",
    page_icon="📮",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');
    .stApp { font-family: 'DM Sans', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%);
        padding: 1.8rem 2.2rem; border-radius: 14px; margin-bottom: 1.5rem; color: white;
    }
    .main-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .main-header p { margin: 0.4rem 0 0 0; opacity: 0.75; font-size: 0.95rem; }
    .fiche-card {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 12px; padding: 1.2rem 1.5rem; margin-bottom: 1rem;
    }
    .fiche-card h4 { margin: 0 0 0.8rem 0; color: #1e293b; font-weight: 600; }
    .badge {
        display: inline-block; padding: 4px 12px; border-radius: 20px;
        font-size: 0.82rem; font-weight: 600; color: white; margin: 2px 4px 2px 0;
    }
    .stat-card {
        text-align: center; background: white; border: 1px solid #e2e8f0;
        border-radius: 10px; padding: 1rem;
    }
    .stat-card .value { font-size: 1.6rem; font-weight: 700; color: #1e293b; }
    .stat-card .label { font-size: 0.8rem; color: #64748b; margin-top: 4px; }
    .kv-row { display: flex; padding: 4px 0; border-bottom: 1px solid #f1f5f9; }
    .kv-key { flex: 0 0 200px; font-weight: 600; color: #475569; font-size: 0.88rem; }
    .kv-val { flex: 1; color: #1e293b; font-size: 0.88rem; word-break: break-word; }
    .status-box {
        padding: 0.8rem 1.2rem; border-radius: 10px; margin: 0.6rem 0;
        font-weight: 500; font-size: 0.9rem;
    }
    .status-success { background: #d1fae5; border-left: 4px solid #10b981; color: #065f46; }
    .status-error { background: #fee2e2; border-left: 4px solid #ef4444; color: #991b1b; }
    .status-info { background: #dbeafe; border-left: 4px solid #3b82f6; color: #1e40af; }
    .status-warn { background: #fef3c7; border-left: 4px solid #f59e0b; color: #92400e; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
defaults = {
    "df_projets": None, "selected_pdp_email": None, "selected_pdp_data": None,
    "mf_token": None, "mf_token_expiry": None, "mf_user_id": None,
    "mf_creds": None, "smtp_config": {}, "send_history": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ─────────────────────────────────────────────
# HELPERS — EMAIL
# ─────────────────────────────────────────────
def normalize_email(raw: str) -> str:
    if not raw or not isinstance(raw, str):
        return ""
    email = raw.strip()
    email = re.sub(r"^mailto:", "", email, flags=re.IGNORECASE)
    email = re.sub(r"\s+", "", email).lower().strip()
    return email


def extract_emails_from_field(raw: str) -> list:
    if not raw or not isinstance(raw, str):
        return []
    return [e.lower().strip() for e in re.findall(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}", raw)]


# ─────────────────────────────────────────────
# HELPERS — DATA
# ─────────────────────────────────────────────
def load_notion_csv(uploaded_file) -> pd.DataFrame:
    try:
        df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
    except Exception:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding="latin-1")

    # Trouver la colonne email PDP
    email_col = None
    for col in df.columns:
        if "email" in col.lower() and "pdp" in col.lower():
            email_col = col
            break
    if not email_col:
        for col in df.columns:
            if "email" in col.lower():
                email_col = col
                break
    df["_email_normalized"] = df[email_col].apply(normalize_email) if email_col else ""
    return df


def get_stade_class(stade: str) -> str:
    s = (stade or "").lower()
    if "contentieux" in s and "pré" not in s:
        return "contentieux"
    if "précontentieux" in s:
        return "precontentieux"
    if "retard" in s:
        return "retard"
    if "financé" in s:
        return "finance"
    if "remboursé" in s:
        return "rembourse"
    return ""


# ─────────────────────────────────────────────
# HELPERS — RENDERING
# ─────────────────────────────────────────────
def render_kv(key: str, val):
    if not val or (isinstance(val, str) and not val.strip()) or (isinstance(val, float) and pd.isna(val)):
        return ""
    val_str = str(val).strip()
    val_str = re.sub(r"\(https://www\.notion\.so/[^\)]+\)", "", val_str).strip()
    val_str = val_str.replace("\n", "<br>")
    return f'<div class="kv-row"><div class="kv-key">{key}</div><div class="kv-val">{val_str}</div></div>'


def render_badge(stade: str) -> str:
    color = STADE_COLORS.get(stade, "#94a3b8")
    icon = STADE_ICONS.get(stade, "📌")
    return f'<span class="badge" style="background:{color}">{icon} {stade}</span>'


def render_section(title: str, row: pd.Series, cols: list) -> str:
    html = f'<div class="fiche-card"><h4>{title}</h4>'
    for col in cols:
        if col in row.index:
            html += render_kv(col, row[col])
    html += "</div>"
    return html


# ─────────────────────────────────────────────
# HELPERS — MERCI FACTEUR API
# ─────────────────────────────────────────────
def mf_get_token(pub: str, sec: str) -> dict:
    r = requests.get(f"{API_BASE}/getToken", headers={
        "Authorization": f"Basic {base64.b64encode(f'{pub}:{sec}'.encode()).decode()}"
    }, timeout=15)
    return r.json()


def mf_ensure_token():
    creds = st.session_state.get("mf_creds")
    if not creds:
        return None
    now = datetime.now()
    if st.session_state.mf_token and st.session_state.mf_token_expiry and now < st.session_state.mf_token_expiry:
        return st.session_state.mf_token
    result = mf_get_token(creds["public"], creds["secret"])
    if "result" in result and "token" in result["result"]:
        st.session_state.mf_token = result["result"]["token"]
        st.session_state.mf_token_expiry = now + timedelta(hours=23)
        return st.session_state.mf_token
    return None


def mf_api(method, endpoint, data=None):
    token = mf_ensure_token()
    if not token:
        return {"error": "Token MF non disponible"}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    url = f"{API_BASE}/{endpoint}"
    try:
        if method == "GET":
            r = requests.get(url, headers=headers, params=data, timeout=30)
        elif method == "POST":
            r = requests.post(url, headers=headers, json=data, timeout=30)
        elif method == "DELETE":
            r = requests.delete(url, headers=headers, params=data, timeout=30)
        else:
            return {"error": f"Méthode non supportée: {method}"}
        return r.json()
    except Exception as e:
        return {"error": str(e)}


def mf_send_letter(user_id, destinataires, expediteur, fichiers_b64,
                   mode_envoi="normal", type_courrier="lettre",
                   recto_verso=False, couleur=True, date_envoi=None,
                   ref_interne=None, ref_unique=None):
    payload = {
        "userId": user_id, "typeCourrier": type_courrier, "modeEnvoi": mode_envoi,
        "rectoVerso": recto_verso, "couleur": couleur,
        "destinataires": destinataires, "jsonExp": expediteur, "base64Files": fichiers_b64,
    }
    if date_envoi:
        payload["dateEnvoi"] = date_envoi
    if ref_interne:
        payload["refInterne"] = ref_interne
    if ref_unique:
        payload["refUnique"] = ref_unique
    return mf_api("POST", "sendCourrier", payload)


# ─────────────────────────────────────────────
# HELPERS — SMTP
# ─────────────────────────────────────────────
def send_email_smtp(smtp_host, smtp_port, smtp_user, smtp_pass, from_addr,
                    to_addr, subject, body_html, attachments=None, use_tls=True):
    try:
        msg = MIMEMultipart("mixed")
        msg["From"] = from_addr
        msg["To"] = to_addr
        msg["Subject"] = subject
        msg.attach(MIMEText(body_html, "html", "utf-8"))
        if attachments:
            for fname, fbytes in attachments:
                att = MIMEApplication(fbytes, Name=fname)
                att["Content-Disposition"] = f'attachment; filename="{fname}"'
                msg.attach(att)
        if use_tls:
            srv = smtplib.SMTP(smtp_host, smtp_port, timeout=15)
            srv.starttls()
        else:
            srv = smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=15)
        srv.login(smtp_user, smtp_pass)
        srv.sendmail(from_addr, to_addr, msg.as_string())
        srv.quit()
        return {"success": True, "message": f"Email envoyé à {to_addr}"}
    except Exception as e:
        return {"success": False, "message": str(e)}


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📮 LPB — Courrier, Email & Fiche PDP</h1>
    <p>Import CSV Notion → Recherche par email → Fiche complète PDP → Envoi lettre physique ou email avec PJ</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Import Notion")
    uploaded_csv = st.file_uploader(
        "Glisse ton export CSV ici", type=["csv"],
        help="Export CSV de la base 'Projets Immobiliers' Notion", key="csv_up",
    )
    if uploaded_csv:
        df = load_notion_csv(uploaded_csv)
        st.session_state.df_projets = df
        n_emails = (df["_email_normalized"] != "").sum()
        st.markdown(f'<div class="status-box status-success">✅ <b>{len(df)}</b> projets — <b>{n_emails}</b> emails PDP</div>', unsafe_allow_html=True)
    elif st.session_state.df_projets is not None:
        st.markdown(f'<div class="status-box status-info">📋 {len(st.session_state.df_projets)} projets en mémoire</div>', unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("## ⚙️ Connexions")

    with st.expander("🏠 Merci Facteur", expanded=False):
        mf_pub = st.text_input("Clé publique", type="password", key="sb_mf_pub")
        mf_sec = st.text_input("Clé secrète", type="password", key="sb_mf_sec")
        if st.button("🔗 Connecter MF", key="sb_mf_btn", use_container_width=True):
            if mf_pub and mf_sec:
                st.session_state.mf_creds = {"public": mf_pub, "secret": mf_sec}
                if mf_ensure_token():
                    st.success("✅ Connecté !")
                else:
                    st.error("❌ Échec")
        mf_uid = st.text_input("User ID MF", value=st.session_state.get("mf_user_id") or "", key="sb_mf_uid")
        if mf_uid:
            st.session_state.mf_user_id = mf_uid
        if st.session_state.mf_token:
            st.markdown('<div class="status-box status-success">🟢 Token actif</div>', unsafe_allow_html=True)

    with st.expander("📧 SMTP (Email)", expanded=False):
        sh = st.text_input("Serveur SMTP", value="smtp.gmail.com", key="sb_sh")
        sp = st.number_input("Port", value=587, key="sb_sp")
        su = st.text_input("Utilisateur", key="sb_su")
        spw = st.text_input("Mot de passe", type="password", key="sb_spw")
        sf = st.text_input("Email expéditeur", key="sb_sf")
        stls = st.checkbox("STARTTLS", value=True, key="sb_stls")
        st.session_state.smtp_config = {
            "host": sh, "port": sp, "user": su, "password": spw,
            "from": sf or su, "tls": stls,
        }
        if su and spw:
            st.markdown('<div class="status-box status-success">🟢 SMTP configuré</div>', unsafe_allow_html=True)


# ─────────────────────────────────────────────
# MAIN TABS
# ─────────────────────────────────────────────
tab_fiche, tab_envoi, tab_suivi, tab_histo = st.tabs([
    "🔍 Recherche & Fiche PDP",
    "📬 Envoyer Courrier / Email",
    "🚚 Suivi Merci Facteur",
    "📋 Historique",
])

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 1 — RECHERCHE & FICHE PDP
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_fiche:
    df = st.session_state.df_projets
    if df is None:
        st.info("👈 Importe ton export CSV Notion dans la sidebar pour commencer.")
    else:
        # ── Stats globales ──
        st.markdown("### 📊 Vue d'ensemble")
        stade_col = "Stade" if "Stade" in df.columns else None
        if stade_col:
            stades = df[stade_col].fillna("").str.strip()
            counts = {
                "Total": len(df),
                "Financés": (stades == "Financé").sum(),
                "Remboursés": (stades == "Remboursé").sum(),
                "Retard": stades.str.contains("Retard", case=False, na=False).sum(),
                "Précontentieux": stades.str.contains("Précontentieux", case=False, na=False).sum(),
                "Contentieux": stades.str.contains("Contentieux", case=False, na=False).sum(),
            }
        else:
            counts = {"Total": len(df)}

        stat_cols = st.columns(len(counts))
        colors = {"Total": "#1e293b", "Financés": "#28a745", "Remboursés": "#6c757d",
                  "Retard": "#f39c12", "Précontentieux": "#e67e22", "Contentieux": "#dc3545"}
        for col, (label, val) in zip(stat_cols, counts.items()):
            c = colors.get(label, "#1e293b")
            col.markdown(f'<div class="stat-card"><div class="value" style="color:{c}">{val}</div><div class="label">{label}</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        # ── Recherche ──
        st.markdown("### 🔎 Recherche PDP par email")
        col_s, col_b = st.columns([4, 1])
        with col_s:
            search_raw = st.text_input("Email du porteur de projet", placeholder="email@domaine.com", key="search_email", label_visibility="collapsed")
        with col_b:
            btn_search = st.button("🔍 Rechercher", use_container_width=True)

        search_norm = normalize_email(search_raw)

        if btn_search and search_norm:
            matches = df[df["_email_normalized"] == search_norm]
            if matches.empty:
                matches = df[df["_email_normalized"].str.contains(search_norm, case=False, na=False)]
            if matches.empty:
                st.markdown(f'<div class="status-box status-warn">⚠️ Aucun projet pour <b>{search_norm}</b></div>', unsafe_allow_html=True)
                st.session_state.selected_pdp_data = None
            else:
                st.session_state.selected_pdp_email = search_norm
                st.session_state.selected_pdp_data = matches

        # ── Fiche PDP ──
        pdp_data = st.session_state.selected_pdp_data
        if pdp_data is not None and not pdp_data.empty:
            first = pdp_data.iloc[0]
            nom_pdp = str(first.get("Nom du porteur", "")).strip()
            societe = str(first.get("Société (SIREN)", "")).strip()
            tel = str(first.get("Téléphone", "")).strip()
            email_pdp = st.session_state.selected_pdp_email

            st.markdown("---")
            st.markdown(f"### 👤 Fiche PDP — {nom_pdp}")

            # 3 colonnes : identité, résumé, alertes
            c1, c2, c3 = st.columns(3)
            with c1:
                html = '<div class="fiche-card"><h4>🪪 Identité</h4>'
                html += render_kv("Nom", nom_pdp)
                html += render_kv("Société (SIREN)", societe)
                html += render_kv("Téléphone", tel)
                html += render_kv("Email", email_pdp)
                notaire = str(first.get("Notaire du PDP", "")).strip()
                if notaire and notaire != "nan":
                    notaire_clean = re.sub(r"mailto:\S+", "", notaire).strip()
                    notaire_emails = extract_emails_from_field(notaire)
                    html += render_kv("Notaire", notaire_clean) if notaire_clean else ""
                    html += render_kv("Email notaire", ", ".join(notaire_emails)) if notaire_emails else ""
                html += '</div>'
                st.markdown(html, unsafe_allow_html=True)

            with c2:
                html = '<div class="fiche-card"><h4>📊 Projets</h4>'
                html += render_kv("Nombre total", str(len(pdp_data)))
                for stade_name, count in pdp_data["Stade"].fillna("").str.strip().value_counts().items():
                    if stade_name:
                        html += f'<div class="kv-row"><div class="kv-key">{render_badge(stade_name)}</div><div class="kv-val">{count}</div></div>'
                html += '</div>'
                st.markdown(html, unsafe_allow_html=True)

            with c3:
                has_cx = pdp_data["Stade"].fillna("").str.contains("Contentieux", case=False, na=False).any()
                has_pcx = pdp_data["Stade"].fillna("").str.contains("Précontentieux", case=False, na=False).any()
                has_ret = pdp_data["Stade"].fillna("").str.contains("Retard", case=False, na=False).any()
                html = '<div class="fiche-card"><h4>🚨 Alertes</h4>'
                if has_cx:
                    html += '<div class="status-box status-error">🔴 CONTENTIEUX en cours</div>'
                if has_pcx:
                    html += '<div class="status-box status-warn">🟠 PRÉCONTENTIEUX en cours</div>'
                if has_ret:
                    html += '<div class="status-box status-warn">🟡 Retard signalé</div>'
                if not (has_cx or has_pcx or has_ret):
                    html += '<div class="status-box status-success">✅ RAS</div>'
                html += '</div>'
                st.markdown(html, unsafe_allow_html=True)

            # ── Détail projets ──
            st.markdown("---")
            st.markdown(f"### 📁 Détail des {len(pdp_data)} projets")

            for idx, (_, row) in enumerate(pdp_data.iterrows()):
                stade = str(row.get("Stade", "")).strip()
                nom_p = str(row.get("Nom du projet", "")).strip()
                id_bo = str(row.get("ID Back-office", "")).strip()
                is_alert = "Contentieux" in stade or "Précontentieux" in stade

                with st.expander(
                    f"{STADE_ICONS.get(stade, '📌')} **{nom_p}** — ID {id_bo} — {stade}",
                    expanded=(idx == 0 or is_alert),
                ):
                    cp1, cp2 = st.columns(2)
                    with cp1:
                        st.markdown(render_section("📋 Projet", row, FICHE_COLS_PROJET), unsafe_allow_html=True)
                        st.markdown(render_section("📅 Dates & Durées", row, FICHE_COLS_DATES), unsafe_allow_html=True)
                    with cp2:
                        st.markdown(render_section("💰 Financier", row, FICHE_COLS_FINANCIER), unsafe_allow_html=True)
                        st.markdown(render_section("📌 Suivi & Garanties", row, FICHE_COLS_SUIVI), unsafe_allow_html=True)

                    # Actions
                    ca1, ca2, ca3 = st.columns(3)
                    with ca1:
                        if st.button(f"📮 Lettre", key=f"fl_{idx}"):
                            st.session_state["pf_nom"] = nom_pdp
                            st.session_state["pf_soc"] = societe.split(" - ")[0] if " - " in societe else societe
                            st.session_state["pf_adr"] = str(row.get("Adresse et lots", "")).strip()
                            st.session_state["pf_ref"] = f"Projet {id_bo} - {nom_p}"
                            st.session_state["pf_mode"] = "lettre"
                            st.info("→ Va dans l'onglet 'Envoyer'")
                    with ca2:
                        if st.button(f"📧 Email", key=f"fe_{idx}"):
                            st.session_state["pf_email"] = email_pdp
                            st.session_state["pf_ref"] = f"Projet {id_bo} - {nom_p}"
                            st.session_state["pf_mode"] = "email"
                            st.info("→ Va dans l'onglet 'Envoyer'")
                    with ca3:
                        st.download_button("📥 CSV", data=row.to_frame().T.to_csv(index=False),
                                           file_name=f"projet_{id_bo}.csv", mime="text/csv", key=f"fx_{idx}")

            st.markdown("---")
            st.download_button("📥 Exporter tous les projets du PDP",
                               data=pdp_data.to_csv(index=False),
                               file_name=f"pdp_{email_pdp.replace('@','_')}.csv",
                               mime="text/csv", key="exp_all_pdp")


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 2 — ENVOI
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_envoi:
    pf_mode = st.session_state.pop("pf_mode", None)
    mode = st.radio("**Mode**", ["📮 Lettre physique (Merci Facteur)", "📧 Email (SMTP)"],
                    index=1 if pf_mode == "email" else 0, horizontal=True, key="env_radio")

    # Prefill
    pf_nom = st.session_state.pop("pf_nom", "")
    pf_soc = st.session_state.pop("pf_soc", "")
    pf_adr = st.session_state.pop("pf_adr", "")
    pf_email = st.session_state.pop("pf_email", "")
    pf_ref = st.session_state.pop("pf_ref", "")

    st.markdown("---")
    st.markdown("### 🎯 Destinataire")
    d1, d2 = st.columns(2)
    with d1:
        dc = st.selectbox("Civilité", ["M", "Mme", ""], key="e_civ")
        dn = st.text_input("Nom *", value=pf_nom, key="e_nom")
        dp = st.text_input("Prénom", key="e_prenom")
        ds = st.text_input("Société", value=pf_soc, key="e_soc")
    with d2:
        da1 = st.text_input("Adresse 1 *", value=pf_adr, key="e_a1")
        da2 = st.text_input("Adresse 2", key="e_a2")
        dcp = st.text_input("Code postal *", key="e_cp")
        dv = st.text_input("Ville *", key="e_ville")
    d3, d4 = st.columns(2)
    with d3:
        dpays = st.text_input("Pays *", value="France", key="e_pays")
    with d4:
        demail = st.text_input("Email destinataire" + (" *" if "Email" in mode else ""),
                               value=pf_email, key="e_email")

    st.markdown("---")

    if "Lettre" in mode:
        st.markdown("### 🏢 Expéditeur")
        e1, e2 = st.columns(2)
        with e1:
            en = st.text_input("Nom exp. *", key="x_nom")
            epr = st.text_input("Prénom exp.", key="x_prenom")
            eso = st.text_input("Société exp.", value="La Première Brique", key="x_soc")
        with e2:
            ea = st.text_input("Adresse exp. *", key="x_adr")
            ecp = st.text_input("CP exp. *", key="x_cp")
            evi = st.text_input("Ville exp. *", key="x_ville")
            epays = st.text_input("Pays exp.", value="France", key="x_pays")

        st.markdown("---")
        st.markdown("### 📝 Contenu")
        o1, o2, o3 = st.columns(3)
        with o1:
            tc = st.selectbox("Type", list(TYPES_COURRIER.keys()), key="e_tc")
        with o2:
            me = st.selectbox("Mode postal", list(MODES_ENVOI.keys()), key="e_me")
        with o3:
            rv = st.checkbox("Recto-verso", key="e_rv")
            coul = st.checkbox("Couleur", value=True, key="e_coul")

        pdfs = st.file_uploader("📎 PDF(s) à envoyer *", type=["pdf"], accept_multiple_files=True, key="e_pdfs")

        r1, r2 = st.columns(2)
        with r1:
            date_e = st.date_input("Date programmée (opt.)", value=None, min_value=datetime.now().date(), key="e_date")
        with r2:
            ref_int = st.text_input("Réf. interne", value=pf_ref, key="e_ri")
            ref_uniq = st.text_input("Réf. anti-doublon (max 200)", key="e_ru", max_chars=200)

    else:
        st.markdown("### 📝 Contenu de l'email")
        subj = st.text_input("Objet *", value=f"Concernant : {pf_ref}" if pf_ref else "", key="e_subj")
        body = st.text_area("Corps (HTML) *", height=250,
                            placeholder="<p>Bonjour,</p>\n<p>...</p>\n<p>Cordialement,<br>LPB</p>", key="e_body")
        atts = st.file_uploader("📎 Pièces jointes (PDF, images, docs...)", accept_multiple_files=True, key="e_atts")

    st.markdown("---")

    cs, cr = st.columns([4, 1])
    with cs:
        if st.button("🚀 ENVOYER", type="primary", use_container_width=True, key="btn_go"):

            if "Lettre" in mode:
                errs = []
                if not dn and not ds:
                    errs.append("Nom ou société dest. requis")
                if not dcp:
                    errs.append("CP requis")
                if not dv:
                    errs.append("Ville requise")
                if not pdfs:
                    errs.append("Au moins 1 PDF requis")
                if not st.session_state.mf_token:
                    errs.append("Non connecté à MF")
                if not st.session_state.mf_user_id:
                    errs.append("User ID MF requis")
                if errs:
                    for e in errs:
                        st.error(f"❌ {e}")
                else:
                    with st.spinner("📮 Envoi..."):
                        fb64 = [{"fileName": f.name, "fileBase64": base64.b64encode(f.read()).decode()} for f in pdfs]
                        res = mf_send_letter(
                            user_id=st.session_state.mf_user_id,
                            destinataires=[{"civilite": dc, "nom": dn, "prenom": dp, "societe": ds,
                                            "adresse1": da1, "adresse2": da2, "adresse3": "",
                                            "cp": dcp, "ville": dv, "pays": dpays, "email": demail}],
                            expediteur={"nom": en, "prenom": epr, "societe": eso,
                                        "adresse1": ea, "adresse2": "", "adresse3": "",
                                        "cp": ecp, "ville": evi, "pays": epays},
                            fichiers_b64=fb64,
                            mode_envoi=MODES_ENVOI[me], type_courrier=TYPES_COURRIER[tc],
                            recto_verso=rv, couleur=coul,
                            date_envoi=str(date_e) if date_e else None,
                            ref_interne=ref_int or None, ref_unique=ref_uniq or None,
                        )
                        if "error" not in res and res.get("result"):
                            st.markdown('<div class="status-box status-success">✅ <b>Courrier envoyé !</b></div>', unsafe_allow_html=True)
                            st.json(res["result"])
                            st.session_state.send_history.append({
                                "type": "📮", "dest": f"{dn} {dp}".strip(), "mode": me,
                                "ref": ref_int, "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                "api_result": res.get("result"),
                            })
                        else:
                            st.markdown(f'<div class="status-box status-error">❌ {json.dumps(res, ensure_ascii=False)}</div>', unsafe_allow_html=True)

            else:
                errs = []
                if not demail:
                    errs.append("Email dest. requis")
                if not subj:
                    errs.append("Objet requis")
                if not body:
                    errs.append("Corps requis")
                smtp = st.session_state.smtp_config
                if not smtp.get("user") or not smtp.get("password"):
                    errs.append("SMTP non configuré")
                if errs:
                    for e in errs:
                        st.error(f"❌ {e}")
                else:
                    with st.spinner("📧 Envoi..."):
                        att_list = [(f.name, f.read()) for f in atts] if atts else None
                        res = send_email_smtp(
                            smtp["host"], smtp["port"], smtp["user"], smtp["password"],
                            smtp["from"], demail, subj, body, att_list, smtp.get("tls", True),
                        )
                        if res["success"]:
                            st.markdown('<div class="status-box status-success">✅ <b>Email envoyé !</b></div>', unsafe_allow_html=True)
                            st.session_state.send_history.append({
                                "type": "📧", "dest": demail, "sujet": subj,
                                "pj": len(att_list) if att_list else 0,
                                "date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                            })
                        else:
                            st.markdown(f'<div class="status-box status-error">❌ {res["message"]}</div>', unsafe_allow_html=True)

    with cr:
        if st.button("🗑️ Reset", use_container_width=True, key="btn_rst"):
            st.rerun()


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 3 — SUIVI MF
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_suivi:
    st.markdown("### 🚚 Suivi Merci Facteur")
    if not st.session_state.mf_token:
        st.info("Connecte-toi à Merci Facteur dans la sidebar.")
    else:
        s1, s2, s3 = st.tabs(["📋 Derniers envois", "🔎 Détail", "📊 Quotas"])
        with s1:
            if st.button("🔄 Charger", key="suiv_load"):
                st.json(mf_api("GET", "listEnvois", {"userId": st.session_state.mf_user_id}).get("result", {}))
        with s2:
            eid = st.text_input("ID envoi", key="suiv_eid")
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                if st.button("📄 Détail", key="suiv_det") and eid:
                    st.json(mf_api("GET", "getEnvoi", {"idEnvoi": eid}).get("result", {}))
            with sc2:
                if st.button("🚚 Suivi", key="suiv_trk") and eid:
                    st.json(mf_api("GET", "getSuiviEnvoi", {"idEnvoi": eid}).get("result", {}))
            with sc3:
                if st.button("📜 Preuves", key="suiv_prf") and eid:
                    st.json(mf_api("GET", "getProof", {"idEnvoi": eid}).get("result", {}))
        with s3:
            if st.button("📊 Quotas", key="suiv_qt"):
                st.json(mf_api("GET", "getQuotaCompte").get("result", {}))


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# TAB 4 — HISTORIQUE
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
with tab_histo:
    st.markdown("### 📋 Historique de la session")
    if not st.session_state.send_history:
        st.info("Aucun envoi dans cette session.")
    else:
        for i, e in enumerate(reversed(st.session_state.send_history)):
            with st.expander(f"{e['type']} → {e.get('dest','?')} — {e['date']}", expanded=(i == 0)):
                st.json(e)
        if st.button("🗑️ Vider", key="histo_clear"):
            st.session_state.send_history = []
            st.rerun()


# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("---")
st.markdown("<div style='text-align:center;opacity:0.4;font-size:0.8rem;'>"
            "📮 LPB Courrier & Fiche PDP — Merci Facteur API v1.2 + SMTP — Données locales uniquement</div>",
            unsafe_allow_html=True)

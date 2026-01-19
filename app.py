
import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
import smtplib
from email.message import EmailMessage
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Analyse patrimoniale", layout="wide")

APP_TITLE = "Analyse patrimoniale"
APP_SUBTITLE = "Remplis le formulaire. À la fin : export PDF + envoi automatique."
DEFAULT_BRAND = "CL Conseils"

TEMPLATE_PATH = Path(__file__).with_name("template.xlsx")

# -------------------- HELPERS --------------------
def load_template():
    xl = pd.ExcelFile(TEMPLATE_PATH)
    df = pd.read_excel(TEMPLATE_PATH, sheet_name=xl.sheet_names[0], header=None).fillna("")
    return xl, df

def get_cell(df, r, c, default=""):
    v = df.iat[r, c]
    if v == "":
        return default
    return v

def set_cell(df, r, c, value):
    df.iat[r, c] = value

def to_xlsx_bytes(xl, df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=xl.sheet_names[0], header=False, index=False)
    out.seek(0)
    return out.read()

def build_pdf_bytes(data: dict, brand: str):
    styles = getSampleStyleSheet()
    story = []
    story.append(Paragraph(f"<b>{brand}</b>", styles["Title"]))
    story.append(Paragraph(f"{APP_TITLE} — Synthèse client", styles["Heading2"]))
    story.append(Spacer(1, 12))

    def section(title, rows):
        story.append(Paragraph(f"<b>{title}</b>", styles["Heading3"]))
        t = Table(rows, colWidths=[220, 280])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.Color(0.98,0.98,0.98)]),
        ]))
        story.append(t)
        story.append(Spacer(1, 10))

    # Sections
    section("Identité", [
        ["Champ", "Valeur"],
        ["Nom", data.get("nom","")],
        ["Prénom", data.get("prenom","")],
        ["Date de naissance", data.get("date_naissance","")],
        ["Lieu de naissance", data.get("lieu_naissance","")],
        ["Adresse", data.get("adresse","")],
        ["Téléphone", data.get("tel","")],
        ["Email", data.get("mail","")],
        ["Situation familiale", data.get("situation_fam","")],
        ["Nombre de parts", str(data.get("parts",""))],
    ])

    section("Budget mensuel", [
        ["Poste", "Montant (€)"],
        ["Salaires mensuels (avant PAS)", str(data.get("salaire",0))],
        ["Déclarant 1 — salaire", str(data.get("dec1",0))],
        ["Déclarant 2 — salaire", str(data.get("dec2",0))],
        ["Revenus locatifs existants (retenus 80%)", str(data.get("rev_loc",0))],
        ["Bien n°1 — loyer retenu", str(data.get("bien1",0))],
        ["Bien n°2 — loyer retenu", str(data.get("bien2",0))],
        ["Emprunt RP / loyer", str(data.get("emprunt_rp",0))],
        ["Charges fixes / abonnements", str(data.get("charges",0))],
        ["Essence", str(data.get("essence",0))],
        ["Impôts", str(data.get("impots",0))],
        ["Crédits — souscripteur ?", "Oui" if data.get("credit_flag",0)==1 else "Non"],
    ])

    exp = data.get("experience", {})
    section("Profil financier — Connaissances / expérience", [["Support", "O/N"]] + [[k, "Oui" if v else "Non"] for k,v in exp.items()])

    # Render to bytes
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, title="Synthèse Analyse patrimoniale")
    doc.build(story)
    buffer.seek(0)
    return buffer.read()

def send_email_smtp(
    smtp_host, smtp_port, smtp_user, smtp_password,
    sender, recipient, subject, body,
    pdf_bytes, pdf_name,
    xlsx_bytes, xlsx_name,
):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)

    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    msg.add_attachment(xlsx_bytes, maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=xlsx_name)

    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# -------------------- UI --------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    .card {border:1px solid #e9e9e9; border-radius:16px; padding:18px; background:#fff;}
    .muted {color:#6b7280;}
    .pill {display:inline-block; padding:6px 10px; border-radius:999px; background:#f3f4f6; font-size:12px; margin-right:6px;}
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(f"## {APP_TITLE}")
st.markdown(f"<span class='muted'>{APP_SUBTITLE}</span>", unsafe_allow_html=True)
st.write("")

with st.sidebar:
    st.markdown("### Paramètres (admin)")
    brand = st.text_input("Nom de marque / Cabinet", value=DEFAULT_BRAND)
    recipient_email = st.text_input("Email de réception (où l'appli envoie les docs)", value="")
    st.markdown("---")
    st.markdown("### Envoi email (SMTP)")
    st.caption("Pour une appli publique, le plus simple est d'utiliser une boîte dédiée type contact@...")
    smtp_host = st.text_input("SMTP host", value="smtp.gmail.com")
    smtp_port = st.number_input("SMTP port", value=587, step=1)
    smtp_user = st.text_input("SMTP user (login)", value="", help="Souvent l'adresse email complète")
    smtp_password = st.text_input("SMTP password / app password", value="", type="password")
    sender_email = st.text_input("Email expéditeur (From)", value=smtp_user)

xl, df = load_template()

# --- Stepper ---
steps = ["Identité", "Budget", "Profil financier", "Validation & export"]
if "step" not in st.session_state:
    st.session_state.step = 0

st.markdown(" ".join([f"<span class='pill'>{'✅' if i < st.session_state.step else ('👉' if i==st.session_state.step else '•')} {s}</span>" for i,s in enumerate(steps)]), unsafe_allow_html=True)
st.write("")

def next_step():
    st.session_state.step = min(st.session_state.step + 1, len(steps)-1)

def prev_step():
    st.session_state.step = max(st.session_state.step - 1, 0)

# -------------------- STEP 1 --------------------
if st.session_state.step == 0:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("Identité")
    c1, c2 = st.columns(2)
    with c1:
        nom = st.text_input("Nom", value=str(get_cell(df, 1, 9, "")))
        prenom = st.text_input("Prénom", value=str(get_cell(df, 2, 9, "")))
        date_naissance = st.date_input("Date de naissance", value=pd.to_datetime(get_cell(df, 3, 9, ""), errors="coerce").date() if str(get_cell(df,3,9,"")) else None)
        lieu_naissance = st.text_input("Lieu de naissance", value=str(get_cell(df, 4, 9, "")))
        situation_fam = st.text_input("Situation familiale", value=str(get_cell(df, 5, 11, "")))
    with c2:
        adresse = st.text_input("Adresse", value=str(get_cell(df, 5, 9, "")))
        tel = st.text_input("Téléphone", value=str(get_cell(df, 7, 9, "")))
        mail = st.text_input("Email", value=str(get_cell(df, 8, 9, "")))
        parts = st.number_input("Nombre de parts", value=float(get_cell(df, 4, 11, 0) or 0), step=1.0)

    set_cell(df, 1, 9, nom)
    set_cell(df, 2, 9, prenom)
    set_cell(df, 3, 9, str(date_naissance) if date_naissance else "")
    set_cell(df, 4, 9, lieu_naissance)
    set_cell(df, 5, 11, situation_fam)
    set_cell(df, 5, 9, adresse)
    set_cell(df, 7, 9, tel)
    set_cell(df, 8, 9, mail)
    set_cell(df, 4, 11, parts)

    st.write("")
    cprev, cnext = st.columns([1,1])
    with cnext:
        st.button("Continuer ➜", on_click=next_step, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- STEP 2 --------------------
elif st.session_state.step == 1:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("Budget mensuel")
    c1, c2 = st.columns(2)
    with c1:
        salaire = st.number_input("Salaires mensuels (avant PAS)", value=float(get_cell(df, 2, 2, 0) or 0))
        dec1 = st.number_input("Déclarant 1 — salaire", value=float(get_cell(df, 3, 2, 0) or 0))
        dec2 = st.number_input("Déclarant 2 — salaire", value=float(get_cell(df, 4, 2, 0) or 0))
        rev_loc = st.number_input("Revenus locatifs existants (retenus 80%)", value=float(get_cell(df, 5, 2, 0) or 0))
        bien1 = st.number_input("Bien n°1 — loyer retenu", value=float(get_cell(df, 7, 2, 0) or 0))
        bien2 = st.number_input("Bien n°2 — loyer retenu", value=float(get_cell(df, 8, 2, 0) or 0))
    with c2:
        credit_flag = st.selectbox("Crédits — souscripteur ?", options=[0,1], index=0 if int(get_cell(df, 12, 2, 0) or 0)==0 else 1,
                                   format_func=lambda x: "Non" if x==0 else "Oui")
        emprunt_rp = st.number_input("Emprunt RP / loyer", value=float(get_cell(df, 13, 2, 0) or 0))
        charges = st.number_input("Charges fixes / abonnements", value=float(get_cell(df, 16, 2, 0) or 0))
        essence = st.number_input("Essence", value=float(get_cell(df, 17, 2, 0) or 0))
        impots = st.number_input("Impôts", value=float(get_cell(df, 18, 2, 0) or 0))

    set_cell(df, 2, 2, salaire)
    set_cell(df, 3, 2, dec1)
    set_cell(df, 4, 2, dec2)
    set_cell(df, 5, 2, rev_loc)
    set_cell(df, 7, 2, bien1)
    set_cell(df, 8, 2, bien2)
    set_cell(df, 12, 2, credit_flag)
    set_cell(df, 13, 2, emprunt_rp)
    set_cell(df, 16, 2, charges)
    set_cell(df, 17, 2, essence)
    set_cell(df, 18, 2, impots)

    st.write("")
    cprev, cnext = st.columns([1,1])
    with cprev:
        st.button("⟵ Retour", on_click=prev_step, use_container_width=True)
    with cnext:
        st.button("Continuer ➜", on_click=next_step, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- STEP 3 --------------------
elif st.session_state.step == 2:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("Profil financier")
    st.caption("Coche ce que tu connais déjà / as déjà expérimenté.")
    categories = [
        "SICAV ou FCP","Obligations","Actions","Trackers - fonds alternatifs",
        "Warrants - Futurs - Options","Gestion directe ou personnelle",
        "Gestion déléguée ou sous mandat","Fonds euros","Supports unités de compte",
        "SCPI","OPCI","FCPI","FIP"
    ]
    # display in 2 columns
    c1, c2 = st.columns(2)
    exp = {}
    for i,cat in enumerate(categories):
        col = c1 if i % 2 == 0 else c2
        with col:
            exp[cat] = st.checkbox(cat, value=False, key=f"exp_{i}")

    # store a simple "O" in col 13, row 2..
    for idx,cat in enumerate(categories, start=2):
        set_cell(df, idx, 13, "O" if exp[cat] else "")

    st.write("")
    cprev, cnext = st.columns([1,1])
    with cprev:
        st.button("⟵ Retour", on_click=prev_step, use_container_width=True)
    with cnext:
        st.button("Continuer ➜", on_click=next_step, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# -------------------- STEP 4 --------------------
else:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("Validation & export")
    st.caption("Tu peux télécharger le PDF + l'Excel, et l'appli peut aussi envoyer automatiquement les deux au cabinet.")

    # Collect data from current session (re-read from df + widgets keys)
    data = {
        "nom": str(get_cell(df, 1, 9, "")),
        "prenom": str(get_cell(df, 2, 9, "")),
        "date_naissance": str(get_cell(df, 3, 9, "")),
        "lieu_naissance": str(get_cell(df, 4, 9, "")),
        "adresse": str(get_cell(df, 5, 9, "")),
        "tel": str(get_cell(df, 7, 9, "")),
        "mail": str(get_cell(df, 8, 9, "")),
        "situation_fam": str(get_cell(df, 5, 11, "")),
        "parts": get_cell(df, 4, 11, ""),
        "salaire": float(get_cell(df, 2, 2, 0) or 0),
        "dec1": float(get_cell(df, 3, 2, 0) or 0),
        "dec2": float(get_cell(df, 4, 2, 0) or 0),
        "rev_loc": float(get_cell(df, 5, 2, 0) or 0),
        "bien1": float(get_cell(df, 7, 2, 0) or 0),
        "bien2": float(get_cell(df, 8, 2, 0) or 0),
        "credit_flag": int(get_cell(df, 12, 2, 0) or 0),
        "emprunt_rp": float(get_cell(df, 13, 2, 0) or 0),
        "charges": float(get_cell(df, 16, 2, 0) or 0),
        "essence": float(get_cell(df, 17, 2, 0) or 0),
        "impots": float(get_cell(df, 18, 2, 0) or 0),
        "experience": {k: st.session_state.get(f"exp_{i}", False) for i,k in enumerate([
            "SICAV ou FCP","Obligations","Actions","Trackers - fonds alternatifs",
            "Warrants - Futurs - Options","Gestion directe ou personnelle",
            "Gestion déléguée ou sous mandat","Fonds euros","Supports unités de compte",
            "SCPI","OPCI","FCPI","FIP"
        ])},
    }

    xlsx_bytes = to_xlsx_bytes(xl, df)
    pdf_bytes = build_pdf_bytes(data, brand=brand)

    c1, c2 = st.columns(2)
    with c1:
        st.download_button("⬇️ Télécharger le PDF", data=pdf_bytes, file_name="analyse_patrimoniale.pdf", mime="application/pdf", use_container_width=True)
        st.download_button("⬇️ Télécharger l'Excel", data=xlsx_bytes, file_name="analyse_patrimoniale.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

    with c2:
        st.markdown("#### Envoi automatique au cabinet")
        if not recipient_email:
            st.warning("Renseigne l'email de réception dans la sidebar (admin).")
        else:
            subject = st.text_input("Objet de l'email", value="Analyse patrimoniale — documents client")
            body = st.text_area("Message", value="Bonjour,\n\nVeuillez trouver en pièce jointe le PDF + l'Excel remplis.\n\nCordialement,")
            ready = all([smtp_host, smtp_port, smtp_user, smtp_password, sender_email, recipient_email])

            if st.button("📩 Envoyer maintenant", use_container_width=True, disabled=not ready):
                try:
                    send_email_smtp(
                        smtp_host=smtp_host,
                        smtp_port=int(smtp_port),
                        smtp_user=smtp_user,
                        smtp_password=smtp_password,
                        sender=sender_email,
                        recipient=recipient_email,
                        subject=subject,
                        body=body,
                        pdf_bytes=pdf_bytes,
                        pdf_name="analyse_patrimoniale.pdf",
                        xlsx_bytes=xlsx_bytes,
                        xlsx_name="analyse_patrimoniale.xlsx",
                    )
                    st.success("Envoyé ✅")
                except Exception as e:
                    st.error(f"Erreur d'envoi : {e}")

    st.write("")
    cprev, _ = st.columns([1,1])
    with cprev:
        st.button("⟵ Retour", on_click=prev_step, use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

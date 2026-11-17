
import streamlit as st
import pandas as pd
import smtplib
from datetime import date
from email.message import EmailMessage
from patrimoine import (
    APP_TITLE, BUDGET_FIELDS, CATEGORIES, IDENTITY_FIELDS,
    build_pdf_bytes, collect_data, get_cell, load_template, set_cell,
    sync_budget_totals, to_xlsx_bytes,
)

# -------------------- CONFIG --------------------
st.set_page_config(page_title="Analyse patrimoniale", layout="wide")

APP_SUBTITLE = "Remplis le formulaire. À la fin : export PDF + envoi automatique."
DEFAULT_BRAND = "CL Conseils"

# -------------------- HELPERS --------------------
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


def eur(x: float) -> str:
    """Format number as French-style euros: 12 345 € (no weird commas)."""
    try:
        x = float(x)
    except Exception:
        return str(x)
    s = f"{x:,.0f}".replace(",", " ").replace(" ", " ")
    return f"{s} €"

# -------------------- UI --------------------
st.markdown(
    """
    <style>
:root{
  --bg:#f7f8fb;
  --card:#ffffff;
  --border:#e7eaf0;
  --text:#0f172a;
  --muted:#64748b;
  --primary:#2563eb;
  --primary2:#1d4ed8;
  --success:#16a34a;
  --warn:#f59e0b;
}
.block-container {padding-top: 1.25rem; padding-bottom: 2rem; background: var(--bg);}
h1,h2,h3 {color: var(--text);}
.card {border:1px solid var(--border); border-radius:18px; padding:18px; background:var(--card); box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);}
.muted {color:var(--muted);}
.pill {display:inline-block; padding:6px 10px; border-radius:999px; background:#eef2ff; color:#1e3a8a; font-size:12px; margin-right:6px; border:1px solid #e0e7ff;}
.kpiwrap{display:flex; gap:12px; flex-wrap:wrap; margin-top:6px;}
.kpi{flex:1; min-width:210px; border:1px solid var(--border); border-radius:16px; padding:14px; background: #ffffff;}
.kpi .label{font-size:12px; color:var(--muted); margin-bottom:6px;}
.kpi .value{font-size:22px; font-weight:800; color:var(--text);}
.kpi.good .value{color: var(--success);}
.kpi.warn .value{color: var(--warn);}
.notice{border-left:6px solid var(--primary); background:#eff6ff; padding:12px 14px; border-radius:14px; color:#0b1b3a; border:1px solid #dbeafe;}
.footerhint{font-size:12px;color:var(--muted);}
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

# The DataFrame belongs to this browser session, never a shared mutable cache.
if "client_df" not in st.session_state:
    st.session_state.client_df = load_template()
df = st.session_state.client_df

# --- Stepper ---
steps = ["Identité", "Budget", "Profil financier", "Validation & export"]
if "step" not in st.session_state:
    st.session_state.step = 0

st.markdown(" ".join([f"<span class='pill'>{'✅' if i < st.session_state.step else ('👉' if i==st.session_state.step else '•')} {s}</span>" for i,s in enumerate(steps)]), unsafe_allow_html=True)
st.write("")

def persist_visible_fields():
    # Navigation callbacks execute before the next render. Save current widget
    # state here too, including edits delivered with the navigation click.
    fields = IDENTITY_FIELDS if st.session_state.step == 0 else BUDGET_FIELDS
    if st.session_state.step in (0, 1):
        for key, (r, c, _) in fields.items():
            if key in ("salaire", "rev_loc"):
                continue
            widget_key = f"client_{key}"
            if widget_key in st.session_state:
                value = st.session_state[widget_key]
                if key == "date_naissance":
                    value = value.isoformat() if value else ""
                set_cell(df, r, c, value)
        if st.session_state.step == 1:
            sync_budget_totals(df)
    elif st.session_state.step == 2:
        for i, cat in enumerate(CATEGORIES):
            set_cell(df, i + 2, 13, "O" if st.session_state.get(f"exp_{i}", False) else "")


def next_step():
    persist_visible_fields()
    st.session_state.step = min(st.session_state.step + 1, len(steps)-1)

def prev_step():
    persist_visible_fields()
    st.session_state.step = max(st.session_state.step - 1, 0)

# -------------------- STEP 1 --------------------
if st.session_state.step == 0:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("Identité")
    c1, c2 = st.columns(2)
    with c1:
        nom = st.text_input("Nom", value=str(get_cell(df, 1, 9, "")), key="client_nom")
        prenom = st.text_input("Prénom", value=str(get_cell(df, 2, 9, "")), key="client_prenom")
        date_naissance = st.date_input("Date de naissance", value=pd.to_datetime(get_cell(df, 3, 9, ""), errors="coerce").date() if str(get_cell(df,3,9,"")) else None, min_value=date(1900, 1, 1), max_value=date.today(), key="client_date_naissance")
        lieu_naissance = st.text_input("Lieu de naissance", value=str(get_cell(df, 4, 9, "")), key="client_lieu_naissance")
        situation_fam = st.text_input("Situation familiale", value=str(get_cell(df, 5, 11, "")), key="client_situation_fam")
    with c2:
        adresse = st.text_input("Adresse", value=str(get_cell(df, 5, 9, "")), key="client_adresse")
        tel = st.text_input("Téléphone", value=str(get_cell(df, 7, 9, "")), key="client_tel")
        mail = st.text_input("Email", value=str(get_cell(df, 8, 9, "")), key="client_mail")
        parts = st.number_input("Nombre de parts", value=float(get_cell(df, 4, 11, 0) or 0), step=1.0, key="client_parts")

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
        dec1 = st.number_input("Déclarant 1 — salaire", value=float(get_cell(df, 3, 2, 0) or 0), key="client_dec1")
        dec2 = st.number_input("Déclarant 2 — salaire", value=float(get_cell(df, 4, 2, 0) or 0), key="client_dec2")
        bien1 = st.number_input("Bien n°1 — loyer hors charges avant décote", value=float(get_cell(df, 7, 2, 0) or 0), key="client_bien1")
        bien2 = st.number_input("Bien n°2 — loyer hors charges avant décote", value=float(get_cell(df, 8, 2, 0) or 0), key="client_bien2")
    with c2:
        credit_flag = st.selectbox("Crédits — souscripteur ?", options=[0,1], index=0 if int(get_cell(df, 12, 2, 0) or 0)==0 else 1,
                                   format_func=lambda x: "Non" if x==0 else "Oui", key="client_credit_flag")
        emprunt_rp = st.number_input("Emprunt RP / loyer", value=float(get_cell(df, 13, 2, 0) or 0), key="client_emprunt_rp")
        charges = st.number_input("Charges fixes / abonnements (hors essence et impôts)", value=float(get_cell(df, 16, 2, 0) or 0), key="client_charges")
        essence = st.number_input("Essence", value=float(get_cell(df, 17, 2, 0) or 0), key="client_essence")
        impots = st.number_input("Impôts", value=float(get_cell(df, 18, 2, 0) or 0), key="client_impots")
        autres_credits = st.number_input("Autres crédits (voiture, conso, etc.)", value=float(get_cell(df, 19, 2, 0) or 0), key="client_autres_credits")
        loisirs_budget = st.number_input("Budget plaisirs / voyages (à réserver chaque mois)", value=float(get_cell(df, 20, 2, 0) or 0), key="client_loisirs_budget")
    
    persist_visible_fields()
    budget = sync_budget_totals(df)
    revenus_retenus = budget["revenus_retenus"]
    charges_et_credits = budget["charges_et_credits"]
    reste_a_vivre = budget["reste_a_vivre"]
    reste_apres_loisirs = budget["reste_apres_loisirs"]
    capacite_endettement_max = budget["capacite_endettement_max"]
    capacite_mensualite_restante = budget["capacite_mensualite_restante"]
    st.caption(f"Total des salaires : {eur(budget['salaire'])} · Loyers retenus à 80 % : {eur(budget['rev_loc'])}")
    st.caption("Ces deux totaux sont calculés à partir des déclarants et des biens, sans double comptage.")

    st.markdown("### Résultats (en direct)")

    # KPI cards
    st.markdown(
        f"""
        <div class="notice"><b>Lecture simple :</b> on prend les salaires à 100% + les loyers à 80% (une seule décote),
        puis on retire charges et crédits existants. La capacité d'emprunt est calculée avec un taux d'endettement de 35%.</div>
        <div class="kpiwrap">
          <div class="kpi">
            <div class="label">Revenus retenus</div>
            <div class="value">{eur(revenus_retenus)}</div>
            <div class="footerhint">Salaires 100% + Loyers 80%</div>
          </div>
          <div class="kpi">
            <div class="label">Charges + crédits existants</div>
            <div class="value">{eur(charges_et_credits)}</div>
            <div class="footerhint">Charges fixes + mensualités en cours</div>
          </div>
          <div class="kpi {'good' if reste_a_vivre>=0 else 'warn'}">
            <div class="label">Reste à vivre</div>
            <div class="value">{eur(reste_a_vivre)}</div>
            <div class="footerhint">Avant budget plaisirs</div>
          </div>
          <div class="kpi {'good' if reste_apres_loisirs>=0 else 'warn'}">
            <div class="label">Épargne possible</div>
            <div class="value">{eur(reste_apres_loisirs)}</div>
            <div class="footerhint">Reste à vivre - plaisirs/voyages</div>
          </div>
          <div class="kpi">
            <div class="label">Capacité d'endettement max (35%)</div>
            <div class="value">{eur(capacite_endettement_max)}</div>
            <div class="footerhint">Mensualités totales max</div>
          </div>
          <div class="kpi good">
            <div class="label">Capacité de mensualité restante</div>
            <div class="value">{eur(capacite_mensualite_restante)}</div>
            <div class="footerhint">Ce qu'il reste pour un nouveau crédit</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )


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
    categories = CATEGORIES
    # display in 2 columns
    c1, c2 = st.columns(2)
    exp = {}
    for i,cat in enumerate(categories):
        col = c1 if i % 2 == 0 else c2
        with col:
            exp[cat] = st.checkbox(cat, value=get_cell(df, i + 2, 13) == "O", key=f"exp_{i}")

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

    # Exports read the same persistent state, including off-screen profile fields.
    sync_budget_totals(df)
    data = collect_data(df)
    xlsx_bytes = to_xlsx_bytes(data, brand=brand)
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

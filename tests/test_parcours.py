"""Exercise real Streamlit reruns and parse the actual download payloads."""

from datetime import date
from io import BytesIO
from pathlib import Path
import re
import smtplib

import openpyxl
from pypdf import PdfReader
import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest
import patrimoine

from patrimoine import calculate_budget, collect_data, load_template

APP_PATH = Path(__file__).resolve().parents[1] / "app.py"


IDENTITY = {
    "nom": "DUPONT-FICTIF", "prenom": "Élodie", "date_naissance": "1984-07-23",
    "lieu_naissance": "Béziers", "adresse": "18 rue des Tests, 34500 Béziers",
    "tel": "0612345678", "mail": "elodie.fictive@example.invalid",
    "situation_fam": "Pacsée", "parts": 2.5,
}
AMOUNTS = {
    "dec1": 3200.0, "dec2": 2100.0, "bien1": 750.0, "bien2": 550.0,
    "emprunt_rp": 1100.0, "charges": 420.0, "essence": 180.0,
    "impots": 350.0, "autres_credits": 240.0, "loisirs_budget": 300.0,
}
# Independent expectations, including every field and every Yes/No answer.
EXPECTED_ROWS = {
    "Cabinet": "CL Conseils", "Nom": "DUPONT-FICTIF", "Prénom": "Élodie",
    "Date de naissance": "1984-07-23", "Lieu de naissance": "Béziers",
    "Adresse": "18 rue des Tests, 34500 Béziers", "Téléphone": "0612345678",
    "Email": "elodie.fictive@example.invalid", "Situation familiale": "Pacsée",
    "Nombre de parts": 2.5, "Total des salaires mensuels (avant PAS)": 5300.0,
    "Déclarant 1 - salaire": 3200.0, "Déclarant 2 - salaire": 2100.0,
    "Total des loyers retenus à 80 %": 1040.0,
    "Bien n°1 - loyer hors charges avant décote": 750.0,
    "Bien n°2 - loyer hors charges avant décote": 550.0,
    "Crédits - souscripteur ?": "Oui", "Emprunt RP / loyer": 1100.0,
    "Charges fixes / abonnements (hors essence et impôts)": 420.0,
    "Essence": 180.0, "Impôts": 350.0,
    "Autres crédits (voiture, conso, etc.)": 240.0, "Budget plaisirs / voyages": 300.0,
    "Revenus retenus": 6340.0, "Charges + crédits existants": 2290.0,
    "Reste à vivre": 4050.0, "Épargne possible": 3750.0,
    "Capacité d'endettement max (35 %)": 2219.0,
    "Capacité de mensualité restante": 879.0,
    "SICAV ou FCP": "Oui", "Obligations": "Non", "Actions": "Oui",
    "Trackers - fonds alternatifs": "Non", "Warrants - Futurs - Options": "Oui",
    "Gestion directe ou personnelle": "Non", "Gestion déléguée ou sous mandat": "Oui",
    "Fonds euros": "Non", "Supports unités de compte": "Oui", "SCPI": "Non",
    "OPCI": "Oui", "FCPI": "Non", "FIP": "Oui",
}


def click(at, label):
    next(button for button in at.button if button.label == label).click().run()
    assert not at.exception
    return at


@pytest.fixture
def downloads(monkeypatch):
    payloads = {}
    original = st.download_button

    def capture(label, *, data, file_name, **kwargs):
        payloads[file_name] = data
        return original(label, data=data, file_name=file_name, **kwargs)

    monkeypatch.setattr(st, "download_button", capture)
    return payloads


def fill_client(at):
    for key, value in IDENTITY.items():
        if key == "date_naissance":
            at.date_input(key="client_date_naissance").set_value(date.fromisoformat(value))
        elif key == "parts":
            at.number_input(key="client_parts").set_value(value)
        else:
            at.text_input(key=f"client_{key}").set_value(value)
    # Deliver unsaved widget edits in the same event as navigation.
    click(at, "Continuer ➜")
    for key, value in AMOUNTS.items():
        at.number_input(key=f"client_{key}").set_value(value)
    at.selectbox(key="client_credit_flag").select(1)
    at.run()
    html = " ".join(markdown.value for markdown in at.markdown)
    for value in ["6 340 €", "2 290 €", "4 050 €", "3 750 €", "2 219 €", "879 €"]:
        assert value in html
    click(at, "Continuer ➜")
    for i in range(13):
        at.checkbox(key=f"exp_{i}").set_value(i % 2 == 0)
    click(at, "Continuer ➜")
    return at


def workbook_rows(payload):
    book = openpyxl.load_workbook(BytesIO(payload), data_only=False)
    assert book.sheetnames == ["Synthèse client"]
    assert not any(cell.data_type == "f" for row in book.active for cell in row)
    return {row[0].value: row[1].value for row in book.active if row[0].value is not None}


def pdf_text(payload):
    return " ".join(" ".join(page.extract_text().split()) for page in PdfReader(BytesIO(payload)).pages)


def assert_documents(downloads, expected):
    rows = workbook_rows(downloads["analyse_patrimoniale.xlsx"])
    text = pdf_text(downloads["analyse_patrimoniale.pdf"])
    for label, value in expected.items():
        assert rows[label] == value, label
        if label == "Cabinet":
            assert str(value) in text
        elif isinstance(value, float):
            # PDF numeric serialization may include .0 and float rounding noise.
            match = re.search(re.escape(label) + r"\s+(-?[\d.]+)", text)
            assert match, label
            assert float(match.group(1)) == pytest.approx(value), label
        else:
            assert f"{label} {value}" in text, label
    assert "LEMARIE" not in text
    assert "LEMARIE" not in str(rows)


def test_complete_journey_round_trips_edits_and_actual_exports(downloads):
    at = fill_client(AppTest.from_file(APP_PATH).run())
    assert_documents(downloads, EXPECTED_ROWS)
    initial = downloads.copy()
    # A sidebar edit triggers a rerun with every client widget off-screen.
    at.sidebar.text_input[0].set_value("CL Conseils & Patrimoine").run()
    assert_documents(downloads, {**EXPECTED_ROWS, "Cabinet": "CL Conseils & Patrimoine"})
    click(at, "⟵ Retour")
    assert all(at.checkbox(key=f"exp_{i}").value == (i % 2 == 0) for i in range(13))
    at.checkbox(key="exp_0").uncheck()
    click(at, "⟵ Retour")
    for key, value in AMOUNTS.items():
        assert at.number_input(key=f"client_{key}").value == value
    assert at.selectbox(key="client_credit_flag").value == 1
    at.number_input(key="client_dec1").set_value(4000.0)
    click(at, "⟵ Retour")
    for key, value in IDENTITY.items():
        if key not in ("parts", "date_naissance"):
            assert at.text_input(key=f"client_{key}").value == value
    assert at.date_input(key="client_date_naissance").value == date(1984, 7, 23)
    assert at.number_input(key="client_parts").value == 2.5
    at.text_input(key="client_nom").set_value("MARTIN-FICTIF")
    click(at, "Continuer ➜")
    click(at, "Continuer ➜")
    click(at, "Continuer ➜")
    expected = {**EXPECTED_ROWS, "Nom": "MARTIN-FICTIF", "Cabinet": "CL Conseils & Patrimoine",
                "Déclarant 1 - salaire": 4000.0, "Total des salaires mensuels (avant PAS)": 6100.0,
                "Revenus retenus": 7140.0, "Reste à vivre": 4850.0, "Épargne possible": 4550.0,
                "Capacité d'endettement max (35 %)": 2499.0,
                "Capacité de mensualité restante": 1159.0, "SICAV ou FCP": "Non"}
    assert_documents(downloads, expected)
    assert downloads != initial


@pytest.mark.parametrize("key,value,expected_revenue,expected_reste,expected_saving,expected_capacity", [
    ("dec1", 3000., 3000., 3000., 3000., 1050.),
    ("dec2", 2000., 2000., 2000., 2000., 700.),
    ("bien1", 1000., 800., 800., 800., 280.),
    ("bien2", 500., 400., 400., 400., 140.),
    ("charges", 120., 0., -120., -120., 0.),
    ("essence", 80., 0., -80., -80., 0.),
    ("impots", 200., 0., -200., -200., 0.),
    ("emprunt_rp", 900., 0., -900., -900., 0.),
    ("autres_credits", 150., 0., -150., -150., 0.),
    ("loisirs_budget", 250., 0., 0., -250., 0.),
])
def test_each_budget_input_updates_in_same_rerun(key, value, expected_revenue, expected_reste, expected_saving, expected_capacity):
    at = AppTest.from_file(APP_PATH).run()
    click(at, "Continuer ➜")
    at.number_input(key=f"client_{key}").set_value(value).run()
    assert not at.exception
    budget = calculate_budget(collect_data(at.session_state["client_df"]))
    assert budget["revenus_retenus"] == expected_revenue
    assert budget["reste_a_vivre"] == expected_reste
    assert budget["reste_apres_loisirs"] == expected_saving
    assert budget["capacite_mensualite_restante"] == expected_capacity
    html = " ".join(markdown.value for markdown in at.markdown)
    assert f'{expected_saving:,.0f} €'.replace(',', ' ') in html


def test_zero_fractional_and_over_indebted_budgets():
    data = collect_data(load_template())
    assert set(calculate_budget(data).values()) == {0.0}
    data.update(AMOUNTS, dec1=3200.25, bien1=750.50)
    assert calculate_budget(data)["revenus_retenus"] == pytest.approx(6340.65)
    data["autres_credits"] = 10000.0
    budget = calculate_budget(data)
    assert budget["reste_a_vivre"] < 0
    assert budget["capacite_mensualite_restante"] == 0


def test_template_is_loaded_once_per_session(monkeypatch, downloads):
    calls = []
    original = patrimoine.load_template

    def counted_load():
        calls.append(True)
        return original()

    monkeypatch.setattr(patrimoine, "load_template", counted_load)
    at = fill_client(AppTest.from_file(APP_PATH).run())
    click(at, "⟵ Retour")
    click(at, "⟵ Retour")
    click(at, "⟵ Retour")
    assert len(calls) == 1
    AppTest.from_file(APP_PATH).run()
    assert len(calls) == 2


def test_sessions_are_isolated_and_empty_values_stay_empty():
    first = AppTest.from_file(APP_PATH).run()
    first.text_input(key="client_nom").set_value("CLIENT UN").run()
    second = AppTest.from_file(APP_PATH).run()
    assert second.text_input(key="client_nom").value == ""
    assert first.session_state["client_df"] is not second.session_state["client_df"]
    first.text_input(key="client_nom").set_value("")
    first.date_input(key="client_date_naissance").set_value(None)
    click(first, "Continuer ➜")
    click(first, "⟵ Retour")
    assert first.text_input(key="client_nom").value == ""
    assert first.date_input(key="client_date_naissance").value is None


def test_text_is_literal_in_excel_and_pdf(downloads):
    at = AppTest.from_file(APP_PATH).run()
    at.text_input(key="client_nom").set_value('=1+2')
    at.text_input(key="client_prenom").set_value('Élodie <b> & Martin')
    click(at, "Continuer ➜")
    click(at, "Continuer ➜")
    click(at, "Continuer ➜")
    rows = workbook_rows(downloads["analyse_patrimoniale.xlsx"])
    text = pdf_text(downloads["analyse_patrimoniale.pdf"])
    assert rows["Nom"] == '=1+2'
    assert rows["Prénom"] == 'Élodie <b> & Martin'
    assert 'Élodie <b> & Martin' in text


def test_email_attachments_match_downloads_without_network(monkeypatch, downloads):
    messages = []

    class FakeSMTP:
        def __init__(self, *args):
            pass
        def __enter__(self):
            return self
        def __exit__(self, *args):
            pass
        def starttls(self):
            pass
        def login(self, *args):
            pass
        def send_message(self, msg):
            messages.append(msg)

    monkeypatch.setattr(smtplib, "SMTP", FakeSMTP)
    at = fill_client(AppTest.from_file(APP_PATH).run())
    for widget in at.sidebar.text_input:
        if widget.label.startswith("Email de réception"):
            widget.set_value("cabinet@example.invalid")
        elif widget.label == "SMTP user (login)":
            widget.set_value("test@example.invalid")
        elif widget.label == "SMTP password / app password":
            widget.set_value("test-password")
        elif widget.label == "Email expéditeur (From)":
            widget.set_value("test@example.invalid")
    at.run()
    click(at, "📩 Envoyer maintenant")
    assert len(messages) == 1
    attachments = {part.get_filename(): part.get_payload(decode=True)
                   for part in messages[0].iter_attachments()}
    assert attachments == downloads
    assert_documents(attachments, EXPECTED_ROWS)

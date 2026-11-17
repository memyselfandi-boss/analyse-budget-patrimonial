"""Client state, budget rules and exports, independent of the Streamlit UI."""

from io import BytesIO
from pathlib import Path
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

APP_TITLE = "Analyse patrimoniale"
TEMPLATE_PATH = Path(__file__).with_name("template.xlsx")

# Internal coordinates retained for compatibility with the existing form.
# Exports use named fields: the old template's labels/merged cells do not match
# several of these coordinates and must not be presented as a filled workbook.
IDENTITY_FIELDS = {
    "nom": (1, 9, "Nom"),
    "prenom": (2, 9, "Prénom"),
    "date_naissance": (3, 9, "Date de naissance"),
    "lieu_naissance": (4, 9, "Lieu de naissance"),
    "adresse": (5, 9, "Adresse"),
    "tel": (7, 9, "Téléphone"),
    "mail": (8, 9, "Email"),
    "situation_fam": (5, 11, "Situation familiale"),
    "parts": (4, 11, "Nombre de parts"),
}
BUDGET_FIELDS = {
    "salaire": (2, 2, "Total des salaires mensuels (avant PAS)"),
    "dec1": (3, 2, "Déclarant 1 - salaire"),
    "dec2": (4, 2, "Déclarant 2 - salaire"),
    "rev_loc": (5, 2, "Total des loyers retenus à 80 %"),
    "bien1": (7, 2, "Bien n°1 - loyer hors charges avant décote"),
    "bien2": (8, 2, "Bien n°2 - loyer hors charges avant décote"),
    "credit_flag": (12, 2, "Crédits - souscripteur ?"),
    "emprunt_rp": (13, 2, "Emprunt RP / loyer"),
    "charges": (16, 2, "Charges fixes / abonnements (hors essence et impôts)"),
    "essence": (17, 2, "Essence"),
    "impots": (18, 2, "Impôts"),
    "autres_credits": (19, 2, "Autres crédits (voiture, conso, etc.)"),
    "loisirs_budget": (20, 2, "Budget plaisirs / voyages"),
}
CATEGORIES = (
    "SICAV ou FCP", "Obligations", "Actions", "Trackers - fonds alternatifs",
    "Warrants - Futurs - Options", "Gestion directe ou personnelle",
    "Gestion déléguée ou sous mandat", "Fonds euros", "Supports unités de compte",
    "SCPI", "OPCI", "FCPI", "FIP",
)
METRIC_LABELS = {
    "revenus_retenus": "Revenus retenus",
    "charges_et_credits": "Charges + crédits existants",
    "reste_a_vivre": "Reste à vivre",
    "reste_apres_loisirs": "Épargne possible",
    "capacite_endettement_max": "Capacité d'endettement max (35 %)",
    "capacite_mensualite_restante": "Capacité de mensualité restante",
}


def get_cell(df, r, c, default=""):
    value = df.iat[r, c]
    return default if value == "" else value


def set_cell(df, r, c, value):
    df.iat[r, c] = value


def load_template():
    """Use the template structure, never its sample client's input values."""
    with pd.ExcelFile(TEMPLATE_PATH) as xl:
        df = pd.read_excel(xl, sheet_name=0, header=None).fillna("").astype(object)
    for key, (r, c, _) in IDENTITY_FIELDS.items():
        set_cell(df, r, c, 0.0 if key == "parts" else "")
    for r, c, _ in BUDGET_FIELDS.values():
        set_cell(df, r, c, 0.0)
    for i in range(len(CATEGORIES)):
        set_cell(df, i + 2, 13, "")
    return df


def collect_data(df):
    data = {key: get_cell(df, r, c) for key, (r, c, _) in IDENTITY_FIELDS.items()}
    data.update({key: float(get_cell(df, r, c, 0) or 0)
                 for key, (r, c, _) in BUDGET_FIELDS.items()})
    data["credit_flag"] = int(data["credit_flag"])
    data["experience"] = {cat: get_cell(df, i + 2, 13) == "O"
                          for i, cat in enumerate(CATEGORIES)}
    return data


def calculate_budget(data):
    """Totals are derived from details; rent is discounted exactly once."""
    salaire = data["dec1"] + data["dec2"]
    rev_loc = (data["bien1"] + data["bien2"]) * 0.80
    revenus = salaire + rev_loc
    charges = data["charges"] + data["essence"] + data["impots"]
    credits = data["emprunt_rp"] + data["autres_credits"]
    reste = revenus - charges - credits
    metrics = {
        "salaire": salaire,
        "rev_loc": rev_loc,
        "revenus_retenus": revenus,
        "charges_et_credits": charges + credits,
        "reste_a_vivre": reste,
        "reste_apres_loisirs": reste - data["loisirs_budget"],
        "capacite_endettement_max": revenus * 0.35,
        "capacite_mensualite_restante": max(0.0, revenus * 0.35 - credits),
    }
    return {key: round(value, 2) for key, value in metrics.items()}


def sync_budget_totals(df):
    budget = calculate_budget(collect_data(df))
    for key in ("salaire", "rev_loc"):
        r, c, _ = BUDGET_FIELDS[key]
        set_cell(df, r, c, budget[key])
    return budget


def export_sections(data):
    """One set of fields for both documents, including all 13 profile answers."""
    budget = calculate_budget(data)
    values = {**data, "salaire": budget["salaire"], "rev_loc": budget["rev_loc"]}
    return [
        ("Identité", [(label, values[key]) for key, (_, _, label) in IDENTITY_FIELDS.items()]),
        ("Budget mensuel", [(label, ("Oui" if values[key] else "Non")
                             if key == "credit_flag" else values[key])
                            for key, (_, _, label) in BUDGET_FIELDS.items()]),
        ("Indicateurs mensuels (euros)", [(label, budget[key])
                                         for key, label in METRIC_LABELS.items()]),
        ("Profil financier", [(cat, "Oui" if data["experience"].get(cat, False) else "Non")
                              for cat in CATEGORIES]),
    ]


def to_xlsx_bytes(data, brand):
    """Export a labelled snapshot, avoiding stale template formulas and labels."""
    rows = [("Cabinet", brand)]
    for title, fields in export_sections(data):
        rows.extend([(None, None), (title, "Valeur"), *fields])
    out = BytesIO()
    # Keep the application's existing pandas/openpyxl export dependency.
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, sheet_name="Synthèse client", header=False, index=False)
        sheet = writer.sheets["Synthèse client"]
        sheet.column_dimensions["A"].width = 64
        sheet.column_dimensions["B"].width = 58
        from openpyxl.styles import Alignment, Font, PatternFill
        titles = {title for title, _ in export_sections(data)} | {"Cabinet"}
        for row in sheet:
            for cell in row:
                # User text beginning with '=' is text, never an Excel formula.
                if isinstance(cell.value, str):
                    cell.data_type = "s"
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            if row[0].value in titles:
                for cell in row:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill("solid", fgColor="1D4ED8")
            if isinstance(row[1].value, (int, float)):
                row[1].number_format = "#,##0.00"
            sheet.row_dimensions[row[0].row].height = 36
        sheet.freeze_panes = "B2"
    return out.getvalue()


def build_pdf_bytes(data, brand):
    styles = getSampleStyleSheet()
    story = [Paragraph(escape(str(brand)), styles["Title"]),
             Paragraph(f"{APP_TITLE} - Synthèse client", styles["Heading2"]), Spacer(1, 12)]
    for title, fields in export_sections(data):
        story.append(Paragraph(title, styles["Heading3"]))
        rows = [["Champ", "Valeur"], *fields]
        rows = [[Paragraph(escape(str(value)), styles["BodyText"]) for value in row]
                for row in rows]
        table = Table(rows, colWidths=[235, 264], repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
            ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.Color(.98, .98, .98)]),
        ]))
        story.extend([table, Spacer(1, 10)])
    out = BytesIO()
    SimpleDocTemplate(out, pagesize=A4, leftMargin=48, rightMargin=48,
                      title="Synthèse Analyse patrimoniale").build(story)
    return out.getvalue()

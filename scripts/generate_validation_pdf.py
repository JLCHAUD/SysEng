"""
Génère le PDF de procédure de validation manuelle — ExoSync / SysEng Cockpits.
Usage : python scripts/generate_validation_pdf.py
Sortie : docs/ExoSync_Validation_Manuelle.pdf
"""

import sys
from pathlib import Path
from datetime import date

ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT))

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Palette ──────────────────────────────────────────────────────────────────
BLUE_DARK   = HexColor("#1a3c6e")
BLUE_MED    = HexColor("#2e6da4")
BLUE_LIGHT  = HexColor("#dce8f5")
RED         = HexColor("#c0392b")
RED_LIGHT   = HexColor("#fdecea")
GREEN       = HexColor("#27ae60")
GREEN_LIGHT = HexColor("#eafaf1")
ORANGE      = HexColor("#e67e22")
ORANGE_LIGHT= HexColor("#fef5e7")
GREY_LIGHT  = HexColor("#f5f5f5")
GREY_MED    = HexColor("#cccccc")
YELLOW      = HexColor("#fff9c4")

OUTPUT = ROOT / "docs" / "ExoSync_Validation_Manuelle.pdf"
OUTPUT.parent.mkdir(parents=True, exist_ok=True)

# ── Styles ────────────────────────────────────────────────────────────────────
ss = getSampleStyleSheet()

def S(name, **kw):
    try:
        base = ss[name]
    except KeyError:
        base = ss["Normal"]
    return ParagraphStyle(name + "_custom_" + str(id(kw)), parent=base, **kw)

style_h1 = S("Heading1", fontSize=18, textColor=BLUE_DARK, spaceBefore=6, spaceAfter=4,
              fontName="Helvetica-Bold")
style_h2 = S("Heading2", fontSize=13, textColor=BLUE_DARK, spaceBefore=14, spaceAfter=4,
              fontName="Helvetica-Bold", borderPad=4)
style_h3 = S("Heading3", fontSize=11, textColor=BLUE_MED, spaceBefore=10, spaceAfter=3,
              fontName="Helvetica-Bold")
style_body = S("Normal", fontSize=9.5, leading=15, spaceAfter=4)
style_code = S("Code", fontSize=8.5, fontName="Courier", backColor=GREY_LIGHT,
               borderWidth=0.5, borderColor=GREY_MED, borderPad=6,
               leftIndent=8, spaceAfter=6)
style_bullet = S("Normal", fontSize=9.5, leading=15, leftIndent=14, spaceAfter=2,
                 bulletIndent=4)
style_caption = S("Normal", fontSize=8, textColor=HexColor("#666666"), spaceAfter=6)
style_note = S("Normal", fontSize=9, leading=14, backColor=YELLOW, borderPad=6,
               leftIndent=6, spaceAfter=6)

def H2(txt):
    return [
        Spacer(1, 0.2*cm),
        Paragraph(txt, style_h2),
        HRFlowable(width="100%", thickness=1.5, color=BLUE_DARK, spaceAfter=4),
    ]

def H3(txt):
    return [Paragraph(txt, style_h3)]

def body(txt):
    return Paragraph(txt, style_body)

def code(txt):
    return Paragraph(txt.replace(" ", "&nbsp;").replace("\n", "<br/>"), style_code)

def bullet(txt):
    return Paragraph(f"• &nbsp;{txt}", style_bullet)

def note(txt):
    return Paragraph(f"<b>Note :</b> {txt}", style_note)

def ok_row(label, detail):
    return [label, detail]


# ── Tables helpers ────────────────────────────────────────────────────────────
def check_table(rows, col_widths=None):
    """Tableau de vérification avec fond alterné."""
    col_widths = col_widths or [4*cm, 12*cm]
    data = [["Check", "Résultat attendu"]] + rows
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    style = TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_DARK),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, BLUE_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ])
    tbl.setStyle(style)
    return tbl


def alert_box(title, content, color, bg):
    tbl = Table([[f"  {title}", content]], colWidths=[3.5*cm, 13*cm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (0,0), color),
        ("BACKGROUND",  (1,0), (1,0), bg),
        ("TEXTCOLOR",   (0,0), (0,0), white),
        ("FONTNAME",    (0,0), (0,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",  (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("BOX",         (0,0), (-1,-1), 1, color),
    ]))
    return tbl


def status_table(rows):
    """Tableau bilan final."""
    data = [["Composant", "Fichier / Commande", "Statut"]] + rows
    tbl = Table(data, colWidths=[5*cm, 8*cm, 3.5*cm])
    style = TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_DARK),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, GREY_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("VALIGN",      (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("ALIGN",       (2,0), (2,-1), "CENTER"),
    ])
    tbl.setStyle(style)
    return tbl


# ── Cover ─────────────────────────────────────────────────────────────────────
def cover_block():
    elems = []
    elems.append(Spacer(1, 1.5*cm))

    cover = Table(
        [[Paragraph(
            "<font color='white'><b>ExoSync / SysEng</b><br/>"
            "<font size=20>Procédure de Validation Manuelle</font><br/>"
            "<font size=11>Cockpits Ingénieur &amp; Dashboard Métier</font></font>",
            ParagraphStyle("cover", fontName="Helvetica", fontSize=16,
                           alignment=TA_CENTER, leading=28)
        )]],
        colWidths=[17*cm]
    )
    cover.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BLUE_DARK),
        ("TOPPADDING", (0,0), (-1,-1), 24),
        ("BOTTOMPADDING", (0,0), (-1,-1), 24),
        ("LEFTPADDING", (0,0), (-1,-1), 16),
        ("RIGHTPADDING", (0,0), (-1,-1), 16),
        ("ROUNDEDCORNERS", [8]),
    ]))
    elems.append(cover)
    elems.append(Spacer(1, 0.6*cm))

    meta = Table([
        ["Date :", date.today().strftime("%d/%m/%Y")],
        ["Version :", "1.0 — premier test end-to-end"],
        ["Projet :", "SysEng Alstom — ExoSync MVP"],
        ["Auteur :", "Jean-Luc Chaudon"],
    ], colWidths=[3*cm, 14*cm])
    meta.setStyle(TableStyle([
        ("FONTNAME",    (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9.5),
        ("TOPPADDING",  (0,0), (-1,-1), 3),
        ("BOTTOMPADDING",(0,0), (-1,-1), 3),
        ("TEXTCOLOR",   (0,0), (0,-1), BLUE_MED),
    ]))
    elems.append(meta)
    return elems


# ── Document ──────────────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        str(OUTPUT),
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm, bottomMargin=2*cm,
    )

    story = []

    # ── Cover ─────────────────────────────────────────────────────────────────
    story += cover_block()
    story.append(Spacer(1, 0.4*cm))

    # ── Sommaire ──────────────────────────────────────────────────────────────
    story += H2("Sommaire")
    toc = [
        ("1", "Pré-requis", "Dépendances Python, tests de non-régression"),
        ("2", "Test automatisé (script demo)", "Cycle complet en une commande"),
        ("3", "Validation cockpit ingénieur", "Vérifications onglet par onglet dans Excel"),
        ("4", "Validation dashboard métier", "KPIs, alertes, filtrage, _Manifeste"),
        ("5", "Simulation de saisie réelle", "Modifier les données et vérifier le cycle push/pull"),
        ("6", "Bilan et prochaines étapes", "État actuel, gaps CLI, registre.json"),
    ]
    toc_data = [["#", "Section", "Contenu"]] + toc
    toc_tbl = Table(toc_data, colWidths=[0.8*cm, 5*cm, 10.7*cm])
    toc_tbl.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_MED),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTNAME",    (0,1), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, BLUE_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("TEXTCOLOR",   (0,1), (0,-1), BLUE_DARK),
    ]))
    story.append(toc_tbl)

    # ── §1 Pré-requis ─────────────────────────────────────────────────────────
    story += H2("1 — Pré-requis")
    story.append(body("Avant de lancer les tests, s'assurer que l'environnement est correct :"))
    story.append(Spacer(1, 0.2*cm))
    story += H3("1.1  Dépendances Python")
    story.append(body("Depuis la racine du projet <b>SysEng</b> :"))
    story.append(code("pip install -r requirements.txt"))
    story.append(body("Librairies clés utilisées : <b>openpyxl</b>, <b>pydantic</b>. "
                      "La génération Excel ne nécessite pas LibreOffice."))

    story += H3("1.2  Suite de tests — non-régression")
    story.append(body("Lancer la suite complète avant tout test manuel :"))
    story.append(code("python -m pytest -q"))
    story.append(body("Résultat attendu : <b>tous les tests passent</b>, zéro échec, zéro erreur."))
    story.append(note("Si des tests échouent avant même de commencer les validations manuelles, "
                      "ne pas continuer — corriger d'abord la régression."))

    story += H3("1.3  Fichiers de configuration requis")
    req_data = [
        ["config/uo_instances.json", "Instances UO réelles du projet (5 UOs minimum)"],
        ["config/acteurs.json", "Profils acteurs — doit contenir Jean-Luc Bernard (USR004)"],
        ["output/cockpits/", "Dossier créé automatiquement par le script"],
        ["output/store.json", "Créé/mis à jour par le script (données simulées)"],
    ]
    story.append(check_table(req_data, [5.5*cm, 11*cm]))

    # ── §2 Test automatisé ────────────────────────────────────────────────────
    story += H2("2 — Test automatisé (script demo)")
    story.append(body("Le script <b>scripts/demo_cockpits.py</b> orchestre le cycle complet "
                      "en une seule commande :"))
    story.append(code("python scripts/demo_cockpits.py"))

    story += H3("2.1  Ce que fait le script")
    steps = [
        ("Étape 1", "Charge les UOs et acteurs réels depuis config/"),
        ("Étape 2", "Génère les 3 cockpits ingénieurs dans output/cockpits/"),
        ("Étape 3", "Injecte 10 valeurs simulées dans output/store.json "
                    "(simule la saisie Excel + sync ExoSync)"),
        ("Étape 4", "Génère Dashboard_Jean-Luc_Bernard.xlsx"),
        ("Étape 5", "Vérifie automatiquement 4 points critiques et affiche ✅ / ❌"),
    ]
    steps_data = [["Étape", "Action"]] + [[f"  {s}", d] for s, d in steps]
    st = Table(steps_data, colWidths=[2.5*cm, 14*cm])
    st.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_MED),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, BLUE_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("FONTNAME",    (0,1), (0,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",   (0,1), (0,-1), BLUE_DARK),
    ]))
    story.append(st)
    story.append(Spacer(1, 0.3*cm))

    story += H3("2.2  Sortie attendue (4 vérifications)")
    story.append(code(
        "  Filtrage (Denis Renard absent)  : ✅\n"
        "  UO-001 (Alice) présente         : ✅\n"
        "  Alerte dérive UO-002 (52h>48h)  : ✅\n"
        "  Avancement UO-001 = 65%         : ✅"
    ))
    story += [
        alert_box("BLOQUANT", "Si une vérification affiche ❌, arrêter — "
                  "le générateur est en erreur.", RED, RED_LIGHT),
        Spacer(1, 0.3*cm),
    ]

    story += H3("2.3  Fichiers générés")
    files_data = [
        ["Cockpit_Alice_Dubois.xlsx",      "~9 Ko", "Cockpit d'Alice (UO-001, UO-002)"],
        ["Cockpit_Bruno_Lecomte.xlsx",     "~9 Ko", "Cockpit de Bruno (UO-003, UO-004)"],
        ["Cockpit_Camille_Vidal.xlsx",     "~9 Ko", "Cockpit de Camille (UO-005)"],
        ["Dashboard_Jean-Luc_Bernard.xlsx","~11 Ko","Dashboard consolidé équipe"],
    ]
    fd = Table([["Fichier", "Taille", "Contenu"]] + files_data,
               colWidths=[6*cm, 2*cm, 8.5*cm])
    fd.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_DARK),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 9),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, BLUE_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(fd)

    # ── §3 Validation cockpit ─────────────────────────────────────────────────
    story += H2("3 — Validation cockpit ingénieur")
    story.append(body("Ouvrir <b>output/cockpits/Cockpit_Alice_Dubois.xlsx</b> dans Excel. "
                      "Vérifier les 3 onglets ci-dessous."))

    story += H3("3.1  Onglet « Mes UOs »")
    checks_uos = [
        ["Ligne 5", "Headers : UO ID | Type UO | Système | Projet | Charge (h) | % Avancement | H réalisées | Date fin | Alerte"],
        ["Lignes 6+", "Une ligne par UO — UO-001 et UO-002 présentes (pas UO-003)"],
        ["Col F (% Avancement)", "Fond JAUNE — zone de saisie ingénieur"],
        ["Col G (H réalisées)", "Fond JAUNE — zone de saisie ingénieur"],
        ["Col I (Alerte)", "Formule Excel visible : =IF(G6>E6,\"⚠ Dérive heures\",...) — pas de fond jaune"],
        ["Charge allouée", "Valeur numérique dans col E (32h pour UO-001, 48h pour UO-002)"],
    ]
    story.append(check_table(checks_uos))

    story += H3("3.2  Onglet « Agenda »")
    checks_ag = [
        ["Structure", "3 sections : « Semaine en cours », « Prochaines échéances », « Points ouverts / Actions »"],
        ["En-têtes de section", "Fond bleu foncé avec titre en blanc"],
        ["Colonnes", "UO ID | Activité | Priorité | Date échéance | Statut | Action"],
        ["Zone saisie Agenda", "Colonne « Action » en fond jaune (saisie libre)"],
        ["Points ouverts", "Lignes vides jaunes — zone à remplir par l'ingénieur"],
    ]
    story.append(check_table(checks_ag))

    story += H3("3.3  Onglet « _Manifeste »")
    checks_mf = [
        ["Ligne 1", "Cellule A1 = MANIFESTE_V=1"],
        ["Ligne 2", "12 headers : TYPE | SCOPE | NOM_GLOBAL | NOM_LOCAL | FEUILLE | TABLEAU | CLE | COLONNES | CELLULE | DIRECTION | FORMULE | COMMENTAIRE"],
        ["Règles PUSH", "2 règles par UO : avancement (col F) et heures_realisees (col G) vers store"],
        ["Règles PULL", "2 règles par UO : charge_allouee et date_fin depuis store"],
        ["COMMENTAIRE", "Colonne 12 remplie — description en français pour chaque règle (>10 chars)"],
        ["Séparation", "Ligne vide entre les règles de chaque UO"],
    ]
    story.append(check_table(checks_mf))

    # ── §4 Validation dashboard ───────────────────────────────────────────────
    story += H2("4 — Validation dashboard métier")
    story.append(body("Ouvrir <b>output/cockpits/Dashboard_Jean-Luc_Bernard.xlsx</b>. "
                      "Le dashboard consolide les 5 UOs de l'équipe (Alice + Bruno + Camille)."))

    story += H3("4.1  Onglet « Synthèse »")
    checks_syn = [
        ["KPIs (ligne 2)", "Nb UOs = 5, Charge totale = 192h, % moyen = environ 0.5, Nb alertes = 2"],
        ["Tableau consolidé", "5 lignes de données à partir de la ligne 6"],
        ["Filtrage", "Denis Renard ABSENT — l'acteur USR004 filtre par ingénieur"],
        ["Colonne Ingénieur", "Nom de l'ingénieur visible sur chaque ligne"],
        ["Colonne Alerte", "UO-002 : « Dérive heures » — UO-003 : potentiellement « Échéance proche »"],
    ]
    story.append(check_table(checks_syn))

    story += H3("4.2  Onglet « Par Ingénieur »")
    checks_pi = [
        ["Structure", "Une section par ingénieur : Alice (2 UOs), Bruno (2 UOs), Camille (1 UO)"],
        ["En-têtes", "Fond bleu avec nom ingénieur + total charge + % avancement moyen"],
        ["Colonnes UO", "UO ID | Type | Système | Projet | Charge | % Avancement | H réalisées | Date fin"],
        ["Lien cockpit", "Cellule « Voir cockpit » avec lien hypertexte vers fichier .xlsx"],
    ]
    story.append(check_table(checks_pi))

    story += H3("4.3  Onglet « Alertes »")
    checks_al = [
        ["UO-002", "Présente en alerte Critique (fond ROUGE) — 52h réalisées > 48h allouées"],
        ["Tri criticité", "Alertes Critique avant Élevée"],
        ["Colonnes", "Ingénieur | UO ID | Type alerte | Détail | Criticité"],
        ["Pas d'alerte", "Si aucune alerte : bandeau vert « Aucune alerte active »"],
    ]
    story.append(check_table(checks_al))

    story += H3("4.4  Onglet « _Manifeste »")
    checks_dmf = [
        ["Ligne 1", "MANIFESTE_V=1"],
        ["Règles", "3 règles PULL par UO : avancement, heures_realisees, points_ouverts"],
        ["Nombre lignes", "5 UOs × 3 règles = 15 règles de données (+ lignes vides entre UOs)"],
        ["COMMENTAIRE", "Colonne 12 remplie pour chaque règle"],
    ]
    story.append(check_table(checks_dmf))

    # ── §5 Simulation de saisie réelle ────────────────────────────────────────
    story += H2("5 — Simulation de saisie réelle")
    story.append(body("Pour aller plus loin, modifier les données simulées et relancer "
                      "le script pour vérifier que le dashboard s'adapte."))

    story += H3("5.1  Modifier les saisies dans le script")
    story.append(body("Éditer <b>scripts/demo_cockpits.py</b>, section <b>SAISIES_SIMULEES</b> "
                      "(ligne ~41) :"))
    story.append(code(
        "SAISIES_SIMULEES = {\n"
        "    # Passer UO-001 en dérive : 38h > 32h allouées\n"
        "    \"uo.UO-001.avancement\": 0.90,\n"
        "    \"uo.UO-001.heures_realisees\": 38.0,   # ← dérive\n"
        "    # ...\n"
        "}"
    ))
    story.append(body("Relancer ensuite :"))
    story.append(code("python scripts/demo_cockpits.py"))
    story.append(body("L'alerte doit apparaître sur UO-001 dans l'onglet Alertes du dashboard."))

    story += H3("5.2  Scénarios de test recommandés")
    scenarios = [
        ["Dérive heures", "heures_realisees > charge_allouee pour une UO",
         "Ligne rouge dans Alertes, alerte dans col I du cockpit"],
        ["Échéance proche", "date_fin < aujourd'hui + 7j (modifier config)",
         "Alerte « Échéance proche » dans Alertes"],
        ["Avancement 0%", "avancement = 0.0 sur une UO à date passée",
         "Potentielle alerte retard dans Alertes"],
        ["Tous OK", "Toutes les heures sous le seuil, avancements cohérents",
         "Onglet Alertes affiche le bandeau vert"],
    ]
    sc_tbl = Table(
        [["Scénario", "Action", "Résultat attendu"]] + scenarios,
        colWidths=[3.5*cm, 6*cm, 7*cm]
    )
    sc_tbl.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,0), BLUE_DARK),
        ("TEXTCOLOR",   (0,0), (-1,0), white),
        ("FONTNAME",    (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",    (0,0), (-1,-1), 8.5),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [white, GREY_LIGHT]),
        ("GRID",        (0,0), (-1,-1), 0.5, GREY_MED),
        ("TOPPADDING",  (0,0), (-1,-1), 5),
        ("BOTTOMPADDING",(0,0), (-1,-1), 5),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("VALIGN",      (0,0), (-1,-1), "TOP"),
    ]))
    story.append(sc_tbl)

    # ── §6 Bilan ──────────────────────────────────────────────────────────────
    story += H2("6 — Bilan et prochaines étapes")
    story.append(body("État du pipeline à l'issue de cette première vague de tests :"))
    story.append(Spacer(1, 0.3*cm))

    bilan = [
        ["Suite de tests (361 tests)", "pytest -q", "✅ PASS"],
        ["Génération cockpit ingénieur", "cockpit_ingenieur_generator.py", "✅ OK"],
        ["Génération dashboard métier", "dashboard_metier_generator.py", "✅ OK"],
        ["Cycle push/pull automatisé", "scripts/demo_cockpits.py", "✅ OK"],
        ["Validation visuelle Excel", "output/cockpits/*.xlsx", "🔲 Manuel"],
        ["CLI generate-cockpit", "main.py → cockpit_generator.py (ancien)", "⚠ Non branché"],
        ["Registre.json", "Nouveaux fichiers non déclarés", "⚠ Non mis à jour"],
    ]
    story.append(status_table(bilan))

    story.append(Spacer(1, 0.4*cm))
    story += H3("Prochaines étapes (CONV-08)")
    next_steps = [
        "Brancher <b>cockpit_ingenieur_generator</b> dans <b>main.py</b> "
        "(nouveau sous-commande <i>generate-cockpit-v2</i> ou remplacement de l'ancien)",
        "Enregistrer les nouveaux fichiers <i>Cockpit_*.xlsx</i> dans <b>config/registre.json</b> "
        "pour que <i>python main.py sync</i> les traite",
        "Test end-to-end complet : saisie manuelle dans Excel → sync → vérification dashboard",
        "Vérifier la confidentialité du repo GitHub (<i>L09U1-CFL2400-CLIM.xlsx</i> peut avoir été poussé)",
    ]
    for s in next_steps:
        story.append(bullet(s))

    story.append(Spacer(1, 0.5*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GREY_MED))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph(
        f"Document généré le {date.today().strftime('%d/%m/%Y')} — ExoSync / SysEng v0.1",
        S("Normal", fontSize=8, textColor=HexColor("#888888"), alignment=TA_CENTER)
    ))

    doc.build(story)
    print(f"PDF généré : {OUTPUT}")
    print(f"Taille     : {OUTPUT.stat().st_size // 1024} Ko")


if __name__ == "__main__":
    build()

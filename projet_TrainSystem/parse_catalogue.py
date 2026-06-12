"""
parse_catalogue.py — Parse 'UO TS.docx' en catalogue structure.
============================================================================
Lit directement le Word (ou un extrait texte [style|B|N]) et produit
catalogue_uo.json :
{
  "L09U1": {
    "lot": 9, "lot_libelle": "...", "uo": 1, "uo_libelle": "...",
    "activites":      [{"id", "designation", "description"}],
    "donnees_entree": ["..."],
    "livrables":      ["..."],
    "criteres":       ["..."],
    "complexite":     ["..."]
  }, ...
}

Structure du Word : titres en GRAS — sections reconnues : Activités,
Données d'entrée, Principaux livrables, Critères d'acceptation,
Niveaux de complexité. Tout autre titre gras = une activité.

Usage : python projet_TrainSystem/parse_catalogue.py ["chemin\\UO TS.docx"]
"""
import json
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SRC = Path(sys.argv[1] if len(sys.argv) > 1
           else r"C:\Users\fabie\Documents\JLC\UO TS.docx")
OUT = Path(__file__).parent / "catalogue_uo.json"

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def extract_docx_lines(path: Path) -> list[str]:
    """Extrait les paragraphes du .docx au format '[style|B|N] texte'."""
    doc = ET.fromstring(zipfile.ZipFile(path).read("word/document.xml"))
    lines = []
    for p in doc.iter(_W + "p"):
        style_el = p.find(f".//{_W}pStyle")
        style = style_el.get(_W + "val") if style_el is not None else ""
        text = "".join(t.text or "" for t in p.iter(_W + "t")).strip()
        if not text:
            continue
        bold = any(
            r.find(f"{_W}rPr/{_W}b") is not None
            and "".join(t.text or "" for t in r.iter(_W + "t")).strip()
            for r in p.iter(_W + "r"))
        num = p.find(f".//{_W}numPr") is not None
        lines.append(f"[{style}|{'B' if bold else '.'}|{'N' if num else '.'}] {text}")
    return lines

RE_LINE = re.compile(r"^\[([^|\]]*)\|([B.])\|([N.])\]\s?(.*)$")
RE_LOT = re.compile(r"LOT\s*(\d+)\s*[-:]?\s*(.+)", re.IGNORECASE)
RE_UO = re.compile(r"UO\s*(\d)\b")
RE_ACT_CODE = re.compile(r"^([A-Z]{2}\.L\d\.\d+)\s*:\s*(.+)$")


def norm(s: str) -> str:
    """minuscules sans accents pour comparer les titres de sections."""
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower().strip()


SECTIONS = {
    "activites": "activites",
    "donnees d'entree": "donnees_entree",
    "donnees d’entree": "donnees_entree",
    "principaux livrables": "livrables",
    "livrables": "livrables",
    "criteres d'acceptation": "criteres",
    "criteres d’acceptation": "criteres",
    "niveaux de complexite": "complexite",
    "niveau de complexite": "complexite",
}
NOISE = ("uncontrolled when printed", "confidential")

catalogue = {}
lot_num, lot_lib = None, ""
uo = None          # dict de l'UO courante
mode = None        # section courante
current = None     # activite courante
act_seq = 0

if SRC.suffix.lower() == ".docx":
    lines = extract_docx_lines(SRC)
else:
    lines = SRC.read_text(encoding="utf-8").splitlines()

for raw in lines:
    m = RE_LINE.match(raw)
    if not m:
        continue
    style, bold, _num, text = m.groups()
    text = text.strip()
    if not text or any(n in norm(text) for n in NOISE):
        continue

    # ── En-tetes de structure (hors tableaux) ────────────────────────────────
    if style == "Paragraphedeliste":
        ml = RE_LOT.search(text)
        if ml and "LOT" in text.upper():
            lot_num, lot_lib = int(ml.group(1)), ml.group(2).strip()
            uo, mode, current = None, None, None
            continue
        mu = RE_UO.search(text)
        if mu and lot_num:
            key = f"L{lot_num:02d}U{mu.group(1)}"
            uo = catalogue.setdefault(key, {
                "lot": lot_num, "lot_libelle": lot_lib,
                "uo": int(mu.group(1)),
                "uo_libelle": ("System Engineer activities" if mu.group(1) == "1"
                               else "Sub-System Engineer activities"),
                "activites": [], "donnees_entree": [], "livrables": [],
                "criteres": [], "complexite": [],
            })
            mode, current, act_seq = None, None, 0
            continue

    if style != "TableParagraph" or uo is None:
        continue

    # ── Lignes de tableau ─────────────────────────────────────────────────────
    n = norm(text)
    if bold == "B":
        if n.startswith("uo1") or n.startswith("uo2"):
            continue
        if n in SECTIONS:
            mode = SECTIONS[n]
            current = None
            continue
        if mode in (None, "activites"):
            # Gras se terminant par ':' (ex. "Remarques :") = sous-titre de
            # description, pas une nouvelle activite
            if n.rstrip(":").strip() in ("remarques", "remarque") or n.endswith(":"):
                if current is not None:
                    current["description"].append(text)
                continue
            # Titre gras inconnu = une activite
            mode = "activites"
            mc = RE_ACT_CODE.match(text)
            if mc:
                current = {"id": mc.group(1), "designation": mc.group(2).strip(),
                           "description": []}
            else:
                act_seq += 1
                current = {"id": f"ACT-{act_seq:02d}", "designation": text,
                           "description": []}
            uo["activites"].append(current)
            continue
        # gras dans une autre section : on le range comme item
        uo[mode].append(text)
        continue

    # Texte normal / puces
    if mode == "activites" and current is not None:
        current["description"].append(text)
    elif mode and mode != "activites":
        uo[mode].append(text)

for u in catalogue.values():
    for act in u["activites"]:
        act["description"] = "\n".join(act["description"])

OUT.write_text(json.dumps(catalogue, ensure_ascii=False, indent=2), encoding="utf-8")

print(f"[OK] {OUT.name}")
hdr = f"{'UO':6s} {'act':>3s} {'DE':>3s} {'liv':>3s} {'crit':>4s} {'cplx':>4s}  lot"
print(" ", hdr)
for key, u in sorted(catalogue.items()):
    print(f"  {key:6s} {len(u['activites']):3d} {len(u['donnees_entree']):3d} "
          f"{len(u['livrables']):3d} {len(u['criteres']):4d} {len(u['complexite']):4d}"
          f"  {u['lot_libelle'][:45]}")

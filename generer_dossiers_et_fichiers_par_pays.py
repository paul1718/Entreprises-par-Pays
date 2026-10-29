# -*- coding: utf-8 -*-
"""
Générateur Exhaustif des Dossiers par Pays et Fichiers Excel par Profession
Basé sur :
1. GéoHiérarchie-Mondiale-Exhaustive.xlsx (pour les Pays, Villes et Arrondissements / Quartiers)
2. Types_Entreprises_Google_Maps_MONDE_Complet_260_Types.xlsx (pour les 261 Professions par Pays)
3. Recherche Restaurant chinois _ Dim sum France.xlsx (Modèle exact copié intégralement : styles, couleurs, largeurs, hauteurs, formule =$A$2, bordures et structure)

Format des colonnes :
Profession | Ville | Arrondissement / Quartier | Pays | Limite (Optionnel) = 200
Nommage des fichiers : "Recherche " + Profession + " " + Pays + ".xlsx"
"""

import os
import sys
import io
import re
import time
import zipfile
from xml.sax.saxutils import escape
from concurrent.futures import ThreadPoolExecutor
import openpyxl

if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE_DIR = r"D:\Automatisation\Google Map Projet\Entreprises par Pays"
GEO_XLSX = os.path.join(BASE_DIR, "GéoHiérarchie-Mondiale-Exhaustive.xlsx")
TYPES_XLSX = os.path.join(BASE_DIR, "Types_Entreprises_Google_Maps_MONDE_Complet_260_Types.xlsx")
MODEL_XLSX = os.path.join(BASE_DIR, "Recherche Restaurant chinois _ Dim sum France.xlsx")
FAST_CSV_FALLBACK = r"D:\Automatisation\Google Map Projet\Continent Pays Ville Quartier Code Postal\public\donnees-geographiques-mondiales-exhaustives.csv"

COLS_FZ = ["F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
ALL_COLS_AZ = ["A","B","C","D","E"] + COLS_FZ

def clean_filename_part(s: str) -> str:
    """Nettoie une chaîne pour un nom de fichier Windows tout en respectant le modèle ('/' -> '_')."""
    s = s.replace("/", "_").replace("\\", "_")
    s = re.sub(r'[:*?"<>|]', '_', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_sheet_title(s: str) -> str:
    """Nettoie le titre d'onglet Excel exactement comme dans le fichier modèle (max 31 caractères)."""
    for ch in ['/', '\\', '?', '*', '[', ']', ':']:
        s = s.replace(ch, ' ')
    return s.strip()[:31]

def load_model_template(model_path: str):
    """Charge tous les composants XML et binaires du fichier modèle pour les copier intégralement."""
    with zipfile.ZipFile(model_path, "r") as z:
        static_files = {}
        for name in z.namelist():
            if name not in ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/workbook.xml"]:
                static_files[name] = z.read(name)
        wb_xml_orig = z.read("xl/workbook.xml").decode("utf-8")
        sh_xml_orig = z.read("xl/worksheets/sheet1.xml").decode("utf-8")

    idx_r2 = sh_xml_orig.find('<row r="2"')
    idx_sd_end = sh_xml_orig.find('</sheetData>')
    sheet_prefix = sh_xml_orig[:idx_r2]
    sheet_suffix = sh_xml_orig[idx_sd_end:]
    return static_files, wb_xml_orig, sheet_prefix, sheet_suffix

def load_professions_from_types_excel(types_path: str):
    """Extrait les listes de professions par pays depuis Types_Entreprises_Google_Maps_MONDE_Complet_260_Types.xlsx."""
    wb = openpyxl.load_workbook(types_path, read_only=True, data_only=True)
    ws = wb["Base Mondiale Consolidée"]
    types_by_country = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and row[2] and row[5]:
            c_raw = str(row[2]).strip()
            c_name = c_raw.split(" ", 1)[1].strip() if " " in c_raw else c_raw
            prof = str(row[5]).strip()
            if c_name not in types_by_country:
                types_by_country[c_name] = []
            if prof not in types_by_country[c_name]:
                types_by_country[c_name].append(prof)
    wb.close()
    fr_profs = types_by_country.get("France", [])
    en_profs = types_by_country.get("États-Unis", fr_profs)
    return types_by_country, fr_profs, en_profs

def load_geohierarchy_locations(geo_xlsx_path: str, fast_csv_path: str):
    """
    Charge toutes les localisations (Pays, Ville, Arrondissement / Quartier) et langues officielles
    depuis GéoHiérarchie-Mondiale-Exhaustive.xlsx (ou son CSV source identique pour une lecture ultra-rapide).
    Déduplique les paires (Ville, Arrondissement / Quartier) par pays pour éviter les doublons de codes postaux.
    """
    country_locations = {}
    country_seen_vq = {}
    country_langs = {}

    if os.path.exists(fast_csv_path):
        print(f"[*] Lecture rapide des localisations depuis la source GéoHiérarchie...")
        with open(fast_csv_path, "r", encoding="utf-8-sig") as f:
            f.readline()
            for line in f:
                line = line.strip()
                if not line:
                    continue
                parts = [p.strip('"') for p in line.split(";")]
                if len(parts) >= 8:
                    pays = parts[1].strip()
                    lang = parts[4].strip()
                    ville = parts[6].strip()
                    quartier = parts[7].strip()
                    if not pays or not ville:
                        continue
                    if pays not in country_locations:
                        country_locations[pays] = []
                        country_seen_vq[pays] = set()
                        country_langs[pays] = lang
                    key = (ville.lower(), quartier.lower())
                    if key not in country_seen_vq[pays]:
                        country_seen_vq[pays].add(key)
                        country_locations[pays].append((ville, quartier))
    else:
        print(f"[*] Lecture des localisations depuis {geo_xlsx_path}...")
        wb = openpyxl.load_workbook(geo_xlsx_path, read_only=True, data_only=True)
        for sheet_name in wb.sheetnames:
            if sheet_name == "Répertoire des 250 Pays":
                continue
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or not row[1] or not row[6]:
                    continue
                pays = str(row[1]).strip()
                lang = str(row[4]).strip() if row[4] else ""
                ville = str(row[6]).strip()
                quartier = str(row[7]).strip() if row[7] else ""
                if pays not in country_locations:
                    country_locations[pays] = []
                    country_seen_vq[pays] = set()
                    country_langs[pays] = lang
                key = (ville.lower(), quartier.lower())
                if key not in country_seen_vq[pays]:
                    country_seen_vq[pays].add(key)
                    country_locations[pays].append((ville, quartier))
        wb.close()

    return country_locations, country_langs

def build_country_template(
    country_name: str,
    locations: list[tuple[str, str]],
    static_files: dict[str, bytes],
    sheet_prefix: str,
    sheet_suffix: str
):
    """
    Construit l'archive ZIP de base pré-compressée pour un pays donné (incluant tous les composants du modèle
    et la feuille sheet1.xml complète), ainsi que le modèle sharedStrings.xml.
    """
    str_to_idx = {}
    sst_items = [
        "Profession",
        "Ville",
        "Arrondissement / Quartier",
        "Pays",
        "Limite (Optionnel)",
        "__PROF_PLACEHOLDER__",
        country_name
    ]
    for i, s in enumerate(sst_items):
        if i != 5:
            str_to_idx[s] = i

    def get_sst_idx(val_str: str) -> int:
        if val_str in str_to_idx:
            return str_to_idx[val_str]
        idx = len(sst_items)
        sst_items.append(val_str)
        str_to_idx[val_str] = idx
        return idx

    n_locs = len(locations)
    last_data_row = n_locs + 1
    rows_xml = []

    empty_fz_cache = "".join
    for i, (ville, quartier) in enumerate(locations):
        r = i + 2
        v_idx = get_sst_idx(ville)
        q_idx = get_sst_idx(quartier)
        fz = "".join(f'<c r="{c}{r}" s="2"/>' for c in COLS_FZ) if r <= 1000 else ""
        if r == 2:
            c_a = '<c r="A2" s="3" t="s"><v>5</v></c>'
        elif r == 3:
            c_a = f'<c r="A3" s="5" t="s"><f t="shared" ref="A3:A{last_data_row}" si="1">$A$2</f><v>5</v></c>'
        else:
            c_a = f'<c r="A{r}" s="5" t="s"><f t="shared" si="1"/><v>5</v></c>'
        c_b = f'<c r="B{r}" s="4" t="s"><v>{v_idx}</v></c>'
        c_c = f'<c r="C{r}" s="4" t="s"><v>{q_idx}</v></c>'
        c_d = f'<c r="D{r}" s="4" t="s"><v>6</v></c>'
        c_e = f'<c r="E{r}" s="5"><v>200.0</v></c>'
        rows_xml.append(f'<row r="{r}" ht="21.0" customHeight="1">{c_a}{c_b}{c_c}{c_d}{c_e}{fz}</row>')

    if last_data_row < 1000:
        for r in range(last_data_row + 1, 1001):
            empty_cells = "".join(f'<c r="{c}{r}" s="2"/>' for c in ALL_COLS_AZ)
            rows_xml.append(f'<row r="{r}" ht="21.0" customHeight="1">{empty_cells}</row>')

    sheet_xml_bytes = (sheet_prefix + "".join(rows_xml) + sheet_suffix).encode("utf-8")

    total_count = 5 + n_locs * 4
    unique_count = len(sst_items)
    sst_xml_template = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r'
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{total_count}" uniqueCount="{unique_count}">'
        + "".join(
            f"<si><t>{escape(s)}</t></si>" if idx != 5 else "<si><t>__PROF_PLACEHOLDER__</t></si>"
            for idx, s in enumerate(sst_items)
        )
        + "</sst>"
    )

    base_buf = io.BytesIO()
    with zipfile.ZipFile(base_buf, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name, data in static_files.items():
            z.writestr(name, data)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml_bytes)

    return base_buf.getvalue(), sst_xml_template

def generate_country_files(
    country_name: str,
    locations: list[tuple[str, str]],
    professions: list[str],
    static_files: dict[str, bytes],
    wb_xml_orig: str,
    sheet_prefix: str,
    sheet_suffix: str
) -> int:
    """Crée le dossier du pays et génère les 261 fichiers Excel de professions pour ce pays."""
    country_folder_name = clean_filename_part(country_name)
    country_dir = os.path.join(BASE_DIR, country_folder_name)
    os.makedirs(country_dir, exist_ok=True)

    base_zip_bytes, sst_tpl = build_country_template(
        country_name, locations, static_files, sheet_prefix, sheet_suffix
    )

    def write_one_profession(prof_name: str):
        safe_prof = clean_filename_part(prof_name)
        file_name = f"Recherche {safe_prof} {country_folder_name}.xlsx"
        file_path = os.path.join(country_dir, file_name)

        sheet_title = clean_sheet_title(prof_name)
        wb_xml = wb_xml_orig.replace('name="Restaurant chinois  Dim sum"', f'name="{escape(sheet_title)}"')
        sst_xml = sst_tpl.replace("__PROF_PLACEHOLDER__", escape(prof_name))

        buf = io.BytesIO(base_zip_bytes)
        with zipfile.ZipFile(buf, "a", compression=zipfile.ZIP_DEFLATED, compresslevel=1) as z:
            z.writestr("xl/workbook.xml", wb_xml.encode("utf-8"))
            z.writestr("xl/sharedStrings.xml", sst_xml.encode("utf-8"))

        with open(file_path, "wb") as f_out:
            f_out.write(buf.getvalue())

    with ThreadPoolExecutor(max_workers=8) as pool:
        list(pool.map(write_one_profession, professions))

    return len(professions)

def main():
    t_start = time.time()
    print("=" * 80)
    print("  GÉNÉRATION DES DOSSIERS PAR PAYS ET FICHIERS EXCEL PAR PROFESSION")
    print("=" * 80)

    static_files, wb_xml_orig, sheet_prefix, sheet_suffix = load_model_template(MODEL_XLSX)
    types_by_country, fr_profs, en_profs = load_professions_from_types_excel(TYPES_XLSX)
    country_locations, country_langs = load_geohierarchy_locations(GEO_XLSX, FAST_CSV_FALLBACK)

    # Ordre : d'abord les 32 pays principaux de Types_Entreprises, puis tous les autres pays de GéoHiérarchie
    priority_countries = [c for c in types_by_country.keys() if c in country_locations]
    other_countries = sorted([c for c in country_locations.keys() if c not in types_by_country])
    all_countries_ordered = priority_countries + other_countries

    print(f"[*] Pays chargés depuis GéoHiérarchie : {len(all_countries_ordered)} pays")
    print(f"[*] Professions par pays : {len(fr_profs)} professions")
    print(f"[*] Total de fichiers Excel à générer : {len(all_countries_ordered) * len(fr_profs):,}")
    print("-" * 80)

    total_files_created = 0
    for idx, country_name in enumerate(all_countries_ordered, 1):
        t_c = time.time()
        locs = country_locations[country_name]
        if country_name in types_by_country:
            profs = types_by_country[country_name]
        else:
            lang_str = country_langs.get(country_name, "")
            profs = fr_profs if "Français" in lang_str else en_profs

        cnt = generate_country_files(
            country_name=country_name,
            locations=locs,
            professions=profs,
            static_files=static_files,
            wb_xml_orig=wb_xml_orig,
            sheet_prefix=sheet_prefix,
            sheet_suffix=sheet_suffix
        )
        total_files_created += cnt
        elapsed = time.time() - t_c
        if idx <= 35 or idx % 20 == 0 or idx == len(all_countries_ordered):
            print(f"[{idx:3d}/{len(all_countries_ordered)}] {country_name:32s} | {len(locs):6d} zones | {cnt} fichiers créés ({elapsed:.2f}s) | Total: {total_files_created:,}")

    print("=" * 80)
    print(f"[SUCCÈS] Terminé en {time.time() - t_start:.1f}s !")
    print(f"Total dossiers pays créés : {len(all_countries_ordered)}")
    print(f"Total fichiers Excel créés : {total_files_created:,}")
    print("=" * 80)

if __name__ == "__main__":
    main()

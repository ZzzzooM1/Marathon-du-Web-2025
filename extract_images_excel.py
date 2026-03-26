"""
Script d'extraction d'images depuis des fichiers Excel - DIGIDEM
----------------------------------------------------------------
Usage:
    python extract_images_excel.py --input "C:/dossier/racine" --output "C:/dossier/images_extraites"

Le script parcourt TOUS les sous-dossiers récursivement,
trouve tous les .xlsx/.xls et extrait toutes les images dedans.
"""

import os
import argparse
from pathlib import Path
from zipfile import ZipFile
import shutil

def extract_images_from_excel(excel_path: Path, output_dir: Path, compteur: list) -> int:
    extracted = 0
    try:
        with ZipFile(excel_path, 'r') as z:
            media_files = [f for f in z.namelist() if f.startswith('xl/media/')]
            if not media_files:
                return 0

            excel_stem = excel_path.stem[:30].replace(' ', '_')

            for media_file in media_files:
                ext = Path(media_file).suffix.lower()
                if ext not in {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp', '.tiff'}:
                    continue

                compteur[0] += 1
                new_name = f"img_{compteur[0]:04d}__{excel_stem}{ext}"
                dest = output_dir / new_name

                with z.open(media_file) as src, open(dest, 'wb') as dst:
                    shutil.copyfileobj(src, dst)

                extracted += 1

    except Exception as e:
        print(f"    Erreur sur {excel_path.name}: {e}")

    return extracted


def run(input_dir: str, output_dir: str):
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    excel_files = sorted(input_path.rglob("*.xlsx")) + sorted(input_path.rglob("*.xls"))

    if not excel_files:
        print(f"Aucun fichier Excel trouvé dans {input_dir}")
        return

    print(f"{len(excel_files)} fichier(s) Excel trouvé(s)\n")

    compteur = [0]
    total_images = 0
    fichiers_avec_images = 0

    for excel_path in excel_files:
        chemin_relatif = excel_path.relative_to(input_path)
        print(f"  {chemin_relatif} ...", end=" ")
        n = extract_images_from_excel(excel_path, output_path, compteur)
        if n > 0:
            fichiers_avec_images += 1
            total_images += n
            print(f"{n} image(s) extraite(s)")
        else:
            print("aucune image")

    print("\n" + "="*55)
    print(f"  Fichiers Excel traités : {len(excel_files)}")
    print(f"  Fichiers avec images   : {fichiers_avec_images}")
    print(f"  Images extraites       : {total_images}")
    print(f"  Dossier de sortie      : {output_path}")
    print("="*55)
    print("\nProchaine étape : lance restructure_dataset.py sur ce dossier")
    print(f'  --input "{output_path}"')


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extraction images depuis Excel")
    parser.add_argument("--input",  required=True, help="Dossier racine contenant les Excel (et sous-dossiers)")
    parser.add_argument("--output", required=True, help="Dossier de sortie pour les images extraites")
    args = parser.parse_args()
    run(args.input, args.output)

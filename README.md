# Entreprises par Pays - Générateur Mondial Excel

Générateur automatique de dossiers par pays et de fichiers Excel de recherche Google Maps par profession et par pays.

## Structure du Projet (Fichiers Sources Uniquement)

- `generer_dossiers_et_fichiers_par_pays.py` : Script autonome ultra-optimisé permettant de générer les dossiers pour 252 pays et 261 fichiers Excel par profession par pays (soit 65 772 fichiers Excel au total).
- `GéoHiérarchie-Mondiale-Exhaustive.xlsx` : Référentiel mondial des continents, pays, villes, arrondissements/quartiers et codes postaux.
- `Types_Entreprises_Google_Maps_MONDE_Complet_260_Types.xlsx` : Répertoire mondial consolidé des 261 types et catégories d'entreprises Google Maps.
- `Recherche Restaurant chinois _ Dim sum France.xlsx` : Fichier modèle original reproduit fidèlement pour chaque pays et profession (styles, en-têtes bleus `#1A73E8`, cellule A2 jaune `#FFFF00`, formules partagées `=$A$2`, limite par défaut de 200).

## Utilisation

```bash
python generer_dossiers_et_fichiers_par_pays.py
```

> **Note importante :** Les dossiers par pays et les 65 772 fichiers Excel générés ne sont jamais versionnés sur GitHub (exclus via `.gitignore`). Seuls les scripts et fichiers sources de référence sont suivis.

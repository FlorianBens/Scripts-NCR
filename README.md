# Astreintes — Excel ➜ CSV / ICS

Ce dépôt contient **deux scripts totalement indépendants**, chacun ayant un rôle précis et autonome :

1. **Un script de lecture de planning Excel**  
   → il détecte les jours d’astreinte et les structure (dates, blocs, CSV)

2. **Un script de génération d’agenda ICS**  
   → il transforme ces dates en événements de calendrier (journées entières)

L’objectif est de partir d’un **planning Excel brut** et d’obtenir :
- des **données exploitables (CSV)**
- un **agenda ICS propre**, importable dans Samsung Calendar, Google Calendar ou Outlook.

---

## Prérequis

- Python **3.10 ou supérieur**
- Dépendance principale :
  - `openpyxl` (lecture des fichiers Excel)

Installation :
```bash
pip install openpyxl

## Script 1 — Lecture du planning Excel

### Rôle

Ce script analyse un fichier Excel de planning (.xlsx) et détecte automatiquement les jours d’astreinte à partir de codes définis (par exemple -14, -X).

Il constitue la base de données du projet : toutes les sorties (CSV, agenda ICS) reposent sur les informations qu’il extrait.

Il permet :
- d’identifier les dates exactes d’astreinte
- de regrouper les jours consécutifs
- de générer un CSV exploitable
- d’alimenter le script de génération ICS

---

### Principe de fonctionnement

Le script repose sur une structure fixe du tableau Excel :

- une ligne contenant les mois
- une ligne contenant les jours du mois
- une ligne cible contenant les codes de planning

Étapes internes :

1. Charger le fichier Excel avec openpyxl
2. Sélectionner la feuille à analyser
3. Parcourir le planning colonne par colonne
4. Pour chaque colonne :
   - lire le mois
   - lire le jour
   - lire la valeur de planning
5. Si la valeur correspond à un code d’astreinte :
   - reconstruire la date complète (année / mois / jour)
   - ajouter la date à la liste des astreintes
6. Une fois toutes les dates collectées :
   - trier chronologiquement
   - regrouper les dates consécutives en blocs
7. Exporter les résultats (CSV et structures internes)

Logique simplifiée :

pour chaque colonne :
  mois = lire ligne mois
  jour = lire



## Script 2 — Génération de l’agenda ICS (journées entières)

### Rôle

Ce script génère un fichier agenda au format ICS importable dans un calendrier (Samsung Calendar, Google Calendar, Outlook, etc.) à partir des jours d’astreinte.

Il crée des événements :
- sur des journées entières (pas d’heures)
- regroupés par périodes consécutives (blocs)
- avec un titre uniforme (ex : "Astreinte")

Le script peut fonctionner :
- à partir des dates / blocs produits par le script 1
- ou à partir d’un fichier intermédiaire (ex : CSV), selon l’implémentation

---

### Principe de fonctionnement

1. Charger la liste des dates (ou blocs de dates) issus de l’extraction
2. Trier les dates et/ou vérifier qu’elles sont cohérentes
3. Regrouper les dates consécutives en blocs (si ce n’est pas déjà fait)
4. Pour chaque bloc :
   - définir la date de début (incluse)
   - définir la date de fin au format ICS (exclusif)
5. Écrire un événement ICS par bloc
6. Générer un fichier .ics complet et valide

---

### Format ICS en journées entières

Pour un événement "journée entière", le format ICS utilise :

- DTSTART;VALUE=DATE:YYYYMMDD
- DTEND;VALUE=DATE:YYYYMMDD

Important : DTEND est toujours exclusif (c’est le jour suivant le dernier jour affiché).

Exemple : astreinte du 05/01/2026 au 08/01/2026 inclus

DTSTART;VALUE=DATE:20260105
DTEND;VALUE=DATE:20260109

---

### Exécution

python make_ics.py

---

### Sortie

Le script produit un fichier ICS, par exemple :

- astreintes.ics

Si l’option overwrite est activée :
- l’ancien fichier est supprimé (s’il existe)
- un nouveau fichier est recréé proprement (évite les doublons à l’import)

---

### Points importants (compatibilité calendrier)

- Encodage recommandé : UTF-8
- Dates obligatoirement au format : YYYYMMDD
- DTEND exclusif obligatoire (sinon erreurs / affichage incorrect)
- Éviter les caractères non échappés dans SUMMARY / DESCRIPTION



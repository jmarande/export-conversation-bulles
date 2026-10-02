# Export Conversation Bulles

Application Python avec interface graphique permettant de transformer des exports de conversations contenus dans des fichiers Excel en documents DOCX présentés sous forme de bulles de discussion.

Le rendu est conçu pour être exploitable dans **LibreOffice Writer** et Microsoft Word.

## Fonctionnalités

- Import de fichiers Excel (`.xlsx` / `.xls`)
- Sélection des colonnes expéditeur, message, direction, date et heure
- Présentation des échanges sous forme de bulles gauche / droite
- Couleur configurable par participant
- Texte configurable, noir par défaut
- Largeur des bulles et espacement réglables
- Conservation des retours à la ligne utiles tout en supprimant les lignes vides parasites
- Les bulles sont maintenues sur une seule page lorsqu'elles peuvent y tenir
- Déduplication optionnelle des messages
- Prise en charge de colonnes additionnelles
- Gestion optionnelle des médias et miniatures vidéo
- Aperçu et filtre des données avant export
- Sauvegarde automatique des paramètres entre deux lancements
- Export DOCX compatible LibreOffice

## Installation

Python 3.11 ou plus récent est recommandé.

```bash
python -m pip install -r requirements.txt
```

## Lancement

```bash
python export_conversation_bulles.py
```

Sous Windows :

```powershell
py export_conversation_bulles.py
```

## Dépendances

- `pandas`
- `openpyxl`
- `customtkinter`
- `python-docx`
- `Pillow`

Pour la génération de miniatures vidéo, `ffmpeg` doit être installé séparément et disponible dans le `PATH`.

## Paramètres

Les réglages de l'application sont conservés automatiquement dans :

```text
~/.export_conversation_bulles_libreoffice_settings.json
```

## Confidentialité

L'application fonctionne localement. Les fichiers Excel et les conversations traitées ne sont pas envoyés vers un service distant par le script.

Avant de publier des captures d'écran, exemples ou fichiers de test, veillez à supprimer toute donnée personnelle ou donnée issue d'un dossier réel.

## État du projet

Version actuelle : **LibreOffice V5**.

Le projet est principalement développé et testé sous Windows avec export vers LibreOffice Writer.

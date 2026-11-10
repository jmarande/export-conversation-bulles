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

## Import Signal (JSON V2 déjà déchiffré)

L'application conserve l'import Excel existant et propose maintenant **Ouvrir Signal JSON…**.
Le fichier `conversations.json` est un export de conversations **déjà décodé**
(par exemple par sigbackup), pas une base `signal.db` chiffrée.

1. Ouvrir le JSON ; sélectionner, si demandé, le dossier `medias/` issu du même export.
2. Choisir une conversation ou « Toutes les conversations » dans le menu déroulant.
3. Vérifier les données en aperçu, les participants, la direction et les chemins médias.
4. Exporter en DOCX Writer/Word. Le document commence par un tableau des contacts liés aux seules conversations retenues ; chaque conversation démarre ensuite sur une nouvelle page avec la liste des participants, la première et la dernière date et le nombre d'éléments. Les images sont insérées, les vidéos ont une miniature
   lorsque ffmpeg est disponible, et les autres fichiers sont référencés.

Les jointures médias utilisent **exclusivement** `attachments[].extractedPath` et le
dossier médias choisi. Aucune correspondance approximative n'est autorisée dans ce mode.
Les médias non inclus dans la sauvegarde sont signalés explicitement, sans les inventer.
Les dates affichées sont converties de l'UTC vers Europe/Paris ; les horodatages ISO UTC
et les références sources restent dans les colonnes importées.
Le SHA-256 du JSON et le nom du snapshot sont ajoutés à la provenance du document.
La déduplication est désactivée pour Signal afin de ne pas éliminer de preuves.

L'import XML n'est **pas** encore pris en charge : il nécessite un exemple et un schéma
de l'export concerné. Ne pas transmettre de clés Signal, de sauvegardes chiffrées ou de
données sensibles sur GitHub. Les jeux de test du dépôt sont entièrement synthétiques.

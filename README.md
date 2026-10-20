# Automatisation du traitement des factures avec Python

Ce projet met en place un robot d'automatisation (RPA) développé en **Python** pour lire, vérifier et classer des factures Excel, produire un rapport récapitulatif et notifier un responsable par e-mail lorsqu'une facture nécessite un contrôle.

## Fonctionnalités

- Lecture automatique des factures `.xlsx` placées directement dans le dossier `Factures/`.
- Extraction des informations : **numéro**, **fournisseur**, **date** et **montant**.
- Vérification des colonnes obligatoires, des valeurs manquantes et du format du montant.
- Classification des factures selon un seuil de **1 000 MAD** :
  - **Normal** si le montant est inférieur ou égal à 1 000 MAD : déplacement vers `Factures/Traitees/`.
  - **À contrôler** si le montant est supérieur à 1 000 MAD : notification par e-mail et déplacement vers `Factures/A_controler/`.
- Génération du fichier `rapport_factures.xlsx` regroupant les factures traitées avec succès.
- Journalisation des opérations et des erreurs dans `log.txt`.

## Technologies utilisées

- **Python 3**
- **Pandas** : lecture et traitement des factures, création du rapport.
- **OpenPyXL** : prise en charge des fichiers Excel `.xlsx`.
- **python-dotenv** : chargement des paramètres depuis un fichier `.env`.
- **smtplib** et **email.mime** : envoi des notifications par SMTP Gmail.
- **os**, **shutil** et **logging** : dossiers, déplacement des fichiers et journalisation.

## Structure du projet

```text
Automatisation du traitement des factures/
├── main.py
├── Factures/
│   ├── facture_exemple.xlsx         # À déposer ici avant un traitement
│   ├── Traitees/                   # Factures normales déjà traitées
│   └── A_controler/                # Factures nécessitant un contrôle
├── rapport_factures.xlsx           # Rapport généré après traitement valide
├── log.txt                         # Journal d'exécution
├── .gitignore
└── README.md
```

> Le dossier ZIP fourni contient déjà des exemples de factures dans `Traitees/` et `A_controler/`, ainsi qu'un fichier `factures_erreur.xlsx` dans `Factures/`. Le script ne parcourt pas les sous-dossiers et ne retraite donc pas automatiquement les factures déjà classées.

## Installation

1. Télécharger ou cloner le projet et ouvrir un terminal dans son dossier.
2. (Recommandé) Créer et activer un environnement virtuel :

   ```bash
   python -m venv venv
   ```

   Sous Windows (PowerShell) :

   ```powershell
   .\venv\Scripts\Activate.ps1
   ```

3. Installer les dépendances :

   ```bash
   pip install pandas openpyxl python-dotenv
   ```

## Configuration de l'envoi d'e-mails

Créer un fichier `.env` à la racine du projet :

```dotenv
EMAIL_EXPEDITEUR=adresse@gmail.com
MOT_DE_PASSE_APPLICATION=mot_de_passe_application
EMAIL_RESPONSABLE=responsable@example.com
```

Pour Gmail, utiliser un **mot de passe d'application** si votre compte le permet, et non votre mot de passe habituel. Ne jamais publier le fichier `.env` sur GitHub : l'ajouter au `.gitignore`.

Le script utilise `smtp.gmail.com` sur le port `587` avec STARTTLS. L'envoi d'e-mail n'est tenté que pour les factures dont le montant dépasse **1 000 MAD**.

## Format attendu des factures

Chaque fichier Excel à traiter doit contenir au minimum les colonnes suivantes (à l'identique) :

| numero | fournisseur | date | montant |
|---|---|---|---:|
| FAC-001 | Fournisseur A | 2026-10-20 | 750 |
| FAC-002 | Fournisseur B | 2026-10-20 | 1500 |

Le script lit **la première ligne de données** de chaque fichier Excel.

## Exécution

1. Placer les factures `.xlsx` à analyser directement dans `Factures/`.
2. Lancer :

   ```bash
   python main.py
   ```

3. Consulter les résultats dans le terminal, dans `Factures/Traitees/`, `Factures/A_controler/`, `rapport_factures.xlsx` et `log.txt`.

**Attention :** le script **déplace** les factures valides vers leurs dossiers respectifs. Si un fichier portant le même nom existe déjà dans le dossier de destination, celui-ci est remplacé. Conserver une copie des factures originales avant un test.

## Gestion des erreurs

Si une colonne obligatoire est absente, une valeur requise manque ou le montant ne peut pas être converti en nombre, le script affiche une erreur et l'inscrit dans `log.txt`. Le fichier concerné reste dans `Factures/` et n'est pas ajouté au rapport.

Si aucune facture valide n'est trouvée, le script affiche un message et ne crée pas de nouveau rapport Excel.

## Objectif pédagogique

Mettre en pratique l'automatisation de tâches administratives répétitives : extraction de données Excel, contrôle de règles métier, génération de rapports, organisation de fichiers et notifications automatiques.

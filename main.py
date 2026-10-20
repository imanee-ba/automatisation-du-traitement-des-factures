import os
import shutil#importer les fonctions pour deplacer les factures dans les dossiers prevus
import smtplib
import logging#importe le module qui enregistre les evenements et les erreurs dans un journal
import pandas as pd

from email.mime.text import MIMEText#
from dotenv import load_dotenv

load_dotenv()# Charge les paramètres du fichier .env dans les variables d’environnement.

# CONFIGURATION DES DOSSIERS

dossier = "Factures"#indique le chemin du dossier contenant les factures à traiter

dossier_traitees = os.path.join(dossier, "Traitees")
dossier_controle = os.path.join(dossier, "A_controler")

os.makedirs(dossier_traitees, exist_ok=True)#cree le dossier "Traitees" s'il n'existe pas deja
os.makedirs(dossier_controle, exist_ok=True)#cree le dossier "A_controler" s'il n'existe pas deja

# CONFIGURATION EMAIL

email_expediteur = os.getenv("EMAIL_EXPEDITEUR")#indique l'adresse email de l'expéditeur
mot_de_passe = os.getenv("MOT_DE_PASSE_APPLICATION")#indique le mot de passe de l'application pour l'expéditeur
email_responsable = os.getenv("EMAIL_RESPONSABLE")#indique l'adresse email du responsable qui recevra les notifications

# CONFIGURATION DU LOG

#definit le fichier , le niveau et le format des messages du journal
logging.basicConfig(
    filename="log.txt",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
#ecrit dans le journal que le robot vient de demarrer
logging.info("Démarrage du robot RPA")

# LISTE DES DONNÉES

#cree une liste vide pour stocker les informations extraites des factures
donnees = []

# PARCOURIR LES FACTURES

#parcourt un par un les elements du dossier factures
for fichier in os.listdir(dossier):
#conserve uniquement les fichiers Excel
    if fichier.endswith(".xlsx"):
#construit le chemin complet du fichier Excel
        chemin = os.path.join(dossier, fichier)
#enregistre une information dans le journal que le robot a detecte une facture
        try:
#enregistre une information dans le journal que le robot a detecte une facture
            logging.info(f"Facture détectée : {fichier}")

            # LIRE LA FACTURE

            facture = pd.read_excel(chemin)


            # VÉRIFIER LES COLONNES
#Enumère les colonnes obligatoires pour le traitement des factures
            colonnes_obligatoires = [
                "numero",
                "fournisseur",
                "date",
                "montant"
            ]
#verifie que toutes les colonnes obligatoires sont présentes dans le fichier Excel
            for colonne in colonnes_obligatoires:
#detecte si une colonne obligatoire est manquante et lève une exception avec un message d'erreur
                if colonne not in facture.columns:
                    #arrete le traitement et affiche un message d'erreur indiquant quelle colonne est manquante
                    raise ValueError(
                        f"La colonne {colonne} est manquante"
                    )


            # EXTRAIRE LES INFORMATIONS
#recupere les champs de la facture à partir de la première ligne du fichier Excel
            numero = facture.loc[0, "numero"]
            fournisseur = facture.loc[0, "fournisseur"]
            date = facture.loc[0, "date"]
            montant = facture.loc[0, "montant"]


            # VÉRIFIER LES INFORMATIONS
#controle les champs extraits de la facture pour s'assurer qu'ils ne sont pas vides ou manquants
            if pd.isna(numero):
                raise ValueError("Numéro de facture manquant")

            if pd.isna(fournisseur):
                raise ValueError("Fournisseur manquant")

            if pd.isna(date):
                raise ValueError("Date manquante")

            if pd.isna(montant):
                raise ValueError("Montant manquant")


            # VÉRIFIER LE MONTANT

            try:
                montant = float(montant)

            except:
                raise ValueError("Montant incorrect")


            # PRISE DE DÉCISION
#clase le statut de la facture en fonction du montant : si le montant est inférieur ou égal à 1000, le statut est "Normal", sinon il est "À contrôler"
            if montant <= 1000:
                statut = "Normal"
            else:
                statut = "À contrôler"


            # AFFICHAGE

            print("------------------------------")
            print("Facture :", fichier)
            print("Numéro :", numero)
            print("Fournisseur :", fournisseur)
            print("Date :", date)
            print("Montant :", montant)
            print("Statut :", statut)


            logging.info(
                f"{numero} - {fournisseur} - "
                f"{montant} MAD - {statut}"
            )


            # AJOUTER AU DATAFRAME
#ajoute les informations extraites de la facture à la liste "donnees" pour créer un rapport ultérieur
            donnees.append({
                "numero": numero,
                "fournisseur": fournisseur,
                "date": date,
                "montant": montant,
                "statut": statut
            })


            # ENVOYER UN EMAIL
            # SI MONTANT > 1000

            if montant > 1000:

                sujet = "Facture à contrôler"

                message = f"""
Bonjour,

Une facture nécessite un contrôle.

Numéro : {numero}
Fournisseur : {fournisseur}
Montant : {montant} MAD

Cordialement,
Robot RPA
"""
#transforme le message en objet MIMEText pour l'envoi par email, en spécifiant le format du message et l'encodage
                msg = MIMEText(
                    message,
                    "plain",
                    "utf-8"
                )

                msg["Subject"] = sujet
                msg["From"] = email_expediteur
                msg["To"] = email_responsable

                try:

                    serveur = smtplib.SMTP(
                        "smtp.gmail.com",
                        587
                    )

                    serveur.ehlo()

                    serveur.starttls()

                    serveur.ehlo()

                    serveur.login(
                        email_expediteur,
                        mot_de_passe
                    )

                    serveur.sendmail(
                        email_expediteur,
                        email_responsable,
                        msg.as_string()
                    )

                    serveur.quit()

                    print(
                        "E-mail envoyé pour la facture :",
                        numero
                    )

                    logging.info(
                        f"E-mail envoyé pour la facture {numero}"
                    )

                except Exception as erreur_email:

                    print(
                        "Erreur lors de l'envoi de l'e-mail :",
                        erreur_email
                    )

                    logging.error(
                        f"Erreur email pour {numero} : "
                        f"{erreur_email}"
                    )


            # CLASSER LA FACTURE

            if statut == "Normal":

                destination = os.path.join(
                    dossier_traitees,
                    fichier
                )

            else:

                destination = os.path.join(
                    dossier_controle,
                    fichier
                )


            # Si un fichier du même nom existe déjà
            if os.path.exists(destination):
                os.remove(destination)


            shutil.move(
                chemin,
                destination
            )


            print(
                "Facture déplacée vers :",
                destination
            )

            logging.info(
                f"{fichier} déplacée vers {destination}"
            )


        # GESTION DES ERREURS

        except Exception as erreur:

            print(
                "Erreur avec le fichier",
                fichier,
                ":",
                erreur
            )

            logging.error(
                f"Erreur avec {fichier} : {erreur}"
            )


# CRÉER LE DATAFRAME

df = pd.DataFrame(donnees)


# CRÉER LE RAPPORT EXCEL

if not df.empty:

    df.to_excel(
        "rapport_factures.xlsx",
        index=False
    )

    print("\nRapport Excel créé avec succès.")

    logging.info(
        "rapport_factures.xlsx créé avec succès"
    )

else:

    print(
        "\nAucune facture valide n'a été trouvée."
    )

    logging.warning(
        "Aucune facture valide trouvée"
    )


# FIN DU TRAITEMENT
# 

logging.info("Fin du traitement RPA")

print("\nTraitement terminé.")
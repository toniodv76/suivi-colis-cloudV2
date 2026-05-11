# Suivi Colis V2 - Cloud + historique + rapport mensuel

## Fonctions
- Ajout colis depuis téléphone / PC
- Affichage écran Raspberry en plein écran
- Statuts : A FAIRE / FAIT
- Tout est enregistré en base PostgreSQL Render
- Historique complet
- Export Excel
- Rapport mensuel envoyé par mail automatiquement le dernier jour du mois à 18h

## Pages
- `/` : page d'ajout / gestion avec code accès
- `/ecran` : affichage magasin sans code
- `/historique` : historique complet avec code accès
- `/export/monthly?code=TONCODE` : export Excel du mois courant
- `/send_report?code=TONCODE` : envoi manuel du rapport mensuel

## Variables Render importantes
- ACCESS_CODE : code d'accès, ex: h2otech
- DATABASE_URL : créée automatiquement si tu utilises render.yaml
- EMAIL_TO : ton adresse mail de réception
- EMAIL_FROM : adresse mail expéditeur
- SMTP_HOST : serveur SMTP, ex: smtp.gmail.com
- SMTP_PORT : 587
- SMTP_USER : identifiant mail
- SMTP_PASSWORD : mot de passe application SMTP

## Gmail
Pour Gmail, il faut créer un mot de passe d'application Google, pas ton mot de passe normal.

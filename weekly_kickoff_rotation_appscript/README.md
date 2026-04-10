# Weekly Kickoff Rotation (Apps Script)

## Contenu
- `CONFIG.gs`
- `MAIN.gs`
- `TOOLS.gs`
- `UTILS.gs`
- `WEEKLY_UI.gs`
- `AUDIT.gs`
- `gsheet_template/*.csv` (template Google Sheet)

## Mise en place rapide
1. Créer un Google Sheet vide.
2. Importer chaque CSV du dossier `gsheet_template` dans un onglet portant le meme nom.
3. Ouvrir Apps Script depuis ce Google Sheet et coller les 5 fichiers `.gs`.
4. Mettre `WEEKLY_ANNOUNCEMENT_CHANNEL_ID`, `LIEN_VISIO_WEEKLY`, `WEEKLY_SLIDES_FOLDER_ID`, `LIEN_REMOTE`, `MANAGERS_SLACK_GROUP_IDS_SALES`, `MANAGERS_SLACK_GROUP_IDS_CARE` dans l'onglet `CONFIG`.
4bis. Configurer le webhook Slack dans `Script Properties` avec la cle `SLACK_WEBHOOK_URL` (ou executer `setSlackWebhookUrl("https://hooks.slack.com/services/...")`).
4ter. `SLACK_POST_CHANNEL_ID` est optionnel si le webhook est deja lie au bon canal.
4bis. Mettre `LOG_LEVEL=ERROR` pour reduire les logs (ou `INFO`/`DEBUG`).
5. Lancer `install()` une fois.
6. Utiliser le menu `IT-INDY` > `Dry Run Weekly` pour verifier le message.
7. (Optionnel) Deployer en Web App et ouvrir `...?page=start` pour le flow "Start Weekly" (Meet/Slide/Remote).
8. (Optionnel) Partager le formulaire d'absence via la Web App: `...?page=absence`.

## Si le script est standalone (non lie au Sheet)
1. Recuperer l'ID du Google Sheet cible (dans l'URL).
2. Dans Apps Script, executer `setSpreadsheetId("VOTRE_SPREADSHEET_ID")`.
3. Puis executer `install()`.

## Notes
- Le script alterne automatiquement semaine `care` / `sales` depuis `ROTATION_START_DATE`.
- `care` utilise l'onglet `ROTATION_CARE` (colonne `slackUserId`).
- `sales` utilise l'onglet `ROTATION_SALES` (colonne `slackUserId`).
- `start` (personne qui lance le weekly) utilise l'onglet `ROTATION_START` (colonne `slackUserId`).
- La rotation est circulaire par equipe: 1 -> 2 -> 3 -> 1, en ignorant les absents.
- Regle fallback: en semaine Sales, si aucun Sales n'est disponible, le script designe la prochaine personne disponible cote Care.
- Le backup est automatique: prochain disponible dans la rotation.
- Pour `ROTATION_START`, pas de backup: une seule personne est affichee dans le message.
- Si `WEEKLY_SLIDES_FOLDER_ID` est renseigne, le message inclut le lien du Google Slide dont le nom contient la date du weekly (`yyyyMMdd`, ex: `20260223`).
- Si `LIEN_VISIO_WEEKLY` est renseigne, le message inclut le lien de visio.
- Si `LIEN_REMOTE` est renseigne, il est ajoute au message Slack.
- Si la date est dans `JOUR_OFF_INDY` ou est un jour ferie FR via `JOUR_FERIES_FRANCAIS_API`, le message n'est pas envoye et le tour n'est pas consomme.
- Format `JOUR_OFF_INDY`: `YYYY-MM-DD,YYYY-MM-DD` (quelques dates par an).
- Formulaire absence: `getAbsenceWebAppUrl()` renvoie le lien direct pour declaration d'absence.
- La colonne `active` accepte `true/false` (insensible a la casse, ex: `TRUE` fonctionne).
- Absences: MVP via onglet `ABSENCES`.
- Lucca (optionnel): utiliser les `Script Properties` (pas l'onglet `CONFIG`):
  - `LUCCA_ENABLED=true`
  - `LUCCA_BASE_URL=https://...`
  - `LUCCA_API_TOKEN=...`

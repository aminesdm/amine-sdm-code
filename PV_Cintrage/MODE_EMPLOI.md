# PV de Cintrage – Remplissage automatique depuis Excel

## Fichiers
- `Donnees_PV.xlsx` : tableau à remplir, **une ligne = un PV**
  (colonnes : Ouvrage, Plan, Designation, Rapport, RFI).
- `PV_Cintrage_Modele.docx` : modèle Word du PV, relié aux colonnes Excel (publipostage).

> Ne pas renommer les en-têtes de la ligne 1 d'Excel : ce sont les noms des champs du Word.

## Configuration (une seule fois)
1. Placer les deux fichiers dans le même dossier.
2. Ouvrir `PV_Cintrage_Modele.docx` dans Word.
3. Onglet **Publipostage** → **Sélection des destinataires** → **Utiliser une liste existante…**
4. Choisir `Donnees_PV.xlsx`, puis la feuille **Donnees$** (cocher « La première ligne contient les en-têtes »).
5. Enregistrer le Word. La liaison est mémorisée : aux prochaines ouvertures, Word demande
   « exécuter la commande SQL… » → répondre **Oui**.

## Utilisation
1. Saisir les nouvelles lignes dans Excel, enregistrer et fermer Excel.
2. Ouvrir le Word → **Aperçu des résultats** pour vérifier (flèches pour passer d'un PV à l'autre).
3. **Terminer et fusionner** → **Modifier des documents individuels** → *Tous*
   (ou *De … à …* pour ne générer que les nouveaux PV) → un document avec tous les PV remplis,
   à enregistrer ou imprimer.

## Remarque
Les colonnes Excel sont au format Texte pour conserver les zéros (`0001`).

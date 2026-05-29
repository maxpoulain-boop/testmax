# Référentiel d'affectation commerciale — HomeServe Énergies Services

Projet Google Sheets de gestion des affectations commerciales par filiale, commune et produit, avec notifications email automatiques.

## État actuel du projet

Le fichier `Referentiel_HomeServe_ES.xlsx` (ou son équivalent Google Sheets) contient :

### Feuilles principales

| Feuille | Rôle |
|---|---|
| `ACCUEIL` | Page de présentation + documentation syntaxe CP |
| `RECHERCHE` | Outil de recherche d'affectation (Filiale + CP + Ville → commerciaux disponibles avec produits) |
| `VUE_REGIONALE` | Pilotage : couverture par région, transferts en cours, top commerciaux |
| `AFFECTATIONS_COMMUNES` | ~9300 lignes : Filiale × Commune × Commercial × Rôle (Principal/Co-affecté) |
| `PRODUITS_COMMERCIAUX` | ~280 lignes : quels produits chaque commercial gère |
| `EXCEPTIONS_PRODUITS` | Cas où un produit spécifique sur une commune est traité par un autre commercial |
| `TRANSFERTS` | Remplacements temporaires/définitifs (absences, départs) avec dates de début/fin |
| `CONTACTS` | Destinataires des emails de notification (mode TEST/PROD) |
| `PARAMETRES` | Listes de référence (filiales, produits, régions, motifs, rôles, etc.) |

### Feuilles techniques (cachées ou grises)

| Feuille | Rôle |
|---|---|
| `_COMMERCIAUX_PAR_FILIALE_H` | Référentiel commerciaux par filiale (horizontal) |
| `_DDL_TRANSFERTS` | 200 lignes de formules pour les listes déroulantes en cascade |
| `DPT_SOURCE` | ~35 000 communes INSEE (référence territoriale) |

## Filiales actuelles (15)

Après fusion ID Energies + PH Energies + Conviflamme + Concept Habitat → **VIK Energies**, restent :

Aujard, Cerise Energies, Chauffage du Nord, DEC Energies, EGS Energies, GD Energies, GEODIS, JCM Confort, Lepretre Energies, Mure Energies, Prigent Abiven, Roussin Energies, SBF Energies, SMT Energies, **VIK Energies**

## Produits (24)

Liste officielle dans PARAMETRES!C2:C25 :
Chauffage solaire, Poêle à granulés, Pompe à chaleur Air/Eau, Système solaire combiné, Pompe à chaleur Air/Air, Chauffe-eau solaire, Chauffe-eau thermodynamique, Isolation des combles perdus, Isolation des planchers bas, Isolation des rampants, Isolation thermique extérieure, Isolation thermique intérieure, Panneaux photovoltaïques, Audit, Rénovation d'ampleur, Ventilation, Chaudière gaz, Chauffe-eau gaz, Chaudière granulés, Chaudière bois, Chaudière électrique, Chauffe-eau électrique, Poêle à bois, Chaudière fioul.

## Logique métier clé

### Système de rôles
- **Principal** : commercial titulaire de la commune
- **Co-affecté** : commercial partageant la zone (cas Cerise, VIK, certaines communes Conviflamme historiques)

### Système de transferts
- Une ligne TRANSFERTS = un remplacement temporaire ou définitif
- Colonne `Actif` doit être `Oui` pour que le transfert prenne effet
- Colonne `CP` accepte syntaxes multiples : vide / `Tous` / `33000` / `33000, 33100` / `33000-33999` / mix `33000-33500, 35000` / sauts de ligne (copier-coller Excel)
- Le matching CP utilise la fonction custom Apps Script `CP_MATCH`
- La formule de RECHERCHE.statut détecte automatiquement si un commercial est en transfert à la date de référence

### Notifications email
- Déclenchées par le passage Actif = Oui
- Email immédiat à la validation
- Rappel automatique 3 jours ouvrés avant la date de début (déclencheur quotidien à 7h)
- Email d'annulation si Actif repasse à Non
- Mode TEST (1 email) / PROD (3 emails) basculable dans la feuille CONTACTS

## Scripts Apps Script installés

### 1. `notifications_transferts.gs`
Le principal. Contient :
- `onEdit` (déclencheur installable) : email à l'activation
- `verifierRappelsQuotidiens` (déclencheur temporel quotidien) : email J-3 ouvrés
- `CP_MATCH(cp, expression)` : custom function utilisée dans les formules RECHERCHE
- `setupTriggers` : à lancer 1× pour installer les déclencheurs
- `testEnvoiEmail`, `afficherDestinataires`, `reinitialiserHistorique` : utilitaires

### 2. `installer_validations.gs`
À relancer si les listes déroulantes sautent (après réimport Excel → Sheets) :
- `installerToutesLesValidations` : recrée toutes les validations nativement dans Google Sheets

### 3. `notes_helper.gs`
Ajoute les info-bulles au survol (sur colonnes B, C, I de TRANSFERTS) :
- `ajouterNotes` / `retirerNotes`

## Points d'attention récurrents

1. **Listes déroulantes Excel ne survivent pas à l'import Google Sheets** → toujours utiliser `installerToutesLesValidations` après import
2. **Les info-bulles Excel ne sont pas importées** non plus → utiliser `ajouterNotes`
3. **Toujours vérifier les validations après modification** du fichier (régression fréquente)
4. **Formules matricielles** : Google Sheets requiert le pattern `MATCH(1, INDEX((cond1)*(cond2), 0), 0)` au lieu de `MATCH(1, (cond1)*(cond2), 0)` qui exige CSE

## Évolutions envisagées (à voir avec l'utilisateur)

- Rotation équitable / round-robin entre co-affectés pour les RDV
- Dashboard de charge par commercial
- Pré-remplir les "A AFFECTER" restants (notamment VIK, Cerise complète, etc.)
- Compléter les commerciaux des filiales encore vides : Aujard, Chauffage du Nord, EGS, GD, JCM Confort, Lepretre, Mure, Prigent Abiven, Roussin, SBF, SMT, GEODIS

## Fichiers du projet

- `Referentiel_HomeServe_ES.xlsx` : fichier de référence (à ne pas modifier directement)
- `notifications_transferts.gs` : script principal
- `installer_validations.gs` : script de récupération validations
- `notes_helper.gs` : script info-bulles
- `README.md` : ce fichier

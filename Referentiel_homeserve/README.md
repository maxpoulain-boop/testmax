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
- Mode TEST (1 email en C9) / PROD (nombre **illimité** d'emails — colonne C à partir de C12, une adresse par ligne) basculable dans la feuille CONTACTS

## Installation (après import Excel → Sheets)

**Séquence complète dans l'ordre :**

1. Ouvrir le Google Sheet
2. `Extensions → Apps Script`
3. Créer/remplacer les 4 fichiers `.gs` avec le contenu de ce dépôt
4. Sauvegarder → fermer et rouvrir le Sheet
5. Le menu **⚙️ Transferts** apparaît dans le ruban
6. **⚙️ Transferts → Installer les déclencheurs**
7. **⚙️ Transferts → Installer les listes déroulantes** (20-30s)
8. **⚙️ Transferts → Ajouter les notes d'aide**
9. Renseigner l'email dans CONTACTS (cellule C9 pour mode TEST)
10. **⚙️ Transferts → Tester l'envoi d'email**
11. **⚙️ Transferts → Diagnostiquer la configuration** → tout doit être ✓
12. (Optionnel) **⚙️ Transferts → Protéger les feuilles de référence**

## Scripts Apps Script installés

### 1. `notifications_transferts.gs` (principal)
- `onOpen` : crée le menu **⚙️ Transferts** dans le ruban
- `gererModificationActif` (déclencheur installable) : email à l'activation d'un transfert
- `verifierRappelsQuotidiens` (déclencheur quotidien 7h) : email J-3 ouvrés
- `CP_MATCH(cp, expression)` : custom function pour les formules RECHERCHE
- `setupTriggers` : installe les déclencheurs (via menu)
- `diagnostiquer` : vérifie feuilles, emails, déclencheurs
- `purgerProprietesObsoletes` : nettoie les flags d'envoi orphelins (auto, chaque jour)
- `protegerFeuillesReference` : protège DPT_SOURCE, _DDL_TRANSFERTS, _COMMERCIAUX_PAR_FILIALE_H
- `testEnvoiEmail`, `afficherDestinataires`, `reinitialiserHistorique`

> **v5** : les flags d'envoi orphelins (suite à suppression de lignes ou
> changement de dates) sont purgés automatiquement chaque jour. Les listes de
> référence sont pilotées par des **plages nommées** (voir ci-dessous) :
> agrandir une liste ne nécessite plus de modifier le code.

> **v4** : le handler de modification a été renommé (`gererModificationActif`)
> pour éviter une double exécution. **Après mise à jour du code, relancer une
> fois ⚙️ Transferts → Installer les déclencheurs.**
>
> **v3** : le tracking des emails est basé sur le contenu du transfert
> (filiale + commercial + date) et non plus sur le numéro de ligne.
> Les insertions/suppressions de lignes ne créent plus de doublons ni de silences.

### 2. `installer_validations.gs`
À relancer si les listes déroulantes sautent (après réimport Excel → Sheets) :
- `installerToutesLesValidations` : recrée toutes les validations nativement
- `installerPlagesNommees` : crée les plages nommées `PARAM_*` à partir de l'objet
  `PARAM` (source unique des adresses). Appelé automatiquement par
  `installerToutesLesValidations`.

**Plages nommées créées** (toutes pointent vers l'onglet PARAMETRES) :

| Plage nommée | Adresse | Contenu |
|---|---|---|
| `PARAM_REGIONS` | colonne A (dès A2) | Régions |
| `PARAM_FILIALES` | colonne B (dès B2) | Filiales |
| `PARAM_PRODUITS` | colonne C (dès C2) | Produits |
| `PARAM_ACTIF` | colonne D (dès D2) | Oui / Non |
| `PARAM_MOTIFS` | colonne E (dès E2) | Motifs |
| `PARAM_PRIORITES` | colonne H (dès H2) | Priorités |
| `PARAM_ROLES` | colonne I (dès I2) | Rôles |
| `PARAM_PRODUITS_T` | colonne J (dès J2) | Produits + "Tous" |

> **Plages auto-détectées (v6)** : chaque plage est calculée de la ligne 2
> jusqu'à la dernière cellule non-vide de sa colonne. Pour agrandir une liste
> (ex : ajouter une filiale), il suffit de l'écrire dans PARAMETRES **puis** de
> relancer "Installer les listes déroulantes" — plus besoin de toucher au code
> ni d'étendre la plage nommée à la main.

### 3. `notes_helper.gs`
Info-bulles au survol (TRANSFERTS colonnes B, C, I + RECHERCHE) :
- `ajouterNotes` / `retirerNotes`

### 4. `generer_affectations.gs`
Génération en masse des affectations par règles (évite la saisie commune par commune) :
- `creerFeuilleSaisie` : crée l'onglet `SAISIE_SECTEURS` (1 ligne par commercial :
  Filiale, Commercial, Type zone `CP`/`COMMUNES`, Zone, Produits)
- `genererAffectations` : déplie les règles contre `DPT_SOURCE` et remplit
  `AFFECTATIONS_COMMUNES` + `PRODUITS_COMMERCIAUX`. Rôle déduit (1 commercial =
  Principal, 2+ = Co-affecté). Colonne `Origine` = `généré` → les lignes manuelles
  sont préservées à chaque régénération. Réutilise `CP_MATCH`. Rapport de conflits.

## Points d'attention récurrents

1. **Listes déroulantes Excel ne survivent pas à l'import Google Sheets** → toujours relancer `installerToutesLesValidations` après import
2. **Les info-bulles Excel ne sont pas importées** → relancer `ajouterNotes`
3. **Toujours vérifier après modification** via **⚙️ Transferts → Diagnostiquer**
4. **Formules matricielles** : Google Sheets requiert `MATCH(1, INDEX((cond1)*(cond2), 0), 0)` — le pattern CSE `MATCH(1, (cond1)*(cond2), 0)` ne fonctionne pas
5. **DPT_SOURCE (35 000 lignes)** : principale source de lenteur — protéger la feuille évite les recalculs accidentels

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
- `generer_affectations.gs` : génération des affectations par règles
- `docs/guide_administrateur.html` : guide propriétaire (imprimable en PDF)
- `docs/guide_utilisateur_transferts.html` : guide gestionnaires de transferts
- `README.md` : ce fichier

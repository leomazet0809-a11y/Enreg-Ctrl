# Enregistrements Contrôles

Application web de saisie des **rapports de contrôle d'entrée** : gammes de contrôle (GAMSERIE), valeurs par article (TS), matrice article / fournisseur, révisions, documents PDF, cartes de contrôle et analyses des dérives.

Tout le projet est enregistré dans un **classeur Excel (.xlsx)** dont vous choisissez l'emplacement.

## Démarrer

Deux façons d'ouvrir l'application, avec **Google Chrome** ou **Microsoft Edge** :

- **En ligne** : https://leomazet0809-a11y.github.io/Enreg-Ctrl/ (le classeur reste sur votre PC, rien n'est envoyé sur internet). Dans Edge, menu *… → Applications → Installer ce site en tant qu'application* pour l'avoir comme un logiciel.
- **Depuis le dossier téléchargé** (bouton vert **Code → Download ZIP** sur GitHub, puis clic droit → *Extraire tout…*) : double-cliquez sur **`Lancer l'application.bat`**. Une petite fenêtre noire s'ouvre (laissez-la ouverte, vous pouvez la réduire) et l'application s'ouvre dans le navigateur à l'adresse `http://localhost:8765`. Fermez la fenêtre noire pour arrêter.

> N'ouvrez pas directement `index.html` par double-clic : ouverte depuis un fichier (adresse `C:/…/index.html`), la page ne peut pas enregistrer dans le classeur (le navigateur ne répond pas à la demande d'écriture).

À l'accueil, choisissez :
- **Créer un nouveau classeur…** : indiquez le dossier et le nom du fichier Excel (par exemple sur un lecteur réseau partagé) ;
- **Ouvrir un classeur existant…** : sélectionnez un classeur déjà créé par l'application ;
- **Reprendre « … »** : rouvre directement le dernier classeur utilisé sur ce poste.

> Le navigateur demande l'autorisation de modifier le fichier : acceptez-la.

## Enregistrer

- Cliquez sur l'icône **disquette** en haut de page,
- ou appuyez sur **Ctrl + S**.

**Ctrl + Maj + S** enregistre sous un autre nom ou dans un autre dossier.

En haut à droite, l'indicateur du classeur affiche son nom et son état :

| Pastille | Signification |
| --- | --- |
| verte — *Enregistré à 14:32* | tout est écrit dans le classeur |
| orange — *Modifications non enregistrées* | pensez à enregistrer |

Tant que le contenu de l'application n'est pas enregistré dans le classeur :

- en quittant un rapport (croix en haut à droite), l'application demande si vous êtes certain de vouloir fermer : *Annuler*, *Fermer sans enregistrer* ou *Enregistrer et fermer* ;
- en fermant l'onglet ou la page du navigateur, le navigateur demande confirmation avant de quitter.

Si l'enregistrement échoue, vérifiez que le classeur **n'est pas ouvert dans Excel** au même moment, puis recommencez.

## Contenu du classeur Excel

| Onglet | Contenu |
| --- | --- |
| Rapports | une ligne par rapport enregistré (article, lot, quantités, conformes / non conformes, décision…) |
| Mesures | toutes les valeurs relevées, pièce par pièce, avec leur résultat |
| GAMSERIE | les gammes de contrôle par famille |
| TS | les valeurs attendues par article (nominal, tolérances, type de spécification) |
| Matrice | la matrice article / fournisseur et les références documentaires |
| Journal | l'historique des révisions |
| _Donnees *(masqué)* | les données de l'application ; **ne pas modifier** |

Un onglet **Mode d'emploi** rappelle ces règles dans le classeur. Vos propres onglets supplémentaires sont conservés à l'enregistrement (sans leur mise en forme).

## Ajouter ou modifier des données dans Excel

Les onglets **GAMSERIE**, **TS** et **Matrice** se modifient directement dans Excel : ajouter des lignes (nouvel article, nouvelle spécification…), corriger une valeur, une tolérance, un fournisseur, etc.

1. Ouvrez le classeur dans Excel, faites vos modifications, **enregistrez et fermez** le classeur.
2. Les données du classeur sont rechargées dans l'application :
   - à l'ouverture de l'application ou au rafraîchissement de la page : le dernier classeur est rechargé directement si le navigateur a gardé l'autorisation d'y accéder (sinon, cliquez sur **Reprendre « … »** à l'accueil) ;
   - quand vous revenez dans l'application (changement de fenêtre ou d'onglet) ;
   - à chaque enregistrement : le classeur est relu juste avant d'écrire, les modifications Excel sont intégrées puis réécrites avec celles de l'application. Si le classeur était bloqué par Excel, il est relu à nouveau dès qu'il se libère.
3. Un court message en bas de l'écran indique les onglets repris (plus de fenêtre à fermer). Enregistrez (**Ctrl + S**) pour mettre le classeur à jour.

Bon à savoir :

- Une gamme (GAMSERIE) ou des valeurs (TS) modifiées dans Excel ouvrent une **révision en brouillon** : validez-la dans *Référentiel → GAMSERIE / TS* pour qu'elle serve aux nouveaux rapports.
- **Identifiant des points (colonne ID)** : chaque ligne de GAMSERIE a un identifiant unique de 6 caractères (ex. `X5JVCH`), créé par l'application et qui ne change jamais. C'est lui qui relie la gamme, les valeurs TS, les rapports et les analyses. Le **Code** (A, B, C…) n'est qu'une étiquette : il peut être vide, valoir « N/A », être en double ou changer d'une révision à l'autre.
  - Ne modifiez pas la colonne ID. Une ligne ajoutée dans GAMSERIE sans ID en reçoit un ; une ligne copiée (ID en double) en reçoit un nouveau.
  - Dans TS, une ligne est rattachée au point par son **ID**. Une ligne sans ID est rattachée par son code, seulement si ce code est unique dans la gamme de la famille.
  - Les classeurs créés avant les identifiants sont repris automatiquement à l'ouverture (même résultat sur tous les postes) : enregistrez ensuite le classeur.
- Dans TS, modifiez **Nominal**, **Tol -**, **Tol +** et **Type spéc.** (Générale / Individuelle). Les colonnes Famille, Code, Type saisie, Révision, Date et Statut sont recalculées par l'application. Mini. et Maxi. (nominal + tolérances) ne figurent plus dans TS : ils sont calculés et affichés dans les rapports, les analyses des dérives et les cartes de contrôle.
- Les onglets **Rapports**, **Mesures** et **Journal** sont en lecture seule : les rapports se saisissent dans l'application.
- Si vous avez aussi des modifications non enregistrées dans l'application, elle vous demande quoi faire : *Intégrer les modifications Excel* (vos rapports en cours sont conservés) ou *Garder la version de l'application*.
- Un classeur Excel préparé à la main, avec des onglets nommés GAMSERIE, TS et/ou Matrice, peut aussi être ouvert : ses données sont reprises comme bases de départ.

## Créer un article et vérifier la cohérence

Tous les articles d'une famille suivent la même gamme (points de contrôle, moyens et méthodes) ; seules les valeurs changent d'un article à l'autre, dans TS.

- **Créer un article** (*Référentiel → TS* ou *Référentiel → Matrice*, bouton *Ajouter un article*) : indiquez sa référence, son fournisseur et sa famille. Ses lignes TS sont préparées automatiquement à partir de la gamme de la famille, avec le type de spécification habituel de la famille : il ne reste qu'à saisir les valeurs. Vous pouvez aussi partir des valeurs d'un article existant de la même famille.
- Changer la famille d'un article dans la Matrice (article en révision brouillon) prépare de la même façon les points manquants.
- **Référentiel → Cohérence** liste ce qui empêche un article d'avoir un rapport complet, avec un bouton pour corriger :
  - article sans famille, ou famille absente de GAMSERIE ;
  - point de la gamme sans valeur dans TS (saisissez « - » si le point ne s'applique pas à l'article) ;
  - valeur TS sur un code qui n'existe pas dans la gamme de la famille.

## Travailler à plusieurs (plusieurs PC, un seul classeur)

1. Placez le classeur dans un **dossier réseau partagé** accessible depuis chaque PC (lecteur réseau de l'entreprise). Évitez un dossier synchronisé OneDrive / SharePoint : en cas d'enregistrements simultanés, OneDrive crée des copies en conflit.
2. Sur chaque PC, ouvrez le classeur dans l'application. Au premier usage, elle propose **« Autoriser l'accès au dossier »** : acceptez et choisissez le dossier qui contient le classeur (une seule fois par PC ; option aussi disponible dans *Paramètres → Sauvegarde*).
3. Chacun enregistre normalement (**Ctrl + S**).

Ce qui se passe à chaque enregistrement :

- Si un autre utilisateur est en train d'enregistrer, le bandeau **« Classeur en cours d'enregistrement par un autre utilisateur »** s'affiche ; votre enregistrement part automatiquement dès que le classeur est disponible.
- Si le classeur est ouvert dans Excel sur un poste, le bandeau **« Classeur ouvert dans Excel »** s'affiche ; l'enregistrement part dès que le classeur est fermé. Si personne ne l'a réellement ouvert (fichier « ~$… » resté après un arrêt brutal d'Excel), cliquez sur *Enregistrer quand même*.
- Avant d'écrire, l'application relit le classeur et **fusionne** ce que les autres ont enregistré : rapports ajoutés, modifiés ou supprimés, gammes, valeurs, matrice, révisions, documents. Si un même élément a été modifié sur deux postes, votre version est gardée (pour un rapport, la plus récente).
- Quand vous revenez dans l'application (changement de fenêtre ou d'onglet), elle relit aussi le classeur et affiche les rapports enregistrés entre-temps par les autres.
- Le rapport en cours de saisie (non encore complet) reste propre à chaque PC.

Conseil : confiez la modification des gammes et des valeurs (Référentiel, ou onglets Excel) à une seule personne, et laissez les contrôleurs saisir les rapports.

## Navigation

- L'écran principal est la liste des **Rapports enregistrés** : cliquez sur un rapport pour l'ouvrir.
- Le bouton **Nouveau rapport** (document +) est à gauche du titre « Rapports enregistrés » : un nouveau rapport s'ouvre uniquement depuis cet écran. Sa **Date** est la date d'ouverture du rapport, renseignée automatiquement (non modifiable).
- **Référentiel** : GAMSERIE, TS, Matrice, Cohérence, Révisions et Documents. Une pastille orange sur son icône indique le nombre d'incohérences.
  - Le bouton **Actualiser** (à côté du titre) recharge toutes les données du classeur Excel dans l'application, **sans rien écrire dans le classeur**. Si des modifications de l'application ne sont pas encore enregistrées, il demande de les garder (les données du classeur sont intégrées autour) ou de les abandonner pour recharger exactement le classeur.
- **Paramètres** (roue crantée) : s'ouvre dans une petite fenêtre en haut à droite, par-dessus l'écran en cours, sans le quitter. Onglets Sauvegarde (classeur, travail à plusieurs, copies) et Affichage (mode clair / sombre, taille du texte). La croix, la touche Échap ou un clic en dehors la referment.
- Le ruban en haut regroupe les icônes **Analyses des dérives** (courbe), **Référentiel** (livre), **Paramètres** (roue crantée), **Enregistrer** (disquette). Depuis un rapport, les mêmes vérifications qu'à sa fermeture s'appliquent.
- Dans un rapport, une pastille rouge indique le nombre de pièces non conformes à droite de chaque spécification, et le nombre de spécifications non conformes à côté du numéro de chaque pièce.
- Les valeurs relevées s'écrivent avec une virgule : un point saisi est remplacé par une virgule dès que la cellule est validée.
- **Analyses des dérives** et **Référentiel** s'ouvrent dans une grande fenêtre flottante au-dessus du menu principal. Tant qu'elle est ouverte, seuls **Enregistrer** et **Paramètres** restent accessibles dans le ruban. La croix, la touche Échap ou un clic sur le fond grisé la referment et ramènent au menu principal.
- La croix en haut à droite d'un rapport ramène à la liste.
- Un rapport commencé mais incomplet (article, lot ou quantité échantillon manquant) apparaît dans un bandeau « Rapport en cours » avec un bouton **Reprendre**.
- Chaque bouton affiche une info-bulle qui précise son action après une seconde de survol.
- L'affichage s'adapte à la taille de l'écran (grand écran, portable, tablette, téléphone).

## Taille du texte

*Paramètres → Affichage* : boutons **A−** et **A+** (de 80 % à 150 %) et retour à la taille normale. Le réglage est mémorisé sur le poste.

## Mode clair / sombre

*Paramètres → Affichage* : choisissez **Clair** ou **Sombre**. Le choix est mémorisé sur le poste ; au premier lancement, l'application suit le réglage de Windows.

## Mettre en ligne (facultatif)

L'application peut aussi être publiée avec **GitHub Pages** (Settings → Pages → branche contenant `index.html`). Elle s'ouvre alors depuis une adresse web, toujours dans Chrome ou Edge, et enregistre de la même façon dans le classeur de votre choix sur le poste.

## Pour information technique

- Application statique : `index.html`, `Lancer l'application.bat` + `outils/serveur-local.ps1` (petit serveur web local en PowerShell, sans installation), `css/app.css`, `js/app.js` (logique métier), `js/classeur.js` (lecture / écriture du classeur), `js/partage.js` (travail à plusieurs), `js/infobulle.js` (info-bulles).
- Lecture et écriture Excel : bibliothèque [SheetJS](https://sheetjs.com) incluse dans `lib/` (fonctionne sans connexion internet).
- Accès au fichier : API *File System Access* de Chrome / Edge. Dans un autre navigateur, chaque enregistrement télécharge une copie du classeur.

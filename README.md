# Enregistrements Contrôles

Application web de saisie des **rapports de contrôle d'entrée** : gammes de contrôle (GAMSERIE), valeurs par article (TS), matrice article / fournisseur, révisions, documents PDF, cartes de contrôle et analyses des dérives.

Tout le projet est enregistré dans un **classeur Excel (.xlsx)** dont vous choisissez l'emplacement.

## Démarrer

1. Téléchargez le dossier de l'application (bouton vert **Code → Download ZIP** sur GitHub), puis décompressez-le.
2. Ouvrez le fichier **`index.html`** avec **Google Chrome** ou **Microsoft Edge** (double-clic, ou clic droit → *Ouvrir avec*).
3. À l'accueil, choisissez :
   - **Créer un nouveau classeur…** : indiquez le dossier et le nom du fichier Excel (par exemple sur un lecteur réseau partagé) ;
   - **Ouvrir un classeur existant…** : sélectionnez un classeur déjà créé par l'application ;
   - **Reprendre « … »** : rouvre directement le dernier classeur utilisé sur ce poste.

> Le navigateur demande l'autorisation de modifier le fichier : acceptez-la.

## Enregistrer

- Cliquez sur le bouton **Enregistrer** (icône disquette) en haut de page,
- ou appuyez sur **Ctrl + S**.

**Ctrl + Maj + S** enregistre sous un autre nom ou dans un autre dossier.

En haut à droite, l'indicateur du classeur affiche son nom et son état :

| Pastille | Signification |
| --- | --- |
| verte — *Enregistré à 14:32* | tout est écrit dans le classeur |
| orange — *Modifications non enregistrées* | pensez à enregistrer |

Si vous fermez l'onglet avec des modifications non enregistrées, le navigateur vous prévient.

Si l'enregistrement échoue, vérifiez que le classeur **n'est pas ouvert dans Excel** au même moment, puis recommencez.

## Contenu du classeur Excel

| Onglet | Contenu |
| --- | --- |
| Rapports | une ligne par rapport enregistré (article, lot, quantités, conformes / non conformes, décision…) |
| Mesures | toutes les valeurs relevées, pièce par pièce, avec leur résultat |
| GAMSERIE | les gammes de contrôle par famille |
| TS | les valeurs attendues par article (nominal, tolérances, mini, maxi) |
| Matrice | la matrice article / fournisseur et les références documentaires |
| Journal | l'historique des révisions |
| _Donnees *(masqué)* | les données de l'application ; **ne pas modifier** |

Un onglet **Mode d'emploi** rappelle ces règles dans le classeur. Vos propres onglets supplémentaires sont conservés à l'enregistrement (sans leur mise en forme).

## Ajouter ou modifier des données dans Excel

Les onglets **GAMSERIE**, **TS** et **Matrice** se modifient directement dans Excel : ajouter des lignes (nouvel article, nouvelle spécification…), corriger une valeur, une tolérance, un fournisseur, etc.

1. Ouvrez le classeur dans Excel, faites vos modifications, **enregistrez et fermez** le classeur.
2. Revenez dans l'application :
   - si l'application était déjà ouverte, elle détecte la modification et reprend les données toute seule ;
   - sinon, ouvrez le classeur depuis l'écran d'accueil.
3. Un message indique les onglets repris. Enregistrez (**Ctrl + S**) pour mettre le classeur à jour.

Bon à savoir :

- Une gamme (GAMSERIE) ou des valeurs (TS) modifiées dans Excel ouvrent une **révision en brouillon** : validez-la dans *Paramètres → GAMSERIE / TS* pour qu'elle serve aux nouveaux rapports.
- Dans TS, modifiez **Nominal**, **Tol -**, **Tol +** et **Type spéc.** (Générale / Individuelle). Les colonnes Mini., Maxi., Spécification, Type saisie, Révision, Date et Statut sont recalculées par l'application.
- Les onglets **Rapports**, **Mesures** et **Journal** sont en lecture seule : les rapports se saisissent dans l'application.
- Si vous avez aussi des modifications non enregistrées dans l'application, elle vous demande quoi faire : *Intégrer les modifications Excel* (vos rapports en cours sont conservés) ou *Garder la version de l'application*.
- Un classeur Excel préparé à la main, avec des onglets nommés GAMSERIE, TS et/ou Matrice, peut aussi être ouvert : ses données sont reprises comme bases de départ.

## Navigation

- L'écran principal est la liste des **Rapports enregistrés** : cliquez sur un rapport pour l'ouvrir, ou sur **Nouveau rapport**.
- Les boutons **Analyses des dérives** et **Paramètres** de cet écran ouvrent les écrans correspondants ; la croix en haut à droite de chaque écran ramène à la liste.
- Un rapport commencé mais incomplet (article, lot ou quantité échantillon manquant) apparaît dans un bandeau « Rapport en cours » avec un bouton **Reprendre**.
- L'affichage s'adapte à la taille de l'écran (grand écran, portable, tablette, téléphone).

## Mode clair / sombre

Le bouton lune / soleil en haut à droite bascule entre le mode clair et le mode sombre. Le choix est mémorisé sur le poste.

## Mettre en ligne (facultatif)

L'application peut aussi être publiée avec **GitHub Pages** (Settings → Pages → branche contenant `index.html`). Elle s'ouvre alors depuis une adresse web, toujours dans Chrome ou Edge, et enregistre de la même façon dans le classeur de votre choix sur le poste.

## Pour information technique

- Application statique : `index.html`, `css/app.css`, `js/app.js` (logique métier), `js/classeur.js` (lecture / écriture du classeur).
- Lecture et écriture Excel : bibliothèque [SheetJS](https://sheetjs.com) incluse dans `lib/` (fonctionne sans connexion internet).
- Accès au fichier : API *File System Access* de Chrome / Edge. Dans un autre navigateur, chaque enregistrement télécharge une copie du classeur.

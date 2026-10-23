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

Conseil : confiez la modification des gammes et des valeurs (Paramètres, ou onglets Excel) à une seule personne, et laissez les contrôleurs saisir les rapports.

## Navigation

- L'écran principal est la liste des **Rapports enregistrés** : cliquez sur un rapport pour l'ouvrir, ou sur **Nouveau rapport**.
- Dans un rapport, une pastille rouge indique le nombre de pièces non conformes à droite de chaque spécification, et le nombre de spécifications non conformes à côté du numéro de chaque pièce.
- Les valeurs relevées s'écrivent avec une virgule : un point saisi est remplacé par une virgule dès que la cellule est validée.
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

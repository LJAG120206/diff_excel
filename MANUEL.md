# Manuel utilisateur — Diff Excel

Diff Excel compare deux versions d’un classeur dans le navigateur, comme un `git diff`. Le calcul reste local : aucun fichier n’est envoyé sur un serveur.

## Prérequis

- **Windows**
- **Node.js en version LTS** (Long Term Support : la version stable recommandée, pas la « Current »)

Node.js fournit deux commandes utilisées par le projet :

- `node` — le moteur JavaScript
- `npm` — l’outil qui installe les dépendances et lance l’application

Avec `lancer.bat`, vous n’avez pas à les taper : le script les trouve tout seul. Ordre de recherche :

1. la variable d’environnement `NODE_PORTABLE` (dossier qui contient `node.exe`)
2. le dossier `node-portable\` à côté de `lancer.bat`
3. Node.js déjà présent dans le PATH Windows (installation classique)

### 1. Ouvrir un terminal

- **Invite de commandes** : touche Windows, tapez `cmd`, Entrée
- ou **PowerShell** : touche Windows, tapez `PowerShell`, Entrée

### 2. Contrôler si Node.js est installé

Dans le terminal, si Node est dans le PATH :

```bat
node -v
npm -v
```

Si vous utilisez déjà un **Node portable**, allez dans le dossier qui contient `node.exe` puis :

```bat
node.exe -v
```

Ou, avec le chemin complet (adaptez-le) :

```bat
"C:\chemin\vers\votre-node-portable\node.exe" -v
```

| Résultat | Signification |
| --- | --- |
| Un numéro s’affiche, par exemple `v22.14.0` | Node.js est présent |
| `'node' n'est pas reconnu…` | Node n’est pas dans le PATH. Ce n’est pas bloquant : utilisez le zip portable ci-dessous, ou pointez `NODE_PORTABLE` vers votre dossier existant |

### 3. Contrôler si c’est bien une LTS

Avec Node dans le PATH :

```bat
node -p "process.release.lts"
```

Avec un Node portable :

```bat
"C:\chemin\vers\votre-node-portable\node.exe" -p "process.release.lts"
```

| Résultat | Signification |
| --- | --- |
| Un nom s’affiche (`Krypton`, `Jod`, etc.) | Vous êtes sur une LTS : c’est bon |
| `undefined` | Vous êtes sur une version « Current », pas LTS. Prenez une archive **LTS** |

La version LTS du moment est celle indiquée en vert sur [https://nodejs.org](https://nodejs.org) (**LTS**, pas *Latest Current*).

### 4. Node portable (recommandé, sans droits administrateur)

C’est la méthode adaptée aux PC d’entreprise : pas d’installeur, pas de droits admin.

**Vous avez déjà un Node portable**

Deux possibilités :

- Copiez (ou dézippez) le contenu dans le dossier `node-portable` **à côté de** `lancer.bat`. `node.exe` peut être directement dans `node-portable\` ou dans un sous-dossier (cas typique de l’archive officielle).
- Ou indiquez une fois pour toutes l’emplacement existant. Dans **cmd** :

```bat
setx NODE_PORTABLE "C:\chemin\vers\votre-node-portable"
```

`setx` écrit la variable pour les sessions suivantes. Fermez le terminal, ouvrez-en un nouveau, puis lancez `lancer.bat`. Le dossier visé doit contenir `node.exe` (et en général `npm.cmd`).

**Vous n’avez pas encore Node (et pas les droits admin)**

C’est possible. Double-cliquez sur `lancer.bat` : s’il ne trouve pas Node, il télécharge tout seul la LTS **portable** (fichier zip, pas l’installeur) dans `node-portable\`, installe les dépendances du projet, puis ouvre le navigateur. Aucune élévation administrateur n’est demandée.

Il faut seulement :

- pouvoir écrire dans le dossier du projet
- un accès internet la première fois (nodejs.org + registre npm)

Les fois suivantes, le Node portable déjà téléchargé est réutilisé.

Si le téléchargement automatique échoue (réseau d’entreprise, site bloqué), faites-le à la main :

1. Allez sur [https://nodejs.org](https://nodejs.org).
2. Ouvrez **Other Downloads** / archives, et téléchargez le **Windows Binary (.zip)** **x64** de la LTS (pas le `.msi`).
3. Créez un dossier `node-portable` à côté de `lancer.bat`.
4. Dézippez l’archive dedans. Vous devez obtenir quelque chose comme :

```text
diff_excel\
  lancer.bat
  node-portable\
    node-v22.xx.x-win-x64\
      node.exe
      npm.cmd
      ...
```

5. Relancez `lancer.bat`. Le script affiche la ligne `Node portable : ...` puis la version.

Ne commitez pas `node-portable` : ce dossier est ignoré par git.

### 5. Installation classique (si vous avez les droits admin)

1. Allez sur [https://nodejs.org](https://nodejs.org).
2. Téléchargez le bouton **LTS** (installeur Windows `.msi`).
3. Laissez les options par défaut (PATH coché).
4. Fermez tous les terminaux, ouvrez-en un nouveau.
5. Vérifiez avec `node -v` puis `node -p "process.release.lts"`.

`lancer.bat` s’en sert automatiquement s’il ne trouve pas de Node portable.

## Lancer l’application

### Méthode recommandée

1. Ouvrez le dossier du projet.
2. Double-cliquez sur `lancer.bat` : une fenêtre noire **doit rester ouverte**.
3. Pas de Node et pas admin : le script télécharge une LTS portable, puis installe les dépendances (quelques minutes la première fois).
4. Le navigateur s’ouvre sur [http://127.0.0.1:5173/](http://127.0.0.1:5173/). Si ce n’est pas le cas, ouvrez cette adresse à la main. Si le serveur tourne déjà, le script n’en lance pas un second : il ouvre seulement le navigateur.

Laissez la fenêtre noire ouverte pendant que vous travaillez. C’est le serveur local.

### Méthode alternative (terminal)

Dans le dossier du projet :

```bat
npm install
npm start
```

`npm start` lance le serveur et ouvre le navigateur.  
`npm run dev` lance le serveur : ouvrez [http://127.0.0.1:5173/](http://127.0.0.1:5173/) si le navigateur ne s’ouvre pas.

## Arrêter l’application

Fermez la fenêtre du script, ou dans le terminal appuyez sur **Ctrl+C**.

## Premier usage

1. Déposez (ou cliquez pour choisir) la **version de base** à gauche (ancien fichier).
2. Déposez la **nouvelle version** à droite.
3. La comparaison se lance dès que les deux fichiers sont chargés.

Formats acceptés : `.xlsx`, `.xls`, `.xlsm`, `.xlsb`, `.ods`, `.csv`, `.tsv`.

Pour tester sans vos fichiers : cliquez sur **charger un exemple**, ou ouvrez `http://localhost:5173/?demo=1`.

## Comparer et fusionner

- **Revue** : liste des changements à accepter ou refuser, feuille par feuille.
- **Grille** : vue type Git (unifié ou côte à côte).
- **Tout garder : ancienne / nouvelle** : tranche toutes les différences d’un coup.
- **Enregistrer le classeur fusionné** : télécharge un troisième fichier. Les mises en forme d’origine (couleurs, polices, bordures, etc.) sont conservées autant que possible. Les changements encore en attente restent ceux du fichier de base.
- **Comparatif coloré** : télécharge un classeur annoté (vert = ajouté, rouge = supprimé).

Raccourcis en vue grille : **n** / **j** = changement suivant, **p** / **k** = précédent, **1** = unifié, **2** = côte à côte.

Le bouton de réinitialisation (icône de retour) ramène à l’écran de chargement.

## En cas de problème

| Symptôme | Que faire |
| --- | --- |
| « Node.js est introuvable » / échec du téléchargement | Relancez `lancer.bat` avec internet. Sinon dézippez à la main le zip Windows LTS dans `node-portable\`, ou définissez `NODE_PORTABLE` vers un `node.exe` déjà présent. |
| `NODE_PORTABLE est defini mais node.exe est introuvable` | La variable pointe vers le mauvais dossier : elle doit viser le dossier **qui contient** `node.exe`, pas un dossier parent. |
| `node -p "process.release.lts"` affiche `undefined` | Ce n’est pas une LTS. Reprenez l’archive **LTS** (zip ou installeur vert). |
| Le navigateur ne s’ouvre pas | Laissez la fenêtre noire ouverte et allez sur [http://127.0.0.1:5173/](http://127.0.0.1:5173/). |
| La fenêtre noire se ferme tout de suite | Relancez `lancer.bat` : la fenêtre est maintenant **gardée ouverte** même en cas d’erreur. Lisez le message affiché. |
| Le port est déjà utilisé | `lancer.bat` ouvre le navigateur sur l’instance déjà lancée. Pour tout redémarrer, fermez l’autre fenêtre Diff Excel. |
| Fichier illisible | Vérifiez le format. Les fichiers protégés par mot de passe ne sont pas pris en charge. |
| La mise en forme manque dans le fichier fusionné | Travaillez de préférence avec des `.xlsx` / `.xlsm`. Les graphiques et macros peuvent être perdus si vous acceptez des changements de la nouvelle version. |

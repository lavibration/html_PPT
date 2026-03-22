# 🚀 Convertisseur HTML vers PPTX (GitHub Powered)

Ce projet permet de transformer du code HTML (compatible Tailwind CSS) en un fichier PowerPoint (.pptx) dont les éléments (textes, formes, images) restent entièrement modifiables.

## 🛠️ Comment ça marche ?
1. **Frontend (GitHub Pages)** : Une interface simple pour coller votre code HTML et prévisualiser le rendu.
2. **Backend (GitHub Actions)** : Un script Python qui utilise Playwright pour calculer le rendu visuel et `python-pptx` pour générer le document PowerPoint.

---

## 🚀 Installation (Pour vous ou pour partager)

### 1. Hébergement sur GitHub Pages
Pour que l'outil soit accessible en ligne (ex: `https://votre-pseudo.github.io/votre-repo/`) :
1. Créez un nouveau dépôt sur GitHub.
2. Poussez les fichiers du projet (`index.html`, `script.py`, `requirements.txt`, `.github/workflows/convert.yml`) vers ce dépôt.
3. Allez dans **Settings** -> **Pages**.
4. Sous "Build and deployment", choisissez la branche `main` et cliquez sur **Save**.

### 2. Création du Token de sécurité (PAT)
Pour que la page web puisse demander à GitHub de générer le fichier, vous avez besoin d'un "Personal Access Token" :
1. Allez dans vos [Paramètres GitHub (Tokens)](https://github.com/settings/tokens).
2. Cliquez sur **Generate new token (classic)**.
3. Donnez-lui un nom (ex: "PPTX-Converter").
4. Cochez uniquement la case **`repo`** (pour permettre le déclenchement des Actions).
5. Copiez le token (il ne sera affiché qu'une seule fois).

---

## 📖 Utilisation

1. Ouvrez l'URL de votre page GitHub Pages.
2. Saisissez le chemin de votre dépôt (ex: `votre-pseudo/nom-du-repo`).
3. Collez votre Token GitHub.
4. Collez votre code HTML dans la zone de texte.
5. Cliquez sur **Générer via GitHub Actions**.
6. **Récupération du fichier** :
   - Allez sur l'onglet **Actions** de votre dépôt GitHub.
   - Cliquez sur le dernier workflow terminé (nommé "Generate PPTX").
   - Téléchargez le fichier dans la section **Artifacts**.

---

## 🤝 Partager l'outil
Si vous voulez partager l'outil à quelqu'un :
- Envoyez-lui simplement l'URL de votre **GitHub Pages**.
- **Important** : L'utilisateur devra posséder son propre Token GitHub et l'utiliser avec son propre dépôt (ou le vôtre si vous lui donnez les droits) pour que cela fonctionne.

---

## 🎨 Spécifications techniques
- **Framework UI** : Tailwind CSS
- **Moteur de rendu** : Playwright (Chromium)
- **Générateur Document** : python-pptx
- **Ratio** : 16:9 (1280x720px)

# TFA Service Park — Guide de déploiement

## Prérequis

- Google Sheet TFA existant avec les feuilles `driver_state` et `damage_components`
- Compte GitHub avec accès au repo du service park
- Clé API Steam (obtenue sur https://steamcommunity.com/dev/apikey)

---

## Étape 1 — Vérifier la structure du Google Sheet

Ouvre le Google Sheet TFA et vérifie que les deux feuilles ont exactement ces colonnes (ordre libre, noms exacts) :

**Feuille `driver_state`**
```
driver_guid | driver_name | car_model | stage_id | validated |
repair_used_min | ballast_kg | restrictor | penalty_seconds | last_updated
```

**Feuille `damage_components`**
```
driver_guid | stage_id | component_id | score | severity |
ballast_kg | restrictor | repair_min | repaired
```

Le code lit les colonnes par leur nom en ligne 1 — l'ordre n'a pas d'importance.

---

## Étape 2 — Créer le projet Apps Script

1. Dans le Google Sheet TFA : **Extensions → Apps Script**
2. Le projet s'ouvre dans un nouvel onglet
3. Supprime le code par défaut (`function myFunction() {}`)
4. Colle intégralement le contenu de `gas/code.gs`
5. Sauvegarde (Ctrl+S) — nomme le projet ex: `TFA Service Park`

---

## Étape 3 — Configurer les Script Properties

Dans Apps Script : **Project Settings** (icône ⚙️ à gauche) → **Script Properties** → **Add script property**

Ajoute ces 4 propriétés :

| Propriété | Valeur |
|---|---|
| `STEAM_SECRET` | Chaîne aléatoire longue — génère avec : `crypto.randomUUID().replace(/-/g,'') + crypto.randomUUID().replace(/-/g,'')` dans la console du navigateur |
| `SHEET_ID` | L'ID du Google Sheet (dans l'URL : `docs.google.com/spreadsheets/d/**ID_ICI**/edit`) |
| `SERVICE_PARK_URL` | URL GitHub Pages (ex: `https://monorg.github.io/tfa-servicepark`) — peut être provisoire, à mettre à jour à l'étape 8 |
| `CURRENT_STAGE_ID` | `etape_01` (à changer à chaque étape) |

> **Important :** `STEAM_SECRET` ne doit jamais être partagé ni versionné. C'est la clé de signature des tokens pilotes.

---

## Étape 4 — Déployer le Web App

1. Dans Apps Script : bouton **Deploy → New deployment**
2. Type : **Web app**
3. Paramètres :
   - **Description** : TFA Service Park v1
   - **Execute as** : Me (ton compte Google)
   - **Who has access** : Anyone
4. Clique **Deploy**
5. Autorise les permissions demandées (accès aux Sheets, aux URLs externes)
6. **Copie l'URL de déploiement** — format : `https://script.google.com/macros/s/AKfycb.../exec`

> À chaque modification du code, tu dois créer un **New deployment** (pas "Manage deployments" → Edit). L'URL reste la même si tu réutilises le même deployment ID.

---

## Étape 5 — Configurer le HTML

Ouvre `service_park_auth.html` et renseigne le bloc CONFIG en haut du fichier :

```javascript
const CONFIG = {
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycb.../exec',  // ← URL étape 4
  CURRENT_STAGE_ID: 'etape_01'
};
```

---

## Étape 6 — Publier sur GitHub Pages

1. Crée un repo GitHub public (ex: `tfa-servicepark`)
2. Push `service_park_auth.html` à la racine sous le nom `index.html` :
   ```bash
   cp service_park_auth.html index.html
   git init
   git add index.html
   git commit -m "TFA Service Park"
   git branch -M main
   git remote add origin https://github.com/TON_ORG/tfa-servicepark.git
   git push -u origin main
   ```
3. Dans le repo GitHub : **Settings → Pages**
4. Source : **Deploy from a branch** → branche `main` → dossier `/` (root)
5. Clique **Save**
6. Attends ~1 minute, puis l'URL GitHub Pages s'affiche (ex: `https://tonorg.github.io/tfa-servicepark`)

---

## Étape 7 — Mettre à jour SERVICE_PARK_URL

Retourne dans Apps Script → **Project Settings → Script Properties** et mets à jour :

| Propriété | Nouvelle valeur |
|---|---|
| `SERVICE_PARK_URL` | URL GitHub Pages obtenue à l'étape 6 |

> Cette URL est utilisée comme `openid.realm` dans le flow Steam — elle doit correspondre exactement à l'URL du site.

---

## Étape 8 — Test du flow complet

### Test de l'auth Steam

1. Ouvre `https://tonorg.github.io/tfa-servicepark`
2. Clique **Se connecter avec Steam**
3. Steam s'ouvre → connecte-toi
4. Tu dois être redirigé vers le service park avec ton nom et tes données chargées

### Test des réparations

1. Toggle quelques composants → vérifie que le budget se met à jour
2. Ouvre le Google Sheet → feuille `damage_components` → la colonne `repaired` doit se mettre à jour en quelques secondes
3. Clique **Valider les réparations** → confirme → vérifie que `validated = TRUE` dans `driver_state`
4. Recharge la page → les toggles doivent être désactivés (état validé chargé)

### En cas d'erreur

- **"Auth failed: signature invalide"** → vérifie que `SERVICE_PARK_URL` dans les Script Properties correspond exactement à l'URL passée à Steam
- **"driver_not_found"** → le Steam ID64 du pilote n'est pas dans la feuille `driver_state` — vérifie la colonne `driver_guid`
- **Erreur CORS** → vérifie que le Web App est déployé avec "Who has access: Anyone" (pas "Anyone with Google account")
- **Toast "Erreur sauvegarde"** → ouvre la console du navigateur (F12) pour voir le détail de l'erreur fetch

---

## Étape 9 — Changer d'étape

Pour chaque nouvelle étape :

1. Dans Apps Script → **Script Properties** → modifier `CURRENT_STAGE_ID` (ex: `etape_02`)
2. Dans `service_park_auth.html` → modifier `CONFIG.CURRENT_STAGE_ID` → push sur GitHub
3. Importer les nouveaux résultats de course dans la feuille `damage_components` (driver_guid + stage_id + composants)
4. Remettre `validated = FALSE` dans `driver_state` pour tous les pilotes

---

## Récapitulatif des URLs à noter

| Élément | Valeur |
|---|---|
| Google Sheet ID | `...` |
| Apps Script Web App URL | `https://script.google.com/macros/s/.../exec` |
| GitHub Pages URL | `https://tonorg.github.io/tfa-servicepark` |
| STEAM_SECRET | (ne pas noter ici — stocké dans Script Properties uniquement) |

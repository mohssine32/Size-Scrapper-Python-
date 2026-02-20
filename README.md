# 👟 Python Size Scrapper

Scraper Python qui extrait automatiquement les **informations produit** et les **guides de taille** depuis des sites e-commerce de mode, puis exporte les données dans un fichier Excel formaté.

## 🎯 Sites supportés

| Marque | Produit | Guide de taille |
|--------|---------|-----------------|
| **Prada** | ✅ | ✅ (EU, UK, US + cm) |
| **Kleman** | ✅ | ✅ (EU, UK, US + cm) |
| **La Bottega Gardiane** | ✅ | ✅ (EU, UK, US, IT + cm) |

## 📋 Fonctionnalités

- **Scraping produit** : titre, genre (Homme/Femme/Unisexe), type (Shoes, Bag, Clothing, Accessory)
- **Scraping guide de taille** : tailles EU/FR, UK, US, IT et longueur du pied en cm
- **Détection automatique** du genre et du type de produit via mots-clés et `dataLayer`
- **Export Excel** stylé avec 2 onglets :
  - *Pages produit* : liste des produits scrapés
  - *Guides de taille* : tableaux de correspondance des tailles
- **Acceptation automatique des cookies**
- **Anti-détection bot** (masquage `navigator.webdriver`, user-agent personnalisé)

## 🛠️ Prérequis

- **Python 3.8+**
- **Playwright** (automatisation navigateur Chromium)
- **openpyxl** (lecture/écriture Excel)

## 📦 Installation

```bash
# Cloner le projet
git clone <url-du-repo>
cd python-size-scrapper

# Installer les dépendances
pip install playwright openpyxl

# Installer le navigateur Chromium pour Playwright
playwright install chromium
```

## 🚀 Utilisation

### Script principal (recommandé)

Scrape le produit **et** le guide de taille en une seule commande :

```bash
python main.py <URL> [Homme|Femme]
```

**Exemples :**

```bash
# Prada
python main.py https://www.prada.com/fr/fr/women/shoes/...

# Kleman - chaussures homme
python main.py https://kleman-france.com/products/padror-th-cognac Homme

# La Bottega Gardiane - chaussures femme
python main.py https://www.labottegardiane.com/... Femme
```

> Le paramètre `Homme|Femme` est optionnel (par défaut : `Homme`). Il est utilisé pour sélectionner le bon tableau de tailles sur les sites Kleman et La Bottega Gardiane.

### Scripts individuels

#### Scraper produit uniquement

```bash
python scraper_produit.py <URL>
```

Extrait le titre, le genre et le type du produit, et affiche les résultats dans le terminal.

#### Scraper guide de taille uniquement

```bash
python scraper_guide_taille.py <URL> [Homme|Femme]
```

Extrait le guide de taille et l'exporte dans le fichier Excel.

#### Export Excel (produit seul)

```bash
python export_excel.py <URL>
```

Scrape les infos produit et les ajoute dans l'onglet *Pages produit* du fichier Excel.

## 📊 Structure du fichier Excel

Le fichier `etudes_de_cas.xlsx` est généré automatiquement avec 2 onglets :

### Onglet 1 — Pages produit

| Nom Produit | Gender | Type | URL | Guide de taille |
|-------------|--------|------|-----|-----------------|
| Derby Padror | Homme | Shoes | https://... | 1 |

### Onglet 2 — Guides de taille

Format horizontal avec correspondance multi-systèmes :

| Systèmes métriques | | Taille 1 | Taille 2 | Taille 3 | ... |
|--------------------|-|----------|----------|----------|-----|
| Marque | EU | 39 | 40 | 41 | ... |
| Royaume-Uni | UK | 5 | 6 | 7 | ... |
| États-Unis | US | 6 | 7 | 8 | ... |
| Longueur pied | | 25 cm | 25.5 cm | 26 cm | ... |

## 📁 Structure du projet

```
python-size-scrapper/
├── main.py                    # Script principal (produit + guide + export)
├── scraper_produit.py         # Scraper d'informations produit (standalone)
├── scraper_guide_taille.py    # Scraper de guide de taille (standalone)
├── export_excel.py            # Export Excel onglet produit (standalone)
├── etudes_de_cas.xlsx         # Fichier Excel généré (après exécution)
└── README.md
```

## ⚙️ Détails techniques

- **Navigateur** : Chromium lancé en mode visible (`headless=False`) pour éviter les blocages anti-bot
- **Locale** : `fr-FR` pour obtenir les pages en français
- **Détection genre/type** : analyse du `dataLayer` JavaScript et du contenu texte de la page
- **Sélecteurs CSS** : spécifiques à chaque marque pour extraire les tableaux de taille

## ⚠️ Notes importantes

- Un navigateur Chromium s'ouvre automatiquement lors du scraping — **c'est normal**
- Ne pas fermer le navigateur manuellement, il se ferme automatiquement à la fin
- Le scraping peut prendre quelques secondes par page (chargement JS + cookies)
- Les sites e-commerce peuvent modifier leur structure HTML, ce qui peut nécessiter une mise à jour des sélecteurs

## 📝 Licence

Usage personnel / éducatif.

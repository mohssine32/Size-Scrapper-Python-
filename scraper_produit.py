"""
Scraper de produit e-commerce
Usage : python scraper_produit.py <URL>
"""

import sys
import time
from playwright.sync_api import sync_playwright

GENDER_KEYWORDS = {
    "Homme":   ["mens", "homme", "man", "uomo", "masculin"],
    "Femme":   ["womens", "femme", "woman", "donna", "feminin"],
    "Unisexe": ["unisex", "unisexe"],
}

TYPE_KEYWORDS = {
    "Shoes":     ["chaussure", "shoe", "basket", "sneaker", "derby", "mocassin", "boot", "scarpe", "footwear", "derbies", "soulier", "espadrille"],
    "Bag":       ["sac", "bag", "pochette", "handbag", "borsa"],
    "Clothing":  ["veste", "manteau", "robe", "pantalon", "jacket", "coat", "dress", "pull", "chemise"],
    "Accessory": ["ceinture", "belt", "foulard", "scarf", "chapeau", "hat"],
}


def guess_gender(text):
    text_lower = text.lower()
    for genre, keywords in GENDER_KEYWORDS.items():
        if any(kw in text_lower for kw in keywords):
            return genre
    return None


def guess_type(text):
    text_lower = text.lower()
    for product_type, keywords in TYPE_KEYWORDS.items():
        if any(kw in text_lower for kw in keywords):
            return product_type
    return None


def scrape_product(url):
    result = {"titre": None, "gender": None, "type": None, "url": url}

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,  # navigateur visible pour eviter le blocage
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                       "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            locale="fr-FR",
        )
        page = context.new_page()

        # Masquer le bot
        page.add_init_script("""
            Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
            window.chrome = { runtime: {} };
        """)

        print(f"\n Chargement de la page (un navigateur va s'ouvrir, c'est normal)...")

        page.goto(url, wait_until="domcontentloaded", timeout=60000)

        # Attendre que le JS charge
        print(" Attente du chargement JavaScript...")
        time.sleep(5)

        # ── Titre via h1 ──────────────────────────────────────
        try:
            result["titre"] = page.eval_on_selector("h1", "el => el.innerText.trim()")
        except Exception:
            pass

        # Si pas de h1, on prend le title de la page
        if not result["titre"]:
            title = page.title()
            result["titre"] = title.split("|")[0].split("-")[0].strip()

        # ── Lire le dataLayer (contient gender, type, etc.) ───
        try:
            datalayer = page.evaluate("() => JSON.stringify(window.dataLayer)")
            if datalayer:
                # Gender depuis dataLayer (u4 = mens/womens)
                if not result["gender"]:
                    result["gender"] = guess_gender(datalayer)

                # Type depuis dataLayer
                if not result["type"]:
                    result["type"] = guess_type(datalayer)
        except Exception:
            pass

        # ── Fallback : lire le texte brut de la page ──────────
        if not result["gender"] or not result["type"]:
            try:
                page_text = page.inner_text("body")
                if not result["gender"]:
                    result["gender"] = guess_gender(page_text)
                if not result["type"]:
                    result["type"] = guess_type(page_text)
            except Exception:
                pass

        browser.close()

    return result


def afficher_resultat(data):
    print("\n" + "=" * 50)
    print("  PRODUIT TROUVE")
    print("=" * 50)
    print(f"  Titre  : {data['titre']  or 'Non trouve'}")
    print(f"  Gender : {data['gender'] or 'Non trouve'}")
    print(f"  Type   : {data['type']   or 'Non trouve'}")
    print(f"  URL    : {data['url']}")
    print("=" * 50 + "\n")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("\nErreur : tu dois fournir une URL en argument.")
        print("Usage : python scraper_produit.py <URL>\n")
        sys.exit(1)

    url = sys.argv[1]

    if not url.startswith("http"):
        print("\nErreur : l'URL doit commencer par http:// ou https://\n")
        sys.exit(1)

    try:
        data = scrape_product(url)
        afficher_resultat(data)
    except Exception as e:
        print(f"\nErreur lors du scraping : {e}\n")
        sys.exit(1)

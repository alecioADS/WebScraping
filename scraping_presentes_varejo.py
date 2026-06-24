# pip install selenium webdriver-manager pandas openpyxl undetected-chromedriver

"""
Web Scraping - Lista de Presentes (Varejo)

Fonte 1: Mercado Livre (página de Ofertas + busca complementar por termos)
Fonte 2: Amazon Brasil (com fallback Magalu -> Shopee se bloquear)
Filtro: produtos entre R$ 100,00 e R$ 150,00
Volume: mínimo 300 / máximo 2000
"""

from __future__ import annotations

import re
import time
import random
from collections import Counter
from urllib.parse import quote_plus, urljoin

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

PRECO_MIN = 100.00
PRECO_MAX = 150.00
MIN_REGISTROS = 300
MAX_REGISTROS = 2000
MAX_PAGINAS = 40
ARQUIVO_SAIDA = "presentes_varejo.xlsx"

TERMOS_ML = [
    "presentes", "decoracao casa", "brinquedos", "relogio masculino",
    "relogio feminino", "perfume importado", "mochila impermeavel",
    "carteira couro", "fone bluetooth", "garrafa termica",
    "luminaria mesa", "kit ferramentas", "jogo de jantar",
    "organizador cozinha", "caixa de som bluetooth", "quadros decorativos",
    "mochila notebook",
]

USER_AGENT_MAC = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)
USER_AGENT_WIN = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
)


def iniciar_driver(stealth: bool = False, headless: bool = False) -> WebDriver:
    """Inicia Chrome WebDriver. stealth=True usa undetected-chromedriver com anti-bot."""
    if stealth:
        import undetected_chromedriver as uc
        options = uc.ChromeOptions()
        if headless:
            options.add_argument("--headless")
        options.add_argument("--start-maximized")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=pt-BR")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument(f"--user-agent={USER_AGENT_WIN}")
        driver = uc.Chrome(options=options)
        try:
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
            )
        except Exception:
            pass
    else:
        options = Options()
        if headless:
            options.add_argument("--headless=new")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=pt-BR")
        options.add_argument(f"--user-agent={USER_AGENT_MAC}")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

    driver.set_page_load_timeout(45)
    return driver


def pagina_bloqueada(driver: WebDriver) -> bool:
    """Detecta bloqueio HTTP 403, captcha, login obrigatório etc."""
    try:
        titulo = (driver.title or "").lower()
        corpo = (driver.page_source or "")[:10000].lower()
        indicadores = (
            "403", "forbidden", "access denied", "acesso negado",
            "request blocked", "bloqueado", "captcha", "cloudflare",
            "comportamento incomum", "please confirm you are a human",
            "página indisponível", "pagina indisponivel",
        )
        return any(i in titulo or i in corpo for i in indicadores)
    except Exception:
        return True


def limpar_preco(texto: str | None) -> float | None:
    """Converte string monetária 'R$ 1.234,56' em float 1234.56."""
    if not texto:
        return None
    texto = texto.replace("\xa0", " ").replace("R$", "").strip()
    match = re.search(r"(\d{1,3}(?:\.\d{3})*|\d+),(\d{2})", texto)
    if match:
        try:
            return float(f"{match.group(1).replace('.', '')}.{match.group(2)}")
        except ValueError:
            pass
    match_int = re.search(r"(\d{1,3}(?:\.\d{3})*|\d+)", texto)
    if match_int:
        try:
            return float(match_int.group(1).replace(".", ""))
        except ValueError:
            pass
    return None


def preco_valido(preco: float | None) -> bool:
    return preco is not None and PRECO_MIN <= preco <= PRECO_MAX


def scroll_gradual(driver: WebDriver, passos: int = 5, delay: float = 0.5) -> None:
    try:
        for i in range(passos):
            driver.execute_script(
                f"window.scrollTo(0, (document.body.scrollHeight / {passos}) * {i + 1});"
            )
            time.sleep(delay)
    except Exception:
        pass


def texto_primeiro(el, seletores) -> str | None:
    for s in seletores:
        try:
            t = el.find_element(By.CSS_SELECTOR, s).text.strip()
            if t:
                return t
        except NoSuchElementException:
            continue
    return None


def atributo_primeiro(el, seletores, atributo) -> str | None:
    for s in seletores:
        try:
            v = el.find_element(By.CSS_SELECTOR, s).get_attribute(atributo)
            if v:
                return v
        except NoSuchElementException:
            continue
    return None


def adicionar(produtos, vistos, descricao, preco, loja, link) -> None:
    if not descricao or not link or not preco_valido(preco):
        return
    link_limpo = link.split("?")[0]
    chave = (loja, link_limpo)
    if chave in vistos:
        return
    produtos.append({
        "Descrição": re.sub(r"\s+", " ", descricao).strip(),
        "Valor": round(float(preco), 2),
        "Loja": loja,
        "Link": link,
    })
    vistos.add(chave)


def remover_duplicados(produtos: list[dict]) -> list[dict]:
    unicos, vistos = [], set()
    for p in produtos:
        chave = (p["Loja"], p["Link"].split("?")[0])
        if chave in vistos:
            continue
        unicos.append(p)
        vistos.add(chave)
    return unicos


def coletar_mercado_livre_ofertas(driver: WebDriver, limite: int) -> list[dict]:
    """Coleta da página de Ofertas do Mercado Livre, com paginação."""
    loja = "Mercado Livre"
    produtos, vistos = [], set()
    url = "https://www.mercadolivre.com.br/ofertas"
    seletor_cards = (
        "li.promotion-item, div.poly-card, "
        "div.ui-search-result__wrapper, li.ui-search-layout__item"
    )

    for pagina in range(1, MAX_PAGINAS + 1):
        if len(produtos) >= limite:
            break
        print(f"[{loja}] Página {pagina}: {url}")
        try:
            driver.get(url)
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_cards))
            )
        except TimeoutException:
            print(f"[{loja}] Timeout na página {pagina}")
            break

        for card in driver.find_elements(By.CSS_SELECTOR, seletor_cards):
            if len(produtos) >= limite:
                break
            try:
                descricao = texto_primeiro(card, [
                    ".promotion-item__title", ".poly-component__title",
                    ".ui-search-item__title", "h2",
                ])
                preco_t = texto_primeiro(card, [
                    ".andes-money-amount__fraction",
                    ".promotion-item__price", ".price-tag-fraction",
                ])
                centavos = texto_primeiro(card, [
                    ".andes-money-amount__cents", ".price-tag-cents",
                ])
                if preco_t and centavos and "," not in preco_t:
                    preco_t = f"{preco_t},{centavos}"
                link = atributo_primeiro(card, ["a[href]"], "href")
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)
            except (NoSuchElementException, WebDriverException):
                continue

        print(f"[{loja}] Acumulado: {len(produtos)}")

        proxima = None
        for s in ("a.andes-pagination__link[title='Seguinte']", "a[aria-label='Seguinte']"):
            try:
                proxima = driver.find_element(By.CSS_SELECTOR, s).get_attribute("href")
                if proxima:
                    break
            except NoSuchElementException:
                continue
        if not proxima or proxima == url:
            print(f"[{loja}] Fim da paginação")
            break
        url = proxima
        time.sleep(random.uniform(1.5, 3))

    return produtos


def coletar_mercado_livre_busca(driver: WebDriver, limite: int, termos: list[str]) -> list[dict]:
    """Busca por múltiplos termos no Mercado Livre para complementar o volume mínimo."""
    loja = "Mercado Livre"
    produtos, vistos = [], set()
    seletor_cards = (
        "li.ui-search-layout__item, div.ui-search-result__wrapper, div.poly-card"
    )

    for termo in termos:
        if len(produtos) >= limite:
            break
        url = f"https://lista.mercadolivre.com.br/{quote_plus(termo)}"
        print(f"[{loja} Busca] '{termo}': {url}")
        try:
            driver.get(url)
            if pagina_bloqueada(driver):
                print(f"[{loja} Busca] Bloqueado em '{termo}', pulando")
                continue
            WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_cards))
            )
        except TimeoutException:
            print(f"[{loja} Busca] Timeout em '{termo}'")
            continue

        antes = len(produtos)
        for card in driver.find_elements(By.CSS_SELECTOR, seletor_cards):
            if len(produtos) >= limite:
                break
            try:
                descricao = texto_primeiro(card, [
                    ".ui-search-item__title", ".poly-component__title", "h2",
                ])
                preco_t = texto_primeiro(card, [
                    ".andes-money-amount__fraction", ".price-tag-fraction",
                ])
                centavos = texto_primeiro(card, [
                    ".andes-money-amount__cents", ".price-tag-cents",
                ])
                if preco_t and centavos and "," not in preco_t:
                    preco_t = f"{preco_t},{centavos}"
                link = atributo_primeiro(card, ["a[href]"], "href")
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)
            except (NoSuchElementException, WebDriverException):
                continue
        print(f"[{loja} Busca] +{len(produtos) - antes}, total: {len(produtos)}")
        time.sleep(random.uniform(2, 4))

    return produtos


def coletar_amazon(driver: WebDriver, limite: int) -> list[dict]:
    """Coleta da Amazon Brasil com filtro de preço já na URL (rh=p_36)."""
    loja = "Amazon"
    produtos, vistos = [], set()
    # rh=p_36:10000-15000 filtra preço entre R$100 e R$150 (valor em centavos)
    base = "https://www.amazon.com.br/s?k=presentes&rh=p_36%3A10000-15000"

    for pagina in range(1, MAX_PAGINAS + 1):
        if len(produtos) >= limite:
            break
        url = f"{base}&page={pagina}"
        print(f"[{loja}] Página {pagina}: {url}")
        try:
            driver.get(url)
            time.sleep(random.uniform(2, 4))
            scroll_gradual(driver)
            if pagina_bloqueada(driver):
                print(f"[{loja}] Bloqueado")
                break
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[data-component-type='s-search-result']")
                )
            )
        except TimeoutException:
            print(f"[{loja}] Timeout página {pagina}")

        cards = driver.find_elements(
            By.CSS_SELECTOR, "div[data-component-type='s-search-result']"
        )
        if not cards:
            break
        for card in cards:
            if len(produtos) >= limite:
                break
            try:
                descricao = texto_primeiro(
                    card, ["h2 span", "span.a-size-base-plus", "span.a-size-medium"]
                )
                preco_int = texto_primeiro(card, ["span.a-price-whole"])
                preco_cent = texto_primeiro(card, ["span.a-price-fraction"]) or "00"
                preco_t = f"{preco_int},{preco_cent}" if preco_int else None
                link = atributo_primeiro(
                    card, ["h2 a[href]", "a.a-link-normal[href]"], "href"
                )
                link = urljoin("https://www.amazon.com.br", link) if link else None
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)
            except Exception:
                continue
        print(f"[{loja}] Acumulado: {len(produtos)}")

        if driver.find_elements(
            By.CSS_SELECTOR, ".s-pagination-next.s-pagination-disabled"
        ):
            print(f"[{loja}] Fim da paginação")
            break

    return produtos


def coletar_magalu(driver: WebDriver, limite: int) -> list[dict]:
    """Coleta do Magalu (fallback por spec; historicamente bloqueia com 403)."""
    loja = "Magalu"
    produtos, vistos = [], set()
    seletor_cards = (
        "li[data-testid='product-card'], div[data-testid='product-card'], "
        "a[data-testid='product-card-container']"
    )

    for pagina in range(1, MAX_PAGINAS + 1):
        if len(produtos) >= limite:
            break
        url = f"https://www.magazineluiza.com.br/busca/presentes/?page={pagina}"
        print(f"[{loja}] Página {pagina}: {url}")
        try:
            driver.get(url)
            time.sleep(random.uniform(2, 4))
            scroll_gradual(driver)
            if pagina_bloqueada(driver):
                print(f"[{loja}] Bloqueado")
                break
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_cards))
            )
        except TimeoutException:
            print(f"[{loja}] Timeout página {pagina}")

        cards = driver.find_elements(By.CSS_SELECTOR, seletor_cards)
        if not cards:
            print(f"[{loja}] Nenhum card encontrado")
            break

        for card in cards:
            if len(produtos) >= limite:
                break
            try:
                descricao = texto_primeiro(
                    card, ["[data-testid='product-title']", "h2", "p"]
                )
                preco_t = texto_primeiro(card, [
                    "[data-testid='price-value']",
                    "[data-testid='price-original']",
                    "p[data-testid*='price']",
                ])
                link = card.get_attribute("href") or atributo_primeiro(
                    card, ["a[href]"], "href"
                )
                link = urljoin("https://www.magazineluiza.com.br", link) if link else None
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)
            except Exception:
                continue
        print(f"[{loja}] Acumulado: {len(produtos)}")

    return produtos


def coletar_shopee(driver: WebDriver, limite: int) -> list[dict]:
    """Coleta da Shopee como segundo fallback se Magalu bloquear."""
    loja = "Shopee"
    produtos, vistos = [], set()
    url = f"https://shopee.com.br/search?keyword={quote_plus('presentes')}"
    print(f"[{loja}] {url}")
    try:
        driver.get(url)
        time.sleep(5)
        try:
            botoes = driver.find_elements(
                By.XPATH, "//*[contains(text(), 'Português (BR)')]"
            )
            if botoes:
                botoes[0].click()
                time.sleep(3)
        except Exception:
            pass
        if pagina_bloqueada(driver):
            print(f"[{loja}] Bloqueado")
            return []
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((
                By.CSS_SELECTOR,
                "a[data-sqe='link'], div.shopee-search-item-result__item",
            ))
        )
        scroll_gradual(driver, passos=8, delay=0.6)
        for card in driver.find_elements(By.CSS_SELECTOR, "a[data-sqe='link']"):
            if len(produtos) >= limite:
                break
            try:
                descricao = texto_primeiro(
                    card, ["[data-sqe='name']", "div[class*='title']"]
                )
                preco_t = texto_primeiro(
                    card, ["div[class*='price']", "span.price"]
                )
                link = card.get_attribute("href")
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)
            except Exception:
                continue
    except Exception as e:
        print(f"[{loja}] Erro: {e}")
    return produtos


def coletar_fonte2(limite: int) -> list[dict]:
    """Fonte 2 conforme spec: Amazon -> Magalu (fallback) -> Shopee (2º fallback)."""
    cadeia = [
        ("Amazon", False, coletar_amazon),
        ("Magalu", True, coletar_magalu),
        ("Shopee", True, coletar_shopee),
    ]
    produtos = []
    for nome, stealth, funcao in cadeia:
        if len(produtos) >= limite:
            break
        restante = limite - len(produtos)
        print(f"\n[Fonte 2] Tentando {nome}... ({restante} produtos a coletar)")
        driver = None
        try:
            driver = iniciar_driver(stealth=stealth)
            coletados = funcao(driver, restante)
            if coletados:
                print(f"[Fonte 2] {nome}: {len(coletados)} produtos coletados")
                produtos.extend(coletados)
            else:
                print(f"[Fonte 2] {nome} não retornou produtos. Próxima opção.")
        except Exception as e:
            print(f"[Fonte 2] Erro em {nome}: {e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass
    return produtos


def exportar_excel(produtos: list[dict], arquivo: str = ARQUIVO_SAIDA) -> None:
    df = pd.DataFrame(produtos, columns=["Descrição", "Valor", "Loja", "Link"])
    df.to_excel(arquivo, engine="openpyxl", index=False)
    print(f"\n[OK] Excel exportado: {arquivo}")


def imprimir_resumo(produtos: list[dict]) -> None:
    contagem = Counter(p["Loja"] for p in produtos)
    print("\n================== RESUMO ==================")
    for loja, total in contagem.items():
        print(f"  {loja}: {total} produto(s)")
    print(f"  Total Geral: {len(produtos)} produto(s)")
    print("============================================")


def main() -> None:
    print("=" * 56)
    print("ROBÔ DE WEB SCRAPING - PRESENTES VAREJO")
    print(f"Filtro: R$ {PRECO_MIN:.2f} a R$ {PRECO_MAX:.2f}")
    print(f"Volume: mínimo {MIN_REGISTROS} / máximo {MAX_REGISTROS}")
    print("=" * 56)

    produtos: list[dict] = []

    # FONTE 1: Mercado Livre - Ofertas
    print("\n--- FONTE 1: Mercado Livre (Ofertas) ---")
    driver = None
    try:
        driver = iniciar_driver(stealth=False)
        produtos.extend(coletar_mercado_livre_ofertas(driver, MAX_REGISTROS // 2))
    except Exception as e:
        print(f"[Erro ML Ofertas] {e}")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    # FONTE 2: Amazon -> Magalu -> Shopee
    restante = max(MAX_REGISTROS - len(produtos), MIN_REGISTROS - len(produtos))
    if restante > 0:
        print(f"\n--- FONTE 2: Amazon (fallback Magalu/Shopee) | {restante} a coletar ---")
        produtos.extend(coletar_fonte2(restante))

    produtos = remover_duplicados(produtos)

    # Complemento ML por busca se não atingiu o mínimo
    if len(produtos) < MIN_REGISTROS:
        falta = MIN_REGISTROS - len(produtos)
        print(f"\n--- COMPLEMENTO: Busca ML por termos | faltam {falta} ---")
        driver = None
        try:
            driver = iniciar_driver(stealth=False)
            produtos.extend(coletar_mercado_livre_busca(driver, falta, TERMOS_ML))
            produtos = remover_duplicados(produtos)
        except Exception as e:
            print(f"[Erro Complemento ML] {e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    if len(produtos) > MAX_REGISTROS:
        produtos = produtos[:MAX_REGISTROS]

    if len(produtos) < MIN_REGISTROS:
        print(
            f"\n[Aviso] Atingiu {len(produtos)}/{MIN_REGISTROS} - possíveis bloqueios de rede."
        )

    if produtos:
        exportar_excel(produtos)
        imprimir_resumo(produtos)
    else:
        print("\n[Falha] Nenhum produto coletado.")


if __name__ == "__main__":
    main()

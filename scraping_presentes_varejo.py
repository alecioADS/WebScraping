# Comando para instalar todas as dependências necessárias:
# pip install selenium webdriver-manager pandas openpyxl undetected-chromedriver

"""
Web Scraping - Lista de Presentes (Varejo)

Descrição:
    Script acadêmico de web scraping que navega automaticamente por páginas
    de e-commerce, percorre a paginação, aplica filtro de preço e exporta
    os dados coletados para uma planilha Excel.

Fontes:
    Fonte 1: Mercado Livre (página de Ofertas + busca complementar por termos)
    Fonte 2: Amazon Brasil (com fallback Magalu -> Shopee se bloquear)

Filtro aplicado: produtos entre R$ 100,00 e R$ 150,00
Volume esperado: mínimo 300 / máximo 2000 registros
"""

# Permite uso de anotações de tipo como 'str | None' sem importar Optional
from __future__ import annotations

# Biblioteca padrão do Python
import re                          # Expressões regulares (para limpar preços)
import time                        # Pausas entre requisições
import random                      # Gera delays aleatórios para simular humano
from collections import Counter    # Conta ocorrências por loja no resumo final
from urllib.parse import quote_plus, urljoin  # Codifica termos para URL e une caminhos

# Bibliotecas de terceiros
import pandas as pd                # Estrutura os dados em DataFrame e exporta Excel

# Módulos do Selenium para automação do navegador
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,   # Elemento HTML não encontrado na página
    TimeoutException,         # Tempo de espera esgotado
    WebDriverException,       # Erros gerais do WebDriver
)
from selenium.webdriver.chrome.options import Options        # Configurações do Chrome
from selenium.webdriver.chrome.service import Service        # Serviço do ChromeDriver
from selenium.webdriver.common.by import By                  # Estratégias de localização (CSS, XPATH...)
from selenium.webdriver.remote.webdriver import WebDriver    # Tipagem do driver
from selenium.webdriver.support import expected_conditions as EC  # Condições de espera
from selenium.webdriver.support.ui import WebDriverWait      # Espera explícita

# Gerencia o download e atualização automática do ChromeDriver
from webdriver_manager.chrome import ChromeDriverManager

# ──────────────────────────────────────────────
# CONSTANTES DE CONFIGURAÇÃO
# ──────────────────────────────────────────────

PRECO_MIN = 100.00          # Preço mínimo do filtro (R$)
PRECO_MAX = 150.00          # Preço máximo do filtro (R$)
MIN_REGISTROS = 300         # Volume mínimo de produtos a coletar
MAX_REGISTROS = 2000        # Volume máximo de produtos a coletar
MAX_PAGINAS = 40            # Limite de páginas por fonte (evita loop infinito)
ARQUIVO_SAIDA = "presentes_varejo.xlsx"  # Nome do arquivo Excel gerado

# Termos usados na busca complementar do Mercado Livre
# Ativados somente se o volume mínimo não for atingido pelas fontes principais
TERMOS_ML = [
    "presentes", "decoracao casa", "brinquedos", "relogio masculino",
    "relogio feminino", "perfume importado", "mochila impermeavel",
    "carteira couro", "fone bluetooth", "garrafa termica",
    "luminaria mesa", "kit ferramentas", "jogo de jantar",
    "organizador cozinha", "caixa de som bluetooth", "quadros decorativos",
    "mochila notebook",
]

# User-Agent realista para simular um navegador comum e reduzir bloqueios
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


# ──────────────────────────────────────────────
# FUNÇÕES AUXILIARES — DRIVER E DETECÇÃO
# ──────────────────────────────────────────────

def iniciar_driver(stealth: bool = False, headless: bool = False) -> WebDriver:
    """
    Inicia e retorna uma instância do Chrome WebDriver.

    Parâmetros:
        stealth  -- usa undetected-chromedriver para contornar proteções anti-bot
        headless -- executa o navegador em segundo plano (sem janela visível)

    Retorna:
        Instância do WebDriver configurada e pronta para uso.
    """
    if stealth:
        # Modo stealth: usa undetected-chromedriver que modifica o Chrome
        # para não ser detectado como bot por sites com proteção avançada
        import undetected_chromedriver as uc
        options = uc.ChromeOptions()

        if headless:
            options.add_argument("--headless")  # Sem janela visível

        options.add_argument("--start-maximized")             # Abre maximizado
        options.add_argument("--no-sandbox")                  # Necessário em alguns sistemas Linux
        options.add_argument("--disable-dev-shm-usage")       # Evita erro de memória compartilhada
        options.add_argument("--lang=pt-BR")                  # Idioma português do Brasil
        options.add_argument("--disable-blink-features=AutomationControlled")  # Oculta flag de automação
        options.add_argument(f"--user-agent={USER_AGENT_WIN}")  # Simula navegador Windows

        driver = uc.Chrome(options=options)

        try:
            # Injeta script JavaScript que remove a propriedade 'webdriver' do navegador
            # Sites detectam automação checando navigator.webdriver == true
            driver.execute_cdp_cmd(
                "Page.addScriptToEvaluateOnNewDocument",
                {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
            )
        except Exception:
            pass  # Se falhar, continua sem o patch (melhor do que travar)

    else:
        # Modo padrão: Selenium comum com ChromeDriver gerenciado pelo webdriver-manager
        options = Options()

        if headless:
            options.add_argument("--headless=new")  # Modo headless moderno do Chrome

        options.add_argument("--start-maximized")
        options.add_argument("--disable-gpu")           # Evita problemas gráficos no headless
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--lang=pt-BR")
        options.add_argument(f"--user-agent={USER_AGENT_MAC}")  # Simula navegador macOS

        # ChromeDriverManager baixa e atualiza automaticamente o driver compatível com o Chrome instalado
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)

    # Define tempo máximo para carregamento de uma página (45 segundos)
    driver.set_page_load_timeout(45)
    return driver


def pagina_bloqueada(driver: WebDriver) -> bool:
    """
    Verifica se a página atual está bloqueada (403, captcha, Cloudflare etc.).

    Analisa o título e os primeiros 10.000 caracteres do HTML da página
    em busca de palavras-chave que indicam bloqueio ou redirecionamento
    para página de verificação.

    Retorna:
        True se a página estiver bloqueada, False caso contrário.
    """
    try:
        titulo = (driver.title or "").lower()          # Título da aba em minúsculas
        corpo = (driver.page_source or "")[:10000].lower()  # Primeiros 10 KB do HTML

        # Palavras-chave que indicam bloqueio ou verificação de acesso
        indicadores = (
            "403", "forbidden", "access denied", "acesso negado",
            "request blocked", "bloqueado", "captcha", "cloudflare",
            "comportamento incomum", "please confirm you are a human",
            "página indisponível", "pagina indisponivel",
        )

        # Retorna True se qualquer indicador aparecer no título ou no corpo da página
        return any(i in titulo or i in corpo for i in indicadores)
    except Exception:
        # Se não conseguiu verificar a página, assume bloqueio por segurança
        return True


# ──────────────────────────────────────────────
# FUNÇÕES AUXILIARES — PREÇO E DADOS
# ──────────────────────────────────────────────

def limpar_preco(texto: str | None) -> float | None:
    """
    Converte uma string monetária brasileira em float.

    Exemplos de entrada aceitos:
        'R$ 1.234,56'  ->  1234.56
        '129,90'       ->  129.90
        'R$100'        ->  100.0

    Retorna:
        Valor como float, ou None se não for possível converter.
    """
    if not texto:
        return None

    # Remove símbolo de moeda, espaços especiais e brancos extras
    texto = texto.replace("\xa0", " ").replace("R$", "").strip()

    # Tenta encontrar padrão com centavos: ex. '1.234,56' ou '129,90'
    match = re.search(r"(\d{1,3}(?:\.\d{3})*|\d+),(\d{2})", texto)
    if match:
        try:
            # Remove pontos de milhar e troca vírgula por ponto decimal
            return float(f"{match.group(1).replace('.', '')}.{match.group(2)}")
        except ValueError:
            pass

    # Fallback: tenta capturar só a parte inteira do valor
    match_int = re.search(r"(\d{1,3}(?:\.\d{3})*|\d+)", texto)
    if match_int:
        try:
            return float(match_int.group(1).replace(".", ""))
        except ValueError:
            pass

    return None  # Não foi possível interpretar o preço


def preco_valido(preco: float | None) -> bool:
    """
    Verifica se o preço está dentro da faixa definida (R$100 a R$150).

    Retorna:
        True se o preço for válido e dentro do intervalo, False caso contrário.
    """
    return preco is not None and PRECO_MIN <= preco <= PRECO_MAX


def scroll_gradual(driver: WebDriver, passos: int = 5, delay: float = 0.5) -> None:
    """
    Rola a página gradualmente do topo até o final.

    Necessário em páginas com carregamento lazy (imagens e cards que só
    aparecem quando o usuário rola a tela até eles).

    Parâmetros:
        passos -- quantas etapas de rolagem dividir a página
        delay  -- tempo de espera (segundos) entre cada etapa
    """
    try:
        for i in range(passos):
            # Rola até a fração (i+1)/passos da altura total da página
            driver.execute_script(
                f"window.scrollTo(0, (document.body.scrollHeight / {passos}) * {i + 1});"
            )
            time.sleep(delay)  # Aguarda o conteúdo carregar antes do próximo passo
    except Exception:
        pass  # Ignora erros de scroll (ex: página sem scrollbar)


# ──────────────────────────────────────────────
# FUNÇÕES AUXILIARES — EXTRAÇÃO DE ELEMENTOS
# ──────────────────────────────────────────────

def texto_primeiro(el, seletores) -> str | None:
    """
    Tenta extrair o texto do primeiro seletor CSS que funcionar no elemento.

    Itera pela lista de seletores em ordem; retorna o primeiro texto não-vazio
    encontrado. Útil porque diferentes versões/layouts de uma mesma página
    podem usar classes CSS distintas.

    Parâmetros:
        el        -- elemento pai onde buscar (WebElement)
        seletores -- lista de seletores CSS a tentar, em ordem de prioridade

    Retorna:
        Texto do elemento encontrado, ou None se nenhum funcionar.
    """
    for s in seletores:
        try:
            t = el.find_element(By.CSS_SELECTOR, s).text.strip()
            if t:
                return t  # Retorna o primeiro texto não-vazio encontrado
        except NoSuchElementException:
            continue  # Seletor não encontrado, tenta o próximo
    return None


def atributo_primeiro(el, seletores, atributo) -> str | None:
    """
    Tenta extrair um atributo HTML do primeiro seletor CSS que funcionar.

    Funciona da mesma forma que texto_primeiro(), mas para atributos
    como 'href', 'src', 'data-id' etc.

    Parâmetros:
        el        -- elemento pai onde buscar (WebElement)
        seletores -- lista de seletores CSS a tentar, em ordem de prioridade
        atributo  -- nome do atributo HTML a extrair (ex: 'href')

    Retorna:
        Valor do atributo encontrado, ou None se nenhum funcionar.
    """
    for s in seletores:
        try:
            v = el.find_element(By.CSS_SELECTOR, s).get_attribute(atributo)
            if v:
                return v  # Retorna o primeiro valor não-vazio encontrado
        except NoSuchElementException:
            continue
    return None


# ──────────────────────────────────────────────
# FUNÇÕES AUXILIARES — GERENCIAMENTO DA LISTA
# ──────────────────────────────────────────────

def adicionar(produtos, vistos, descricao, preco, loja, link) -> None:
    """
    Valida e adiciona um produto à lista, evitando duplicatas.

    Antes de inserir, verifica se:
        - A descrição e o link não estão vazios
        - O preço está dentro da faixa permitida (R$100 a R$150)
        - O produto (mesma loja + mesmo link) ainda não foi adicionado

    A URL é limpa (sem parâmetros de rastreamento ?utm=...) para
    comparação de duplicatas mais precisa.

    Parâmetros:
        produtos  -- lista onde o produto será inserido
        vistos    -- conjunto de chaves (loja, link) já processados
        descricao -- título do produto
        preco     -- preço como float
        loja      -- nome da loja de origem
        link      -- URL do produto
    """
    # Descarta produto se falta informação essencial ou preço fora da faixa
    if not descricao or not link or not preco_valido(preco):
        return

    # Remove parâmetros de URL (ex: ?utm_source=...) para evitar duplicatas
    link_limpo = link.split("?")[0]
    chave = (loja, link_limpo)

    # Se já foi visto este produto nesta loja, ignora
    if chave in vistos:
        return

    # Adiciona o produto com os dados padronizados
    produtos.append({
        "Descrição": re.sub(r"\s+", " ", descricao).strip(),  # Remove espaços extras
        "Valor": round(float(preco), 2),                        # Arredonda para 2 casas decimais
        "Loja": loja,
        "Link": link,
    })
    vistos.add(chave)  # Marca como visto para evitar duplicata futura


def remover_duplicados(produtos: list[dict]) -> list[dict]:
    """
    Remove produtos duplicados de uma lista consolidada.

    Usa a combinação (Loja + Link sem parâmetros) como chave única.
    Chamado após juntar resultados de múltiplas fontes.

    Retorna:
        Nova lista somente com produtos únicos, mantendo a ordem original.
    """
    unicos, vistos = [], set()
    for p in produtos:
        chave = (p["Loja"], p["Link"].split("?")[0])
        if chave in vistos:
            continue  # Produto repetido, ignora
        unicos.append(p)
        vistos.add(chave)
    return unicos


# ──────────────────────────────────────────────
# FUNÇÕES DE COLETA — MERCADO LIVRE
# ──────────────────────────────────────────────

def coletar_mercado_livre_ofertas(driver: WebDriver, limite: int) -> list[dict]:
    """
    Coleta produtos da página de Ofertas do Mercado Livre com paginação automática.

    Navega página a página clicando em 'Seguinte' até atingir o limite
    de produtos ou o máximo de páginas configurado.

    Parâmetros:
        driver -- instância ativa do WebDriver
        limite -- número máximo de produtos a coletar nesta função

    Retorna:
        Lista de dicionários com os dados dos produtos coletados.
    """
    loja = "Mercado Livre"
    produtos, vistos = [], set()
    url = "https://www.mercadolivre.com.br/ofertas"  # URL inicial da seção de ofertas

    # Seletores CSS dos cards de produto (múltiplas versões do layout do ML)
    seletor_cards = (
        "li.promotion-item, div.poly-card, "
        "div.ui-search-result__wrapper, li.ui-search-layout__item"
    )

    for pagina in range(1, MAX_PAGINAS + 1):
        if len(produtos) >= limite:
            break  # Já atingiu o limite desejado

        print(f"[{loja}] Página {pagina}: {url}")

        try:
            driver.get(url)  # Navega para a URL da página atual
            # Aguarda até que os cards de produto estejam presentes no DOM
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_cards))
            )
        except TimeoutException:
            # Se a página demorar mais de 20s para carregar, encerra a coleta
            print(f"[{loja}] Timeout na página {pagina}")
            break

        # Percorre todos os cards de produto encontrados na página
        for card in driver.find_elements(By.CSS_SELECTOR, seletor_cards):
            if len(produtos) >= limite:
                break  # Interrompe o loop interno se já atingiu o limite
            try:
                # Extrai o título do produto (tenta diferentes seletores por compatibilidade)
                descricao = texto_primeiro(card, [
                    ".promotion-item__title", ".poly-component__title",
                    ".ui-search-item__title", "h2",
                ])

                # Extrai a parte inteira do preço (ex: '129' de 'R$ 129,90')
                preco_t = texto_primeiro(card, [
                    ".andes-money-amount__fraction",
                    ".promotion-item__price", ".price-tag-fraction",
                ])

                # Extrai os centavos do preço (ex: '90' de 'R$ 129,90')
                centavos = texto_primeiro(card, [
                    ".andes-money-amount__cents", ".price-tag-cents",
                ])

                # Monta o preço completo 'inteiro,centavos' se ainda não tiver vírgula
                if preco_t and centavos and "," not in preco_t:
                    preco_t = f"{preco_t},{centavos}"

                # Extrai o link do produto
                link = atributo_primeiro(card, ["a[href]"], "href")

                # Valida e adiciona o produto à lista
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)

            except (NoSuchElementException, WebDriverException):
                continue  # Card com estrutura inesperada; pula e continua

        print(f"[{loja}] Acumulado: {len(produtos)}")

        # Tenta encontrar o link da próxima página
        proxima = None
        for s in ("a.andes-pagination__link[title='Seguinte']", "a[aria-label='Seguinte']"):
            try:
                proxima = driver.find_element(By.CSS_SELECTOR, s).get_attribute("href")
                if proxima:
                    break
            except NoSuchElementException:
                continue

        # Se não encontrou próxima página (ou é a mesma URL), encerra a paginação
        if not proxima or proxima == url:
            print(f"[{loja}] Fim da paginação")
            break

        url = proxima
        time.sleep(random.uniform(1.5, 3))  # Delay aleatório para simular comportamento humano

    return produtos


def coletar_mercado_livre_busca(driver: WebDriver, limite: int, termos: list[str]) -> list[dict]:
    """
    Busca produtos no Mercado Livre usando múltiplos termos para complementar o volume.

    Chamada somente se as fontes principais não atingirem o mínimo de registros.
    Para cada termo, acessa a URL de busca e extrai os produtos da primeira página.

    Parâmetros:
        driver -- instância ativa do WebDriver
        limite -- número máximo de produtos a coletar no total
        termos -- lista de palavras-chave para pesquisar

    Retorna:
        Lista de dicionários com os dados dos produtos coletados.
    """
    loja = "Mercado Livre"
    produtos, vistos = [], set()

    # Seletores compatíveis com diferentes versões da página de resultados do ML
    seletor_cards = (
        "li.ui-search-layout__item, div.ui-search-result__wrapper, div.poly-card"
    )

    for termo in termos:
        if len(produtos) >= limite:
            break  # Já coletou o suficiente

        # Monta a URL de busca codificando o termo para uso em URL (ex: espaços viram %20)
        url = f"https://lista.mercadolivre.com.br/{quote_plus(termo)}"
        print(f"[{loja} Busca] '{termo}': {url}")

        try:
            driver.get(url)

            # Verifica se a página foi bloqueada antes de tentar extrair dados
            if pagina_bloqueada(driver):
                print(f"[{loja} Busca] Bloqueado em '{termo}', pulando")
                continue

            # Aguarda os cards de resultado aparecerem no DOM
            WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, seletor_cards))
            )
        except TimeoutException:
            print(f"[{loja} Busca] Timeout em '{termo}'")
            continue  # Pula para o próximo termo

        antes = len(produtos)  # Quantidade antes de processar esta página (para log)

        for card in driver.find_elements(By.CSS_SELECTOR, seletor_cards):
            if len(produtos) >= limite:
                break
            try:
                # Extrai título do produto
                descricao = texto_primeiro(card, [
                    ".ui-search-item__title", ".poly-component__title", "h2",
                ])

                # Extrai parte inteira e centavos do preço
                preco_t = texto_primeiro(card, [
                    ".andes-money-amount__fraction", ".price-tag-fraction",
                ])
                centavos = texto_primeiro(card, [
                    ".andes-money-amount__cents", ".price-tag-cents",
                ])

                # Monta o preço completo se necessário
                if preco_t and centavos and "," not in preco_t:
                    preco_t = f"{preco_t},{centavos}"

                link = atributo_primeiro(card, ["a[href]"], "href")
                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)

            except (NoSuchElementException, WebDriverException):
                continue

        # Exibe quantos produtos foram coletados com este termo específico
        print(f"[{loja} Busca] +{len(produtos) - antes}, total: {len(produtos)}")
        time.sleep(random.uniform(2, 4))  # Delay antes de buscar o próximo termo

    return produtos


# ──────────────────────────────────────────────
# FUNÇÕES DE COLETA — FONTE 2 (Amazon + Fallbacks)
# ──────────────────────────────────────────────

def coletar_amazon(driver: WebDriver, limite: int) -> list[dict]:
    """
    Coleta produtos da Amazon Brasil com filtro de preço aplicado na URL.

    Usa o parâmetro 'rh=p_36:10000-15000' que filtra produtos entre
    R$100 e R$150 diretamente na busca da Amazon (valores em centavos).

    Parâmetros:
        driver -- instância ativa do WebDriver
        limite -- número máximo de produtos a coletar

    Retorna:
        Lista de dicionários com os dados dos produtos coletados.
    """
    loja = "Amazon"
    produtos, vistos = [], set()

    # URL base: busca por 'presentes' com filtro de preço R$100-R$150
    # rh=p_36:10000-15000 → filtro de preço em centavos (100,00 a 150,00)
    base = "https://www.amazon.com.br/s?k=presentes&rh=p_36%3A10000-15000"

    for pagina in range(1, MAX_PAGINAS + 1):
        if len(produtos) >= limite:
            break

        url = f"{base}&page={pagina}"  # Adiciona o número da página à URL
        print(f"[{loja}] Página {pagina}: {url}")

        try:
            driver.get(url)
            time.sleep(random.uniform(2, 4))  # Delay antes de verificar o conteúdo
            scroll_gradual(driver)             # Rola a página para carregar todos os itens

            # Verifica se a Amazon bloqueou o acesso
            if pagina_bloqueada(driver):
                print(f"[{loja}] Bloqueado")
                break

            # Aguarda os cards de resultado da busca aparecerem
            WebDriverWait(driver, 20).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "div[data-component-type='s-search-result']")
                )
            )
        except TimeoutException:
            print(f"[{loja}] Timeout página {pagina}")
            # Não interrompe — tenta processar o que foi carregado até o momento

        # Coleta todos os cards de produto visíveis na página
        cards = driver.find_elements(
            By.CSS_SELECTOR, "div[data-component-type='s-search-result']"
        )
        if not cards:
            break  # Nenhum resultado encontrado; provavelmente fim das páginas

        for card in cards:
            if len(produtos) >= limite:
                break
            try:
                # Extrai o título do produto (diferentes seletores por tamanho de texto)
                descricao = texto_primeiro(
                    card, ["h2 span", "span.a-size-base-plus", "span.a-size-medium"]
                )

                # A Amazon separa o preço em parte inteira e centavos em elementos distintos
                preco_int = texto_primeiro(card, ["span.a-price-whole"])
                preco_cent = texto_primeiro(card, ["span.a-price-fraction"]) or "00"

                # Monta o preço no formato '129,90' para passar pela função limpar_preco()
                preco_t = f"{preco_int},{preco_cent}" if preco_int else None

                # Extrai o link relativo e converte para URL absoluta
                link = atributo_primeiro(
                    card, ["h2 a[href]", "a.a-link-normal[href]"], "href"
                )
                link = urljoin("https://www.amazon.com.br", link) if link else None

                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)

            except Exception:
                continue  # Item com estrutura incomum; pula e continua

        print(f"[{loja}] Acumulado: {len(produtos)}")

        # Verifica se o botão "Próxima página" está desabilitado (fim da paginação)
        if driver.find_elements(
            By.CSS_SELECTOR, ".s-pagination-next.s-pagination-disabled"
        ):
            print(f"[{loja}] Fim da paginação")
            break

    return produtos


def coletar_magalu(driver: WebDriver, limite: int) -> list[dict]:
    """
    Coleta produtos do Magazine Luiza (fallback da Fonte 2).

    Ativado somente se a Amazon bloquear o scraping. O Magalu historicamente
    também pode retornar erros 403, por isso é tratado como fallback.

    Parâmetros:
        driver -- instância ativa do WebDriver
        limite -- número máximo de produtos a coletar

    Retorna:
        Lista de dicionários com os dados dos produtos coletados.
    """
    loja = "Magalu"
    produtos, vistos = [], set()

    # Seletores do Magalu (usa atributos data-testid para identificar os cards)
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
            time.sleep(random.uniform(2, 4))  # Aguarda carregamento inicial
            scroll_gradual(driver)            # Garante carregamento lazy

            # Verifica bloqueio antes de tentar extrair dados
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
            break  # Provavelmente chegou ao fim das páginas

        for card in cards:
            if len(produtos) >= limite:
                break
            try:
                # Extrai o título do produto
                descricao = texto_primeiro(
                    card, ["[data-testid='product-title']", "h2", "p"]
                )

                # Extrai o preço (tenta diferentes variantes do seletor)
                preco_t = texto_primeiro(card, [
                    "[data-testid='price-value']",
                    "[data-testid='price-original']",
                    "p[data-testid*='price']",
                ])

                # O Magalu pode ter o link diretamente no card (se for um <a>) ou num filho
                link = card.get_attribute("href") or atributo_primeiro(
                    card, ["a[href]"], "href"
                )
                # Converte link relativo para URL absoluta
                link = urljoin("https://www.magazineluiza.com.br", link) if link else None

                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)

            except Exception:
                continue

        print(f"[{loja}] Acumulado: {len(produtos)}")

    return produtos


def coletar_shopee(driver: WebDriver, limite: int) -> list[dict]:
    """
    Coleta produtos da Shopee (segundo fallback da Fonte 2).

    Ativado somente se tanto Amazon quanto Magalu bloquearem o scraping.
    A Shopee carrega conteúdo via JavaScript e pode exibir um seletor
    de idioma na primeira visita.

    Parâmetros:
        driver -- instância ativa do WebDriver
        limite -- número máximo de produtos a coletar

    Retorna:
        Lista de dicionários com os dados dos produtos coletados.
    """
    loja = "Shopee"
    produtos, vistos = [], set()

    # Monta URL de busca por 'presentes' na Shopee
    url = f"https://shopee.com.br/search?keyword={quote_plus('presentes')}"
    print(f"[{loja}] {url}")

    try:
        driver.get(url)
        time.sleep(5)  # A Shopee leva mais tempo para renderizar o JavaScript inicial

        try:
            # Tenta fechar o seletor de idioma se aparecer (primeira visita)
            botoes = driver.find_elements(
                By.XPATH, "//*[contains(text(), 'Português (BR)')]"
            )
            if botoes:
                botoes[0].click()  # Clica em 'Português (BR)' para fechar o modal
                time.sleep(3)      # Aguarda o modal fechar e a página recarregar
        except Exception:
            pass  # Se o seletor não aparecer, continua normalmente

        # Verifica se a página foi bloqueada
        if pagina_bloqueada(driver):
            print(f"[{loja}] Bloqueado")
            return []

        # Aguarda os cards de produto aparecerem
        WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((
                By.CSS_SELECTOR,
                "a[data-sqe='link'], div.shopee-search-item-result__item",
            ))
        )

        # Rola a página em 8 passos para garantir carregamento de todos os itens lazy
        scroll_gradual(driver, passos=8, delay=0.6)

        for card in driver.find_elements(By.CSS_SELECTOR, "a[data-sqe='link']"):
            if len(produtos) >= limite:
                break
            try:
                # Extrai o nome do produto
                descricao = texto_primeiro(
                    card, ["[data-sqe='name']", "div[class*='title']"]
                )

                # Extrai o preço (Shopee usa classes dinâmicas com hash, por isso usa *=)
                preco_t = texto_primeiro(
                    card, ["div[class*='price']", "span.price"]
                )

                # O link está diretamente no elemento <a> pai
                link = card.get_attribute("href")

                adicionar(produtos, vistos, descricao, limpar_preco(preco_t), loja, link)

            except Exception:
                continue

    except Exception as e:
        print(f"[{loja}] Erro: {e}")

    return produtos


def coletar_fonte2(limite: int) -> list[dict]:
    """
    Orquestra a coleta da Fonte 2 com cadeia de fallback automático.

    Tenta as fontes na ordem: Amazon -> Magalu -> Shopee.
    Se uma fonte retornar produtos suficientes, as seguintes não são ativadas.
    Cada fonte usa sua própria instância do driver (abertura e fechamento isolados).

    Parâmetros:
        limite -- número máximo de produtos a coletar no total

    Retorna:
        Lista consolidada de produtos das fontes que foram ativadas.
    """
    # Define a cadeia de fallback: (nome, usar_stealth, função_de_coleta)
    cadeia = [
        ("Amazon", False, coletar_amazon),   # Fonte principal da Fonte 2
        ("Magalu", True, coletar_magalu),    # Fallback 1: stealth porque bloqueia frequentemente
        ("Shopee", True, coletar_shopee),    # Fallback 2: stealth pela proteção anti-bot
    ]

    produtos = []

    for nome, stealth, funcao in cadeia:
        if len(produtos) >= limite:
            break  # Já coletou o suficiente; não precisa acionar as próximas fontes

        restante = limite - len(produtos)  # Quantos produtos ainda faltam
        print(f"\n[Fonte 2] Tentando {nome}... ({restante} produtos a coletar)")

        driver = None
        try:
            driver = iniciar_driver(stealth=stealth)      # Abre nova instância do browser
            coletados = funcao(driver, restante)           # Executa a coleta nesta fonte

            if coletados:
                print(f"[Fonte 2] {nome}: {len(coletados)} produtos coletados")
                produtos.extend(coletados)
            else:
                # Fonte não retornou produtos; tenta a próxima da cadeia
                print(f"[Fonte 2] {nome} não retornou produtos. Próxima opção.")

        except Exception as e:
            print(f"[Fonte 2] Erro em {nome}: {e}")
        finally:
            # Garante que o driver seja fechado mesmo em caso de erro
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    return produtos


# ──────────────────────────────────────────────
# FUNÇÕES DE SAÍDA — EXCEL E RESUMO
# ──────────────────────────────────────────────

def exportar_excel(produtos: list[dict], arquivo: str = ARQUIVO_SAIDA) -> None:
    """
    Exporta a lista de produtos para um arquivo Excel (.xlsx).

    Cria um DataFrame do pandas com as colunas na ordem correta
    e salva usando openpyxl como engine, sem incluir o índice do DataFrame.

    Parâmetros:
        produtos -- lista de dicionários com os dados dos produtos
        arquivo  -- nome (ou caminho) do arquivo de saída
    """
    # Cria o DataFrame garantindo a ordem das colunas
    df = pd.DataFrame(produtos, columns=["Descrição", "Valor", "Loja", "Link"])

    # Salva em formato .xlsx (engine openpyxl é necessário para este formato)
    df.to_excel(arquivo, engine="openpyxl", index=False)

    print(f"\n[OK] Excel exportado: {arquivo}")


def imprimir_resumo(produtos: list[dict]) -> None:
    """
    Exibe no console um resumo dos produtos coletados agrupados por loja.

    Usa Counter para contar quantos produtos vieram de cada loja
    e imprime o total por loja e o total geral.

    Parâmetros:
        produtos -- lista final de produtos após deduplicação
    """
    # Conta quantos produtos vieram de cada loja
    contagem = Counter(p["Loja"] for p in produtos)

    print("\n================== RESUMO ==================")
    for loja, total in contagem.items():
        print(f"  {loja}: {total} produto(s)")
    print(f"  Total Geral: {len(produtos)} produto(s)")
    print("============================================")


# ──────────────────────────────────────────────
# FUNÇÃO PRINCIPAL — ORQUESTRAÇÃO
# ──────────────────────────────────────────────

def main() -> None:
    """
    Função principal que orquestra todas as etapas do web scraping.

    Fluxo de execução:
        1. Coleta da Fonte 1: Mercado Livre Ofertas
        2. Coleta da Fonte 2: Amazon -> Magalu -> Shopee (com fallback)
        3. Remove duplicatas da lista consolidada
        4. Complemento via busca no ML se ainda não atingiu o mínimo
        5. Trunca a lista se ultrapassou o máximo
        6. Exporta para Excel e imprime o resumo
    """
    # Cabeçalho informativo exibido ao iniciar o script
    print("=" * 56)
    print("ROBÔ DE WEB SCRAPING - PRESENTES VAREJO")
    print(f"Filtro: R$ {PRECO_MIN:.2f} a R$ {PRECO_MAX:.2f}")
    print(f"Volume: mínimo {MIN_REGISTROS} / máximo {MAX_REGISTROS}")
    print("=" * 56)

    produtos: list[dict] = []  # Lista global que acumula todos os produtos

    # ── ETAPA 1: Mercado Livre — Página de Ofertas ──
    print("\n--- FONTE 1: Mercado Livre (Ofertas) ---")
    driver = None
    try:
        driver = iniciar_driver(stealth=False)  # Modo padrão (sem stealth) para o ML

        # Coleta até metade do máximo para deixar espaço para a Fonte 2
        produtos.extend(coletar_mercado_livre_ofertas(driver, MAX_REGISTROS // 2))

    except Exception as e:
        print(f"[Erro ML Ofertas] {e}")
    finally:
        # Fecha o driver mesmo em caso de erro (requisito obrigatório do trabalho)
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

    # ── ETAPA 2: Amazon com fallback Magalu/Shopee ──
    # Calcula quantos produtos ainda precisam ser coletados
    restante = max(MAX_REGISTROS - len(produtos), MIN_REGISTROS - len(produtos))

    if restante > 0:
        print(f"\n--- FONTE 2: Amazon (fallback Magalu/Shopee) | {restante} a coletar ---")
        produtos.extend(coletar_fonte2(restante))

    # ── ETAPA 3: Remove duplicatas da lista consolidada ──
    produtos = remover_duplicados(produtos)

    # ── ETAPA 4: Complemento por busca no ML (só se não atingiu o mínimo) ──
    if len(produtos) < MIN_REGISTROS:
        falta = MIN_REGISTROS - len(produtos)
        print(f"\n--- COMPLEMENTO: Busca ML por termos | faltam {falta} ---")
        driver = None
        try:
            driver = iniciar_driver(stealth=False)

            # Busca pelos termos definidos em TERMOS_ML até completar o mínimo
            produtos.extend(coletar_mercado_livre_busca(driver, falta, TERMOS_ML))

            # Remove duplicatas novamente após adicionar os resultados da busca
            produtos = remover_duplicados(produtos)

        except Exception as e:
            print(f"[Erro Complemento ML] {e}")
        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

    # ── ETAPA 5: Trunca a lista se ultrapassou o máximo permitido ──
    if len(produtos) > MAX_REGISTROS:
        produtos = produtos[:MAX_REGISTROS]  # Mantém apenas os primeiros MAX_REGISTROS

    # ── AVISO: Informa se não atingiu o mínimo (possível bloqueio de rede) ──
    if len(produtos) < MIN_REGISTROS:
        print(
            f"\n[Aviso] Atingiu {len(produtos)}/{MIN_REGISTROS} - possíveis bloqueios de rede."
        )

    # ── ETAPA 6: Exporta e exibe o resumo ──
    if produtos:
        exportar_excel(produtos)    # Gera o arquivo presentes_varejo.xlsx
        imprimir_resumo(produtos)   # Exibe totais por loja no console
    else:
        print("\n[Falha] Nenhum produto coletado.")


# Ponto de entrada do script: executa main() somente quando rodado diretamente
# (não executa se o arquivo for importado como módulo por outro script)
if __name__ == "__main__":
    main()

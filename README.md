# Web Scraping — Lista de Presentes (Varejo)

Trabalho acadêmico de Web Scraping em Python utilizando Selenium WebDriver para coleta automatizada de produtos de e-commerce dentro de uma faixa de preço.

---

## Objetivo

Navegar automaticamente por páginas de e-commerce, percorrer a paginação, aplicar filtro de preço e extrair dados estruturados, exportando o resultado para um arquivo Excel (`.xlsx`).

---

## Fontes de Dados

| Fonte | URL | Papel |
|---|---|---|
| **Mercado Livre — Ofertas** | `mercadolivre.com.br/ofertas` | Fonte principal (Fonte 1) |
| **Mercado Livre — Busca** | `lista.mercadolivre.com.br/{termo}` | Complemento caso não atinja o mínimo |
| **Amazon Brasil** | `amazon.com.br` | Fonte 2 |
| **Magazine Luiza** | `magazineluiza.com.br` | Fallback 1 (se Amazon bloquear) |
| **Shopee** | `shopee.com.br` | Fallback 2 (se Magalu bloquear) |

---

## Filtros e Volume

- **Faixa de preço:** R$ 100,00 a R$ 150,00
- **Mínimo de registros:** 300
- **Máximo de registros:** 2.000
- **Limite de páginas por fonte:** 40

---

## Dados Extraídos

Cada produto coletado possui quatro campos:

| Campo | Tipo | Descrição |
|---|---|---|
| `Descrição` | `str` | Título do produto na loja |
| `Valor` | `float` | Preço atual em reais |
| `Loja` | `str` | Origem do produto (Mercado Livre / Amazon / Magalu / Shopee) |
| `Link` | `str` | URL direta do produto |

---

## Requisitos

- Python 3.10+
- Google Chrome instalado na máquina

### Instalação das dependências

```bash
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install selenium webdriver-manager pandas openpyxl undetected-chromedriver
```

---

## Como Executar

```bash
python scraping_presentes_varejo.py
```

O script abre o navegador automaticamente (modo visível por padrão), percorre as páginas e, ao finalizar, gera o arquivo `presentes_varejo.xlsx` na mesma pasta.

---

## Estrutura do Projeto

```
TrabalhoWebScraping/
│
├── scraping_presentes_varejo.py   # Script principal de web scraping
├── prompt_web_scraping_selenium.md # Especificação / prompt da atividade
├── requirements.txt                # Dependências do projeto
├── README.md                       # Este arquivo
└── .gitignore                      # Arquivos ignorados pelo git
```

---

## Arquitetura do Script

```
main()
 ├── FONTE 1 → coletar_mercado_livre_ofertas()
 │               └── paginação automática até o limite
 │
 ├── FONTE 2 → coletar_fonte2()  [cadeia de fallback]
 │               ├── coletar_amazon()
 │               ├── coletar_magalu()   [fallback 1]
 │               └── coletar_shopee()   [fallback 2]
 │
 ├── COMPLEMENTO → coletar_mercado_livre_busca()
 │                  └── busca por termos se mínimo não atingido
 │
 ├── remover_duplicados()
 ├── exportar_excel()
 └── imprimir_resumo()
```

### Principais Funções

| Função | Responsabilidade |
|---|---|
| `iniciar_driver()` | Inicia o ChromeDriver (modo normal ou stealth anti-bot) |
| `pagina_bloqueada()` | Detecta bloqueios HTTP 403, captcha e Cloudflare |
| `limpar_preco()` | Converte `"R$ 1.234,56"` em `1234.56` |
| `preco_valido()` | Verifica se o preço está na faixa R$100–R$150 |
| `scroll_gradual()` | Rola a página aos poucos para carregar conteúdo lazy |
| `adicionar()` | Valida e insere produto na lista, evitando duplicatas |
| `remover_duplicados()` | Remove duplicatas por URL antes da exportação |
| `exportar_excel()` | Salva DataFrame no arquivo `.xlsx` |
| `imprimir_resumo()` | Exibe totais por loja no console |

---

## Tecnologias Utilizadas

| Biblioteca | Uso |
|---|---|
| `selenium` | Automação do navegador Chrome |
| `webdriver-manager` | Gerencia o ChromeDriver automaticamente |
| `pandas` | Estrutura os dados em DataFrame |
| `openpyxl` | Engine para exportação `.xlsx` |
| `undetected-chromedriver` | Modo stealth para contornar bloqueios anti-bot |

---

## Tratamento de Erros e Anti-bloqueio

- **Esperas explícitas** com `WebDriverWait` + `expected_conditions` em vez de `sleep` fixo
- **Detecção de bloqueio** (403, captcha, Cloudflare) com fallback automático para a próxima fonte
- **Exceções tratadas** (`NoSuchElementException`, `TimeoutException`, `WebDriverException`) — itens mal formatados são pulados sem interromper o robô
- **User-Agent realista** rotacionado entre Mac e Windows
- **Delays aleatórios** entre requisições para simular comportamento humano
- **Modo stealth** (undetected-chromedriver) ativado nas fontes mais restritivas (Magalu, Shopee)
- **`driver.quit()` garantido** em bloco `try/finally` mesmo em caso de erro

---

## Saída Esperada no Console

```
========================================================
ROBÔ DE WEB SCRAPING - PRESENTES VAREJO
Filtro: R$ 100.00 a R$ 150.00
Volume: mínimo 300 / máximo 2000
========================================================

--- FONTE 1: Mercado Livre (Ofertas) ---
[Mercado Livre] Página 1: https://www.mercadolivre.com.br/ofertas
[Mercado Livre] Acumulado: 12
...

--- FONTE 2: Amazon (fallback Magalu/Shopee) ---
[Amazon] Página 1: ...
...

================== RESUMO ==================
  Mercado Livre: 210 produto(s)
  Amazon: 190 produto(s)
  Total Geral: 400 produto(s)
============================================

[OK] Excel exportado: presentes_varejo.xlsx
```

---

## Observações

- O arquivo `presentes_varejo.xlsx` gerado **não é versionado** (listado no `.gitignore`).
- O ambiente virtual `venv/` também não é versionado.
- Alguns sites (Amazon, Shopee) podem bloquear o scraping dependendo da rede/IP — nesse caso o robô ativa automaticamente o fallback.
- O script foi desenvolvido e testado no macOS com Python 3.14 e Google Chrome 125+.

---

## Disciplina

Análise e Desenvolvimento de Sistemas — Programação em Python
Trabalho: Web Scraping com Selenium

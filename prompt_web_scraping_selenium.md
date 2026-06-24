# Prompt para o Claude Code — Web Scraping com Selenium

> Cole o conteúdo abaixo (a partir de "INÍCIO DO PROMPT") no Claude Code.

---

## INÍCIO DO PROMPT

Crie uma solução completa de **Web Scraping em Python** para uma atividade acadêmica, seguindo rigorosamente as especificações abaixo. Gere todo o código em um único script bem comentado e pronto para rodar.

### Objetivo
Navegar automaticamente por páginas de e-commerce, percorrer a paginação, aplicar filtro de preço e extrair dados estruturados, exportando o resultado para um arquivo Excel.

### Tema — Lista de Presentes (Varejo)
- **Fonte 1:** Mercado Livre (seção de **Ofertas** — `https://www.mercadolivre.com.br/ofertas`)
- **Fonte 2:** Amazon Brasil (use **Magalu** como fallback caso a Amazon bloqueie o scraping)
- **Filtro:** apenas produtos com valor entre **R$ 100,00 e R$ 150,00** (filtrar no código, validando o preço extraído)
- **Dados a extrair por produto:**
  - `Descrição` (título do produto)
  - `Valor` (preço atual, em float)
  - `Loja` (Mercado Livre / Amazon / Magalu)
  - `Link` (URL direta do produto)

### Requisitos técnicos obrigatórios
- Usar **Selenium WebDriver** para a extração.
- Usar **`webdriver-manager`** para gerenciar o ChromeDriver automaticamente.
- Usar **Pandas** para estruturar os dados em um DataFrame.
- Usar **`openpyxl`** como engine para gerar o arquivo `.xlsx`.
- Incluir no topo do script um comentário com o comando de instalação:
  `pip install selenium webdriver-manager pandas openpyxl`

### Regras de execução
- **Volume de dados:** mínimo **300** e máximo **2000** registros, somando as duas fontes. O robô deve **percorrer a paginação** até atingir o volume desejado, com um **limite máximo de páginas** para evitar loop infinito.
- O navegador **deve fechar garantidamente ao final** — usar `try/finally` com `driver.quit()`.
- Tratar exceções de elementos não encontrados (`NoSuchElementException`, `TimeoutException`) para que o robô não pare se um item estiver mal formatado; nesses casos, pular o item e continuar.
- Usar **esperas explícitas** (`WebDriverWait` + `expected_conditions`) em vez de `time.sleep()` fixo sempre que possível.
- Limpar e converter o preço (remover "R$", pontos de milhar e converter vírgula decimal) antes de aplicar o filtro.
- Imprimir no console o progresso (página atual, quantidade acumulada de registros).

### Estrutura do Excel
- Arquivo de saída: `presentes_varejo.xlsx`
- Cabeçalhos claros: **Descrição | Valor | Loja | Link**
- Salvar o DataFrame final via `df.to_excel(..., engine="openpyxl", index=False)`.
- Ao final, imprimir o total de registros extraídos por loja e o total geral.

### Organização do código
- Funções separadas por responsabilidade (ex.: `iniciar_driver()`, `coletar_mercado_livre()`, `coletar_amazon()`, `limpar_preco()`, `exportar_excel()`).
- Bloco `if __name__ == "__main__":` para orquestrar a execução.
- Comentários explicativos em português.

## FIM DO PROMPT

# Script para Captura de Cotação do Dólar e Geração de Relatório

Este script realiza a captura da cotação do dólar de um site específico, gera um relatório em formato Word com as informações extraídas e converte esse relatório para PDF. O script utiliza a biblioteca `selenium` para automação do navegador, `docx` para manipulação de documentos Word e `docx2pdf` para conversão para PDF.

## Download do Arquivo

O arquivo executável pode ser baixado a partir do seguinte link: [Download](https://drive.google.com/file/d/1FCq95Cj4u9I9FbBfAccA1qG8VM11WkDR/view?usp=sharing)


## Pré-requisitos

Antes de executar o script, certifique-se de que você tenha os seguintes pré-requisitos instalados:

- **Python 3.x**: O código é escrito em Python e requer Python 3.x.
- **Bibliotecas Python**: Você precisa instalar as bibliotecas necessárias. Instale as dependências usando o arquivo 'requirements.txt':

    ```bash
    pip install -r requirements.txt
    ```

- **ChromeDriver**: O `selenium` requer o ChromeDriver para interagir com o navegador Google Chrome. Baixe o ChromeDriver compatível com sua versão do Chrome e coloque-o em um diretório incluído no PATH do sistema.

## Funcionamento do Código

O script é dividido nas seguintes partes principais:

### 1. Funções

- **`obter_data_hora()`**: Obtém a data e hora no fuso horário de São Paulo e retorna no formato `dd/mm/aaaa`. Caso ocorra um erro, retorna uma mensagem padrão.

- **`fazer_print(driver)`**: Faz um print da tela do navegador e salva a imagem como `screenshot.png`.

- **`extrair_informacoes(driver)`**: Extrai a cotação do dólar, data/hora e URL da página atual. Utiliza o XPath para localizar o elemento da cotação. Retorna essas informações ou valores padrão caso o elemento não seja encontrado.

- **`criar_word(titulo, cotacao, data, url, caminho_imagem, autor)`**: Cria um documento Word com o título, cotação, data, URL e uma imagem do print. Formata o título e adiciona um parágrafo com o link para o site.

- **`converter_word_para_pdf(arquivo_docx, arquivo_pdf)`**: Converte o documento Word gerado para PDF e salva como `cotacao_dolar.pdf`.

### 2. Função Principal

- **`main()`**: Configura o WebDriver do Chrome, acessa o site da cotação do dólar, extrai informações, faz um print da tela, cria o documento Word e converte-o para PDF. Por fim, fecha o navegador.

## Execução

Para executar o script, salve o código em um arquivo chamado `cotacao_dolar.py` e execute o seguinte comando no terminal:

```bash
python cotacao_dolar.py
```

Certifique-se de que o ChromeDriver está disponível no PATH e que todas as bibliotecas Python necessárias estão instaladas.

## Mensagens de Progresso

Durante a execução, o script imprime mensagens para informar o usuário sobre o andamento das tarefas:

- Início do processo
- Acesso ao site
- Extração de informações
- Criação do documento Word
- Conversão para PDF
- Finalização do processo

Essas mensagens ajudam a acompanhar o progresso e identificar possíveis problemas.

## Notas

- O XPath utilizado no script pode precisar ser ajustado se o layout do site mudar.
- O script salva o print da tela e os arquivos gerados no mesmo diretório onde o script é executado.

Para mais informações, consulte a [documentação do Selenium](https://www.selenium.dev/pt-br/documentation/) e da [biblioteca python-docx](https://python-docx.readthedocs.io/en/latest/).

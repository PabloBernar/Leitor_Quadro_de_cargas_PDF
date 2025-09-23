# Leitor de Cargas de PDF

Uma aplicação desktop desenvolvida em Python com `customtkinter` para extrair dados de arquivos PDF de cargas e exportá-los para o formato CSV. Ideal para automatizar a coleta de informações de documentos padronizados.

## Funcionalidades

*   **Seleção Múltipla de PDFs:** Permite selecionar um ou mais arquivos PDF para processamento em lote.
*   **Extração de Dados:** Analisa o conteúdo dos PDFs para identificar e extrair informações relevantes sobre cargas (Tipo de aparelho, Subtipo, Datas de Início/Fim, Dias, Oper, Quantidade de Aparelhos, Horas, Potência, DIC, Qtd. Fat).
*   **Exportação para CSV:** Salva os dados extraídos em arquivos CSV, facilitando a análise e integração com outras ferramentas.
*   **Interface Moderna:** Utiliza `customtkinter` para uma experiência de usuário moderna e responsiva.
*   **Processamento em Segundo Plano:** A extração de PDFs é realizada em uma thread separada para manter a interface do usuário responsiva.

## Como Usar

1.  **Instalação das Dependências:**
    Certifique-se de ter o Python instalado. Em seguida, instale as bibliotecas necessárias:
    ```bash
    pip install customtkinter pdfplumber pandas
    ```

2.  **Executar a Aplicação:**
    Navegue até o diretório onde o arquivo `extrator_pdf.py` está salvo e execute-o:
    ```bash
    python extrator_pdf.py
    ```

3.  **Interface da Aplicação:**
    *   **"Selecionar PDF(s)"**: Clique neste botão para abrir uma janela de diálogo e escolher os arquivos PDF que deseja processar. Você pode selecionar múltiplos arquivos.
    *   **"Extrair para CSV"**: Após selecionar os PDFs, clique neste botão. Será solicitado que você escolha uma pasta de destino onde os arquivos CSV resultantes serão salvos.
    *   **Log Box**: Acompanhe o progresso da extração e quaisquer mensagens de erro ou sucesso na caixa de log na parte inferior da janela.

## Estrutura do Projeto

*   `extrator_pdf.py`: O script principal da aplicação, contendo a lógica da interface gráfica e as funções de extração de PDF.

## Contato

Desenvolvido por @pablo.bernar. Conecte-se no LinkedIn:
[https://www.linkedin.com/in/pablo-bernar/](https://www.linkedin.com/in/pablo-bernar/)

# DocCheck_OS_PDF 📋🚀

Uma ferramenta automatizada em Python para auditar, validar assinaturas, gerar relatórios e organizar (renomear) arquivos PDF de Ordens de Serviço (OS).

## 📌 Sobre o Projeto

Este script foi desenvolvido para otimizar o fluxo de trabalho administrativo, eliminando a necessidade de abrir e verificar manualmente centenas de arquivos PDF. Ele analisa o conteúdo textual e visual dos documentos para garantir que as **Ordens de Serviço** estejam preenchidas corretamente e assinadas, gerando um relatório detalhado em Excel e oferecendo a opção de padronizar a nomenclatura dos arquivos.

## ✨ Funcionalidades Principais

* **🔍 Extração Inteligente de Dados:**
    * Identifica automaticamente **Matrícula** (via Regex), **Nome** e **Função** do funcionário.
    * Verifica se a "Descrição da Função" está preenchida corretamente.
* **✍️ Validação Avançada de Assinaturas:**
    * Detecta assinaturas digitais (DocuSign, ICP-Brasil, etc.).
    * Detecta assinaturas manuais (Tablet/Caneta) identificando anotações do tipo *Ink*, *Stamp* e **vetores curvos (desenhos)**.
* **📊 Relatórios em Excel:**
    * Gera automaticamente o arquivo `Relatorio_Completo_OS.xlsx` com o status de cada documento.
* **files Organização de Arquivos (Renomeação):**
    * Funcionalidade interativa ao final do processo.
    * Renomeia os arquivos para o padrão: `MATRICULA - NomeOriginal.pdf`.
    * Marca arquivos problemáticos com o prefixo `ERROR -`.

## 🛠️ Tecnologias Utilizadas

* [Python 3.x](https://www.python.org/)
* [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) - Para leitura robusta de PDFs e análise vetorial.
* [Pandas](https://pandas.pydata.org/) - Para estruturação de dados e exportação para Excel.
* **OS/Re** - Bibliotecas nativas para manipulação de sistema e expressões regulares.

## ⚙️ Pré-requisitos e Instalação

1. **Clone o repositório:**

   git clone [https://github.com/seu-usuario/OS-Auditor-Manager.git](https://github.com/seu-usuario/OS-Auditor-Manager.git)
   cd OS-Auditor-Manager


2.  **Instale as dependências:**

    ```
    pip install pymupdf pandas openpyxl
    ```

3.  **Configuração (Opcional):**
    No início do script `main.py`, você pode ajustar as constantes:

      * `MATRICULA_MIN` e `MATRICULA_MAX` (Intervalo de matrículas válidas).
      * `IGNORAR_VALIDACAO_ASSINATURA` (Para fins de teste).

## 🚀 Como Usar

1.  Coloque o script na mesma pasta onde estão os arquivos **.pdf** das Ordens de Serviço.
2.  Execute o script:
    ```bash
    python main.py
    ```
3.  O script irá processar todos os arquivos e gerar o `Relatorio_Completo_OS.xlsx`.
4.  Ao final, ele perguntará no terminal:
    > *"Deseja renomear os arquivos conforme as matrículas encontradas? (S/N)"*
5.  Digite `S` para confirmar a renomeação automática baseada nos dados extraídos.

## 📝 Licença

Este projeto está sob a licença MIT. Sinta-se à vontade para contribuir\!

```

### Dica Extra para o GitHub:
Na área de **"About"** (lado direito da página principal do repositório), coloque uma frase curta como:

> *"Automação Python para validar assinaturas (vetoriais e digitais), extrair dados de OS e renomear PDFs em lote."*
```

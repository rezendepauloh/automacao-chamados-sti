# 📄 Conversão de PDFs para Imagens PNG

Este documento descreve o funcionamento do utilitário de conversão de documentos PDF em imagens PNG de alta resolução para exibição no sistema.

---

## 🎯 1. Objetivo do Utilitário (`converter_pdf.py`)

No sistema Bancada STI, alguns anexos, fluxogramas e documentos de portarias/fiscalizações chegam em formato `.pdf`. Para que sejam pré-visualizados dinamicamente na interface sem plugins adicionais, o script [`util/arquivos/converter_pdf.py`](file:///home/paulogoncalves/PythonProjects/automated-OTRS-and-CitSmart/util/arquivos/converter_pdf.py) converte a primeira página de cada PDF localizado na pasta `uploads/` em uma imagem `.png`.

O utilitário utiliza a biblioteca de alta performance `PyMuPDF` (`fitz`), aplicando uma matriz de escala de 2.0x para preservar a legibilidade e nitidez do texto.

---

## 🚀 2. Como Executar via Bash

Certifique-se de que os PDFs estejam salvos no diretório `uploads/` e execute:

```bash
python3 util/arquivos/converter_pdf.py
```

### Executar dentro do contêiner Docker:
```bash
docker exec -i $(docker ps -q -f name=bancada_streamlit_app) python util/arquivos/converter_pdf.py
```

As imagens geradas serão salvas com a extensão `.png` no mesmo diretório de origem do arquivo PDF.

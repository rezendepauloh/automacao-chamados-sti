# Guia de Utilitários, Conversão de Arquivos e Validações

> **Scripts em:** `util/arquivos/` e `util/diagnosticos/`

---

## 🎯 1. Conversão de PDFs para Imagens PNG (`converter_pdf.py`)

No Sistema Bancada STI, alguns anexos, fluxogramas e documentos de portarias/fiscalizações chegam em formato `.pdf`. Para que sejam pré-visualizados dinamicamente na interface sem plugins adicionais, o utilitário converte a primeira página de cada PDF localizado na pasta `uploads/` em uma imagem `.png`.

O utilitário utiliza a biblioteca `PyMuPDF` (`fitz`), aplicando uma matriz de escala de 2.0x para preservar a legibilidade e nitidez do texto.

### Como Executar:
Certifique-se de que os PDFs estejam salvos no diretório `uploads/` e execute dentro do contêiner Docker:
```bash
docker exec bancada_streamlit_app python util/arquivos/converter_pdf.py
```

---

## 📅 2. Validação de Eventos do FullCalendar (`test_json_events.py`)

A aba de Calendário Geral agrega em uma única linha do tempo interativa:
- Chamados de TI (com data de criação ISO, autor, localidade e histórico de notas);
- Escalas de plantão matutino e semanal;
- Férias, garantias e viagens técnicas.

O script `util/diagnosticos/test_json_events.py` simula a carga de chamados do banco de dados e valida se todos os atributos esperados pelo frontend JavaScript do FullCalendar (como `id`, `title`, `start`, `backgroundColor`, `extendedProps`) estão devidamente formatados e sem erros de serialização JSON.

### Como Executar:
```bash
docker exec bancada_streamlit_app python util/diagnosticos/test_json_events.py
```


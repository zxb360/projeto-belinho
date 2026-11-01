# Processador local de holerites

Protótipo em Streamlit para importar holerites PDF, extrair localmente linhas no
formato `código descrição valor`, revisar competências e gerar uma pasta com os
PDFs, planilha XLSX consolidada e minuta DOCX.

## Executar

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

No Windows, ative o ambiente com `.venv\Scripts\activate` e execute
`streamlit run app.py`.

## Fluxo

1. Informe nome e matrícula e carregue um ou mais PDFs.
2. Confira as competências sugeridas pelo nome do arquivo (por exemplo,
   `01_2021.pdf` ou `2021-01.pdf`); corrija os campos se necessário.
3. Revise as rubricas reconhecidas e escolha a pasta de saída.
4. Opcionalmente, carregue um modelo DOCX com os marcadores
   `{{NOME_CLIENTE}}`, `{{MATRICULA}}`, `{{VALOR_TOTAL_DIFERENCAS}}` e
   `{{TABELA_CALCULOS}}`.
5. Gere e revise os arquivos em `DADOS_PROCESSADOS/[Matrícula] - [Nome]/`.

## Limites deste protótipo

- Não faz OCR: PDFs digitalizados como imagem precisam de uma etapa OCR futura.
- A leitura financeira é local e baseada em linhas textuais com código, descrição
  e valor no padrão brasileiro. Confira todos os dados extraídos.
- O total apresentado é a soma bruta das rubricas reconhecidas; não é cálculo
  jurídico de diferenças ou valor devido.
- Não há chamadas ao Gemini, envio para nuvem ou integração OAuth com Drive.
  Para espelhar os arquivos, selecione uma pasta já sincronizada pelo Google
  Drive para Desktop.
- O modelo DOCX é opcional. Se não for fornecido, o programa cria uma minuta
  simples editável; a peça deve ser revisada por profissional antes de uso.
- Esta versão é o protótipo executável pelo Streamlit. Um instalador ou `.exe`
  de dois cliques para Windows exige empacotamento e validação em Windows.

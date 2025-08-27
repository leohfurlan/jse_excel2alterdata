# excel2alterdata

Pipeline em Python para padronizar planilhas diversas de "livro caixa" no layout do **Alterdata**.

## Instalação
```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

> Observação: para arquivos **.xls** (formato antigo) é necessário `xlrd` (já listado nos requisitos).
> Para **.xlsx/.xlsm**, o `openpyxl` será usado.

## Uso
Coloque seus arquivos de clientes (xls/xlsx/csv) em `samples/` e rode:
```bash
python excel2alterdata.py samples output config/mapping.yaml
```

Saídas geradas em `output/`:
- `alterdata_output.xlsx` e `alterdata_output.csv` (ordem e nomes de colunas do Alterdata)
- `mapeamento_colunas.xlsx` (log de como cada coluna foi mapeada em cada aba)
- `inconsistencias.xlsx` (quando houver)

## Ajustando mapeamentos
Edite `config/mapping.yaml` para incluir novos sinônimos por coluna. O fuzzy match ajuda,
mas manter uma lista de sinônimos por cliente/layout deixa o processo mais confiável.

## Integração com n8n
1. Node **Watch (Folder/Drive/S3)** para detectar novos arquivos.
2. Node **Execute Command**: `python excel2alterdata.py "/pasta_entrada" "/pasta_saida"`
3. Node **IF**: se existir `inconsistencias.xlsx` → enviar para revisão (WhatsApp/Email).
4. Caso contrário → entregar `alterdata_output.xlsx/.csv` para import no Alterdata.

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
excel2alterdata.py
------------------
Padroniza planilhas diversas de "livro caixa" no layout exigido pelo Alterdata.
- Auto-detecção de colunas por sinônimos e fuzzy match
- Normalização de datas (pt-BR) e números (pt-BR)
- Geração de CSV/XLSX com formatação de data e valor no padrão brasileiro.
"""
import sys, re, pathlib, json
import pandas as pd
from rapidfuzz import process, fuzz
from dateutil import parser
import yaml

ALTERDATA_COLS = [
    "CodLancAutom",
    "Conta Débito",
    "Conta Crédito",
    "Data",
    "Valor",
    "CodHistórico",
    "Número de Documento",
    "Imóvel",
    "Tipo de documento",
    "CPF/CNPJ",
]

def norm(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"[áàâãä]", "a", s)
    s = re.sub(r"[éêëè]", "e", s)
    s = re.sub(r"[íìïî]", "i", s)
    s = re.sub(r"[óôõöò]", "o", s)
    s = re.sub(r"[úùüû]", "u", s)
    s = re.sub(r"ç", "c", s)
    s = re.sub(r"[^a-z0-9]+", "_", s)
    return s.strip("_")

def load_mapping(path="config/mapping.yaml"):
    with open(path, "r", encoding="utf-8") as f:
        m = yaml.safe_load(f) or {}
    syn = {k: [norm(x) for x in v] for k,v in m.get("synonyms",{}).items()}
    rules = m.get("posting_rules", {"enabled": False})
    req = m.get("required_columns", ALTERDATA_COLS)
    return req, syn, rules

def detect_and_promote_header(df, synonyms):
    known = set()
    for k, arr in synonyms.items():
        known.add(norm(k))
        known.update(norm(a) for a in arr)

    max_scan = min(20, len(df))
    best_idx = -1
    best_score = -1
    for i in range(max_scan):
        row = df.iloc[i]
        non_empty = [c for c in row.tolist() if str(c).strip() and str(c).strip().lower() not in ["nan", "none"]]
        if len(non_empty) < 3:
            continue
        
        score = sum(1 for c in non_empty if norm(c) in known)
        
        if score > best_score:
            best_score = score
            best_idx = i

    if best_idx != -1 and best_score >= 2:
        new_cols = df.iloc[best_idx].astype(str).tolist()
        df2 = df.iloc[best_idx+1:].copy()
        df2.columns = new_cols
        df2 = df2.dropna(axis=1, how="all").dropna(axis=0, how="all")
        return df2, best_idx
    return df, None

def detect_column(colnames, target, synonyms):
    syns = set(synonyms.get(target, [])) | {norm(target)}
    normed = {c: norm(c) for c in colnames}
    for c, n in normed.items():
        if n in syns:
            return c
    choices = list(normed.values())
    if not choices: return None
    best, score, _ = process.extractOne(norm(target), choices, scorer=fuzz.token_sort_ratio)
    if score >= 86:
        for original, n in normed.items():
            if n == best:
                return original
    return None

def detect_payment_date_column(colnames, df):
    priority_tokens = ["pag", "pagto", "pagamento", "baixa", "liquid", "quit"]
    candidates = []
    for c in colnames:
        n = norm(c)
        score = sum(3 for t in priority_tokens if t in n)
        if score > 0 and c in df.columns:
            sample = df[c].head(200)
            ok = sum(1 for v in sample if parse_date_ptbr(v) is not None)
            total = len(sample)
            rate = (ok / total) if total else 0
            candidates.append((score, rate, c))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]

def parse_brl_number(x):
    if pd.isna(x): 
        return None
    s = str(x).strip().replace("R$", "").strip()
    if ',' in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except (ValueError, TypeError):
        try:
            cleaned_s = re.sub(r"[^0-9\.]", "", s.replace(",", "."))
            return float(cleaned_s)
        except (ValueError, TypeError):
            return None

def parse_date_ptbr(x):
    if pd.isna(x): return None
    s = str(x).strip()
    if s.isdigit() and len(s) == 5:
        try:
            return pd.to_datetime(float(s), unit='D', origin='1899-12-30').date()
        except (ValueError, TypeError): pass
    for dayfirst in (True, False):
        try:
            d = parser.parse(s, dayfirst=dayfirst)
            return d.date()
        except (parser.ParserError, TypeError): pass
    return None

def only_digits(s):
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return None
    return re.sub(r"[^0-9]", "", str(s))

def apply_posting_rules(std, df, synonyms_rules):
    if not synonyms_rules.get("enabled", False): return std
    source_aliases = [norm(x) for x in synonyms_rules.get("source_single_account_synonyms", [])]
    default_deb = str(synonyms_rules.get("default_debit_account", "")).strip()
    default_cred= str(synonyms_rules.get("default_credit_account", "")).strip()
    src_col = next((c for c in df.columns if norm(c) in source_aliases), None)
    if src_col is None or std["Conta Débito"].notna().any() or std["Conta Crédito"].notna().any():
        return std

    contas = df[src_col].astype(str).str.strip()
    valor = std["Valor"].fillna(0).astype(float)
    deb_col, cred_col = [], []
    for cta, v in zip(contas, valor):
        if v > 0:
            deb_col.append(cta or None); cred_col.append(default_cred or None)
        elif v < 0:
            deb_col.append(default_deb or None); cred_col.append(cta or None)
        else:
            deb_col.append(None); cred_col.append(None)

    std["Conta Débito"] = std["Conta Débito"].where(std["Conta Débito"].notna(), pd.Series(deb_col, index=std.index))
    std["Conta Crédito"] = std["Conta Crédito"].where(std["Conta Crédito"].notna(), pd.Series(cred_col, index=std.index))
    std["Valor"] = std["Valor"].abs()
    return std

def read_any_excel(path: pathlib.Path):
    suffix = path.suffix.lower()
    if suffix in [".xlsx", ".xlsm"]:
        return pd.ExcelFile(path, engine="openpyxl")
    elif suffix == ".xls":
        return pd.ExcelFile(path, engine="xlrd")
    elif suffix == ".csv":
        df = pd.read_csv(path, sep=None, engine="python", dtype=str, header=None)
        pseudo = type("PseudoExcel", (), {})()
        pseudo.sheet_names = ["CSV"]
        pseudo.parse = lambda _, **kwargs: df.copy()
        return pseudo
    else:
        raise ValueError(f"Formato não suportado: {path.suffix}")

def combine_valor(df):
    cand_val = next((c for c in df.columns if norm(c) in {"valor","valor_total","vlr","movimento","valor_liquido","valor_bruto"}), None)
    if cand_val:
        return df[cand_val].apply(parse_brl_number)
    deb_c = next((c for c in df.columns if "deb" in norm(c)), None)
    cred_c = next((c for c in df.columns if "cred" in norm(c)), None)
    if deb_c or cred_c:
        deb = df[deb_c].apply(parse_brl_number) if deb_c else 0.0
        cred = df[cred_c].apply(parse_brl_number) if cred_c else 0.0
        return (deb.fillna(0) - cred.fillna(0)).astype(float)
    return None

def process_file(path, req, synonyms, rules):
    try:
        xl = read_any_excel(path)
    except Exception as e:
        return None, pd.DataFrame([{"arquivo": str(path), "aba": "-", "erro": f"Falha leitura: {e}"}]), None

    frames, mapping_rows, issues = [], [], []
    for sheet in getattr(xl, "sheet_names", []):
        try:
            temp_df = xl.parse(sheet, dtype=str, header=None)
            if temp_df.empty: continue
            df, _ = detect_and_promote_header(temp_df, synonyms)
            df.columns = [str(c).strip() for c in df.columns]
        except Exception as e:
            issues.append({"arquivo": str(path), "aba": sheet, "erro": f"Falha ao ler ou processar cabeçalho: {e}"})
            continue
        
        colmap = {t: detect_column(list(df.columns), t, synonyms) for t in req}
        
        pick_data = detect_payment_date_column(list(df.columns), df)
        if pick_data is not None:
            colmap["Data"] = pick_data

        std = pd.DataFrame(columns=req)
        temp_data = {t: df[src] for t, src in colmap.items() if src and src in df}
        if not temp_data:
            issues.append({"arquivo": str(path), "aba": sheet, "erro": "Nenhuma coluna reconhecida."})
            continue
        std = pd.DataFrame(temp_data)

        for t in req:
            if t not in std.columns:
                std[t] = None

        std["Data"] = std["Data"].apply(parse_date_ptbr)
        if "Número de Documento" in std:
            std["Número de Documento"] = std["Número de Documento"].astype(str).str.strip()
        if "CPF/CNPJ" in std:
            std["CPF/CNPJ"] = std["CPF/CNPJ"].apply(only_digits)
        
        std["Valor"] = combine_valor(df) if "Valor" not in std or std["Valor"].isna().all() else std["Valor"].apply(parse_brl_number)

        mapping_rows.append({"arquivo": str(path.name), "aba": sheet, **{f"{t} <-": (colmap.get(t) or "(não encontrado)") for t in req}})

        if "Data" in std and "Valor" in std and std["Data"].isna().all() and std["Valor"].isna().all():
            issues.append({"arquivo": str(path), "aba": sheet, "erro": "Aba sem linhas válidas (Data e Valor vazios)."})
            continue

        std = apply_posting_rules(std, df, rules)
        frames.append(std.assign(__arquivo__=path.name, __aba__=sheet))

    out = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    mapdf = pd.DataFrame(mapping_rows) if mapping_rows else pd.DataFrame()

    if not out.empty:
        for col in req + ["__arquivo__", "__aba__"]:
            if col not in out.columns:
                out[col] = None
        out = out[req + ["__arquivo__", "__aba__"]]

    err = pd.DataFrame(issues)
    return out, err, mapdf

def main(in_dir="samples", out_dir="output", mapping="config/mapping.yaml"):
    out_dir = pathlib.Path(out_dir); out_dir.mkdir(parents=True, exist_ok=True)
    req, synonyms, rules = load_mapping(mapping)

    all_out, all_err, all_map = [], [], []
    files = list(pathlib.Path(in_dir).glob("*.*"))
    for p in files:
        if p.suffix.lower() not in [".xlsx",".xls",".xlsm",".csv"]: 
            continue
        out, err, mapdf = process_file(p, req, synonyms, rules)
        if out is not None and not out.empty: all_out.append(out)
        if err is not None and not err.empty: all_err.append(err)
        if mapdf is not None and not mapdf.empty: all_map.append(mapdf)

    final = pd.concat(all_out, ignore_index=True) if all_out else pd.DataFrame()
    errors= pd.concat(all_err, ignore_index=True) if all_err else pd.DataFrame()
    maps  = pd.concat(all_map, ignore_index=True) if all_map else pd.DataFrame()

    if not final.empty:
        for col in ALTERDATA_COLS:
            if col not in final.columns:
                final[col] = None
        
        output_path_xlsx = out_dir / "alterdata_output.xlsx"
        writer = pd.ExcelWriter(output_path_xlsx, engine='xlsxwriter', date_format='dd/mm/yyyy')
        
        final_excel = final.copy()
        final_excel['Data'] = pd.to_datetime(final_excel['Data'])
        
        final_excel.to_excel(writer, index=False, sheet_name='Lançamentos', columns=ALTERDATA_COLS + ["__arquivo__", "__aba__"])
        
        workbook = writer.book
        worksheet = writer.sheets['Lançamentos']
        
        br_money_format = workbook.add_format({'num_format': '#,##0.00'})
        
        try:
            valor_col_index = final_excel.columns.get_loc('Valor')
            worksheet.set_column(valor_col_index, valor_col_index, 15, br_money_format)
        except KeyError:
            pass

        writer.close()

        final[ALTERDATA_COLS].to_csv(out_dir/"alterdata_output.csv", index=False, sep=";", encoding="utf-8-sig", decimal=',')

    if not errors.empty:
        errors.to_excel(out_dir/"inconsistencias.xlsx", index=False)
    if not maps.empty:
        maps.to_excel(out_dir/"mapeamento_colunas.xlsx", index=False)

    # Monta o resumo para ser retornado
    summary = {
        "arquivos_encontrados": len(files),
        "linhas_exportadas": int(final.shape[0]) if not final.empty else 0,
        "tem_inconsistencias": (not errors.empty),
        "saidas": {
            "xlsx": str(out_dir/"alterdata_output.xlsx"),
            "csv": str(out_dir/"alterdata_output.csv"),
            "inconsistencias": str(out_dir/"inconsistencias.xlsx"),
            "mapeamento": str(out_dir/"mapeamento_colunas.xlsx"),
        }
    }
    # Imprime o JSON para compatibilidade com o modo CLI
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    
    # <-- MUDANÇA: Retorna o resumo e o dataframe de erros
    return summary, errors


if __name__ == "__main__":
    in_dir = sys.argv[1] if len(sys.argv)>1 else "samples"
    out_dir = sys.argv[2] if len(sys.argv)>2 else "output"
    mapping = sys.argv[3] if len(sys.argv)>3 else "config/mapping.yaml"
    main(in_dir, out_dir, mapping)

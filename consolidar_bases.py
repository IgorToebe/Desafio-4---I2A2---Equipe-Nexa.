import os
import sys
import unicodedata
import re
import argparse
from datetime import datetime, date
from typing import Optional, List, Dict

import pandas as pd

try:
    import holidays  # type: ignore
except Exception:
    holidays = None  # Will handle absence gracefully

# Integração: após consolidar, executa o vr_agent para gerar o VR final
try:
    # Importa a função de execução do agente de VR
    from vr_agent import run_vr_agent  # type: ignore
except Exception:
    run_vr_agent = None  # type: ignore

# ------------------------------
# Config
# ------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# Nova pasta onde ficarão todos os arquivos .xlsx
BASE_XLSX_DIR = os.path.join(BASE_DIR, "xlsx")

def _xlsx_path(nome: str) -> str:
    # Caminho relativo usado pelo leitor/escritor (que concatena BASE_DIR)
    return os.path.join("xlsx", nome)

ARQUIVOS = {
    "ATIVOS": _xlsx_path("ATIVOS.xlsx"),
    "ADMISSAO": _xlsx_path("ADMISSÃO ABRIL.xlsx"),
    "FERIAS": _xlsx_path("FÉRIAS.xlsx"),
    "DESLIGADOS": _xlsx_path("DESLIGADOS.xlsx"),
    "SINDICATO_VALOR": _xlsx_path("Base sindicato x valor.xlsx"),
    "DIAS_UTEIS": _xlsx_path("Base dias uteis.xlsx"),
    # Arquivos auxiliares para exclusões (opcional)
    "AFASTAMENTOS": _xlsx_path("AFASTAMENTOS.xlsx"),
    "EXTERIOR": _xlsx_path("EXTERIOR.xlsx"),
    # Outros arquivos presentes, mas não utilizados diretamente: APRENDIZ.xlsx, ESTÁGIO.xlsx, VR MENSAL 05.2025.xlsx
}

ARQUIVO_SAIDA = os.path.join("xlsx", "BaseConsolidada.xlsx")

# Mapeamento opcional de colunas para padronização
COL_MATRICULA = "MATRICULA"
COL_SINDICATO = "Sindicato"  # nome esperado da coluna de sindicato
COL_ESTADO = "ESTADO"  # nome esperado da coluna de UF/Estado (em ATIVOS e SINDICATO_VALOR)
COL_STATUS = "DESC. SITUACAO"  # coluna em ATIVOS para status
COL_CARGO = "TITULO DO CARGO"  # coluna em ATIVOS para cargo
COL_DATA_DEMISSAO = "DATA DEMISSÃO"
COL_DIAS_FERIAS = "DIAS DE FÉRIAS"

# Parâmetros de feriados
# Municipais (preencher conforme necessário): {"São Paulo": ["2025-01-25", ...]}
FERIADOS_MUNICIPAIS: Dict[str, List[str]] = {}
COL_MUNICIPIO = "MUNICIPIO"  # se existir na base ATIVOS
COL_UF = "UF"  # alternativa para ESTADO

# Feriados fornecidos pelo usuário (2025): nacionais e estaduais por UF.
# Datas em formato ISO (YYYY-MM-DD) para facilitar o parse.
FERIADOS_CONFIG: Dict[str, object] = {
    "anos": {
        2025: {
            "nacionais": [
                "2025-01-01",  # Confraternização Universal
                "2025-03-04",  # Carnaval (ponto facultativo)
                "2025-04-18",  # Sexta-feira Santa
                "2025-04-21",  # Tiradentes
                "2025-05-01",  # Dia do Trabalho
                "2025-06-19",  # Corpus Christi (ponto facultativo)
                "2025-09-07",  # Independência do Brasil
                "2025-10-12",  # Nossa Senhora Aparecida
                "2025-11-02",  # Finados
                "2025-11-15",  # Proclamação da República
                "2025-11-20",  # Dia da Consciência Negra (como nacional na sua lista)
                "2025-12-25",  # Natal
            ],
            "estaduais": {
                # Rio Grande do Sul
                "RS": [
                    "2025-09-20",  # Revolução Farroupilha
                ],
                # Paraná
                "PR": [
                    "2025-12-19",  # Emancipação Política do Paraná
                ],
                # São Paulo
                "SP": [
                    "2025-07-09",  # Revolução Constitucionalista
                ],
                # Rio de Janeiro
                "RJ": [
                    "2025-03-04",  # Carnaval (estadual RJ)
                    "2025-04-23",  # Dia de São Jorge
                    "2025-11-20",  # Consciência Negra (estadual RJ)
                ],
            },
        }
    }
}

# Mapeamento manual opcional Sindicato -> UF (preencha conforme sua realidade)
SINDICATO_TO_UF: Dict[str, str] = {
    # Exemplo:
     "SITEPD PR - SIND DOS TRAB EM EMPR PRIVADAS DE PROC DE DADOS DE CURITIBA E REGIAO METROPOLITANA": "PR",
     "SINDPPD RS - SINDICATO DOS TRAB. EM PROC. DE DADOS RIO GRANDE DO SUL": "RS",
     "SINDPD SP - SIND.TRAB.EM PROC DADOS E EMPR.EMPRESAS PROC DADOS ESTADO DE SP.": "SP",
     "SINDPD RJ - SINDICATO PROFISSIONAIS DE PROC DADOS DO RIO DE JANEIRO": "RJ"
}

# Palavras-chave para inferência de UF a partir do nome do sindicato
UF_KEYWORDS: Dict[str, List[str]] = {
    "SP": ["SP", "SAO PAULO", "SÃO PAULO", "PAULISTA"],
    "RJ": ["RJ", "RIO DE JANEIRO", "FLUMINENSE"],
    "RS": ["RS", "RIO GRANDE DO SUL", "GAUCHO", "GAÚCHO"],
    "PR": ["PR", "PARANA", "PARANÁ"],
    # Adicione outros estados se necessário
}

# ------------------------------
# Utilitários
# ------------------------------

def _strip_accents(s: str) -> str:
    try:
        return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    except Exception:
        return s

def _normalize_upper_noaccents_spaces(s: Optional[str]) -> str:
    if s is None:
        return ""
    try:
        txt = str(s)
        txt = _strip_accents(txt)
        txt = txt.upper().strip()
        # colapsa espaços múltiplos
        txt = " ".join(txt.split())
        return txt
    except Exception:
        return str(s) if s is not None else ""

_UF_BY_NAME: Dict[str, str] = {
    # nomes normalizados (sem acento, maiúsculos)
    "AC": "AC", "ACRE": "AC",
    "AL": "AL", "ALAGOAS": "AL",
    "AP": "AP", "AMAPA": "AP",
    "AM": "AM", "AMAZONAS": "AM",
    "BA": "BA", "BAHIA": "BA",
    "CE": "CE", "CEARA": "CE",
    "DF": "DF", "DISTRITO FEDERAL": "DF",
    "ES": "ES", "ESPIRITO SANTO": "ES",
    "GO": "GO", "GOIAS": "GO",
    "MA": "MA", "MARANHAO": "MA",
    "MT": "MT", "MATO GROSSO": "MT",
    "MS": "MS", "MATO GROSSO DO SUL": "MS",
    "MG": "MG", "MINAS GERAIS": "MG",
    "PA": "PA", "PARA": "PA",
    "PB": "PB", "PARAIBA": "PB",
    "PR": "PR", "PARANA": "PR",
    "PE": "PE", "PERNAMBUCO": "PE",
    "PI": "PI", "PIAUI": "PI",
    "RJ": "RJ", "RIO DE JANEIRO": "RJ",
    "RN": "RN", "RIO GRANDE DO NORTE": "RN",
    "RS": "RS", "RIO GRANDE DO SUL": "RS",
    "RO": "RO", "RONDONIA": "RO",
    "RR": "RR", "RORAIMA": "RR",
    "SC": "SC", "SANTA CATARINA": "SC",
    "SP": "SP", "SAO PAULO": "SP",
    "SE": "SE", "SERGIPE": "SE",
    "TO": "TO", "TOCANTINS": "TO",
}

def normalizar_uf(valor: Optional[str]) -> Optional[str]:
    if not isinstance(valor, str):
        return None
    s = valor.strip().upper()
    if not s:
        return None
    # remove "ESTADO DE/DO/DA/DAS/" e pontuações comuns
    s = s.replace("ESTADO DE ", "").replace("ESTADO DO ", "").replace("ESTADO DA ", "").replace("ESTADO DAS ", "")
    # remove prefixos como "UF ", "UF-"
    if s.startswith("UF "):
        s = s[3:].strip()
    if s.startswith("UF-"):
        s = s[3:].strip()
    s = _strip_accents(s)
    s = s.replace("-", " ").replace("/", " ").replace(".", " ")
    s = " ".join(s.split())
    # mapeia nomes para UF, mantendo UF se já for código
    return _UF_BY_NAME.get(s, _UF_BY_NAME.get(s.replace(" ESTADO", ""), _UF_BY_NAME.get(s.replace(" EST.", ""), None)))

def ler_excel(caminho: str, **kwargs) -> pd.DataFrame:
    caminho_abs = os.path.join(BASE_DIR, caminho)
    if not os.path.exists(caminho_abs):
        print(f"[AVISO] Arquivo não encontrado: {caminho_abs}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(caminho_abs, engine="openpyxl", **kwargs)
        return df
    except Exception as e:
        print(f"[ERRO] Falha ao ler {caminho_abs}: {e}")
        return pd.DataFrame()


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    # Tira espaços nas bordas e normaliza nomes duplicados
    df.columns = [str(c).strip() for c in df.columns]

    # Remove colunas sem cabeçalho (ex.: 'Unnamed: 3')
    cols_drop = [c for c in df.columns if str(c).strip().lower().startswith("unnamed")]
    if cols_drop:
        df = df.drop(columns=cols_drop, errors="ignore")
    return df


def padronizar_matricula(df: pd.DataFrame, col=COL_MATRICULA) -> pd.DataFrame:
    if df.empty or col not in df.columns:
        return df
    # Converte para string e remove espaços; mantém zeros à esquerda se houver
    df[col] = (
        df[col]
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)  # remove decimal toString comum do Excel
    )
    return df


def to_datetime_safe(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


def detectar_coluna(df: pd.DataFrame, candidatos: List[str]) -> Optional[str]:
    for c in candidatos:
        if c in df.columns:
            return c
    # tenta case-insensitive
    lower_map = {str(col).lower(): col for col in df.columns}
    for c in candidatos:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None


def merge_seguro(left: pd.DataFrame, right: pd.DataFrame, on: str, suffix: str) -> pd.DataFrame:
    if right.empty:
        return left
    if on not in left.columns or on not in right.columns:
        print(f"[AVISO] Merge ignorado: chave '{on}' ausente.")
        return left
    # Evita colisões de nomes
    col_conflicts = set(left.columns).intersection(set(right.columns)) - {on}
    right_renamed = right.rename(columns={c: f"{c}{suffix}" for c in col_conflicts})
    return left.merge(right_renamed, on=on, how="left")


# ------------------------------
# Feriados e dias úteis
# ------------------------------

def obter_feriados_brasil(ano: int, uf: Optional[str] = None) -> List[date]:
    datas: List[date] = []
    if holidays is None:
        # Biblioteca não instalada; sem feriados automáticos
        return datas
    try:
        # Subdivisão por UF (ex.: 'SP', 'RJ') se disponível
        br = holidays.Brazil(years=ano, subdiv=uf) if uf else holidays.Brazil(years=ano)
        datas = list(br.keys())
    except Exception as e:
        print(f"[AVISO] Falha ao carregar feriados do Brazil({ano}, uf={uf}): {e}")
    return datas


def obter_feriados_config(ano: int, uf: Optional[str]) -> List[date]:
    """Retorna feriados a partir da configuração fornecida pelo usuário para o ano/UF informados."""
    out: List[date] = []
    try:
        anos = FERIADOS_CONFIG.get("anos", {})  # type: ignore[assignment]
        cfg = anos.get(ano)
        if cfg:
            nacs = cfg.get("nacionais", [])
            out.extend([pd.to_datetime(x).date() for x in nacs])
            if uf:
                uf = uf.upper()
                est = cfg.get("estaduais", {}).get(uf, [])
                out.extend([pd.to_datetime(x).date() for x in est])
    except Exception:
        pass
    return out


def obter_feriados_municipais(municipio: Optional[str]) -> List[date]:
    if not municipio:
        return []
    datas_str = FERIADOS_MUNICIPAIS.get(municipio, [])
    saida: List[date] = []
    for s in datas_str:
        try:
            saida.append(pd.to_datetime(s).date())
        except Exception:
            pass
    return saida


def contar_feriados_no_mes(ano: int, mes: int, uf: Optional[str], municipio: Optional[str]) -> int:
    # Usa primeiro a configuração fornecida, e também faz união com biblioteca (se disponível)
    feriados_cfg = [d for d in obter_feriados_config(ano, uf) if d.year == ano and d.month == mes]
    feriados_lib = [d for d in obter_feriados_brasil(ano, uf) if d.year == ano and d.month == mes]
    feriados_municipais = [d for d in obter_feriados_municipais(municipio) if d.year == ano and d.month == mes]
    total = len(set(feriados_cfg + feriados_lib + feriados_municipais))
    return total


# ------------------------------
# Pipeline principal
# ------------------------------

def carregar_bases() -> Dict[str, pd.DataFrame]:
    df_ativos = normalizar_colunas(ler_excel(ARQUIVOS["ATIVOS"]))
    df_adm = normalizar_colunas(ler_excel(ARQUIVOS["ADMISSAO"]))
    df_ferias = normalizar_colunas(ler_excel(ARQUIVOS["FERIAS"]))
    df_desl = normalizar_colunas(ler_excel(ARQUIVOS["DESLIGADOS"]))
    df_sind_valor = normalizar_colunas(ler_excel(ARQUIVOS["SINDICATO_VALOR"]))
    # skiprows=1 para linha extra de cabeçalho
    df_dias_uteis = normalizar_colunas(ler_excel(ARQUIVOS["DIAS_UTEIS"], skiprows=1))
    # Normaliza cabeçalho incorreto 'SINDICADO' -> 'Sindicato'
    if not df_dias_uteis.empty:
        ren_map = {}
        for c in df_dias_uteis.columns:
            if str(c).strip().upper() == "SINDICADO":
                ren_map[c] = "Sindicato"
        if ren_map:
            df_dias_uteis = df_dias_uteis.rename(columns=ren_map)

    # Auxiliares (para filtros)
    df_afast = normalizar_colunas(ler_excel(ARQUIVOS["AFASTAMENTOS"]))
    df_exterior = normalizar_colunas(ler_excel(ARQUIVOS["EXTERIOR"]))

    # Padroniza matrícula
    for df in [df_ativos, df_adm, df_ferias, df_desl, df_afast, df_exterior]:
        padronizar_matricula(df)

    # Datas
    to_datetime_safe(df_desl, [COL_DATA_DEMISSAO])
    # Tenta converter quaisquer colunas com 'DATA' no nome
    for df in [df_ativos, df_adm, df_ferias, df_desl]:
        data_cols = [c for c in df.columns if "DATA" in str(c).upper() or "DATE" in str(c).upper()]
        to_datetime_safe(df, data_cols)

    # Limpeza específica: strip em valores de ESTADO na base sindicato x valor
    col_estado_sind = detectar_coluna(df_sind_valor, ["estado", COL_ESTADO, COL_UF])
    if col_estado_sind and not df_sind_valor.empty:
        df_sind_valor[col_estado_sind] = df_sind_valor[col_estado_sind].astype(str).str.strip()
        # normaliza para UF
        df_sind_valor["UF_SIND_VAL"] = df_sind_valor[col_estado_sind].apply(normalizar_uf)
        # Detecta e normaliza coluna de valor (VR por dia)
        possiveis_valor = [
            "VALOR", "VALOR DIA", "VALOR_DIA", "VALOR VR", "VALOR_VR", "VR", "VR DIA", "VR_DIA", "VALOR BASE", "VALOR_BASE"
        ]
        col_valor = detectar_coluna(df_sind_valor, possiveis_valor)
        if col_valor and col_valor != "VALOR_DIARIO_VR":
            df_sind_valor.rename(columns={col_valor: "VALOR_DIARIO_VR"}, inplace=True)

    return {
        "ATIVOS": df_ativos,
        "ADMISSAO": df_adm,
        "FERIAS": df_ferias,
        "DESLIGADOS": df_desl,
        "SIND_VALOR": df_sind_valor,
        "DIAS_UTEIS": df_dias_uteis,
        "AFASTAMENTOS": df_afast,
        "EXTERIOR": df_exterior,
    }


def executar_merges(bases: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = bases["ATIVOS"].copy()

    # Derivar 'estado' a partir de 'Sindicato' se a coluna de estado/UF não existir
    def derivar_uf_por_sindicato(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in.empty:
            return df_in
        col_sind = detectar_coluna(df_in, ["Sindicato", COL_SINDICATO, "SINDICATO", "SINDICADO"])
        if not col_sind:
            return df_in
        # já existe alguma coluna de estado/UF?
        col_estado_exist = detectar_coluna(df_in, ["estado", COL_ESTADO, COL_UF])
        if col_estado_exist:
            return df_in
        # cria nova coluna 'estado' inferida
        def infer_uf(nome: str) -> Optional[str]:
            if not isinstance(nome, str):
                return None
            nome_up = nome.strip().upper()
            # mapeamento explícito
            for k, uf in SINDICATO_TO_UF.items():
                if k.upper() in nome_up:
                    return uf
            # heurística por palavras-chave
            for uf, terms in UF_KEYWORDS.items():
                for t in terms:
                    if t.upper() in nome_up:
                        return uf
            return None

        df_in["estado"] = df_in[col_sind].apply(infer_uf)
        # normaliza para código UF
        df_in["estado"] = df_in["estado"].apply(normalizar_uf)
        return df_in

    df = derivar_uf_por_sindicato(df)

    # Merge com admissão
    df = merge_seguro(df, bases["ADMISSAO"], on=COL_MATRICULA, suffix="_ADM")

    # Merge com férias (exclui coluna de status para evitar criar 'DESC. SITUACAO_FER')
    df_ferias = bases["FERIAS"].copy()
    if not df_ferias.empty:
        # Remove qualquer variação do nome da coluna de status
        def _norm_name(n: str) -> str:
            s = _normalize_upper_noaccents_spaces(str(n))
            return s.replace(".", "")
        cols_to_drop = [c for c in df_ferias.columns if _norm_name(c) == "DESC SITUACAO"]
        if cols_to_drop:
            df_ferias = df_ferias.drop(columns=cols_to_drop, errors="ignore")
    # Executa o merge; se df_ferias estiver vazio, merge_seguro não altera o resultado
    df = merge_seguro(df, df_ferias, on=COL_MATRICULA, suffix="_FER")

    # Merge com desligados (DATA DEMISSÃO)
    df = merge_seguro(df, bases["DESLIGADOS"], on=COL_MATRICULA, suffix="_DESL")

    # Incluir também quem está somente em DESLIGADOS
    df_desl_total = bases.get("DESLIGADOS", pd.DataFrame())
    if not df_desl_total.empty and COL_MATRICULA in df_desl_total.columns:
        # Matrículas já presentes na base (de ATIVOS após merges)
        mats_presentes = set(df[COL_MATRICULA].astype(str).str.strip()) if COL_MATRICULA in df.columns else set()
        df_desl_only = df_desl_total[~df_desl_total[COL_MATRICULA].astype(str).str.strip().isin(mats_presentes)].copy()
        if not df_desl_only.empty:
            # Garante coluna de status preenchida como 'Desligado'
            col_status_desl = detectar_coluna(df_desl_only, [COL_STATUS, "DESC. SITUAÇÃO", "desc. situacao", "DESC SITUACAO"]) or COL_STATUS
            if col_status_desl not in df_desl_only.columns:
                df_desl_only[COL_STATUS] = "Desligado"
            else:
                df_desl_only[col_status_desl] = "Desligado"

            # Tentar derivar UF a partir do Sindicato para permitir merges seguintes
            df_desl_only = derivar_uf_por_sindicato(df_desl_only)

            # Concatena mantendo todas as colunas conhecidas; deixa sort=False para performance
            df = pd.concat([df, df_desl_only], ignore_index=True, sort=False)

    # Merge com Sindicato x Valor (chave: estado/UF)
    df_sind_valor = bases["SIND_VALOR"]
    if not df_sind_valor.empty:
        col_estado_base = detectar_coluna(df, ["estado", COL_ESTADO, COL_UF])
        col_estado_sind = detectar_coluna(df_sind_valor, ["UF_SIND_VAL", "estado", COL_ESTADO, COL_UF])
        if col_estado_base and col_estado_sind:
            # normaliza ambos lados para UF em colunas auxiliares
            df["UF_BASE_TMP"] = df[col_estado_base].apply(normalizar_uf)
            df_sind_valor["UF_SIND_TMP"] = df_sind_valor[col_estado_sind].apply(normalizar_uf)
            # garante nome padrão da coluna de valor (suporta nomes antigos)
            col_valor_sv = detectar_coluna(df_sind_valor, ["VALOR_DIARIO_VR", "VALOR_BASE_VR", "VALOR_VR", "VALOR", "VALOR DIA", "VALOR_DIA", "VR", "VALOR BASE", "VALOR_BASE"])
            if col_valor_sv and col_valor_sv != "VALOR_DIARIO_VR":
                df_sind_valor.rename(columns={col_valor_sv: "VALOR_DIARIO_VR"}, inplace=True)
            df = df.merge(
                df_sind_valor,
                left_on="UF_BASE_TMP",
                right_on="UF_SIND_TMP",
                how="left",
                suffixes=("", "_SIND"),
            )
            # se existir VALOR também no lado esquerdo, prioriza VALOR_DIARIO_VR do sindicato
            if "VALOR_DIARIO_VR_SIND" in df.columns and "VALOR_DIARIO_VR" not in df.columns:
                df.rename(columns={"VALOR_DIARIO_VR_SIND": "VALOR_DIARIO_VR"}, inplace=True)
            elif "VALOR_DIARIO_VR_SIND" in df.columns:
                # cria coluna canônica coalescida
                df["VALOR_DIARIO_VR"] = df["VALOR_DIARIO_VR"].fillna(df["VALOR_DIARIO_VR_SIND"]) if "VALOR_DIARIO_VR" in df.columns else df["VALOR_DIARIO_VR_SIND"]

            # parser robusto de moeda/decimal
            def _parse_valor(v):
                s = str(v).strip()
                if not s or s.lower() in {"nan", "none"}:
                    return float("nan")
                s = re.sub(r"[^0-9,.-]", "", s)
                # casos com vírgula e ponto
                if "," in s and "." in s:
                    if re.match(r"^\d{1,3}(\.\d{3})+,\d{1,2}$", s):
                        s = s.replace(".", "").replace(",", ".")
                    elif re.match(r"^\d{1,3}(,\d{3})+\.\d{1,2}$", s):
                        s = s.replace(",", "")
                    else:
                        # heurística: último separador define decimal
                        last_comma = s.rfind(",")
                        last_dot = s.rfind(".")
                        if last_comma > last_dot:
                            s = s.replace(".", "").replace(",", ".")
                        else:
                            s = s.replace(",", "")
                elif "," in s:
                    # assume vírgula como decimal
                    s = s.replace(".", "")
                    s = s.replace(",", ".")
                elif "." in s:
                    # se muitos pontos, mantem apenas o último como decimal
                    if s.count(".") > 1:
                        parts = s.split(".")
                        s = "".join(parts[:-1]) + "." + parts[-1]
                try:
                    return float(s)
                except Exception:
                    return float("nan")

            if "VALOR_DIARIO_VR" in df.columns:
                df["VALOR_DIARIO_VR"] = df["VALOR_DIARIO_VR"].map(_parse_valor)
            # limpa colunas temporárias
            for c in ["UF_BASE_TMP", "UF_SIND_TMP"]:
                if c in df.columns:
                    df.drop(columns=c, inplace=True)
        else:
            print("[AVISO] Não foi possível identificar colunas de ESTADO/UF para o merge com Sindicato x Valor.")

    # Merge com Dias Úteis (chave: Sindicato)
    df_dias_uteis = bases["DIAS_UTEIS"]
    if not df_dias_uteis.empty:
        # Remover quaisquer colunas "ajustadas" de dias úteis do lado direito (evita duplicar no resultado)
        df_diu = df_dias_uteis.copy()
        def _norm_name(n: str) -> str:
            s = str(n).replace("_", " ").replace(".", " ")
            s = _normalize_upper_noaccents_spaces(s)
            return s
        norm_map = {c: _norm_name(c) for c in df_diu.columns}
        cols_to_drop = [c for c, n in norm_map.items() if ("DIAS UTEIS" in n and "AJUST" in n)]
        if cols_to_drop:
            df_diu = df_diu.drop(columns=cols_to_drop, errors="ignore")

        # Renomear a coluna base de dias úteis para o nome canônico solicitado
        # Tenta encontrar por candidatos comuns e, se não achar, por normalização contendo "DIAS UTEIS"
        col_dias = detectar_coluna(df_diu, [
            "DIAS UTEIS", "DIAS ÚTEIS", "QTD DIAS UTEIS", "QTD DIAS ÚTEIS"
        ])
        if not col_dias:
            for c, n in norm_map.items():
                if "DIAS UTEIS" in n:
                    col_dias = c
                    break
        if col_dias and col_dias != "TOTAL_DIAS_UTEIS_SINDICATO":
            df_diu.rename(columns={col_dias: "TOTAL_DIAS_UTEIS_SINDICATO"}, inplace=True)

        col_sind_base = detectar_coluna(df, ["Sindicato", COL_SINDICATO, "SINDICATO", "SINDICADO"])  # prioriza nome informado
        col_sind_dias = detectar_coluna(df_diu, ["Sindicato", COL_SINDICATO, "SINDICATO", "SINDICADO"])  # prioriza nome informado
    # remove prints de debug
        if col_sind_base and col_sind_dias:
            # Normaliza sindicato para evitar diferenças de caixa e espaços
            df[col_sind_base] = df[col_sind_base].astype(str).str.strip().str.upper()
            df_diu[col_sind_dias] = df_diu[col_sind_dias].astype(str).str.strip().str.upper()
            df = df.merge(df_diu, left_on=col_sind_base, right_on=col_sind_dias, how="left", suffixes=("", "_DIAS"))
            # Pós-merge: remover quaisquer colunas 'ajustadas' vindas da direita (sufixo _DIAS)
            to_drop_post = [c for c in df.columns if c.endswith("_DIAS") and ("AJUST" in _norm_name(c))]
            if to_drop_post:
                df = df.drop(columns=to_drop_post, errors="ignore")
        else:
            print("[AVISO] Não foi possível identificar a coluna 'Sindicato' para o merge de Dias Úteis.")

    # Consolidar coluna de UF: manter apenas 'estado' no resultado final
    def _consolidar_estado(df_in: pd.DataFrame) -> pd.DataFrame:
        if "estado" not in df_in.columns:
            df_in["estado"] = pd.NA
        cand_cols = ["estado", "ESTADO", "UF_SIND_VAL", COL_UF]
        presentes = [c for c in cand_cols if c in df_in.columns]
        if presentes:
            def pick(row):
                for c in ["estado", "ESTADO", "UF_SIND_VAL", COL_UF]:
                    if c in row and pd.notna(row[c]) and str(row[c]).strip():
                        return row[c]
                return pd.NA
            df_in["estado"] = df_in.apply(pick, axis=1)
            # normaliza para código UF
            df_in["estado"] = df_in["estado"].apply(normalizar_uf)
        # remove duplicadas solicitadas
        for drop_c in ["ESTADO", "UF_SIND_VAL"]:
            if drop_c in df_in.columns:
                df_in.drop(columns=drop_c, inplace=True)
        return df_in

    df = _consolidar_estado(df)

    # Consolidar coluna de Cargo: manter apenas 'TITULO DO CARGO'
    def _consolidar_cargo(df_in: pd.DataFrame) -> pd.DataFrame:
        col_titulo = detectar_coluna(df_in, ["TITULO DO CARGO", "TÍTULO DO CARGO"]) or "TITULO DO CARGO"
        col_cargo = detectar_coluna(df_in, ["CARGO"])  # alguns arquivos trazem apenas 'CARGO'
        # Se só existe CARGO, renomeia para TITULO DO CARGO
        if col_cargo and col_titulo not in df_in.columns:
            df_in.rename(columns={col_cargo: "TITULO DO CARGO"}, inplace=True)
            # normaliza conteúdo
            df_in["TITULO DO CARGO"] = df_in["TITULO DO CARGO"].map(_normalize_upper_noaccents_spaces)
            return df_in
        # Se ambos existem, preenche lacunas no título com CARGO e remove CARGO
        if col_titulo in df_in.columns and col_cargo:
            mask = df_in[col_titulo].isna() | (df_in[col_titulo].astype(str).str.strip() == "")
            df_in.loc[mask, col_titulo] = df_in.loc[mask, col_cargo]
            df_in.drop(columns=[col_cargo], inplace=True)
        # normaliza conteúdo do título
        if col_titulo in df_in.columns:
            df_in[col_titulo] = df_in[col_titulo].map(_normalize_upper_noaccents_spaces)
        return df_in

    df = _consolidar_cargo(df)

    return df


def aplicar_filtros_exclusao(df: pd.DataFrame, bases: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if df.empty:
        return df

    df_filtrado = df.copy()

    # 1) Excluir cargos: Diretores, Estagiários, Aprendizes
    col_cargo = detectar_coluna(df_filtrado, [COL_CARGO])
    if col_cargo:
        padrao = r"DIRETOR|DIRETORES|ESTAGIARIO|ESTAGIÁRI|ESTAGIARIO\(A\)|APRENDIZ|APRENDIZES"
        df_filtrado = df_filtrado[~df_filtrado[col_cargo].astype(str).str.upper().str.contains(padrao, regex=True, na=False)]

    # 2) Excluir status: Afastados em geral (inclui licença maternidade e auxílio-doença)
    col_status = detectar_coluna(df_filtrado, [COL_STATUS])
    if col_status:
        padrao_status = (
            r"AFAST|LICENCA MATERN|LICENÇA MATERN|MATERNIDADE|AFASTADOS EM GERAL|"
            r"AUXILIO DOENCA|AUXÍLIO DOENÇA|AUXILIO-DOENCA|AUXÍLIO-DOENÇA|INSS"
        )
        df_filtrado = df_filtrado[~df_filtrado[col_status].astype(str).str.upper().str.contains(padrao_status, regex=True, na=False)]

    # 2b) Regra explícita: qualquer matrícula presente em AFASTAMENTOS não deve aparecer
    df_afast = bases.get("AFASTAMENTOS", pd.DataFrame())
    if not df_afast.empty:
        col_mat_afast = COL_MATRICULA if COL_MATRICULA in df_afast.columns else detectar_coluna(df_afast, ["MATRICULA", "MATRÍCULA", "Matricula"]) or COL_MATRICULA
        if col_mat_afast in df_afast.columns and COL_MATRICULA in df_filtrado.columns:
            excl = set(df_afast[col_mat_afast].astype(str).str.strip())
            df_filtrado = df_filtrado[~df_filtrado[COL_MATRICULA].astype(str).str.strip().isin(excl)]

    # 3) Excluir profissionais no exterior
    # Tenta por coluna na base principal
    possiveis_cols_exterior = ["LOCAL TRABALHO", "LOCALIDADE", "PAIS", "PAÍS", "EXTERIOR"]
    col_exterior = detectar_coluna(df_filtrado, possiveis_cols_exterior)
    if col_exterior:
        df_filtrado = df_filtrado[~df_filtrado[col_exterior].astype(str).str.upper().str.contains(r"EXTERIOR|EXTERNO|INTERNACIONAL|OUTRO PAIS|OUTRO PAÍS", regex=True, na=False)]
    else:
        # fallback pela lista do arquivo EXTERIOR.xlsx
        df_ext = bases.get("EXTERIOR", pd.DataFrame())
        if not df_ext.empty and COL_MATRICULA in df_ext.columns:
            excl = set(df_ext[COL_MATRICULA].astype(str).str.strip())
            df_filtrado = df_filtrado[~df_filtrado[COL_MATRICULA].astype(str).str.strip().isin(excl)]

    return df_filtrado


def validar_corrigir(df: pd.DataFrame, ref_ano: Optional[int] = None, ref_mes: Optional[int] = None) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()

    # 1) Validar/corrigir datas conhecidas
    for col in [c for c in out.columns if "DATA" in str(c).upper() or "DATE" in str(c).upper()]:
        out[col] = pd.to_datetime(out[col], errors="coerce")

    # 2) Preencher ausências em colunas críticas
    if COL_DIAS_FERIAS in out.columns:
        out[COL_DIAS_FERIAS] = pd.to_numeric(out[COL_DIAS_FERIAS], errors="coerce").fillna(0).astype(int)

    if COL_DATA_DEMISSAO in out.columns:
        # já está em datetime; nada a preencher além de manter NaT
        pass

    # 3) Validar férias: se houver colunas de início/fim, garantir dias >= 0
    possiveis_inicio = ["INICIO FERIAS", "INÍCIO FÉRIAS", "DATA INICIO FERIAS", "DATA INÍCIO FÉRIAS"]
    possiveis_fim = ["FIM FERIAS", "FIM FÉRIAS", "DATA FIM FERIAS", "DATA FIM FÉRIAS"]
    c_inicio = detectar_coluna(out, possiveis_inicio)
    c_fim = detectar_coluna(out, possiveis_fim)
    if c_inicio and c_fim:
        out[c_inicio] = pd.to_datetime(out[c_inicio], errors="coerce")
        out[c_fim] = pd.to_datetime(out[c_fim], errors="coerce")
        delta = (out[c_fim] - out[c_inicio]).dt.days
        out["DIAS_FERIAS_CALC"] = delta.clip(lower=0).fillna(0).astype(int)
        # Se COL_DIAS_FERIAS existe e estiver vazio, preenche com calculado
        if COL_DIAS_FERIAS in out.columns:
            mask_nan = out[COL_DIAS_FERIAS].isna() | (out[COL_DIAS_FERIAS] <= 0)
            out.loc[mask_nan, COL_DIAS_FERIAS] = out.loc[mask_nan, "DIAS_FERIAS_CALC"]

    # 4) Aplicar feriados para ajustar Dias Úteis (se base de dias úteis possui a coluna numérica)
    # Suporte a diferentes nomes de coluna
    possiveis_dias_uteis = ["TOTAL_DIAS_UTEIS_SINDICATO", "DIAS UTEIS", "DIAS ÚTEIS", "QTD DIAS UTEIS", "QTD DIAS ÚTEIS"]
    col_dias_uteis = detectar_coluna(out, possiveis_dias_uteis)
    if col_dias_uteis:
        # Determinar referência de mês/ano
        if ref_ano is not None and ref_mes is not None:
            ano_ref, mes_ref = int(ref_ano), int(ref_mes)
        else:
            hoje = pd.Timestamp.today()
            ano_ref, mes_ref = hoje.year, hoje.month
        # Se a base tiver colunas de referência, tente usá-las
        for cand in ["ANO", "ANO_REF", "ANO REFERENCIA", "ANOREFERENCIA"]:
            if cand in out.columns:
                try:
                    ano_ref = int(pd.to_numeric(out[cand], errors="coerce").dropna().mode().iloc[0])
                except Exception:
                    pass
                break
        for cand in ["MES", "MÊS", "MES_REF", "MÊS REF", "MES REFERENCIA", "MÊS REFERENCIA", "MÊS REFERÊNCIA"]:
            if cand in out.columns:
                try:
                    mes_ref = int(pd.to_numeric(out[cand], errors="coerce").dropna().mode().iloc[0])
                except Exception:
                    pass
                break

        # UF/Município (inclui 'estado' derivado do Sindicato)
        uf_col = detectar_coluna(out, ["estado", COL_ESTADO, COL_UF])
        mun_col = detectar_coluna(out, [COL_MUNICIPIO])

        VALID_UFS = {"AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"}

        def ajustar(row):
            uf_val = str(row.get(uf_col, "")).strip().upper() if uf_col else None
            uf = uf_val if (uf_val and uf_val in VALID_UFS) else None
            mun = str(row.get(mun_col, "")).strip() if mun_col else None
            try:
                base = int(pd.to_numeric(row[col_dias_uteis], errors="coerce"))
            except Exception:
                return row[col_dias_uteis]
            n_fer = contar_feriados_no_mes(ano_ref, mes_ref, uf, mun)
            return max(base - n_fer, 0)

        out["DIAS_UTEIS_AJUSTADOS"] = out.apply(ajustar, axis=1)

    return out


def salvar(df: pd.DataFrame, caminho: str) -> None:
    if df is None or df.empty:
        print("[AVISO] DataFrame vazio. Nada foi salvo.")
        return
    caminho_abs = os.path.join(BASE_DIR, caminho)
    try:
        # Remove coluna calculada que não deve constar no arquivo final
        df_to_save = df.copy()
        def _norm_name(n: str) -> str:
            s = str(n).replace("_", " ").replace(".", " ")
            return _normalize_upper_noaccents_spaces(s)
        cols_drop_final = [c for c in df_to_save.columns if _norm_name(c) == "DIAS UTEIS AJUSTADOS"]
        if cols_drop_final:
            df_to_save = df_to_save.drop(columns=cols_drop_final, errors="ignore")

        # Se já existir um arquivo anterior, remove antes de salvar o novo
        if os.path.exists(caminho_abs):
            try:
                os.remove(caminho_abs)
            except Exception as e_rm:
                print(f"[ERRO] Não foi possível remover o arquivo existente: {caminho_abs}. Feche o arquivo no Excel e tente novamente. Detalhe: {e_rm}")
                raise

        # Excel é binário; parâmetro de encoding não se aplica
        df_to_save.to_excel(caminho_abs, index=False, engine="openpyxl")
        print(f"[OK] Base consolidada atualizada em: {caminho_abs}")
    except Exception as e:
        err = str(e)
        print(f"[ERRO] Falha ao substituir/salvar Excel: {err}")
        # Propaga erro para interromper execução e evitar mensagem de sucesso no final
        raise


def _parse_args():
    parser = argparse.ArgumentParser(description="Consolidar bases e gerar VR")
    parser.add_argument("--inicio", type=str, default=None, help="Data de início (YYYY-MM-DD)")
    parser.add_argument("--fim", type=str, default=None, help="Data de fim (YYYY-MM-DD)")
    return parser.parse_args()


def main() -> int:
    args = _parse_args()
    dt_inicio: Optional[date] = None
    dt_fim: Optional[date] = None
    try:
        if args.inicio:
            dt_inicio = datetime.fromisoformat(args.inicio).date()
        if args.fim:
            dt_fim = datetime.fromisoformat(args.fim).date()
    except Exception as e:
        print(f"[AVISO] Datas inválidas em parâmetros --inicio/--fim: {e}")

    print("[INFO] Iniciando consolidação de bases...")
    bases = carregar_bases()

    # Verificações básicas
    if bases["ATIVOS"].empty:
        print("[ERRO] Base principal 'ATIVOS.xlsx' não encontrada ou vazia. Processo abortado.")
        return 2

    print("[INFO] Executando merges...")
    df = executar_merges(bases)

    print("[INFO] Aplicando filtros de exclusão...")
    df = aplicar_filtros_exclusao(df, bases)

    print("[INFO] Validando e corrigindo dados...")
    # Usa mês/ano de referência baseado na data de fim, se fornecida
    if dt_fim is not None:
        df = validar_corrigir(df, ref_ano=dt_fim.year, ref_mes=dt_fim.month)
    else:
        df = validar_corrigir(df)

    print("[INFO] Salvando resultado...")
    salvar(df, ARQUIVO_SAIDA)

    # Mensagem final comunicando a atualização do arquivo
    print("[CONCLUIDO] BaseConsolidada.xlsx foi atualizada com sucesso.")

    # Dispara o agente de VR em seguida
    base_abs = os.path.join(BASE_DIR, ARQUIVO_SAIDA)
    template_abs = os.path.join(BASE_DIR, "xlsx", "VR MENSAL 05.2025.xlsx")
    # Define nome dinâmico para o arquivo final com base na Data de Fim: VR_MENSAL_MM.YYYY_FINAL.xlsx
    if dt_fim is not None:
        out_name = f"VR_MENSAL_{dt_fim.month:02d}.{dt_fim.year}_FINAL.xlsx"
    else:
        out_name = "VR_MENSAL_05.2025_FINAL.xlsx"
    output_abs = os.path.join(BASE_DIR, "xlsx", out_name)

    print("[INFO] Iniciando geração do arquivo de VR com vr_agent...")
    if run_vr_agent is None:
        print("[ERRO] vr_agent não pôde ser importado. Verifique se 'vr_agent.py' está presente e sem erros.")
        return 3
    try:
        # Define competência (MM/YYYY) a partir de dt_fim, se disponível
        competencia = None
        if dt_fim is not None:
            competencia = f"{dt_fim.month:02d}/{dt_fim.year}"
        # Datas de vigência e cutoff
        vigencia_start = dt_inicio.isoformat() if dt_inicio else None
        desligamento_cutoff = dt_fim.isoformat() if dt_fim else None

        result = run_vr_agent(
            base_path=base_abs,
            template_path=template_abs,
            output_path=output_abs,
            competencia=competencia,
            vigencia_start=vigencia_start,
            desligamento_cutoff=desligamento_cutoff,
        )
        # run_vr_agent retorna um texto final; quando o arquivo é salvo, começa com 'OUTPUT_SAVED:'
        if isinstance(result, str) and result.startswith("OUTPUT_SAVED:"):
            print(f"[CONCLUIDO] VR gerado com sucesso: {result.split(':', 1)[1]}")
            return 0
        else:
            print(f"[AVISO] Execução do vr_agent finalizou sem confirmação de arquivo salvo. Retorno: {result}")
            return 1
    except Exception as e:
        print(f"[ERRO] Falha ao executar vr_agent: {e}")
        return 4


if __name__ == "__main__":
    sys.exit(main())

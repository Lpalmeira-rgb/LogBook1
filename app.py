from datetime import datetime, timedelta
import io
import random
from typing import Dict, List, Optional, Tuple
import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# =============================================================================
# CONFIGURAÇÃO DA PÁGINA E CONSTANTES
# =============================================================================
st.set_page_config(
    page_title="Sistema de Logbook ROV",
    page_icon="⚓",
    layout="wide",
)

COLUNAS_EXIBICAO = [
    "Campo",
    "Operação",
    "Data Real (Log)",
    "Colaborador",
    "ROV Alocado",
    "Início Lançamento",
    "Fim Lançamento",
    "Total Horas (h)",
]

NOMES_PADRAO_PILOTOS = [
    "Piloto Lucas Alves",
    "Piloto Gabriel Lima",
    "Piloto Mateus Costa",
    "Piloto André Santos",
    "Piloto Extra 1",
    "Piloto Extra 2",
    "Piloto Extra 3",
    "Piloto Extra 4",
    "Piloto Extra 5",
    "Piloto Extra 6",
]

ROVS_DISPONIVEIS = ["UHD68", "UHD69"]


# =============================================================================
# FUNÇÕES UTILITÁRIAS DE HORÁRIO E PARSING
# =============================================================================
def extrair_horario(texto: str) -> Optional[Tuple[str, str]]:
    """Extrai uma tupla ('HH:MM', 'HH:MM') de uma string no formato 'HH:MM - HH:MM'."""
    if not texto or not isinstance(texto, str) or "-" not in texto:
        return None
    partes = texto.split("-")
    if len(partes) != 2:
        return None
    inicio, fim = partes[0].strip(), partes[1].strip()
    if len(inicio) == 5 and len(fim) == 5 and ":" in inicio and ":" in fim:
        try:
            datetime.strptime(inicio, "%H:%M")
            datetime.strptime(fim, "%H:%M")
            return (inicio, fim)
        except ValueError:
            return None
    return None


def parse_lista_mergulhos_texto(texto_celula: str) -> List[Tuple[str, str]]:
    """Converte múltiplos intervalos separados por vírgula ou ponto-e-vírgula."""
    if not texto_celula or not isinstance(texto_celula, str):
        return []
    blocos = texto_celula.replace(";", ",").split(",")
    return [h for b in blocos if (h := extrair_horario(b.strip())) is not None]


def obter_limites_turno(data_base_str: str, shift_noite: bool) -> Tuple[datetime, datetime]:
    """Retorna (início_turno, fim_turno) para Dia (06:00-18:00) ou Noite (18:00-06:00)."""
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
    if not shift_noite:
        return base_dt.replace(hour=6, minute=0), base_dt.replace(hour=18, minute=0)
    return base_dt.replace(hour=18, minute=0), base_dt.replace(hour=6, minute=0) + timedelta(days=1)


def parse_horario_turno(horastr: str, datastr: str, shift_noite: bool) -> datetime:
    """Converte 'HH:MM' em datetime considerando virada de meia-noite no turno noite."""
    dt = datetime.strptime(f"{datastr} {horastr}", "%Y-%m-%d %H:%M")
    if shift_noite and dt.hour < 12:
        dt += timedelta(days=1)
    return dt


def aplicar_limite_turno(
    dt_i: datetime, dt_f: datetime, data_base_str: str, shift_noite: bool
) -> Tuple[Optional[datetime], Optional[datetime]]:
    """Recorta o intervalo de mergulho para dentro das fronteiras válidas do turno."""
    limite_inicio, limite_fim = obter_limites_turno(data_base_str, shift_noite)
    eff_start = max(dt_i, limite_inicio)
    eff_end = min(dt_f, limite_fim)
    return (eff_start, eff_end) if eff_start < eff_end else (None, None)


def eh_piloto_extra(nome_p: str, idx_original: int) -> bool:
    """Identifica se o piloto é extra (posição >= 5 ou palavra 'extra' no nome)."""
    return idx_original >= 4 or "extra" in nome_p.lower()


# =============================================================================
# MOTOR DE DISTRIBUIÇÃO DE HORAS E ALOCAÇÃO
# =============================================================================
def distribuir_minutos_inteiros(
    nomes: List[str], total_minutos: int, pesos: Optional[List[float]] = None
) -> Dict[str, int]:
    """Distribui 'total_minutos' em inteiros exatos sem perda por arredondamento."""
    n = len(nomes)
    if n == 0 or total_minutos <= 0:
        return {nome: 0 for nome in nomes}

    if not pesos or sum(pesos) <= 0:
        base, sobra = divmod(total_minutos, n)
        return {nomes[i]: base + (1 if i < sobra else 0) for i in range(n)}

    pesos_limpos = [max(0.0, float(w)) for w in pesos]
    ativos = [i for i, w in enumerate(pesos_limpos) if w > 0]
    if not ativos:
        base, sobra = divmod(total_minutos, n)
        return {nomes[i]: base + (1 if i < sobra else 0) for i in range(n)}

    valores_ativos = [pesos_limpos[i] for i in ativos]
    if max(valores_ativos) - min(valores_ativos) < 1e-6:
        base, sobra = divmod(total_minutos, len(ativos))
        resultado = {nome: 0 for nome in nomes}
        for rank, idx in enumerate(ativos):
            resultado[nomes[idx]] = base + (1 if rank < sobra else 0)
        return resultado

    total_w = sum(valores_ativos)
    resultado = {nome: 0 for nome in nomes}
    restante = total_minutos
    ultimo_idx = ativos[-1]

    for idx in ativos:
        if idx == ultimo_idx:
            resultado[nomes[idx]] = restante
        else:
            fatia = int(round(total_minutos * (pesos_limpos[idx] / total_w)))
            fatia = max(0, min(restante, fatia))
            resultado[nomes[idx]] = fatia
            restante -= fatia

    return resultado


def calcular_metas_pilotos(
    pilotos_lista: List[str],
    total_minutos_disponiveis: int,
    custom_pcts: Dict[str, float],
    pct_extra_dia: Optional[float] = None,
    lista_extras_turno: Optional[List[str]] = None,
) -> Dict[str, int]:
    """Calcula a meta em minutos de cada piloto respeitando regras de extra ou customização."""
    if not pilotos_lista or total_minutos_disponiveis <= 0:
        return {p: 0 for p in pilotos_lista}

    if pct_extra_dia is not None and lista_extras_turno:
        extras_no_bloco = [p for p in pilotos_lista if p in lista_extras_turno]
        titulares_no_bloco = [p for p in pilotos_lista if p not in lista_extras_turno]

        if extras_no_bloco and titulares_no_bloco:
            pct_efetivo = max(0.0, min(100.0, float(pct_extra_dia)))
            min_extras = int(round(total_minutos_disponiveis * (pct_efetivo / 100.0)))
            min_titulares = total_minutos_disponiveis - min_extras

            metas = {}
            metas.update(distribuir_minutos_inteiros(extras_no_bloco, min_extras))
            metas.update(distribuir_minutos_inteiros(titulares_no_bloco, min_titulares))
            return metas

    if custom_pcts and any(custom_pcts.get(p, 0) > 0 for p in pilotos_lista):
        pesos = [custom_pcts.get(p, 0.0) for p in pilotos_lista]
        return distribuir_minutos_inteiros(pilotos_lista, total_minutos_disponiveis, pesos)

    return distribuir_minutos_inteiros(pilotos_lista, total_minutos_disponiveis)


def planejar_alocacao_single_rov(
    titulares_metas: Dict[str, int],
    extras_metas: Dict[str, int],
    cap_68: int,
    cap_69: int,
    fixos_68: List[str],
    fixos_69: List[str],
) -> Dict[str, List[Tuple[str, int]]]:
    """
    Planeja a alocação diária em Lançamento Único garantindo que:
    1. Cada piloto titular seja lançado em APENAS UM ROV (UHD68 ou UHD69), nunca nos dois.
    2. Se os tempos dos ROVs permitirem encaixe exato, até o piloto extra fica em APENAS UM ROV.
    3. Caso seja matematicamente necessário dividir um piloto entre os dois ROVs para fechar a
       divisão igualitária, APENAS o piloto extra fará o lançamento nos dois ROVs.
    """
    alocacoes: Dict[str, List[Tuple[str, int]]] = {}
    rem_68 = cap_68
    rem_69 = cap_69

    # 1. Aloca primeiro quem tem ROV fixo definido manualmente
    titulares_livres: List[Tuple[str, int]] = []
    for p, meta in titulares_metas.items():
        if meta <= 0:
            continue
        if p in fixos_68:
            qtd = min(meta, rem_68)
            if qtd > 0:
                alocacoes.setdefault(p, []).append(("UHD68", qtd))
                rem_68 -= qtd
            if meta - qtd > 0:
                alocacoes.setdefault(p, []).append(("UHD69", meta - qtd))
                rem_69 = max(0, rem_69 - (meta - qtd))
        elif p in fixos_69:
            qtd = min(meta, rem_69)
            if qtd > 0:
                alocacoes.setdefault(p, []).append(("UHD69", qtd))
                rem_69 -= qtd
            if meta - qtd > 0:
                alocacoes.setdefault(p, []).append(("UHD68", meta - qtd))
                rem_68 = max(0, rem_68 - (meta - qtd))
        else:
            titulares_livres.append((p, meta))

    # 2. Encontra a melhor partição dos titulares livres entre UHD68 e UHD69
    # de modo que cada titular caiba 100% inteiro em um único ROV
    random.shuffle(titulares_livres)
    n = len(titulares_livres)
    soma_titulares = sum(m for _, m in titulares_livres)

    melhor_subconjunto_68: Optional[List[int]] = None
    melhor_score = None

    for mask in range(1 << n):
        soma_68 = 0
        for i in range(n):
            if mask & (1 << i):
                soma_68 += titulares_livres[i][1]
        soma_69 =

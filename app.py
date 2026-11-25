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
        soma_69 = soma_titulares - soma_68

        if soma_68 <= rem_68 and soma_69 <= rem_69:
            sobra_68 = rem_68 - soma_68
            sobra_69 = rem_69 - soma_69
            extra_precisa_dividir = 1 if (sobra_68 > 0 and sobra_69 > 0) else 0
            qtd_68 = bin(mask).count("1")
            qtd_69 = n - qtd_68
            desequilibrio_qtd = abs(qtd_68 - qtd_69)
            score = (extra_precisa_dividir, desequilibrio_qtd)
            if melhor_score is None or score < melhor_score:
                melhor_score = score
                melhor_subconjunto_68 = [i for i in range(n) if (mask & (1 << i))]

    if melhor_subconjunto_68 is not None:
        set_68 = set(melhor_subconjunto_68)
        for i, (p, meta) in enumerate(titulares_livres):
            if i in set_68:
                alocacoes.setdefault(p, []).append(("UHD68", meta))
                rem_68 -= meta
            else:
                alocacoes.setdefault(p, []).append(("UHD69", meta))
                rem_69 -= meta
    else:
        for p, meta in sorted(titulares_livres, key=lambda x: -x[1]):
            if meta <= rem_68 and rem_68 >= rem_69:
                alocacoes.setdefault(p, []).append(("UHD68", meta))
                rem_68 -= meta
            elif meta <= rem_69:
                alocacoes.setdefault(p, []).append(("UHD69", meta))
                rem_69 -= meta
            else:
                primeiro, segundo = ("UHD68", "UHD69") if rem_68 >= rem_69 else ("UHD69", "UHD68")
                disp_1 = rem_68 if primeiro == "UHD68" else rem_69
                if disp_1 > 0:
                    alocacoes.setdefault(p, []).append((primeiro, disp_1))
                    if primeiro == "UHD68":
                        rem_68 = 0
                    else:
                        rem_69 = 0
                restante_p = meta - disp_1
                if restante_p > 0:
                    alocacoes.setdefault(p, []).append((segundo, restante_p))
                    if segundo == "UHD68":
                        rem_68 = max(0, rem_68 - restante_p)
                    else:
                        rem_69 = max(0, rem_69 - restante_p)

    # 3. Aloca os pilotos extras por último nas sobras de UHD68 e UHD69.
    # Se couber em 1 ROV, fica em 1 ROV; se precisar fechar os dois ROVs, faz a ponte.
    extras_lista = list(extras_metas.items())
    random.shuffle(extras_lista)

    for p_extra, meta_ex in extras_lista:
        if meta_ex <= 0:
            continue
        if meta_ex <= rem_68 and (rem_69 == 0 or rem_68 <= rem_69):
            alocacoes.setdefault(p_extra, []).append(("UHD68", meta_ex))
            rem_68 -= meta_ex
        elif meta_ex <= rem_69 and (rem_68 == 0 or rem_69 <= rem_68):
            alocacoes.setdefault(p_extra, []).append(("UHD69", meta_ex))
            rem_69 -= meta_ex
        elif meta_ex <= rem_68:
            alocacoes.setdefault(p_extra, []).append(("UHD68", meta_ex))
            rem_68 -= meta_ex
        elif meta_ex <= rem_69:
            alocacoes.setdefault(p_extra, []).append(("UHD69", meta_ex))
            rem_69 -= meta_ex
        else:
            restante_ex = meta_ex
            if rem_68 > 0:
                usar_68 = min(restante_ex, rem_68)
                alocacoes.setdefault(p_extra, []).append(("UHD68", usar_68))
                rem_68 -= usar_68
                restante_ex -= usar_68
            if restante_ex > 0 and rem_69 > 0:
                usar_69 = min(restante_ex, rem_69)
                alocacoes.setdefault(p_extra, []).append(("UHD69", usar_69))
                rem_69 -= usar_69
                restante_ex -= usar_69

    return alocacoes


def calcular_tempo_total_rov(
    mergulhos_str: List[Tuple[str, str]], data_base_str: str, shift_noite: bool
) -> Tuple[int, datetime]:
    """Soma os minutos válidos dentro do turno e ancora o início para caber no limite."""
    limite_inicio, limite_fim = obter_limites_turno(data_base_str, shift_noite)
    total_minutos = 0
    primeiros_inicios = []

    for i, f in mergulhos_str:
        dt_i = parse_horario_turno(i, data_base_str, shift_noite)
        dt_f = parse_horario_turno(f, data_base_str, shift_noite)
        if dt_f <= dt_i:
            dt_f += timedelta(days=1)

        s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, shift_noite)
        if s and e:
            total_minutos += int((e - s).total_seconds() / 60)
            primeiros_inicios.append(s)

    max_duracao_turno = int((limite_fim - limite_inicio).total_seconds() / 60)
    total_minutos = min(total_minutos, max_duracao_turno)

    inicio_base = min(primeiros_inicios) if primeiros_inicios else limite_inicio
    ultimo_inicio_possivel = limite_fim - timedelta(minutes=total_minutos)
    if inicio_base > ultimo_inicio_possivel:
        inicio_base = max(limite_inicio, ultimo_inicio_possivel)

    return total_minutos, inicio_base


def encontrar_janela_no_turno(
    agenda_pessoa: List[Tuple[datetime, datetime]],
    inicio_preferido: datetime,
    duracao_min: int,
    limite_inicio: datetime,
    limite_fim: datetime,
) -> Tuple[datetime, datetime]:
    """Encontra um intervalo sem sobreposição estritamente dentro de [limite_inicio, limite_fim]."""
    dur_td = timedelta(minutes=duracao_min)
    ultimo_inicio_valido = max(limite_inicio, limite_fim - dur_td)

    def livre(s: datetime, e: datetime) -> bool:
        if s < limite_inicio or e > limite_fim:
            return False
        return all(max(s, os) >= min(e, oe) for os, oe in agenda_pessoa)

    cand = max(limite_inicio, min(inicio_preferido, ultimo_inicio_valido))
    if livre(cand, cand + dur_td):
        return cand, cand + dur_td

    pontos_candidatos = [limite_inicio, ultimo_inicio_valido]
    for os, oe in agenda_pessoa:
        pontos_candidatos.extend([oe, os - dur_td])

    pontos_validos = sorted(
        {max(limite_inicio, min(p, ultimo_inicio_valido)) for p in pontos_candidatos},
        key=lambda p: abs((p - cand).total_seconds()),
    )

    for p in pontos_validos:
        if livre(p, p + dur_td):
            return p, p + dur_td

    return cand, min(limite_fim, cand + dur_td)


def criar_registro(
    campo: str,
    operacao: str,
    data_base_str: str,
    b_inicio: datetime,
    b_fim: datetime,
    shift_noite: bool,
    modo_str: str,
    colaborador: str,
    cargo: str,
    rov: str,
    duracao_min: int,
) -> dict:
    """Padroniza a estrutura de cada linha gerada no logbook."""
    return {
        "Campo": campo,
        "Operação": operacao,
        "Data Lançamento": data_base_str,
        "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
        "Turno": "Noite" if shift_noite else "Dia",
        "Modo": modo_str,
        "Colaborador": colaborador,
        "Cargo": cargo,
        "ROV Alocado": rov,
        "Início Lançamento": b_inicio.strftime("%H:%M"),
        "Fim Lançamento": b_fim.strftime("%H:%M"),
        "Duração (min)": duracao_min,
        "Total Horas (h)": round(duracao_min / 60.0, 2),
    }


def processar_um_dia(
    data_base_str: str,
    mergulhos_68_inputs: List[Tuple[str, str]],
    mergulhos_69_inputs: List[Tuple[str, str]],
    campo_emb: str,
    nome_op: str,
    shift_noite: bool,
    modo_unico: bool,
    s_68: str,
    s_69: str,
    lista_pilotos: List[str],
    modo_alloc: str,
    alloc_man: Dict[str, str],
    pct_sup: float,
    pct_pilotos_cust: Dict[str, float],
    pct_extra_dia: Optional[float] = None,
) -> Tuple[List[dict], Dict[str, float]]:
    """Processa todos os lançamentos de um único turno/dia."""
    limite_inicio, limite_fim = obter_limites_turno(data_base_str, shift_noite)
    supervisores = {"UHD68": s_68, "UHD69": s_69}

    titulares = [p for idx, p in enumerate(lista_pilotos) if not eh_piloto_extra(p, idx)]
    extras = [p for idx, p in enumerate(lista_pilotos) if eh_piloto_extra(p, idx)]

    pilotos_ativos_dia = lista_pilotos.copy()
    if pct_extra_dia is not None and float(pct_extra_dia) <= 0.0 and extras:
        pilotos_ativos_dia = [p for p in lista_pilotos if p not in extras]

    if pct_pilotos_cust and any(pct_pilotos_cust.get(p, 0) > 0 for p in pilotos_ativos_dia):
        pilotos_ativos_dia = [
            p for p in pilotos_ativos_dia if pct_pilotos_cust.get(p, 0) > 0
        ]

    registros: List[dict] = []
    minutos_dia: Dict[str, int] = {p: 0 for p in [s_68, s_69] + lista_pilotos}

    # -------------------------------------------------------------------------
    # MODO 1: LANÇAMENTO ÚNICO (CONSOLIDADO NO TURNO)
    # -------------------------------------------------------------------------
    if modo_unico:
        minutos_68, inicio_68 = calcular_tempo_total_rov(
            mergulhos_68_inputs, data_base_str, shift_noite
        )
        minutos_69, inicio_69 = calcular_tempo_total_rov(
            mergulhos_69_inputs, data_base_str, shift_noite
        )

        rov_configs = [
            ("UHD68", minutos_68, s_68, inicio_68),
            ("UHD69", minutos_69, s_69, inicio_69),
        ]

        rov_pools = {}
        total_pool_pilotos = 0

        for rov, tot_min, sup_fixo, dt_start in rov_configs:
            if tot_min <= 0:
                continue

            sup_min = max(0, int(round(tot_min * (pct_sup / 100.0))))
            p_min = tot_min - sup_min
            total_pool_pilotos += p_min

            if sup_min > 0:
                b_fim_sup = min(limite_fim, dt_start + timedelta(minutes=sup_min))
                minutos_dia[sup_fixo] += sup_min
                registros.append(
                    criar_registro(
                        campo_emb, nome_op, data_base_str, dt_start, b_fim_sup,
                        shift_noite, "Lançamento Único", sup_fixo, "Supervisor", rov, sup_min
                    )
                )

            rov_pools[rov] = {
                "p_min_restante": p_min,
                "cursor": dt_start + timedelta(minutes=sup_min),
            }

        metas_pilotos = calcular_metas_pilotos(
            pilotos_ativos_dia,
            total_pool_pilotos,
            pct_pilotos_cust,
            pct_extra_dia=pct_extra_dia,
            lista_extras_turno=extras,
        )

        cap_68 = rov_pools["UHD68"]["p_min_restante"] if "UHD68" in rov_pools else 0
        cap_69 = rov_pools["UHD69"]["p_min_restante"] if "UHD69" in rov_pools else 0

        fixos_68: List[str] = []
        fixos_69: List[str] = []
        titulares_metas: Dict[str, int] = {}
        extras_metas: Dict[str, int] = {}

        for p in pilotos_ativos_dia:
            meta_p = metas_pilotos.get(p, 0)
            if meta_p <= 0:
                continue

            if modo_alloc.startswith("Manual"):
                escolha = alloc_man.get(p, "Ambos (UHD68 e UHD69)")
                if escolha == "UHD68":
                    fixos_68.append(p)
                    titulares_metas[p] = meta_p
                elif escolha == "UHD69":
                    fixos_69.append(p)
                    titulares_metas[p] = meta_p
                else:
                    extras_metas[p] = meta_p
            else:
                if p in extras:
                    extras_metas[p] = meta_p
                else:
                    titulares_metas[p] = meta_p

        plano_rov = planejar_alocacao_single_rov(
            titulares_metas=titulares_metas,
            extras_metas=extras_metas,
            cap_68=cap_68,
            cap_69=cap_69,
            fixos_68=fixos_68,
            fixos_69=fixos_69,
        )

        agenda_pilotos: Dict[str, List[Tuple[datetime, datetime]]] = {
            p: [] for p in pilotos_ativos_dia
        }

        pilotos_ordenados_exec = sorted(
            plano_rov.keys(),
            key=lambda p: (len(plano_rov[p]), 0 if p in titulares else 1),
        )

        for p in pilotos_ordenados_exec:
            for rov, alocado in plano_rov[p]:
                if alocado <= 0 or rov not in rov_pools:
                    continue
                pool = rov_pools[rov]
                b_inicio, b_fim = encontrar_janela_no_turno(
                    agenda_pilotos[p],
                    pool["cursor"],
                    alocado,
                    limite_inicio,
                    limite_fim,
                )

                pool["cursor"] = min(limite_fim, pool["cursor"] + timedelta(minutes=alocado))
                pool["p_min_restante"] = max(0, pool["p_min_restante"] - alocado)
                agenda_pilotos[p].append((b_inicio, b_fim))
                minutos_dia[p] += alocado

                registros.append(
                    criar_registro(
                        campo_emb, nome_op, data_base_str, b_inicio, b_fim,
                        shift_noite, "Lançamento Único", p, "Piloto", rov, alocado
                    )
                )

    # -------------------------------------------------------------------------
    # MODO 2: HORÁRIOS EXATOS (MÚLTIPLOS LANÇAMENTOS POR MERGULHO)
    # -------------------------------------------------------------------------
    else:
        lista_mergulhos = []
        for rov, inputs in [("UHD68", mergulhos_68_inputs), ("UHD69", mergulhos_69_inputs)]:
            for i, f in inputs:
                dt_i = parse_horario_turno(i, data_base_str, shift_noite)
                dt_f = parse_horario_turno(f, data_base_str, shift_noite)
                if dt_f <= dt_i:
                    dt_f += timedelta(days=1)
                s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, shift_noite)
                if s and e:
                    lista_mergulhos.append((rov, s, e))

        lista_mergulhos.sort(key=lambda m: m[1])

        total_min_pilotos_turno = sum(
            max(
                0,
                int((dt_f - dt_i).total_seconds() / 60)
                - max(0, int(round(int((dt_f - dt_i).total_seconds() / 60) * (pct_sup / 100.0)))),
            )
            for _, dt_i, dt_f in lista_mergulhos
        )

        metas_turno_m2 = calcular_metas_pilotos(
            pilotos_ativos_dia,
            total_min_pilotos_turno,
            pct_pilotos_cust,
            pct_extra_dia=pct_extra_dia,
            lista_extras_turno=extras,
        )
        minutos_alocados_m2 = {p: 0 for p in pilotos_ativos_dia}
        agenda_ocupacao: Dict[str, List[Tuple[datetime, datetime]]] = {
            p: [] for p in minutos_dia
        }

        def tem_conflito(pessoa: str, inicio: datetime, fim: datetime) -> bool:
            return any(max(inicio, s) < min(fim, e) for s, e in agenda_ocupacao[pessoa])

        preferencia_rov_m2: Dict[str, List[str]] = {}
        if modo_alloc.startswith("Manual"):
            for p in pilotos_ativos_dia:
                esc = alloc_man.get(p, "Ambos (UHD68 e UHD69)")
                preferencia_rov_m2[p] = (
                    [esc] if esc in ROVS_DISPONIVEIS else ROVS_DISPONIVEIS.copy()
                )
        else:
            tit_emb = [p for p in titulares if p in pilotos_ativos_dia]
            random.shuffle(tit_emb)
            meio = (len(tit_emb) + 1) // 2
            for p in tit_emb[:meio]:
                preferencia_rov_m2[p] = ["UHD68"]
            for p in tit_emb[meio:]:
                preferencia_rov_m2[p] = ["UHD69"]
            for p in [e for e in extras if e in pilotos_ativos_dia]:
                preferencia_rov_m2[p] = ROVS_DISPONIVEIS.copy()

        for rov, dt_inicio, dt_fim in lista_mergulhos:
            duracao_min = int((dt_fim - dt_inicio).total_seconds() / 60)
            sup_fixo = supervisores[rov]

            sup_minutos = max(0, int(round(duracao_min * (pct_sup / 100.0))))
            pilotos_minutos = duracao_min - sup_minutos

            def prioridade_piloto(p: str):
                habilitado = 0 if rov in preferencia_rov_m2.get(p, []) else 1
                saldo_restante = metas_turno_m2.get(p, 0) - minutos_alocados_m2.get(p, 0)
                precisa_horas = 0 if saldo_restante > 0 else 1
                return (habilitado, precisa_horas, -saldo_restante, minutos_dia[p], random.random())

            habilitados_rov = [
                p for p in pilotos_ativos_dia if rov in preferencia_rov_m2.get(p, [])
            ]
            candidatos = habilitados_rov if habilitados_rov else pilotos_ativos_dia
            pilotos_ordenados = sorted(candidatos, key=prioridade_piloto)

            qtd_pilotos_mergulho = min(
                len(pilotos_ordenados), 3 if len(pilotos_ordenados) >= 4 else 2
            )

            pilotos_disponiveis = [
                p for p in pilotos_ordenados if not tem_conflito(p, dt_inicio, dt_fim)
            ][:qtd_pilotos_mergulho]

            if not pilotos_disponiveis and pilotos_ordenados:
                pilotos_disponiveis = [pilotos_ordenados[0]]

            def dividir_bloco_equilibrado(lista_disp: List[str], mins_totais: int) -> Dict[str, int]:
                saldos = {
                    p: max(1, metas_turno_m2.get(p, 0) - minutos_alocados_m2.get(p, 0))
                    for p in lista_disp
                }
                return calcular_metas_pilotos(lista_disp, mins_totais, saldos)

            slots = []
            sup_dt_e = dt_inicio + timedelta(minutes=sup_minutos)
            if sup_minutos > 0 and not tem_conflito(sup_fixo, dt_inicio, sup_dt_e):
                slots.append((sup_fixo, sup_minutos, "Supervisor"))
            else:
                pilotos_minutos = duracao_min

            metas_bloco = dividir_bloco_equilibrado(pilotos_disponiveis, pilotos_minutos)
            for p in pilotos_disponiveis:
                slots.append((p, metas_bloco[p], "Piloto"))

            cursor_tempo = dt_inicio
            for colaborador, dur_m, cargo in slots:
                if dur_m <= 0:
                    continue
                b_inicio = max(limite_inicio, cursor_tempo)
                b_fim = min(limite_fim, b_inicio + timedelta(minutes=dur_m))
                cursor_tempo = b_fim
                dur_efetiva = int((b_fim - b_inicio).total_seconds() / 60)
                if dur_efetiva <= 0:
                    continue

                agenda_ocupacao[colaborador].append((b_inicio, b_fim))
                minutos_dia[colaborador] += dur_efetiva
                if cargo == "Piloto":
                    minutos_alocados_m2[colaborador] = (
                        minutos_alocados_m2.get(colaborador, 0) + dur_efetiva
                    )

                registros.append(
                    criar_registro(
                        campo_emb, nome_op, data_base_str, b_inicio, b_fim,
                        shift_noite, "Horários Exatos", colaborador, cargo, rov, dur_efetiva
                    )
                )

    horas_dia = {p: round(m / 60.0, 2) for p, m in minutos_dia.items()}
    return registros, horas_dia


# =============================================================================
# EXPORTAÇÃO EXCEL FORMATADA E COMPONENTES DE RESULTADO REUTILIZÁVEIS
# =============================================================================
@st.cache_data(show_spinner=False)
def gerar_excel_bytes(df: pd.DataFrame) -> bytes:
    """Gera arquivo .xlsx com largura de colunas ajustada automaticamente."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Logbook")
        ws = writer.sheets["Logbook"]
        for col_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(
                [len(str(col_name))] + [len(str(val)) for val in df[col_name].values]
            )
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 40)
    return output.getvalue()


def renderizar_painel_resultados(
    df_resultado: pd.DataFrame,
    horas_resumo: Dict[str, float],
    nome_arquivo_base: str,
    titulo_resumo: str,
    coluna_horas_resumo: str,
    pilotos_lista: List[str],
    habilitar_filtros: bool = False,
):
    """Renderiza métricas KPI, filtros (opcional), tabela, download Excel e resumo de horas."""
    df_exibir = df_resultado.copy()
    filtro_colaboradores: List[str] = []

    if habilitar_filtros:
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            colaboradores_unicos = sorted(df_exibir["Colaborador"].unique().tolist())
            filtro_colaboradores = st.multiselect(
                "🔍 Filtrar Tabela por Piloto / Colaborador (deixe vazio para ver todos):",
                options=colaboradores_unicos,
                default=[],
                key="filtro_piloto_massa",
            )
        with col_f2:
            rovs_unicos = sorted(df_exibir["ROV Alocado"].unique().tolist())
            filtro_rov = st.multiselect(
                "🚤 Filtrar por ROV:",
                options=rovs_unicos,
                default=[],
                key="filtro_rov_massa",
            )

        if filtro_colaboradores:
            df_exibir = df_exibir[df_exibir["Colaborador"].isin(filtro_colaboradores)]
        if filtro_rov:
            df_exibir = df_exibir[df_exibir["ROV Alocado"].isin(filtro_rov)]

    total_horas = round(df_exibir["Total Horas (h)"].sum(), 2)
    horas_pilotos_ativos = [
        h for colab, h in horas_resumo.items() if colab in pilotos_lista and h > 0
    ]
    media_pilotos = (
        round(sum(horas_pilotos_ativos) / len(horas_pilotos_ativos), 2)
        if horas_pilotos_ativos
        else 0.0
    )
    diferenca_max = (
        round(max(horas_pilotos_ativos) - min(horas_pilotos_ativos), 2)
        if len(horas_pilotos_ativos) >= 2
        else 0.0
    )

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total de Lançamentos", f"{len(df_exibir)} linhas")
    k2.metric("Total Horas na Tabela", f"{total_horas:.2f} h")
    k3.metric("Média por Piloto Ativo", f"{media_pilotos:.2f} h")
    k4.metric("Variação Máx. entre Pilotos", f"{diferenca_max:.2f} h")

    st.dataframe(df_exibir, use_container_width=True, hide_index=True)

    sufixo = (
        f"_{'_'.join(c.replace(' ', '_') for c in filtro_colaboradores[:2])}"
        if filtro_colaboradores
        else ""
    )
    st.download_button(
        label="📥 Baixar Planilha Excel (.xlsx)",
        data=gerar_excel_bytes(df_exibir),
        file_name=f"{nome_arquivo_base}{sufixo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.subheader(titulo_resumo)
    df_saldos = pd.DataFrame(
        list(horas_resumo.items()),
        columns=["Colaborador", coluna_horas_resumo],
    )
    if filtro_colaboradores:
        df_saldos = df_saldos[df_saldos["Colaborador"].isin(filtro_colaboradores)]
    st.dataframe(df_saldos, use_container_width=True, hide_index=True)


# =============================================================================
# INTERFACE PRINCIPAL E NAVEGAÇÃO
# =============================================================================
st.title("⚓ Sistema de Lançamento de Horários ROV")
st.markdown(
    "Gerencie lançamentos de turno **Individual** ou **Em Massa (Vários Dias)**, "
    "com consolidação automática ou divisão exata por mergulho."
)

if "tela_ativa" not in st.session_state:
    st.session_state["tela_ativa"] = "individual"

col_nav1, col_nav2, _ = st.columns([2, 2, 4])
with col_nav1:
    if st.button(
        "📅 Lançamento Diário (Individual)",
        type="primary" if st.session_state["tela_ativa"] == "individual" else "secondary",
        use_container_width=True,
    ):
        st.session_state["tela_ativa"] = "individual"
        st.rerun()

with col_nav2:
    if st.button(
        "📚 Lançamento em Massa (Vários Dias)",
        type="primary" if st.session_state["tela_ativa"] == "massa" else "secondary",
        use_container_width=True,
    ):
        st.session_state["tela_ativa"] = "massa"
        st.rerun()

st.divider()

# =============================================================================
# BARRA LATERAL (CONFIGURAÇÕES DA EQUIPE E REGRAS)
# =============================================================================
st.sidebar.header("⚙️ Configurações da Equipe e Regras")

campo_embarcacao = st.sidebar.text_input("Campo de Atuação da Embarcação", "Campo de Búzios")
nome_operacao = st.sidebar.text_input("Nome da Operação", "Inspeção Subsea / Intervenção")

turno_selecionado = st.sidebar.selectbox(
    "Selecione o Turno", ["Dia (06:00 - 18:00)", "Noite (18:00 - 06:00)"]
)

modo_lancamento = st.sidebar.radio(
    "Modo de Lançamento por Operador",
    [
        "1 - Lançamento Único (Consolidado no Turno)",
        "Múltiplos - Horários Exatos dos Mergulhos",
    ],
    help=(
        "Escolha '1' para consolidar o tempo do turno com divisão exata entre "
        "os pilotos em apenas um ROV, ou 'Múltiplos' para dividir cada mergulho individualmente."
    ),
)

is_turno_noite = "Noite" in turno_selecionado
is_modo_unico = "1" in modo_lancamento

st.sidebar.subheader("👨‍✈️ Supervisores (Fixos)")
sup_68 = st.sidebar.text_input("Supervisor UHD68", "Sup. João Silva")
sup_69 = st.sidebar.text_input("Supervisor UHD69", "Sup. Marcos Rocha")

st.sidebar.subheader("👨‍💻 Pilotos (Equipe Dinâmica)")
qtd_pilotos = st.sidebar.number_input(
    "Quantidade de Pilotos no Turno", min_value=1, max_value=12, value=5, step=1
)

pilotos: List[str] = []
for i in range(int(qtd_pilotos)):
    p_default = NOMES_PADRAO_PILOTOS[i] if i < len(NOMES_PADRAO_PILOTOS) else f"Piloto {i+1}"
    label_p = f"Piloto {i+1} (Extra)" if i >= 4 else f"Piloto {i+1}"
    p_nome = st.sidebar.text_input(label_p, p_default, key=f"piloto_input_{i}")
    if p_nome.strip():
        pilotos.append(p_nome.strip())

st.sidebar.subheader("🚤 Alocação nos ROVs")
modo_alocacao_rov = st.sidebar.radio(
    "Forma de Alocação dos Pilotos",
    [
        "Aleatória (1 ROV por Piloto; Extra faz ponte se necessário)",
        "Manual (Escolher ROV por Piloto)",
    ],
)

alocacao_manual: Dict[str, str] = {}
if modo_alocacao_rov.startswith("Manual"):
    st.sidebar.markdown("**Defina o ROV de cada piloto:**")
    for idx, p in enumerate(pilotos):
        idx_padrao = 0 if idx < 2 else (1 if idx < 4 else 2)
        alocacao_manual[p] = st.sidebar.selectbox(
            f"ROV para {p}",
            ["UHD68", "UHD69", "Ambos (UHD68 e UHD69)"],
            index=idx_padrao,
            key=f"alloc_man_{idx}_{p}",
        )

st.sidebar.subheader("⚖️ Porcentagens da Divisão")
pct_supervisor = st.sidebar.number_input(
    "Porcentagem do Supervisor (%)",
    min_value=0.0,
    max_value=50.0,
    value=10.0,
    step=1.0,
    help="Porcentagem do tempo total de mergulho atribuída ao supervisor fixo.",
)

tipo_divisao_pilotos = st.sidebar.radio(
    "Divisão entre os Pilotos",
    ["Igualitária (Divisão Equitativa)", "Personalizada por Piloto (%)"],
)

pct_pilotos_custom: Dict[str, float] = {}
soma_pct_pilotos = 100.0
excede_100_pct = False

if tipo_divisao_pilotos == "Personalizada por Piloto (%)":
    st.sidebar.markdown("**Defina a porcentagem de cada piloto (Máx. 100% no total):**")
    val_padrao = round(100.0 / len(pilotos), 1) if pilotos else 25.0

    if st.sidebar.button("🔄 Distribuir 100% Igualmente", use_container_width=True):
        if pilotos:
            base_pct = round(100.0 / len(pilotos), 1)
            for idx_p, p_item in enumerate(pilotos):
                if idx_p == len(pilotos) - 1:
                    st.session_state[f"pct_{p_item}"] = round(
                        100.0 - base_pct * (len(pilotos) - 1), 1
                    )
                else:
                    st.session_state[f"pct_{p_item}"] = base_pct
            st.rerun()

    valores_atuais = {
        p: float(st.session_state.get(f"pct_{p}", val_padrao)) for p in pilotos
    }

    for p in pilotos:
        k = f"pct_{p}"
        soma_outros = sum(v for nome_p, v in valores_atuais.items() if nome_p != p)
        max_permitido = max(0.0, round(100.0 - soma_outros, 1))
        valor_inicial = min(valores_atuais[p], max_permitido)

        novo_val = st.sidebar.number_input(
            f"% para {p} (disponível: {max_permitido:.1f}%)",
            min_value=0.0,
            max_value=max_permitido,
            value=valor_inicial,
            step=1.0,
            key=k,
        )
        pct_pilotos_custom[p] = novo_val
        valores_atuais[p] = novo_val

    soma_pct_pilotos = round(sum(pct_pilotos_custom.values()), 1)
    restante_pct = round(100.0 - soma_pct_pilotos, 1)

    if soma_pct_pilotos > 100.0:
        excede_100_pct = True
        st.sidebar.error(
            f"🚫 Soma das porcentagens em **{soma_pct_pilotos:.1f}%** (limite máximo: 100%)."
        )
    elif soma_pct_pilotos < 100.0:
        st.sidebar.warning(
            f"⚠️ Total alocado: **{soma_pct_pilotos:.1f}%** (faltam **{restante_pct:.1f}%**)."
        )
    else:
        st.sidebar.success("✅ Total alocado: **100.0%** (Exato)")

# =============================================================================
# TELA 1: LANÇAMENTO DIÁRIO (INDIVIDUAL)
# =============================================================================
if st.session_state["tela_ativa"] == "individual":
    st.subheader("📅 Lançamento Diário Individual")

    data_logbook = st.date_input(
        "Data de Lançamento (Início do Turno)", datetime.now().date(), key="data_indiv"
    )

    col1, col2 = st.columns(2)
    padrao_68 = (
        ["18:00 - 20:00", "02:00 - 06:00", "", "", ""]
        if is_turno_noite
        else ["06:30 - 11:00", "14:00 - 17:30", "", "", ""]
    )
    padrao_69 = (
        ["19:00 - 01:00", "", "", "", ""]
        if is_turno_noite
        else ["08:00 - 12:30", "15:00 - 18:00", "", "", ""]
    )

    with col1:
        st.markdown("#### 🔴 Mergulhos ROV UHD68")
        inputs_68 = [
            st.text_input(f"Mergulho {idx+1}", padrao_68[idx], key=f"d68_{idx+1}")
            for idx in range(5)
        ]

    with col2:
        st.markdown("#### 🔵 Mergulhos ROV UHD69")
        inputs_69 = [
            st.text_input(f"Mergulho {idx+1}", padrao_69[idx], key=f"d69_{idx+1}")
            for idx in range(5)
        ]

    if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
        if not pilotos:
            st.error("⚠️ Adicione pelo menos 1 piloto na barra lateral antes de gerar.")
            st.stop()

        if excede_100_pct:
            st.error(
                f"🚫 A soma das porcentagens está em {soma_pct_pilotos:.1f}%. Ajuste para não ultrapassar 100%."
            )
            st.stop()

        mergulhos_68_inputs = [h for m in inputs_68 if (h := extrair_horario(m)) is not None]
        mergulhos_69_inputs = [h for m in inputs_69 if (h := extrair_horario(m)) is not None]

        if not mergulhos_68_inputs and not mergulhos_69_inputs:
            st.warning(
                "⚠️ Nenhum horário válido encontrado. Preencha no formato HH:MM - HH:MM (ex: 06:30 - 11:00)."
            )
            st.stop()

        data_base_str = data_logbook.strftime("%Y-%m-%d")
        registros, horas_do_dia = processar_um_dia(
            data_base_str,
            mergulhos_68_inputs,
            mergulhos_69_inputs,
            campo_embarcacao,
            nome_operacao,
            is_turno_noite,
            is_modo_unico,
            sup_68,
            sup_69,
            pilotos,
            modo_alocacao_rov,
            alocacao_manual,
            pct_supervisor,
            pct_pilotos_custom,
            pct_extra_dia=None,
        )

        if registros:
            st.session_state["df_exibicao"] = pd.DataFrame(registros)[COLUNAS_EXIBICAO]
            st.session_state["horas_do_dia"] = horas_do_dia
            st.session_state["data_base_str"] = data_base_str
        else:
            st.warning("⚠️ Nenhum lançamento válido gerado dentro do horário do turno selecionado.")

    if "df_exibicao" in st.session_state:
        st.success("✅ Tabela do Logbook Gerada!")
        renderizar_painel_resultados(
            df_resultado=st.session_state["df_exibicao"],
            horas_resumo=st.session_state["horas_do_dia"],
            nome_arquivo_base=f"Logbook_ROV_{st.session_state['data_base_str']}",
            titulo_resumo="📊 Resumo de Horas do Turno (Verificação de Equilíbrio)",
            coluna_horas_resumo="Total Horas no Dia (h)",
            pilotos_lista=pilotos,
            habilitar_filtros=False,
        )

# =============================================================================
# TELA 2: LANÇAMENTO EM MASSA (VÁRIOS DIAS)
# =============================================================================
else:
    st.subheader("📚 Lançamento em Massa (Vários Dias Consecutivos)")
    st.markdown(
        "Configure o **período inicial e a quantidade de dias**. Na coluna **% Piloto Extra**, "
        "defina a porcentagem diária destinada ao piloto extra (ex.: `20` para 20%, `0` para folga)."
    )

    c_m1, c_m2, c_m3 = st.columns([2, 2, 3])
    with c_m1:
        data_inicio_massa = st.date_input(
            "Data Inicial do Período", datetime.now().date(), key="data_ini_massa"
        )
    with c_m2:
        qtd_dias_massa = st.number_input(
            "Quantidade de Dias para Lançar",
            min_value=2,
            max_value=60,
            value=7,
            step=1,
            key="qtd_dias_massa",
        )
    with c_m3:
        usar_pct_extra_diario = st.checkbox(
            "Variar % do Piloto Extra dia a dia na tabela",
            value=True,
            help="Usa a coluna '% Piloto Extra' de cada linha para calcular a cota diária do(s) extra(s).",
        )

    padrao_str_68 = (
        "18:00 - 20:00, 02:00 - 06:00" if is_turno_noite else "06:30 - 11:00, 14:00 - 17:30"
    )
    padrao_str_69 = (
        "19:00 - 01:00" if is_turno_noite else "08:00 - 12:30, 15:00 - 18:00"
    )
    pct_extra_sugerido = round(100.0 / len(pilotos), 1) if pilotos else 20.0

    linhas_grade = [
        {
            "Data": (data_inicio_massa + timedelta(days=d)).strftime("%Y-%m-%d"),
            "Operação": nome_operacao,
            "Mergulhos UHD68": padrao_str_68,
            "Mergulhos UHD69": padrao_str_69,
            "% Piloto Extra": pct_extra_sugerido,
        }
        for d in range(int(qtd_dias_massa))
    ]

    st.markdown("#### ✏️ Grade Editável de Horários e Porcentagem Diária do Extra")
    df_editado = st.data_editor(
        pd.DataFrame(linhas_grade),
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Data": st.column_config.TextColumn("Data (AAAA-MM-DD)", required=True),
            "Operação": st.column_config.TextColumn("Operação"),
            "Mergulhos UHD68": st.column_config.TextColumn("Mergulhos UHD68 (HH:MM - HH:MM)"),
            "Mergulhos UHD69": st.column_config.TextColumn("Mergulhos UHD69 (HH:MM - HH:MM)"),
            "% Piloto Extra": st.column_config.NumberColumn(
                "% Piloto Extra",
                min_value=0.0,
                max_value=100.0,
                step=1.0,
                format="%.1f%%",
                help="Porcentagem do tempo líquido destinada ao(s) piloto(s) extra(s) neste dia.",
            ),
        },
        key="editor_massa",
    )

    if st.button("🚀 Gerar Todos os Lançamentos em Massa", type="primary"):
        if not pilotos:
            st.error("⚠️ Adicione pelo menos 1 piloto na barra lateral antes de gerar.")
            st.stop()

        if excede_100_pct:
            st.error(
                f"🚫 A soma das porcentagens está em {soma_pct_pilotos:.1f}%. Ajuste para não ultrapassar 100%."
            )
            st.stop()

        todos_registros: List[dict] = []
        horas_acumuladas_periodo: Dict[str, float] = {p: 0.0 for p in [sup_68, sup_69] + pilotos}
        dias_processados = 0

        for _, row in df_editado.iterrows():
            raw_data = row.get("Data")
            if pd.isna(raw_data) or not str(raw_data).strip():
                continue

            data_str = str(raw_data).strip()
            try:
                datetime.strptime(data_str, "%Y-%m-%d")
            except ValueError:
                st.warning(f"⚠️ Data ignorada por formato inválido: '{data_str}' (use AAAA-MM-DD).")
                continue

            raw_op = row.get("Operação")
            op_dia = str(raw_op).strip() if pd.notna(raw_op) and str(raw_op).strip() else nome_operacao

            m68_lista = parse_lista_mergulhos_texto(
                str(row.get("Mergulhos UHD68")) if pd.notna(row.get("Mergulhos UHD68")) else ""
            )
            m69_lista = parse_lista_mergulhos_texto(
                str(row.get("Mergulhos UHD69")) if pd.notna(row.get("Mergulhos UHD69")) else ""
            )
            if not m68_lista and not m69_lista:
                continue

            pct_ex_val = None
            if usar_pct_extra_diario:
                try:
                    raw_val = row.get("% Piloto Extra", pct_extra_sugerido)
                    pct_ex_val = max(
                        0.0,
                        min(100.0, float(raw_val) if pd.notna(raw_val) else pct_extra_sugerido),
                    )
                except (ValueError, TypeError):
                    pct_ex_val = pct_extra_sugerido

            regs_dia, hrs_dia = processar_um_dia(
                data_str,
                m68_lista,
                m69_lista,
                campo_embarcacao,
                op_dia,
                is_turno_noite,
                is_modo_unico,
                sup_68,
                sup_69,
                pilotos,
                modo_alocacao_rov,
                alocacao_manual,
                pct_supervisor,
                pct_pilotos_custom,
                pct_extra_dia=pct_ex_val,
            )

            if regs_dia:
                todos_registros.extend(regs_dia)
                dias_processados += 1
                for colab, h_val in hrs_dia.items():
                    horas_acumuladas_periodo[colab] = round(
                        horas_acumuladas_periodo.get(colab, 0.0) + h_val, 2
                    )

        if todos_registros:
            st.session_state["df_massa_exibicao"] = pd.DataFrame(todos_registros)[COLUNAS_EXIBICAO]
            st.session_state["horas_massa"] = horas_acumuladas_periodo
            st.session_state["dias_proc_massa"] = dias_processados
        else:
            st.warning(
                "⚠️ Nenhum lançamento válido gerado. Verifique se os horários correspondem ao turno."
            )

    if "df_massa_exibicao" in st.session_state:
        dias_proc = st.session_state["dias_proc_massa"]
        df_massa = st.session_state["df_massa_exibicao"]
        st.success(
            f"✅ Lançamento em Massa Concluído! ({dias_proc} dias processados | {len(df_massa)} linhas geradas)"
        )
        renderizar_painel_resultados(
            df_resultado=df_massa,
            horas_resumo=st.session_state["horas_massa"],
            nome_arquivo_base=f"Logbook_ROV_Massa_{data_inicio_massa.strftime('%Y-%m-%d')}_{dias_proc}dias",
            titulo_resumo="📊 Resumo Acumulado de Horas no Período",
            coluna_horas_resumo="Total Horas Acumuladas no Período (h)",
            pilotos_lista=pilotos,
            habilitar_filtros=True,
        )

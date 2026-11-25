from datetime import datetime, timedelta
import io
import random
import openpyxl
import pandas as pd
import streamlit as st


# =============================================================================
# FUNÇÕES AUXILIARES E DE CÁLCULO
# =============================================================================
def extrair_horario(texto):
    if not texto or not isinstance(texto, str) or "-" not in texto:
        return None
    partes = texto.split("-")
    if len(partes) != 2:
        return None
    i, f = partes[0].strip(), partes[1].strip()
    if len(i) == 5 and len(f) == 5 and ":" in i and ":" in f:
        try:
            datetime.strptime(i, "%H:%M")
            datetime.strptime(f, "%H:%M")
            return (i, f)
        except ValueError:
            return None
    return None


def parse_lista_mergulhos_texto(texto_celula):
    if not texto_celula or not isinstance(texto_celula, str):
        return []
    blocos = texto_celula.replace(";", ",").split(",")
    resultado = []
    for b in blocos:
        h = extrair_horario(b.strip())
        if h:
            resultado.append(h)
    return resultado


def parse_horario_turno(horastr: str, datastr: str, shift_noite: bool) -> datetime:
    dt = datetime.strptime(f"{datastr} {horastr}", "%Y-%m-%d %H:%M")
    if shift_noite and dt.hour < 12:
        dt += timedelta(days=1)
    return dt


def aplicar_limite_turno(
    dt_i: datetime, dt_f: datetime, data_base_str: str, shift_noite: bool
):
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
    if not shift_noite:
        limite_inicio = base_dt.replace(hour=6, minute=0)
        limite_fim = base_dt.replace(hour=18, minute=0)
    else:
        limite_inicio = base_dt.replace(hour=18, minute=0)
        limite_fim = base_dt.replace(hour=6, minute=0) + timedelta(days=1)

    eff_start = max(dt_i, limite_inicio)
    eff_end = min(dt_f, limite_fim)

    if eff_start < eff_end:
        return eff_start, eff_end
    return None, None


def eh_piloto_extra(nome_p, idx_original):
    return idx_original >= 4 or "extra" in nome_p.lower()


def calcular_metas_pilotos(
    pilotos_lista,
    total_minutos_disponiveis,
    custom_pcts,
    pct_extra_dia=None,
    lista_extras_turno=None,
):
    n_p = len(pilotos_lista)
    if n_p == 0 or total_minutos_disponiveis <= 0:
        return {p: 0 for p in pilotos_lista}

    if pct_extra_dia is not None and lista_extras_turno:
        extras_no_bloco = [p for p in pilotos_lista if p in lista_extras_turno]
        titulares_no_bloco = [p for p in pilotos_lista if p not in lista_extras_turno]

        if extras_no_bloco and titulares_no_bloco:
            pct_efetivo = max(0.0, min(100.0, float(pct_extra_dia)))
            min_total_extras = int(round(total_minutos_disponiveis * (pct_efetivo / 100.0)))
            min_total_titulares = total_minutos_disponiveis - min_total_extras

            metas = {}
            n_ex = len(extras_no_bloco)
            base_ex = min_total_extras // n_ex
            sobra_ex = min_total_extras % n_ex
            for idx, p in enumerate(extras_no_bloco):
                metas[p] = base_ex + (1 if idx < sobra_ex else 0)

            n_tit = len(titulares_no_bloco)
            base_tit = min_total_titulares // n_tit
            sobra_tit = min_total_titulares % n_tit
            fatias_tit = [base_tit + (1 if idx < sobra_tit else 0) for idx in range(n_tit)]
            if n_tit >= 2 and base_tit > 20:
                for i in range(0, n_tit - 1, 2):
                    jitter = random.randint(-2, 2)
                    fatias_tit[i] += jitter
                    fatias_tit[i + 1] -= jitter
            for idx, p in enumerate(titulares_no_bloco):
                metas[p] = fatias_tit[idx]

            return metas

    if custom_pcts and any(custom_pcts.get(p, 0) > 0 for p in pilotos_lista):
        weights = [custom_pcts.get(p, 1.0) for p in pilotos_lista]
        total_w = sum(weights) if sum(weights) > 0 else 1.0
        fatias = []
        restante = total_minutos_disponiveis
        for idx in range(n_p - 1):
            bloco = int(round(total_minutos_disponiveis * (weights[idx] / total_w)))
            bloco = max(0, min(restante, bloco))
            fatias.append(bloco)
            restante -= bloco
        fatias.append(restante)
        return {pilotos_lista[i]: fatias[i] for i in range(n_p)}

    base_p = total_minutos_disponiveis // n_p
    sobra = total_minutos_disponiveis % n_p
    fatias = [base_p + (1 if i < sobra else 0) for i in range(n_p)]
    if n_p >= 2 and base_p > 20:
        for i in range(0, n_p - 1, 2):
            jitter = random.randint(-3, 3)
            fatias[i] += jitter
            fatias[i + 1] -= jitter
    return {pilotos_lista[i]: fatias[i] for i in range(n_p)}


def gerar_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Logbook")
    return output.getvalue()


def calcular_tempo_total_rov(mergulhos_str, data_base_str, shift_noite, shift_start):
    total_minutos = 0
    primeiros_inicios = []
    for i, f in mergulhos_str:
        dt_i = parse_horario_turno(i, data_base_str, shift_noite)
        dt_f = parse_horario_turno(f, data_base_str, shift_noite)
        if dt_f <= dt_i:
            dt_f += timedelta(days=1)

        s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, shift_noite)
        if s and e:
            duracao = int((e - s).total_seconds() / 60)
            total_minutos += duracao
            primeiros_inicios.append(s)

    inicio_base = min(primeiros_inicios) if primeiros_inicios else shift_start
    return total_minutos, inicio_base


def processar_um_dia(
    data_base_str,
    mergulhos_68_inputs,
    mergulhos_69_inputs,
    campo_emb,
    nome_op,
    shift_noite,
    modo_unico,
    s_68,
    s_69,
    lista_pilotos,
    modo_alloc,
    alloc_man,
    pct_sup,
    pct_pilotos_cust,
    pct_extra_dia=None,
):
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
    shift_start = base_dt.replace(hour=18 if shift_noite else 6, minute=0)
    supervisores = {"UHD68": s_68, "UHD69": s_69}

    titulares = [
        p for idx, p in enumerate(lista_pilotos) if not eh_piloto_extra(p, idx)
    ]
    extras = [
        p for idx, p in enumerate(lista_pilotos) if eh_piloto_extra(p, idx)
    ]

    pilotos_ativos_dia = lista_pilotos.copy()
    if pct_extra_dia is not None and float(pct_extra_dia) <= 0.0 and extras:
        pilotos_ativos_dia = [p for p in lista_pilotos if p not in extras]

    registros = []
    horas_dia = {p: 0.0 for p in [s_68, s_69] + lista_pilotos}

    if modo_unico:
        minutos_68, inicio_68 = calcular_tempo_total_rov(
            mergulhos_68_inputs, data_base_str, shift_noite, shift_start
        )
        minutos_69, inicio_69 = calcular_tempo_total_rov(
            mergulhos_69_inputs, data_base_str, shift_noite, shift_start
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
                b_fim_sup = dt_start + timedelta(minutes=sup_min)
                hrs_sup = round(sup_min / 60.0, 2)
                horas_dia[sup_fixo] += hrs_sup
                registros.append({
                    "Campo": campo_emb,
                    "Operação": nome_op,
                    "Data Lançamento": data_base_str,
                    "Data Real (Log)": dt_start.strftime("%Y-%m-%d"),
                    "Turno": "Noite" if shift_noite else "Dia",
                    "Modo": "Lançamento Único",
                    "Colaborador": sup_fixo,
                    "Cargo": "Supervisor",
                    "ROV Alocado": rov,
                    "Início Lançamento": dt_start.strftime("%H:%M"),
                    "Fim Lançamento": b_fim_sup.strftime("%H:%M"),
                    "Duração (min)": sup_min,
                    "Total Horas (h)": hrs_sup,
                })

            rov_pools[rov] = {
                "p_min_restante": p_min,
                "cursor": dt_start + timedelta(minutes=sup_min),
            }

        rovs_ativos = sorted(rov_pools.keys(), key=lambda r: rov_pools[r]["cursor"])
        metas_pilotos = calcular_metas_pilotos(
            pilotos_ativos_dia,
            total_pool_pilotos,
            pct_pilotos_cust,
            pct_extra_dia=pct_extra_dia,
            lista_extras_turno=extras,
        )

        rovs_permitidos = {}
        if modo_alloc.startswith("Manual"):
            for p in pilotos_ativos_dia:
                escolha = alloc_man.get(p, "Ambos (UHD68 e UHD69)")
                if escolha == "UHD68":
                    rovs_permitidos[p] = ["UHD68"]
                elif escolha == "UHD69":
                    rovs_permitidos[p] = ["UHD69"]
                else:
                    rovs_permitidos[p] = rovs_ativos.copy()
        else:
            titulares_emb = [p for p in titulares if p in pilotos_ativos_dia]
            extras_emb = [p for p in extras if p in pilotos_ativos_dia]
            random.shuffle(titulares_emb)
            random.shuffle(extras_emb)

            if len(rovs_ativos) == 2:
                meio = (len(titulares_emb) + 1) // 2
                for p in titulares_emb[:meio]:
                    rovs_permitidos[p] = [rovs_ativos[0]]
                for p in titulares_emb[meio:]:
                    rovs_permitidos[p] = [rovs_ativos[1]]
                for p in extras_emb:
                    rovs_permitidos[p] = rovs_ativos.copy()
            else:
                for p in pilotos_ativos_dia:
                    rovs_permitidos[p] = rovs_ativos.copy()

        ordem_pilotos = [p for p in pilotos_ativos_dia if len(rovs_permitidos[p]) == 1]
        random.shuffle(ordem_pilotos)
        pilotos_ponte = [p for p in pilotos_ativos_dia if len(rovs_permitidos[p]) > 1]
        random.shuffle(pilotos_ponte)
        fila_pilotos = ordem_pilotos + pilotos_ponte

        agenda_pilotos = {p: [] for p in pilotos_ativos_dia}

        def tem_conflito_horario(pessoa, s, e):
            for os, oe in agenda_pilotos[pessoa]:
                if max(s, os) < min(e, oe):
                    return True
            return False

        for p in fila_pilotos:
            necessario = metas_pilotos.get(p, 0)
            if necessario <= 0:
                continue

            lista_rovs_tentativa = rovs_permitidos[p] + [
                r for r in rovs_ativos if r not in rovs_permitidos[p]
            ]

            for rov in lista_rovs_tentativa:
                if necessario <= 0:
                    break
                if rov not in rov_pools:
                    continue
                pool = rov_pools[rov]
                disponivel = pool["p_min_restante"]
                if disponivel <= 0:
                    continue

                alocado = min(necessario, disponivel)
                b_inicio = pool["cursor"]
                b_fim = b_inicio + timedelta(minutes=alocado)

                if tem_conflito_horario(p, b_inicio, b_fim):
                    ultimo_fim = max(e for _, e in agenda_pilotos[p])
                    if ultimo_fim > b_inicio:
                        b_inicio = ultimo_fim
                        b_fim = b_inicio + timedelta(minutes=alocado)

                pool["cursor"] = b_fim
                pool["p_min_restante"] -= alocado
                necessario -= al

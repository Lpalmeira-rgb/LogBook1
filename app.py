from datetime import datetime, timedelta
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)

st.title("⚓ Sistema de Lançamento de Horários ROV")
st.markdown(
    "Escolha entre **Lançamento Único (Consolidado)** ou **Horários Exatos"
    " (Múltiplos Lançamentos)** com trava de turno, equipe dinâmica de"
    " pilotos, cota de supervisor e porcentagens personalizáveis."
)

# --- SIDEBAR: CONFIGURAÇÕES DO LANÇAMENTO ---
st.sidebar.header("⚙️ Configurações do Lançamento")

campo_embarcacao = st.sidebar.text_input(
    "Campo de Atuação da Embarcação", "Campo de Búzios"
)

data_logbook = st.sidebar.date_input(
    "Data de Lançamento (Início do Turno)", datetime.now()
)
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
        "Escolha '1' para consolidar todo o tempo do turno em um único"
        " lançamento por pessoa, ou 'Múltiplos' para dividir cada mergulho"
        " individualmente."
    ),
)

is_turno_noite = "Noite" in turno_selecionado
is_modo_unico = "1" in modo_lancamento

# --- CADASTRO DA EQUIPE ---
st.sidebar.subheader("👨‍✈️ Supervisores (Fixos)")
sup_68 = st.sidebar.text_input("Supervisor UHD68", "Sup. João Silva")
sup_69 = st.sidebar.text_input("Supervisor UHD69", "Sup. Marcos Rocha")

st.sidebar.subheader("👨‍💻 Pilotos (Equipe Dinâmica)")
qtd_pilotos = st.sidebar.number_input(
    "Quantidade de Pilotos no Turno", min_value=1, max_value=12, value=4, step=1
)

nomes_padrao = [
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

pilotos = []
for i in range(int(qtd_pilotos)):
    p_default = nomes_padrao[i] if i < len(nomes_padrao) else f"Piloto {i+1}"
    p_nome = st.sidebar.text_input(f"Piloto {i+1}", p_default, key=f"piloto_input_{i}")
    if p_nome.strip():
        pilotos.append(p_nome.strip())

# --- AJUSTE DE PORCENTAGENS DE DIVISÃO ---
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

pct_pilotos_custom = {}
if tipo_divisao_pilotos == "Personalizada por Piloto (%)":
    st.sidebar.markdown("**Defina a porcentagem/peso de cada piloto:**")
    val_padrao = round(100.0 / len(pilotos), 1) if pilotos else 25.0
    for p in pilotos:
        pct_pilotos_custom[p] = st.sidebar.number_input(
            f"% para {p}", min_value=0.0, max_value=100.0, value=val_padrao, step=1.0, key=f"pct_{p}"
        )

# --- ENTRADA DE MERGULHOS DO DIA ---
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
    st.subheader("🔴 Mergulhos ROV UHD68")
    m_68_1 = st.text_input("Mergulho 1", padrao_68[0], key="d68_1")
    m_68_2 = st.text_input("Mergulho 2", padrao_68[1], key="d68_2")
    m_68_3 = st.text_input("Mergulho 3", padrao_68[2], key="d68_3")
    m_68_4 = st.text_input("Mergulho 4", padrao_68[3], key="d68_4")
    m_68_5 = st.text_input("Mergulho 5", padrao_68[4], key="d68_5")

with col2:
    st.subheader("🔵 Mergulhos ROV UHD69")
    m_69_1 = st.text_input("Mergulho 1", padrao_69[0], key="d69_1")
    m_69_2 = st.text_input("Mergulho 2", padrao_69[1], key="d69_2")
    m_69_3 = st.text_input("Mergulho 3", padrao_69[2], key="d69_3")
    m_69_4 = st.text_input("Mergulho 4", padrao_69[3], key="d69_4")
    m_69_5 = st.text_input("Mergulho 5", padrao_69[4] if len(padrao_69) > 4 else "", key="d69_5")


def extrair_horario(texto):
    if "-" in texto:
        partes = texto.split("-")
        i, f = partes[0].strip(), partes[1].strip()
        if len(i) == 5 and len(f) == 5:
            return (i, f)
    return None


def parse_horario_turno(
    horastr: str, datastr: str, shift_noite: bool
) -> datetime:
    dt = datetime.strptime(f"{datastr} {horastr}", "%Y-%m-%d %H:%M")
    if shift_noite and dt.hour < 12:
        dt += timedelta(days=1)
    return dt


def aplicar_limite_turno(
    dt_i: datetime, dt_f: datetime, data_base_str: str, shift_noite: bool
):
    """Aplica a trava do turno (06:00 - 18:00 ou 18:00 - 06:00)."""
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


def calcular_fatias_pilotos(pilotos_alocados, minutos_totais, custom_pcts):
    n_p = len(pilotos_alocados)
    if n_p == 0:
        return []

    if custom_pcts and any(custom_pcts.get(p, 0) > 0 for p in pilotos_alocados):
        weights = [custom_pcts.get(p, 1.0) for p in pilotos_alocados]
        total_w = sum(weights) if sum(weights) > 0 else 1.0
        fatias = []
        restante = minutos_totais
        for idx in range(n_p - 1):
            bloco = int(round(minutos_totais * (weights[idx] / total_w)))
            bloco = max(10, min(restante - 10 * (n_p - 1 - idx), bloco))
            fatias.append(bloco)
            restante -= bloco
        fatias.append(restante)
        return fatias
    else:
        base_p = minutos_totais // n_p
        fatias = []
        restante = minutos_totais
        for idx in range(n_p - 1):
            v = random.randint(-4, 4) if n_p > 1 else 0
            bloco = max(15, min(restante - 15 * (n_p - 1 - idx), base_p + v))
            fatias.append(bloco)
            restante -= bloco
        fatias.append(restante)
        return fatias


# --- PROCESSAMENTO DO LOGBOOK ---
if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
    data_base_str = data_logbook.strftime("%Y-%m-%d")
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
    shift_start = base_dt.replace(hour=18 if is_turno_noite else 6, minute=0)

    supervisores = {"UHD68": sup_68, "UHD69": sup_69}

    pilotos_embaralhados = pilotos.copy()
    random.shuffle(pilotos_embaralhados)

    mergulhos_68_inputs = [
        extrair_horario(m)
        for m in [m_68_1, m_68_2, m_68_3, m_68_4, m_68_5]
        if extrair_horario(m)
    ]
    mergulhos_69_inputs = [
        extrair_horario(m)
        for m in [m_69_1, m_69_2, m_69_3, m_69_4, m_69_5]
        if extrair_horario(m)
    ]

    registros = []
    horas_do_dia = {p: 0.0 for p in [sup_68, sup_69] + pilotos}

    # =========================================================================
    # MODO 1: LANÇAMENTO ÚNICO CONSOLIDADO POR OPERADOR
    # =========================================================================
    if is_modo_unico:

        def calcular_tempo_total_rov(mergulhos_str):
            total_minutos = 0
            primeiros_inicios = []
            for i, f in mergulhos_str:
                dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
                dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
                if dt_f <= dt_i:
                    dt_f += timedelta(days=1)

                s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, is_turno_noite)
                if s and e:
                    duracao = int((e - s).total_seconds() / 60)
                    total_minutos += duracao
                    primeiros_inicios.append(s)

            inicio_base = min(primeiros_inicios) if primeiros_inicios else shift_start
            return total_minutos, inicio_base

        minutos_68, inicio_68 = calcular_tempo_total_rov(mergulhos_68_inputs)
        minutos_69, inicio_69 = calcular_tempo_total_rov(mergulhos_69_inputs)

        total_minutos_rov = {"UHD68": minutos_68, "UHD69": minutos_69}
        inicio_base_rov = {"UHD68": inicio_68, "UHD69": inicio_69}

        active_rovs = [r for r, m in total_minutos_rov.items() if m > 0]

        if len(active_rovs) == 2:
            meio = len(pilotos_embaralhados) // 2
            alocacao_rov = [
                ("UHD68", total_minutos_rov["UHD68"], sup_68, pilotos_embaralhados[:meio]),
                ("UHD69", total_minutos_rov["UHD69"], sup_69, pilotos_embaralhados[meio:]),
            ]
        elif len(active_rovs) == 1:
            r_ativo = active_rovs[0]
            alocacao_rov = [
                (r_ativo, total_minutos_rov[r_ativo], supervisores[r_ativo], pilotos_embaralhados)
            ]
        else:
            alocacao_rov = []

        for rov, tot_min, sup_fixo, pilotos_rov in alocacao_rov:
            if tot_min <= 0:
                continue

            sup_minutos = max(0, int(round(tot_min * (pct_supervisor / 100.0))))
            pilotos_minutos = tot_min - sup_minutos

            fatias_pilotos = calcular_fatias_pilotos(pilotos_rov, pilotos_minutos, pct_pilotos_custom)

            slots = []
            if sup_minutos > 0:
                slots.append((sup_fixo, sup_minutos, "Supervisor"))

            for idx, p in enumerate(pilotos_rov):
                slots.append((p, fatias_pilotos[idx], "Piloto"))

            cursor_tempo = inicio_base_rov[rov]
            for colaborador, dur_m, cargo in slots:
                if dur_m <= 0:
                    continue
                b_inicio = cursor_tempo
                b_fim = cursor_tempo + timedelta(minutes=dur_m)
                cursor_tempo = b_fim

                hrs = round(dur_m / 60.0, 2)
                horas_do_dia[colaborador] += hrs

                registros.append({
                    "Campo": campo_embarcacao,
                    "Data Lançamento": data_base_str,
                    "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
                    "Turno": "Noite" if is_turno_noite else "Dia",
                    "Modo": "Lançamento Único",
                    "Colaborador": colaborador,
                    "Cargo": cargo,
                    "ROV Alocado": rov,
                    "Início Lançamento": b_inicio.strftime("%H:%M"),
                    "Fim Lançamento": b_fim.strftime("%H:%M"),
                    "Duração (min)": dur_m,
                    "Total Horas (h)": hrs,
                })

    # =========================================================================
    # MODO 2: HORÁRIOS EXATOS DOS MERGULHOS (MÚLTIPLOS LANÇAMENTOS)
    # =========================================================================
    else:
        lista_mergulhos = []
        for i, f in mergulhos_68_inputs:
            dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
            dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
            if dt_f <= dt_i:
                dt_f += timedelta(days=1)
            s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, is_turno_noite)
            if s and e:
                lista_mergulhos.append(("UHD68", s, e))

        for i, f in mergulhos_69_inputs:
            dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
            dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
            if dt_f <= dt_i:
                dt_f += timedelta(days=1)
            s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, is_turno_noite)
            if s and e:
                lista_mergulhos.append(("UHD69", s, e))

        lista_mergulhos.sort(key=lambda m: m[1])

        agenda_ocupacao = {p: [] for p in horas_do_dia.keys()}

        def tem_conflito(pessoa, inicio, fim):
            for s, e in agenda_ocupacao[pessoa]:
                if max(inicio, s) < min(fim, e):
                    return True
            return False

        for rov, dt_inicio, dt_fim in lista_mergulhos:
            duracao_min = int((dt_fim - dt_inicio).total_seconds() / 60)
            sup_fixo = supervisores[rov]

            sup_minutos = max(0, int(round(duracao_min * (pct_supervisor / 100.0))))
            pilotos_minutos = duracao_min - sup_minutos

            pilotos_candidatos = pilotos_embaralhados.copy()
            random.shuffle(pilotos_candidatos)
            pilotos_ordenados = sorted(
                pilotos_candidatos, key=lambda p: (horas_do_dia[p], random.random())
            )

            qtd_pilotos_no_mergulho = min(len(pilotos), 2)
            pilotos_disponiveis = []
            for p in pilotos_ordenados:
                if len(pilotos_disponiveis) >= qtd_pilotos_no_mergulho:
                    break
                if not tem_conflito(p, dt_inicio, dt_fim):
                    pilotos_disponiveis.append(p)

            if not pilotos_disponiveis:
                pilotos_disponiveis = [pilotos_ordenados[0]]

            fatias_pilotos = calcular_fatias_pilotos(pilotos_disponiveis, pilotos_minutos, pct_pilotos_custom)

            slots = []
            sup_dt_s = dt_inicio
            sup_dt_e = dt_inicio + timedelta(minutes=sup_minutos)

            if sup_minutos > 0 and not tem_conflito(sup_fixo, sup_dt_s, sup_dt_e):
                slots.append((sup_fixo, sup_minutos, "Supervisor"))
            else:
                pilotos_minutos = duracao_min
                fatias_pilotos = calcular_fatias_pilotos(pilotos_disponiveis, pilotos_minutos, pct_pilotos_custom)

            for idx, p in enumerate(pilotos_disponiveis):
                slots.append((p, fatias_pilotos[idx], "Piloto"))

            cursor_tempo = dt_inicio
            for colaborador, dur_m, cargo in slots:
                if dur_m <= 0:
                    continue
                b_inicio = cursor_tempo
                b_fim = cursor_tempo + timedelta(minutes=dur_m)
                cursor_tempo = b_fim

                agenda_ocupacao[colaborador].append((b_inicio, b_fim))
                hrs = round(dur_m / 60.0, 2)
                horas_do_dia[colaborador] += hrs

                registros.append({
                    "Campo": campo_embarcacao,
                    "Data Lançamento": data_base_str,
                    "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
                    "Turno": "Noite" if is_turno_noite else "Dia",
                    "Modo": "Horários Exatos",
                    "Colaborador": colaborador,
                    "Cargo": cargo,
                    "ROV Alocado": rov,
                    "Início Lançamento": b_inicio.strftime("%H:%M"),
                    "Fim Lançamento": b_fim.strftime("%H:%M"),
                    "Duração (min)": dur_m,
                    "Total Horas (h)": hrs,
                })

    df_res = pd.DataFrame(registros)

    st.success("✅ Tabela do Logbook Gerada!")
    st.dataframe(df_res, use_container_width=True)

    st.subheader("📊 Resumo de Horas do Turno")
    df_saldos = pd.DataFrame(
        list(horas_do_dia.items()), columns=["Colaborador", "Total Horas no Dia (h)"]
    )
    st.dataframe(df_saldos, use_container_width=True)

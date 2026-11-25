from datetime import datetime, timedelta
import io
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)


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
                necessario -= alocado
                agenda_pilotos[p].append((b_inicio, b_fim))

                hrs = round(alocado / 60.0, 2)
                horas_dia[p] += hrs

                registros.append({
                    "Campo": campo_emb,
                    "Operação": nome_op,
                    "Data Lançamento": data_base_str,
                    "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
                    "Turno": "Noite" if shift_noite else "Dia",
                    "Modo": "Lançamento Único",
                    "Colaborador": p,
                    "Cargo": "Piloto",
                    "ROV Alocado": rov,
                    "Início Lançamento": b_inicio.strftime("%H:%M"),
                    "Fim Lançamento": b_fim.strftime("%H:%M"),
                    "Duração (min)": alocado,
                    "Total Horas (h)": hrs,
                })

    else:
        lista_mergulhos = []
        for i, f in mergulhos_68_inputs:
            dt_i = parse_horario_turno(i, data_base_str, shift_noite)
            dt_f = parse_horario_turno(f, data_base_str, shift_noite)
            if dt_f <= dt_i:
                dt_f += timedelta(days=1)
            s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, shift_noite)
            if s and e:
                lista_mergulhos.append(("UHD68", s, e))

        for i, f in mergulhos_69_inputs:
            dt_i = parse_horario_turno(i, data_base_str, shift_noite)
            dt_f = parse_horario_turno(f, data_base_str, shift_noite)
            if dt_f <= dt_i:
                dt_f += timedelta(days=1)
            s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, shift_noite)
            if s and e:
                lista_mergulhos.append(("UHD69", s, e))

        lista_mergulhos.sort(key=lambda m: m[1])

        agenda_ocupacao = {p: [] for p in horas_dia.keys()}
        rovs_operados_por_piloto = {p: set() for p in pilotos_ativos_dia}

        def tem_conflito(pessoa, inicio, fim):
            for s, e in agenda_ocupacao[pessoa]:
                if max(inicio, s) < min(fim, e):
                    return True
            return False

        preferencia_rov_m2 = {}
        if modo_alloc.startswith("Manual"):
            for p in pilotos_ativos_dia:
                esc = alloc_man.get(p, "Ambos (UHD68 e UHD69)")
                preferencia_rov_m2[p] = (
                    ["UHD68", "UHD69"] if "Ambos" in esc else [esc]
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
                preferencia_rov_m2[p] = ["UHD68", "UHD69"]

        for rov, dt_inicio, dt_fim in lista_mergulhos:
            duracao_min = int((dt_fim - dt_inicio).total_seconds() / 60)
            sup_fixo = supervisores[rov]

            sup_minutos = max(0, int(round(duracao_min * (pct_sup / 100.0))))
            pilotos_minutos = duracao_min - sup_minutos

            def prioridade_piloto(p):
                habilitado = 0 if rov in preferencia_rov_m2.get(p, []) else 1
                eh_ex = 0 if p in extras else 1
                cruzamento = 0
                if (
                    len(rovs_operados_por_piloto[p]) > 0
                    and rov not in rovs_operados_por_piloto[p]
                ):
                    cruzamento = 0 if p in extras else 1
                return (habilitado, cruzamento, horas_dia[p], eh_ex, random.random())

            pilotos_ordenados = sorted(pilotos_ativos_dia, key=prioridade_piloto)

            qtd_pilotos_no_mergulho = min(
                len(pilotos_ativos_dia), 3 if len(pilotos_ativos_dia) >= 5 else 2
            )
            pilotos_disponiveis = []
            for p in pilotos_ordenados:
                if len(pilotos_disponiveis) >= qtd_pilotos_no_mergulho:
                    break
                if not tem_conflito(p, dt_inicio, dt_fim):
                    pilotos_disponiveis.append(p)

            if not pilotos_disponiveis:
                pilotos_disponiveis = [pilotos_ordenados[0]]

            metas_bloco = calcular_metas_pilotos(
                pilotos_disponiveis,
                pilotos_minutos,
                pct_pilotos_cust,
                pct_extra_dia=pct_extra_dia,
                lista_extras_turno=extras,
            )
            fatias_pilotos = [metas_bloco[p] for p in pilotos_disponiveis]

            slots = []
            sup_dt_s = dt_inicio
            sup_dt_e = dt_inicio + timedelta(minutes=sup_minutos)

            if sup_minutos > 0 and not tem_conflito(sup_fixo, sup_dt_s, sup_dt_e):
                slots.append((sup_fixo, sup_minutos, "Supervisor"))
            else:
                pilotos_minutos = duracao_min
                metas_bloco = calcular_metas_pilotos(
                    pilotos_disponiveis,
                    pilotos_minutos,
                    pct_pilotos_cust,
                    pct_extra_dia=pct_extra_dia,
                    lista_extras_turno=extras,
                )
                fatias_pilotos = [metas_bloco[p] for p in pilotos_disponiveis]

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
                if cargo == "Piloto":
                    rovs_operados_por_piloto[colaborador].add(rov)

                hrs = round(dur_m / 60.0, 2)
                horas_dia[colaborador] += hrs

                registros.append({
                    "Campo": campo_emb,
                    "Operação": nome_op,
                    "Data Lançamento": data_base_str,
                    "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
                    "Turno": "Noite" if shift_noite else "Dia",
                    "Modo": "Horários Exatos",
                    "Colaborador": colaborador,
                    "Cargo": cargo,
                    "ROV Alocado": rov,
                    "Início Lançamento": b_inicio.strftime("%H:%M"),
                    "Fim Lançamento": b_fim.strftime("%H:%M"),
                    "Duração (min)": dur_m,
                    "Total Horas (h)": hrs,
                })

    return registros, horas_dia


# =============================================================================
# INTERFACE PRINCIPAL
# =============================================================================
st.title("⚓ Sistema de Lançamento de Horários ROV")
st.markdown(
    "Escolha entre **Lançamento Único (Consolidado)** ou **Horários Exatos"
    " (Múltiplos Lançamentos)**, com modo individual ou **Lançamento em Massa"
    " (Vários Dias)** permitindo variar a porcentagem diária do piloto extra."
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

# --- SIDEBAR ---
st.sidebar.header("⚙️ Configurações da Equipe e Regras")

campo_embarcacao = st.sidebar.text_input(
    "Campo de Atuação da Embarcação", "Campo de Búzios"
)

nome_operacao = st.sidebar.text_input(
    "Nome da Operação", "Inspeção Subsea / Intervenção"
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
        "Escolha '1' para consolidar o tempo do turno com divisão matemática"
        " exata entre todos os pilotos, ou 'Múltiplos' para dividir cada"
        " mergulho individualmente."
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
    label_p = f"Piloto {i+1} (Extra)" if i >= 4 else f"Piloto {i+1}"
    p_nome = st.sidebar.text_input(label_p, p_default, key=f"piloto_input_{i}")
    if p_nome.strip():
        pilotos.append(p_nome.strip())

st.sidebar.subheader("🚤 Alocação nos ROVs")
modo_alocacao_rov = st.sidebar.radio(
    "Forma de Alocação dos Pilotos",
    [
        "Aleatória (Priorizando Extra nos 2 ROVs)",
        "Manual (Escolher ROV por Piloto)",
    ],
)

alocacao_manual = {}
if modo_alocacao_rov.startswith("Manual"):
    st.sidebar.markdown("**Defina o ROV de cada piloto:**")
    for idx, p in enumerate(pilotos):
        if idx < 2:
            idx_padrao = 0
        elif idx < 4:
            idx_padrao = 1
        else:
            idx_padrao = 2
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
    help=(
        "Porcentagem do tempo total de mergulho atribuída ao supervisor fixo."
    ),
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
            f"% para {p}",
            min_value=0.0,
            max_value=100.0,
            value=val_padrao,
            step=1.0,
            key=f"pct_{p}",
        )

# =============================================================================
# TELA 1: LANÇAMENTO DIÁRIO (INDIVIDUAL)
# =============================================================================
if st.session_state["tela_ativa"] == "individual":
    st.subheader("📅 Lançamento Diário Individual")

    data_logbook = st.date_input(
        "Data de Lançamento (Início do Turno)",
        datetime.now().date(),
        key="data_indiv",
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
        m_68_1 = st.text_input("Mergulho 1", padrao_68[0], key="d68_1")
        m_68_2 = st.text_input("Mergulho 2", padrao_68[1], key="d68_2")
        m_68_3 = st.text_input("Mergulho 3", padrao_68[2], key="d68_3")
        m_68_4 = st.text_input("Mergulho 4", padrao_68[3], key="d68_4")
        m_68_5 = st.text_input("Mergulho 5", padrao_68[4], key="d68_5")

    with col2:
        st.markdown("#### 🔵 Mergulhos ROV UHD69")
        m_69_1 = st.text_input("Mergulho 1", padrao_69[0], key="d69_1")
        m_69_2 = st.text_input("Mergulho 2", padrao_69[1], key="d69_2")
        m_69_3 = st.text_input("Mergulho 3", padrao_69[2], key="d69_3")
        m_69_4 = st.text_input("Mergulho 4", padrao_69[3], key="d69_4")
        m_69_5 = st.text_input("Mergulho 5", padrao_69[4], key="d69_5")

    if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
        if not pilotos:
            st.error("⚠️ Adicione pelo menos 1 piloto na barra lateral antes de gerar.")
            st.stop()

        data_base_str = data_logbook.strftime("%Y-%m-%d")

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

        if not mergulhos_68_inputs and not mergulhos_69_inputs:
            st.warning(
                "⚠️ Nenhum horário de mergulho válido foi encontrado. Preencha no"
                " formato HH:MM - HH:MM (ex: 06:30 - 11:00)."
            )
            st.stop()

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
            df_completo = pd.DataFrame(registros)
            colunas_visiveis = [
                "Campo",
                "Operação",
                "Data Real (Log)",
                "Colaborador",
                "ROV Alocado",
                "Início Lançamento",
                "Fim Lançamento",
                "Total Horas (h)",
            ]
            st.session_state["df_exibicao"] = df_completo[colunas_visiveis]
            st.session_state["horas_do_dia"] = horas_do_dia
            st.session_state["data_base_str"] = data_base_str
        else:
            st.warning(
                "⚠️ Nenhum lançamento válido foi gerado dentro do horário do turno selecionado."
            )

    if "df_exibicao" in st.session_state:
        df_exibicao = st.session_state["df_exibicao"]
        horas_do_dia = st.session_state["horas_do_dia"]
        data_base_str = st.session_state["data_base_str"]

        st.success("✅ Tabela do Logbook Gerada!")
        st.dataframe(df_exibicao, use_container_width=True)

        excel_bytes = gerar_excel_bytes(df_exibicao)
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Logbook_ROV_{data_base_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.subheader("📊 Resumo de Horas do Turno (Verificação de Igualdade)")
        df_saldos = pd.DataFrame(
            list(horas_do_dia.items()),
            columns=["Colaborador", "Total Horas no Dia (h)"],
        )
        st.dataframe(df_saldos, use_container_width=True)

# =============================================================================
# TELA 2: LANÇAMENTO EM MASSA (VÁRIOS DIAS)
# =============================================================================
else:
    st.subheader("📚 Lançamento em Massa (Vários Dias Consecutivos)")
    st.markdown(
        "Configure o **período inicial e a quantidade de dias**. Na coluna"
        " **% Piloto Extra**, você pode variar a porcentagem diária destinada ao"
        " piloto extra (ex.: `20` para 20%, `10` para 10%, `0` se não atuar naquele dia)."
    )

    c_m1, c_m2, c_m3 = st.columns([2, 2, 3])
    with c_m1:
        data_inicio_massa = st.date_input(
            "Data Inicial do Período",
            datetime.now().date(),
            key="data_ini_massa",
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
            help=(
                "Quando ativado, usa a coluna '% Piloto Extra' de cada linha para"
                " calcular a cota do(s) piloto(s) extra(s) naquele dia específico."
            ),
        )

    padrao_str_68 = (
        "18:00 - 20:00, 02:00 - 06:00"
        if is_turno_noite
        else "06:30 - 11:00, 14:00 - 17:30"
    )
    padrao_str_69 = (
        "19:00 - 01:00" if is_turno_noite else "08:00 - 12:30, 15:00 - 18:00"
    )

    pct_extra_sugerido = round(100.0 / len(pilotos), 1) if pilotos else 20.0

    linhas_grade = []
    for d in range(int(qtd_dias_massa)):
        dt_dia = data_inicio_massa + timedelta(days=d)
        linhas_grade.append({
            "Data": dt_dia.strftime("%Y-%m-%d"),
            "Operação": nome_operacao,
            "Mergulhos UHD68": padrao_str_68,
            "Mergulhos UHD69": padrao_str_69,
            "% Piloto Extra": pct_extra_sugerido,
        })

    df_grade_input = pd.DataFrame(linhas_grade)

    st.markdown("#### ✏️ Grade Editável de Horários e Porcentagem Diária do Extra")
    df_editado = st.data_editor(
        df_grade_input,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Data": st.column_config.TextColumn(
                "Data (AAAA-MM-DD)", required=True
            ),
            "Operação": st.column_config.TextColumn("Operação"),
            "Mergulhos UHD68": st.column_config.TextColumn(
                "Mergulhos UHD68 (HH:MM - HH:MM)"
            ),
            "Mergulhos UHD69": st.column_config.TextColumn(
                "Mergulhos UHD69 (HH:MM - HH:MM)"
            ),
            "% Piloto Extra": st.column_config.NumberColumn(
                "% Piloto Extra",
                min_value=0.0,
                max_value=100.0,
                step=1.0,
                format="%.1f%%",
                help=(
                    "Porcentagem do tempo líquido de pilotos destinada ao(s)"
                    " piloto(s) extra(s) neste dia."
                ),
            ),
        },
        key="editor_massa",
    )

    if st.button("🚀 Gerar Todos os Lançamentos em Massa", type="primary"):
        if not pilotos:
            st.error("⚠️ Adicione pelo menos 1 piloto na barra lateral antes de gerar.")
            st.stop()

        todos_registros = []
        horas_acumuladas_periodo = {p: 0.0 for p in [sup_68, sup_69] + pilotos}
        dias_processados = 0

        for _, row in df_editado.iterrows():
            raw_data = row.get("Data")
            if pd.isna(raw_data) or not str(raw_data).strip():
                continue
            data_str = str(raw_data).strip()
            try:
                datetime.strptime(data_str, "%Y-%m-%d")
            except ValueError:
                st.warning(
                    f"⚠️ Data ignorada por formato inválido: '{data_str}' (use AAAA-MM-DD)."
                )
                continue

            raw_op = row.get("Operação")
            op_dia = (
                str(raw_op).strip()
                if pd.notna(raw_op) and str(raw_op).strip()
                else nome_operacao
            )

            raw_68 = row.get("Mergulhos UHD68")
            raw_69 = row.get("Mergulhos UHD69")
            m68_lista = parse_lista_mergulhos_texto(
                str(raw_68) if pd.notna(raw_68) else ""
            )
            m69_lista = parse_lista_mergulhos_texto(
                str(raw_69) if pd.notna(raw_69) else ""
            )

            if not m68_lista and not m69_lista:
                continue

            pct_ex_val = None
            if usar_pct_extra_diario:
                try:
                    raw_val = row.get("% Piloto Extra", pct_extra_sugerido)
                    pct_ex_val = float(raw_val) if pd.notna(raw_val) else pct_extra_sugerido
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
            df_massa_completo = pd.DataFrame(todos_registros)
            colunas_visiveis = [
                "Campo",
                "Operação",
                "Data Real (Log)",
                "Colaborador",
                "ROV Alocado",
                "Início Lançamento",
                "Fim Lançamento",
                "Total Horas (h)",
            ]
            st.session_state["df_massa_exibicao"] = df_massa_completo[colunas_visiveis]
            st.session_state["horas_massa"] = horas_acumuladas_periodo
            st.session_state["dias_proc_massa"] = dias_processados
        else:
            st.warning(
                "⚠️ Nenhum lançamento válido foi gerado. Verifique se os horários na"
                " tabela correspondem ao turno selecionado."
            )

    if "df_massa_exibicao" in st.session_state:
        df_massa_exibicao = st.session_state["df_massa_exibicao"]
        horas_massa = st.session_state["horas_massa"]
        dias_proc = st.session_state["dias_proc_massa"]

        st.success(
            f"✅ Lançamento em Massa Concluído! ({dias_proc} dias processados |"
            f" {len(df_massa_exibicao)} linhas geradas)"
        )

        # --- FILTRO POR PILOTO / COLABORADOR ---
        col_f1, col_f2 = st.columns([3, 1])
        with col_f1:
            colaboradores_unicos = sorted(
                df_massa_exibicao["Colaborador"].unique().tolist()
            )
            filtro_colaboradores = st.multiselect(
                "🔍 Filtrar Tabela por Piloto / Colaborador (deixe vazio para ver todos):",
                options=colaboradores_unicos,
                default=[],
                key="filtro_piloto_massa",
            )
        with col_f2:
            rovs_unicos = sorted(df_massa_exibicao["ROV Alocado"].unique().tolist())
            filtro_rov = st.multiselect(
                "🚤 Filtrar por ROV:",
                options=rovs_unicos,
                default=[],
                key="filtro_rov_massa",
            )

        df_massa_filtrada = df_massa_exibicao.copy()
        if filtro_colaboradores:
            df_massa_filtrada = df_massa_filtrada[
                df_massa_filtrada["Colaborador"].isin(filtro_colaboradores)
            ]
        if filtro_rov:
            df_massa_filtrada = df_massa_filtrada[
                df_massa_filtrada["ROV Alocado"].isin(filtro_rov)
            ]

        total_horas_filtro = round(df_massa_filtrada["Total Horas (h)"].sum(), 2)
        st.caption(
            f"Exibindo **{len(df_massa_filtrada)}** de"
            f" **{len(df_massa_exibicao)}** registros | Total de horas na seleção:"
            f" **{total_horas_filtro} h**"
        )

        st.dataframe(df_massa_filtrada, use_container_width=True)

        excel_massa_bytes = gerar_excel_bytes(df_massa_filtrada)
        data_ini_str = data_inicio_massa.strftime("%Y-%m-%d")
        sufixo_arquivo = (
            f"_{'_'.join([c.replace(' ', '_') for c in filtro_colaboradores[:2]])}"
            if filtro_colaboradores
            else ""
        )
        st.download_button(
            label="📥 Baixar Planilha Consolidada / Filtrada (.xlsx)",
            data=excel_massa_bytes,
            file_name=f"Logbook_ROV_Massa_{data_ini_str}_{dias_proc}dias{sufixo_arquivo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.subheader("📊 Resumo Acumulado de Horas no Período")
        df_saldos_massa = pd.DataFrame(
            list(horas_massa.items()),
            columns=["Colaborador", "Total Horas Acumuladas no Período (h)"],
        )
        if filtro_colaboradores:
            df_saldos_massa = df_saldos_massa[
                df_saldos_massa["Colaborador"].isin(filtro_colaboradores)
            ]
        st.dataframe(df_saldos_massa, use_container_width=True)

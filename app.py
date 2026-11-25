from datetime import datetime, timedelta
import io
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
    " pilotos, priorização de piloto extra para atuar nos dois ROVs e"
    " alocação aleatória ou manual por veículo."
)

# --- SIDEBAR: CONFIGURAÇÕES DO LANÇAMENTO ---
st.sidebar.header("⚙️ Configurações do Lançamento")

campo_embarcacao = st.sidebar.text_input(
    "Campo de Atuação da Embarcação", "Campo de Búzios"
)

nome_operacao = st.sidebar.text_input(
    "Nome da Operação", "Inspeção Subsea / Intervenção"
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
        "Escolha '1' para consolidar o tempo do turno com divisão matemática"
        " exata entre todos os pilotos, ou 'Múltiplos' para dividir cada"
        " mergulho individualmente."
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

# --- MODO DE ALOCAÇÃO NOS VEÍCULOS ---
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

# --- AJUSTE DE PORCENTAGENS DE DIVISÃO ---
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
  m_69_5 = st.text_input("Mergulho 5", padrao_69[4], key="d69_5")


def extrair_horario(texto):
  if not texto or "-" not in texto:
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


def calcular_metas_pilotos(
    pilotos_lista, total_minutos_disponiveis, custom_pcts
):
  """Calcula a cota exata de minutos de cada piloto no pool global do turno."""
  n_p = len(pilotos_lista)
  if n_p == 0 or total_minutos_disponiveis <= 0:
    return {p: 0 for p in pilotos_lista}

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
  else:
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


def eh_piloto_extra(nome_p, idx_original):
  return idx_original >= 4 or "extra" in nome_p.lower()


# --- PROCESSAMENTO DO LOGBOOK ---
if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
  if not pilotos:
    st.error(
        "⚠️ Adicione pelo menos 1 piloto na barra lateral antes de gerar."
    )
    st.stop()

  data_base_str = data_logbook.strftime("%Y-%m-%d")
  base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
  shift_start = base_dt.replace(hour=18 if is_turno_noite else 6, minute=0)

  supervisores = {"UHD68": sup_68, "UHD69": sup_69}

  titulares = [
      p for idx, p in enumerate(pilotos) if not eh_piloto_extra(p, idx)
  ]
  extras = [p for idx, p in enumerate(pilotos) if eh_piloto_extra(p, idx)]

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

  registros = []
  horas_do_dia = {p: 0.0 for p in [sup_68, sup_69] + pilotos}

  # =========================================================================
  # MODO 1: LANÇAMENTO ÚNICO CONSOLIDADO
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

    rov_configs = [
        ("UHD68", minutos_68, sup_68, inicio_68),
        ("UHD69", minutos_69, sup_69, inicio_69),
    ]

    rov_pools = {}
    total_pool_pilotos = 0

    for rov, tot_min, sup_fixo, dt_start in rov_configs:
      if tot_min <= 0:
        continue

      sup_min = max(0, int(round(tot_min * (pct_supervisor / 100.0))))
      p_min = tot_min - sup_min
      total_pool_pilotos += p_min

      if sup_min > 0:
        b_fim_sup = dt_start + timedelta(minutes=sup_min)
        hrs_sup = round(sup_min / 60.0, 2)
        horas_do_dia[sup_fixo] += hrs_sup
        registros.append({
            "Campo": campo_embarcacao,
            "Operação": nome_operacao,
            "Data Lançamento": data_base_str,
            "Data Real (Log)": dt_start.strftime("%Y-%m-%d"),
            "Turno": "Noite" if is_turno_noite else "Dia",
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
        pilotos, total_pool_pilotos, pct_pilotos_custom
    )

    rovs_permitidos = {}
    if modo_alocacao_rov.startswith("Manual"):
      for p in pilotos:
        escolha = alocacao_manual.get(p, "Ambos (UHD68 e UHD69)")
        if escolha == "UHD68":
          rovs_permitidos[p] = ["UHD68"]
        elif escolha == "UHD69":
          rovs_permitidos[p] = ["UHD69"]
        else:
          rovs_permitidos[p] = rovs_ativos.copy()
    else:
      titulares_emb = titulares.copy()
      extras_emb = extras.copy()
      random.shuffle(titulares_emb)
      random.shuffle(extras_emb)

      if len(rovs_ativos) == 2:
        meio = (len(titulares_emb) + 1) // 2
        grupo_r1 = titulares_emb[:meio]
        grupo_r2 = titulares_emb[meio:]
        for p in grupo_r1:
          rovs_permitidos[p] = [rovs_ativos[0]]
        for p in grupo_r2:
          rovs_permitidos[p] = [rovs_ativos[1]]
        for p in extras_emb:
          rovs_permitidos[p] = rovs_ativos.copy()
      else:
        for p in pilotos:
          rovs_permitidos[p] = rovs_ativos.copy()

    ordem_pilotos = [p for p in pilotos if len(rovs_permitidos[p]) == 1]
    random.shuffle(ordem_pilotos)
    pilotos_ponte = [p for p in pilotos if len(rovs_permitidos[p]) > 1]
    random.shuffle(pilotos_ponte)
    fila_pilotos = ordem_pilotos + pilotos_ponte

    agenda_pilotos = {p: [] for p in pilotos}

    def tem_conflito_horario(pessoa, s, e):
      for os, oe in agenda_pilotos[pessoa]:
        if max(s, os) < min(e, oe):
          return True
      return False

    for p in fila_pilotos:
      necessario = metas_pilotos[p]
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
        horas_do_dia[p] += hrs

        registros.append({
            "Campo": campo_embarcacao,
            "Operação": nome_operacao,
            "Data Lançamento": data_base_str,
            "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
            "Turno": "Noite" if is_turno_noite else "Dia",
            "Modo": "Lançamento Único",
            "Colaborador": p,
            "Cargo": "Piloto",
            "ROV Alocado": rov,
            "Início Lançamento": b_inicio.strftime("%H:%M"),
            "Fim Lançamento": b_fim.strftime("%H:%M"),
            "Duração (min)": alocado,
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
    rovs_operados_por_piloto = {p: set() for p in pilotos}

    def tem_conflito(pessoa, inicio, fim):
      for s, e in agenda_ocupacao[pessoa]:
        if max(inicio, s) < min(fim, e):
          return True
      return False

    preferencia_rov_m2 = {}
    if modo_alocacao_rov.startswith("Manual"):
      for p in pilotos:
        esc = alocacao_manual.get(p, "Ambos (UHD68 e UHD69)")
        preferencia_rov_m2[p] = (
            ["UHD68", "UHD69"] if "Ambos" in esc else [esc]
        )
    else:
      tit_emb = titulares.copy()
      random.shuffle(tit_emb)
      meio = (len(tit_emb) + 1) // 2
      for p in tit_emb[:meio]:
        preferencia_rov_m2[p] = ["UHD68"]
      for p in tit_emb[meio:]:
        preferencia_rov_m2[p] = ["UHD69"]
      for p in extras:
        preferencia_rov_m2[p] = ["UHD68", "UHD69"]

    for rov, dt_inicio, dt_fim in lista_mergulhos:
      duracao_min = int((dt_fim - dt_inicio).total_seconds() / 60)
      sup_fixo = supervisores[rov]

      sup_minutos = max(0, int(round(duracao_min * (pct_supervisor / 100.0))))
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
        return (habilitado, cruzamento, horas_do_dia[p], eh_ex, random.random())

      pilotos_ordenados = sorted(pilotos, key=prioridade_piloto)

      qtd_pilotos_no_mergulho = min(len(pilotos), 3 if len(pilotos) >= 5 else 2)
      pilotos_disponiveis = []
      for p in pilotos_ordenados:
        if len(pilotos_disponiveis) >= qtd_pilotos_no_mergulho:
          break
        if not tem_conflito(p, dt_inicio, dt_fim):
          pilotos_disponiveis.append(p)

      if not pilotos_disponiveis:
        pilotos_disponiveis = [pilotos_ordenados[0]]

      metas_bloco = calcular_metas_pilotos(
          pilotos_disponiveis, pilotos_minutos, pct_pilotos_custom
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
            pilotos_disponiveis, pilotos_minutos, pct_pilotos_custom
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
        horas_do_dia[colaborador] += hrs

        registros.append({
            "Campo": campo_embarcacao,
            "Operação": nome_operacao,
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

    df_exibicao = df_completo[colunas_visiveis]

    st.session_state["df_exibicao"] = df_exibicao
    st.session_state["horas_do_dia"] = horas_do_dia
    st.session_state["data_base_str"] = data_base_str
  else:
    st.warning(
        "⚠️ Nenhum lançamento válido foi gerado dentro do horário do turno"
        " selecionado."
    )

# --- EXIBIÇÃO DOS RESULTADOS E DOWNLOAD DA PLANILHA ---
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
      list(horas_do_dia.items()), columns=["Colaborador", "Total Horas no Dia (h)"]
  )
  st.dataframe(df_saldos, use_container_width=True)

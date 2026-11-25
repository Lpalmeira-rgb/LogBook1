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
    " pilotos, cota de supervisor e divisão rigorosamente igualitária entre"
    " todos os pilotos."
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
  p_nome = st.sidebar.text_input(
      f"Piloto {i+1}", p_default, key=f"piloto_input_{i}"
  )
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
  m_69_5 = st.text_input(
      "Mergulho 5", padrao_69[4] if len(padrao_69) > 4 else "", key="d69_5"
  )


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
      bloco = max(5, min(restante - 5 * (n_p - 1 - idx), bloco))
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
  # MODO 1: LANÇAMENTO ÚNICO CONSOLIDADO COM POOL IGUALITÁRIO GLOBAL
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

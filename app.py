from datetime import datetime, timedelta
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)

st.title("⚓ Sistema de Lançamento Único de Horários ROV")
st.markdown(
    "Consolidação de mergulhos por turno com **1 único registro por"
    " colaborador**, cota de 10% para o supervisor e alocação randomizada dos"
    " pilotos."
)

# --- SIDEBAR: CONFIGURAÇÕES DO LANÇAMENTO ---
st.sidebar.header("⚙️ Configurações do Lançamento")

data_logbook = st.sidebar.date_input(
    "Data de Lançamento (Início do Turno)", datetime.now()
)
turno_selecionado = st.sidebar.selectbox(
    "Selecione o Turno", ["Dia (06:00 - 18:00)", "Noite (18:00 - 06:00)"]
)

is_turno_noite = "Noite" in turno_selecionado

st.sidebar.subheader("Supervisores (Fixos - Cota 10%)")
sup_68 = st.sidebar.text_input("Supervisor UHD68", "Sup. João Silva")
sup_69 = st.sidebar.text_input("Supervisor UHD69", "Sup. Marcos Rocha")

st.sidebar.subheader("Pilotos (Flutuantes - Cota 90%)")
p1 = st.sidebar.text_input("Piloto 1", "Piloto Lucas Alves")
p2 = st.sidebar.text_input("Piloto 2", "Piloto Gabriel Lima")
p3 = st.sidebar.text_input("Piloto 3", "Piloto Mateus Costa")
p4 = st.sidebar.text_input("Piloto 4", "Piloto André Santos")

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


# --- PROCESSAMENTO DO LANÇAMENTO ÚNICO ---
if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
  data_base_str = data_logbook.strftime("%Y-%m-%d")
  base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")

  if not is_turno_noite:
    shift_start = base_dt.replace(hour=6, minute=0)
  else:
    shift_start = base_dt.replace(hour=18, minute=0)

  supervisores = {"UHD68": sup_68, "UHD69": sup_69}

  # Randomização dos pilotos a cada clique do botão
  pilotos = [p1, p2, p3, p4]
  random.shuffle(pilotos)

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

  # 1. Função para calcular o tempo total e horário inicial de cada ROV
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

  # Alocação dos pilotos randomizados
  if len(active_rovs) == 2:
    alocacao_rov = [
        ("UHD68", total_minutos_rov["UHD68"], sup_68, pilotos[:2]),
        ("UHD69", total_minutos_rov["UHD69"], sup_69, pilotos[2:]),
    ]
  elif len(active_rovs) == 1:
    r_ativo = active_rovs[0]
    alocacao_rov = [
        (r_ativo, total_minutos_rov[r_ativo], supervisores[r_ativo], pilotos)
    ]
  else:
    alocacao_rov = []

  registros = []
  horas_do_dia = {p: 0.0 for p in [sup_68, sup_69] + pilotos}

  for rov, tot_min, sup_fixo, pilotos_rov in alocacao_rov:
    if tot_min <= 0:
      continue

    # Cota de 10% do Supervisor
    sup_minutos = max(15, int(round(tot_min * 0.10)))
    pilotos_minutos = tot_min - sup_minutos

    # Divisão dos 90% restantes entre os pilotos alocados ao veículo
    n_p = len(pilotos_rov)
    base_p = pilotos_minutos // n_p
    fatias_pilotos = []
    restante = pilotos_minutos
    for idx in range(n_p - 1):
      v = random.randint(-4, 4)
      bloco = max(20, min(restante - 20 * (n_p - 1 - idx), base_p + v))
      fatias_pilotos.append(bloco)
      restante -= bloco
    fatias_pilotos.append(restante)

    # Blocos sequenciais de lançamento único
    slots = [(sup_fixo, sup_minutos, "Supervisor")]
    for idx, p in enumerate(pilotos_rov):
      slots.append((p, fatias_pilotos[idx], "Piloto"))

    cursor_tempo = inicio_base_rov[rov]
    for

from datetime import datetime, timedelta
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)

st.title("⚓ Sistema de Divisão Diária de Horários ROV")
st.markdown(
    "Divisão equilibrada de pilotagem para o dia de lançamento. "
    "Supervisores com 10% do tempo do veículo e equilíbrio entre os 4"
    " pilotos."
)

# --- SIDEBAR: CADASTRO DE EQUIPE E SELEÇÃO DE TURNO ---
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
    ["18:30 - 22:30", "01:00 - 05:00", "", "", ""]
    if is_turno_noite
    else ["06:30 - 11:00", "14:00 - 17:30", "", "", ""]
)
padrao_69 = (
    ["19:00 - 03:30", "", "", "", ""]
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
    dt += timedelta(days=1)  # Ajusta madrugadas para a data do dia seguinte
  return dt


# --- CALCULO E RATEIO EQUILIBRADO DO DIA ---
if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
  supervisores = {"UHD68": sup_68, "UHD69": sup_69}
  pilotos = [p1, p2, p3, p4]

  # Dicionário de rastreamento EXCLUSIVO para as horas do próprio dia
  horas_do_dia = {p: 0.0 for p in [sup_68, sup_69] + pilotos}

  mergulhos_68 = [
      extrair_horario(m)
      for m in [m_68_1, m_68_2, m_68_3, m_68_4, m_68_5]
      if extrair_horario(m)
  ]
  mergulhos_69 = [
      extrair_horario(m)
      for m in [m_69_1, m_69_2, m_69_3, m_69_4, m_69_5]
      if extrair_horario(m)
  ]

  data_base_str = data_logbook.strftime("%Y-%m-%d")

  lista_mergulhos = []
  for i, f in mergulhos_68:
    dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
    dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
    if dt_f <= dt_i:
      dt_f += timedelta(days=1)
    lista_mergulhos.append(("UHD68", dt_i, dt_f))

  for i, f in mergulhos_69:
    dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
    dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
    if dt_f <= dt_i:
      dt_f += timedelta(days=1)
    lista_mergulhos.append(("UHD69", dt_i, dt_f))

  lista_mergulhos.sort(key=lambda m: m[1])  # Ordena cronologicamente

  agenda_ocupacao = {p: [] for p in horas_do_dia.keys()}
  registros = []

  def tem_conflito(pessoa, inicio, fim):
    for s, e in agenda_ocupacao[pessoa]:
      if max(inicio, s) < min(fim, e):
        return True
    return False

  for rov, dt_inicio, dt_fim in lista_mergulhos:
    duracao_min = int((dt_fim - dt_inicio).total_seconds() / 60)
    sup_fixo = supervisores[rov]

    # REGRA: Supervisor tem exatamente 10% do tempo de pilotagem do seu ROV
    sup_minutos = max(15, int(round(duracao_min * 0.10)))
    pilotos_minutos = duracao_min - sup_minutos

    # Seleciona os pilotos que possuem MENOS horas acumuladas NO PRÓPRIO DIA
    pilotos_ordenados = sorted(pilotos, key=lambda p: horas_do_dia[p])

    qtd_pilotos = 2
    pilotos_disponiveis = []
    for p in pilotos_ordenados:
      if len(pilotos_disponiveis) >= qtd_pilotos:
        break
      if not tem_conflito(p, dt_inicio, dt_fim):
        pilotos_disponiveis.append(p)

    if not pilotos_disponiveis:
      pilotos_disponiveis = [pilotos_ordenados[0]]

    # Divisão dos 90% do tempo entre os pilotos do bloco
    n_p = len(pilotos_disponiveis)
    base_p = pilotos_minutos // n_p
    fatias_pilotos = []
    restante = pilotos_minutos
    for idx in range(n_p - 1):
      v = random.randint(-4, 4)
      bloco = max(20, min(restante - 20 * (n_p - 1 - idx), base_p + v))
      fatias_pilotos.append(bloco)
      restante -= bloco
    fatias_pilotos.append(restante)

    # Montagem dos horários contínuos
    slots = []
    sup_dt_s = dt_inicio
    sup_dt_e = dt_inicio + timedelta(minutes=sup_minutos)

    if not tem_conflito(sup_fixo, sup_dt_s, sup_dt_e):
      slots.append((sup_fixo, sup_minutos, "Supervisor"))
    else:
      pilotos_minutos = duracao_min
      fatias_pilotos = [duracao_min // n_p] * n_p
      fatias_pilotos[-1] += duracao_min % n_p

    for idx, p in enumerate(pilotos_disponiveis):
      slots.append((p, fatias_pilotos[idx], "Piloto"))

    cursor_tempo = dt_inicio
    for colaborador, dur_m, cargo in slots:
      b_inicio = cursor_tempo
      b_fim = cursor_tempo + timedelta(minutes=dur_m)
      cursor_tempo = b_fim

      agenda_ocupacao[colaborador].append((b_inicio, b_fim))
      hrs = round(dur_m / 60.0, 2)
      horas_do_dia[colaborador] += hrs

      registros.append({
          "Data Lançamento": data_base_str,
          "Data Real (Log)": b_inicio.strftime("%Y-%m-%d"),
          "Turno": "Noite" if is_turno_noite else "Dia",
          "Colaborador": colaborador,
          "Cargo": cargo,
          "ROV Alocado": rov,
          "Início": b_inicio.strftime("%H:%M"),
          "Fim": b_fim.strftime("%H:%M"),
          "Duração (min)": dur_m,
          "Total Horas (h)": hrs,
      })

  df_res = pd.DataFrame(registros)

  st.success("✅ Tabela do Logbook Gerada!")
  st.dataframe(df_res, use_container_width=True)

  # Resumo das horas totais trabalhadas estritamente no dia
  st.subheader("📊 Total de Horas Pilotadas no Dia por Colaborador")
  df_saldos = pd.DataFrame(
      list(horas_do_dia.items()), columns=["Colaborador", "Total Horas no Dia (h)"]
  )
  st.dataframe(df_saldos, use_container_width=True)

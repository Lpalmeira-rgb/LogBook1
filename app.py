from datetime import datetime, timedelta
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)

st.title("⚓ Sistema de Divisão e Lançamento de Horários ROV")
st.markdown(
    "Edite as informações da equipe e os horários de mergulho para gerar o"
    " rateio diário."
)

# --- SIDEBAR: CADASTRO E EQUIPE ---
st.sidebar.header("📋 Cadastro da Equipe & Saldo")

data_logbook = st.sidebar.date_input("Data do Lançamento", datetime.now())

st.sidebar.subheader("Supervisores (Fixos)")
sup_68 = st.sidebar.text_input("Supervisor UHD68", "Sup. João Silva")
saldo_sup_68 = st.sidebar.number_input(
    f"Saldo Horas - {sup_68}", value=40.0, step=0.5
)

sup_69 = st.sidebar.text_input("Supervisor UHD69", "Sup. Marcos Rocha")
saldo_sup_69 = st.sidebar.number_input(
    f"Saldo Horas - {sup_69}", value=38.0, step=0.5
)

st.sidebar.subheader("Pilotos (Flutuantes)")
p1 = st.sidebar.text_input("Piloto 1", "Piloto Lucas Alves")
saldo_p1 = st.sidebar.number_input(f"Saldo Horas - {p1}", value=10.0, step=0.5)

p2 = st.sidebar.text_input("Piloto 2", "Piloto Gabriel Lima")
saldo_p2 = st.sidebar.number_input(f"Saldo Horas - {p2}", value=26.0, step=0.5)

p3 = st.sidebar.text_input("Piloto 3", "Piloto Mateus Costa")
saldo_p3 = st.sidebar.number_input(f"Saldo Horas - {p3}", value=18.0, step=0.5)

p4 = st.sidebar.text_input("Piloto 4", "Piloto André Santos")
saldo_p4 = st.sidebar.number_input(f"Saldo Horas - {p4}", value=11.0, step=0.5)

# --- CORPO PRINCIPAL: ENTRADA DE MERGULHOS ---
col1, col2 = st.columns(2)

with col1:
  st.subheader("🔴 Mergulhos ROV UHD68")
  m_68_1 = st.text_input("Mergulho 1 (Início - Fim)", "06:30 - 11:00", key="d68_1")
  m_68_2 = st.text_input("Mergulho 2 (Início - Fim)", "14:00 - 17:30", key="d68_2")
  m_68_3 = st.text_input("Mergulho 3 (Início - Fim)", "20:00 - 23:00", key="d68_3")
  m_68_4 = st.text_input("Mergulho 4 (Início - Fim)", "", key="d68_4")
  m_68_5 = st.text_input("Mergulho 5 (Início - Fim)", "", key="d68_5")

with col2:
  st.subheader("🔵 Mergulhos ROV UHD69")
  m_69_1 = st.text_input("Mergulho 1 (Início - Fim)", "08:00 - 12:30", key="d69_1")
  m_69_2 = st.text_input("Mergulho 2 (Início - Fim)", "15:00 - 18:00", key="d69_2")
  m_69_3 = st.text_input("Mergulho 3 (Início - Fim)", "", key="d69_3")
  m_69_4 = st.text_input("Mergulho 4 (Início - Fim)", "", key="d69_4")
  m_69_5 = st.text_input("Mergulho 5 (Início - Fim)", "", key="d69_5")


def extrair_horarios(texto):
  if "-" in texto:
    partes = texto.split("-")
    i, f = partes[0].strip(), partes[1].strip()
    if len(i) == 5 and len(f) == 5:
      return (i, f)
  return None


# --- PROCESSAMENTO DO ALGORITMO ---
if st.button("🚀 Calcular e Gerar Divisão de Horários", type="primary"):
  supervisores = {"UHD68": sup_68, "UHD69": sup_69}
  pilotos = [p1, p2, p3, p4]
  horas_acumuladas = {
      sup_68: saldo_sup_68,
      sup_69: saldo_sup_69,
      p1: saldo_p1,
      p2: saldo_p2,
      p3: saldo_p3,
      p4: saldo_p4,
  }

  mergulhos_68 = [
      extrair_horarios(m)
      for m in [m_68_1, m_68_2, m_68_3, m_68_4, m_68_5]
      if extrair_horarios(m)
  ]
  mergulhos_69 = [
      extrair_horarios(m)
      for m in [m_69_1, m_69_2, m_69_3, m_69_4, m_69_5]
      if extrair_horarios(m)
  ]

  # Cálculo
  data_str = data_logbook.strftime("%Y-%m-%d")
  lista_mergulhos = []
  for i, f in mergulhos_68:
    lista_mergulhos.append(("UHD68", i, f))
  for i, f in mergulhos_69:
    lista_mergulhos.append(("UHD69", i, f))

  lista_mergulhos.sort(
      key=lambda m: datetime.strptime(f"{data_str} {m[1]}", "%Y-%m-%d %H:%M")
  )

  agenda_ocupacao = {p: [] for p in horas_acumuladas.keys()}
  registros = []

  for rov, str_inicio, str_fim in lista_mergulhos:
    dt_inicio = datetime.strptime(f"{data_str} {str_inicio}", "%Y-%m-%d %H:%M")
    dt_fim = datetime.strptime(f"{data_str} {str_fim}", "%Y-%m-%d %H:%M")
    if dt_fim <= dt_inicio:
      dt_fim += timedelta(days=1)

    duracao_total_min = int((dt_fim - dt_inicio).total_seconds() / 60)
    sup_fixo = supervisores[rov]

    pilotos_ordenados = sorted(
        pilotos, key=lambda p: horas_acumuladas.get(p, 0)
    )
    diff_h = horas_acumuladas[pilotos_ordenados[-1]] - horas_acumuladas[
        pilotos_ordenados[0]
    ]

    qtd_ops = 4 if diff_h > 12.0 else (3 if diff_h > 6.0 else 2)

    equipe = []
    # Checar supervisor
    sup_livre = True
    for s, e in agenda_ocupacao[sup_fixo]:
      if max(dt_inicio, s) < min(dt_fim, e):
        sup_livre = False
        break
    if sup_livre:
      equipe.append(sup_fixo)

    for piloto in pilotos_ordenados:
      if len(equipe) >= qtd_ops:
        break
      p_livre = True
      for s, e in agenda_ocupacao[piloto]:
        if max(dt_inicio, s) < min(dt_fim, e):
          p_livre = False
          break
      if p_livre:
        equipe.append(piloto)

    if not equipe:
      equipe = [pilotos_ordenados[0]]

    n_membros = len(equipe)
    base_min = duracao_total_min // n_membros
    fatias = []
    restante = duracao_total_min
    for i in range(n_membros - 1):
      v = random.randint(-8, 8)
      bloco = max(30, min(restante - 30 * (n_membros - 1 - i), base_min + v))
      fatias.append(bloco)
      restante -= bloco
    fatias.append(restante)

    cursor = dt_inicio
    for idx, membro in enumerate(equipe):
      dur = fatias[idx]
      b_i, b_f = cursor, cursor + timedelta(minutes=dur)
      cursor = b_f

      agenda_ocupacao[membro].append((b_i, b_f))
      hrs = round(dur / 60.0, 2)
      horas_acumuladas[membro] += hrs

      ponto_medio = b_i + timedelta(minutes=dur / 2)
      turno = (
          "Turno 1 (06:00 - 18:00)"
          if 6 <= ponto_medio.hour < 18
          else "Turno 2 (18:00 - 06:00)"
      )

      registros.append({
          "Data": data_str,
          "Colaborador": membro,
          "Cargo": "Supervisor" if membro in supervisores.values() else "Piloto",
          "ROV Alocado": rov,
          "Início": b_i.strftime("%H:%M"),
          "Fim": b_f.strftime("%H:%M"),
          "Horas Pilotadas": hrs,
          "Turno": turno,
      })

  df_res = pd.DataFrame(registros)

  st.success("✅ Divisão de Horários Calculada com Sucesso!")
  st.dataframe(df_res, use_container_width=True)

  # Resumo do Novo Saldo
  st.subheader("📊 Saldo Atualizado de Horas")
  df_saldos = pd.DataFrame(
      list(horas_acumuladas.items()), columns=["Colaborador", "Novo Saldo (h)"]
  )
  st.dataframe(df_saldos, use_container_width=True)

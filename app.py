from datetime import datetime, timedelta
import io
import os
import random
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Sistema de Logbook ROV", page_icon="⚓", layout="wide"
)

ARQUIVO_HISTORICO = "historico_logbook.csv"


# --- FUNÇÕES DE GERENCIAMENTO DE HISTÓRICO ---
def carregar_historico() -> pd.DataFrame:
  if os.path.exists(ARQUIVO_HISTORICO):
    try:
      df = pd.read_csv(ARQUIVO_HISTORICO)
      # Garante coluna de ID único para exclusão individual
      if "ID_Registro" not in df.columns:
        df.insert(0, "ID_Registro", range(len(df)))
      return df
    except Exception:
      pass
  return pd.DataFrame(columns=[
      "ID_Registro",
      "Data Lançamento",
      "Data Real (Log)",
      "Turno",
      "Colaborador",
      "Cargo",
      "ROV Alocado",
      "Início",
      "Fim",
      "Duração (min)",
      "Total Horas (h)",
  ])


def salvar_no_historico(novos_registros: pd.DataFrame):
  df_existente = carregar_historico()

  # Define novos IDs sequenciais
  proximo_id = (
      df_existente["ID_Registro"].max() + 1 if not df_existente.empty else 0
  )
  novos_registros.insert(
      0,
      "ID_Registro",
      range(proximo_id, proximo_id + len(novos_registros)),
  )

  df_consolidado = pd.concat(
      [df_existente, novos_registros], ignore_index=True
  )
  df_consolidado.to_csv(ARQUIVO_HISTORICO, index=False)


def excluir_itens_especificos(ids_para_remover: list):
  df_existente = carregar_historico()
  df_atualizado = df_existente[
      ~df_existente["ID_Registro"].isin(ids_para_remover)
  ]
  df_atualizado.to_csv(ARQUIVO_HISTORICO, index=False)


def limpar_historico_completo():
  if os.path.exists(ARQUIVO_HISTORICO):
    os.remove(ARQUIVO_HISTORICO)


st.title("⚓ Sistema de Divisão e Histórico de Logbook ROV")

# --- ORGANIZAÇÃO EM ABAS ---
tab_novo, tab_historico = st.tabs(
    ["🚀 Novo Lançamento Diário", "📜 Histórico Acumulado & Relatórios"]
)

# ==============================================================================
# ABA 1: NOVO LANÇAMENTO DIÁRIO (LANÇAMENTO ÚNICO POR PILOTO)
# ==============================================================================
with tab_novo:
  st.markdown(
      "Geração de lançamento **único e contínuo** por colaborador no turno,"
      " proporcional ao tempo total de mergulho de cada ROV."
  )

  # --- SIDEBAR: CONFIGURAÇÕES DA JORNADA ---
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
      ["18:30 - 20:00", "02:00 - 06:00", "", "", ""]
      if is_turno_noite
      else ["06:30 - 11:00", "14:00 - 17:30", "", "", ""]
  )
  padrao_69 = (
      ["19:00 - 03:00", "", "", "", ""]
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
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")
    if not shift_noite:
      limite_inicio = base_dt.replace(hour=6

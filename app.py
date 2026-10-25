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
      # Garante coluna de ID único para exclusao individual
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
  novos_registros.insert(0, "ID_Registro", range(proximo_id, proximo_id + len(novos_registros)))

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

  # --- PROCESSAMENTO DO BOTÃO DE LANÇAMENTO UNIFICADO ---
  if st.button("🚀 Gerar Lançamentos do Logbook", type="primary"):
    supervisores = {"UHD68": sup_68, "UHD69": sup_69}

    pilotos = [p1, p2, p3, p4]
    random.shuffle(pilotos)  # Embaralha os pilotos a cada clique

    horas_do_dia = {p: 0.0 for p in [sup_68, sup_69] + pilotos}
    agenda_ocupacao = {p: [] for p in horas_do_dia.keys()}

    data_base_str = data_logbook.strftime("%Y-%m-%d")
    base_dt = datetime.strptime(data_base_str, "%Y-%m-%d")

    if not is_turno_noite:
      shift_start = base_dt.replace(hour=6, minute=0)
      shift_end = base_dt.replace(hour=18, minute=0)
    else:
      shift_start = base_dt.replace(hour=18, minute=0)
      shift_end = base_dt.replace(hour=6, minute=0) + timedelta(days=1)

    # Função para agregar o tempo total de mergulho de cada ROV no turno
    def consolidar_mergulhos_rov(raw_dives):
      duracao_total_min = 0
      primeiro_inicio = None
      for i, f in raw_dives:
        if i and f:
          dt_i = parse_horario_turno(i, data_base_str, is_turno_noite)
          dt_f = parse_horario_turno(f, data_base_str, is_turno_noite)
          if dt_f <= dt_i:
            dt_f += timedelta(days=1)
          s, e = aplicar_limite_turno(dt_i, dt_f, data_base_str, is_turno_noite)
          if s and e:
            dur = int((e - s).total_seconds() / 60)
            duracao_total_min += dur
            if primeiro_inicio is None or s < primeiro_inicio:
              primeiro_inicio = s
      return duracao_total_min, primeiro_inicio

    m_68_list = [
        extrair_horario(m)
        for m in [m_68_1, m_68_2, m_68_3, m_68_4, m_68_5]
        if extrair_horario(m)
    ]
    m_69_list = [
        extrair_horario(m)
        for m in [m_69_1, m_69_2, m_69_3, m_69_4, m_69_5]
        if extrair_horario(m)
    ]

    tot_68, start_68 = consolidar_mergulhos_rov(m_68_list)
    tot_69, start_69 = consolidar_mergulhos_rov(m_69_list)

    rov_info = []
    if tot_68 > 0:
      rov_info.append(("UHD68", tot_68, start_68, supervisores["UHD68"]))
    if tot_69 > 0:
      rov_info.append(("UHD69", tot_69, start_69, supervisores["UHD69"]))

    # Ordena para processar primeiro o ROV que iniciou o mergulho antes
    rov_info.sort(key=lambda x: x[2])

    def tem_conflito(pessoa, inicio, fim):
      for s, e in agenda_ocupacao[pessoa]:
        if max(inicio, s) < min(fim, e):
          return True
      return False

    pilotos_pool = pilotos.copy()
    registros = []

    for rov, tot_min, primeiro_inicio, sup_fixo in rov_info:
      sup_minutos = max(15, int(round(tot_min * 0.10))) if tot_min > 0 else 0
      pilotos_minutos = tot_min - sup_minutos

      num_pilotos = 2 if len(rov_info) == 2 else min(4, len(pilotos_pool))
      pilotos_selecionados = pilotos_pool[:num_pilotos]

      # Remove pilotos alocados do pool para garantir a divisão entre os veiculos
      for p in pilotos_selecionados:
        if p in pilotos_pool and len(rov_info) == 2:
          pilotos_pool.remove(p)

      # Divisão contínua dos 90% do tempo entre os pilotos do veículo
      n_p = len(pilotos_selecionados)
      base_p = pilotos_minutos // n_p
      fatias_pilotos = []
      restante = pilotos_minutos
      for idx in range(n_p - 1):
        v = random.randint(-4, 4)
        bloco = max(15, min(restante - 15 * (n_p - 1 - idx), base_p + v))
        fatias_pilotos.append(bloco)
        restante -= bloco
      fatias_pilotos.append(restante)

      # Define o ponto de partida do bloco contínuo garantindo limite do turno
      bloco_inicio = min(
          primeiro_inicio, shift_end - timedelta(minutes=tot_min)
      )
      bloco_inicio = max(bloco_inicio, shift_start)

      slots = [(sup_fixo, sup_minutos, "Supervisor")]
      for idx, p in enumerate(pilotos_selecionados):
        slots.append((p, fatias_pilotos[idx], "Piloto"))

      cursor_tempo = bloco_inicio
      for colaborador, dur_m, cargo in slots:
        b_inicio = cursor_tempo
        b_fim = cursor_tempo + timedelta(minutes=dur_m)

        # Ajuste de segurança caso haja sobreposição de agenda
        while tem_conflito(colaborador, b_inicio, b_fim) and b_fim <= shift_end:
          b_inicio += timedelta(minutes=15)
          b_fim += timedelta(minutes=15)

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

    salvar_no_historico(df_res)

    st.success(
        "✅ Lançamentos Únicos Gerados e Salvos no Histórico com Sucesso!"
    )
    st.dataframe(df_res, use_container_width=True)

    st.subheader("📊 Total de Horas Pilotadas no Turno por Colaborador")
    df_saldos = pd.DataFrame(
        list(horas_do_dia.items()),
        columns=["Colaborador", "Total Horas no Dia (h)"],
    )
    st.dataframe(df_saldos, use_container_width=True)


# ==============================================================================
# ABA 2: HISTÓRICO ACUMULADO & GERENCIAMENTO
# ==============================================================================
with tab_historico:
  st.subheader("📜 Histórico Geral de Lançamentos Salvos")

  df_hist = carregar_historico()

  if df_hist.empty or len(df_hist) == 0:
    st.info(
        "Nenhum histórico registrado até o momento. Faça o primeiro lançamento"
        " na aba 'Novo Lançamento Diário'."
    )
  else:
    # --- FILTROS DE CONSULTA ---
    col_f1, col_f2, col_f3 = st.columns(3)

    with col_f1:
      colaboradores_unicos = ["Todos"] + sorted(
          df_hist["Colaborador"].dropna().unique().tolist()
      )
      filtro_colab = st.selectbox("Filtrar por Colaborador", colaboradores_unicos)

    with col_f2:
      rovs_unicos = ["Todos"] + sorted(
          df_hist["ROV Alocado"].dropna().unique().tolist()
      )
      filtro_rov = st.selectbox("Filtrar por ROV", rovs_unicos)

    with col_f3:
      turnos_unicos = ["Todos"] + sorted(
          df_hist["Turno"].dropna().unique().tolist()
      )
      filtro_turno = st.selectbox("Filtrar por Turno", turnos_unicos)

    df_filtrado = df_hist.copy()
    if filtro_colab != "Todos":
      df_filtrado = df_filtrado[df_filtrado["Colaborador"] == filtro_colab]
    if filtro_rov != "Todos":
      df_filtrado = df_filtrado[df_filtrado["ROV Alocado"] == filtro_rov]
    if filtro_turno != "Todos":
      df_filtrado = df_filtrado[df_filtrado["Turno"] == filtro_turno]

    st.dataframe(
        df_filtrado.drop(columns=["ID_Registro"], errors="ignore"),
        use_container_width=True,
    )

    # --- RESUMO TOTAL ACUMULADO ---
    st.subheader("📊 Total Acumulado de Horas de Pilotagem")
    df_totais = (
        df_hist.groupby(["Colaborador", "Cargo"])["Total Horas (h)"]
        .sum()
        .reset_index()
    )
    df_totais.columns = [
        "Colaborador",
        "Cargo",
        "Total Horas Acumuladas no Sistema (h)",
    ]
    st.dataframe(df_totais, use_container_width=True)

    # --- GERENCIAMENTO E EXCLUSÃO DO HISTÓRICO ---
    st.divider()
    st.subheader("🛠️ Opções de Gerenciamento e Exclusão")

    col_exc1, col_exc2 = st.columns(2)

    with col_exc1:
      st.markdown("### 🗑️ Exclusão de Registros Individuais")
      # Formata opções para fácil identificação do usuário
      opcoes_exclusao = {
          row[
              "ID_Registro"
          ]: f"[{row['Data Lançamento']} - {row['Turno']}] {row['Colaborador']} ({row['ROV Alocado']}: {row['Início']}-{row['Fim']} | {row['Total Horas (h)']}h)": row[
              "ID_Registro"
          ]
          for _, row in df_hist.iterrows()
      }

      itens_selecionados = st.multiselect(
          "Selecione um ou mais registros para remover:",
          options=list(opcoes_exclusao.keys()),
          format_func=lambda x: opcoes_exclusao[x],
      )

      if st.button("❌ Excluir Registros Selecionados", type="secondary"):
        if itens_selecionados:
          excluir_itens_especificos(itens_selecionados)
          st.success("Registros selecionados excluídos com sucesso!")
          st.rerun()
        else:
          st.warning("Selecione ao menos um registro na lista para excluir.")

    with col_exc2:
      st.markdown("### 💥 Exclusão Completa do Histórico")
      confirmar_limpeza = st.checkbox(
          "⚠️ Confirmo que desejo apagar TODO o histórico do sistema"
      )
      if st.button("🗑️ Apagar Todo o Histórico", disabled=not confirmar_limpeza):
        limpar_historico_completo()
        st.success("Todo o histórico foi apagado com sucesso!")
        st.rerun()

    # --- DOWNLOAD EM EXCEL ---
    st.divider()
    buffer_excel = io.BytesIO()
    with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
      df_hist.drop(columns=["ID_Registro"], errors="ignore").to_excel(
          writer, sheet_name="Historico Logbook", index=False
      )
      df_totais.to_excel(writer, sheet_name="Totais Acumulados", index=False)
    buffer_excel.seek(0)

    st.download_button(
        label="📥 Baixar Histórico Completo em Excel (.xlsx)",
        data=buffer_excel,
        file_name="Historico_Logbook_ROV.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

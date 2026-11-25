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
    " (Múltiplos Lançamentos)**, com modo individual ou **Lançamento em Massa"
    " (Vários Dias)** permitindo variar a porcentagem diária do piloto extra."
)

# --- CONTROLE DE NAVEGAÇÃO ENTRE ABAS / TELAS ---
if "tela_ativa" not in st.session_state:
  st.session_state["tela_ativa"] = "individual"

col_nav1, col_nav2, _ = st.columns([2, 2, 4])
with col_nav1:
  if st.button(
      "📅 Lançamento Diário (Individual)",
      type="primary"
      if st.session_state["tela_ativa"] == "individual"
      else "secondary",
      use_container_width=True,
  ):
    st.session_state["tela_ativa"] = "individual"
    st.rerun()

with col_nav2:
  if st.button(
      "📚 Lançamento em Massa (Vários Dias)",
      type="primary"
      if st.session_state["tela_ativa"] == "massa"
      else "secondary",
      use_container_width=True,
  ):
    st.session_state["tela_ativa"] = "massa"
    st.rerun()

st.divider()

# --- SIDEBAR: CONFIGURAÇÕES GERAIS (COMUNS AOS DOIS MODOS) ---
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


# =============================================================================
# FUNÇÕES DE CÁLCULO REUTILIZÁVEIS (DIA ÚNICO E MASSA)
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

  # Caso 1: Porcentagem diária específica para Piloto Extra (Lançamento em Massa)
  if pct_extra_dia is not None and lista_extras_turno:
    extras_no_bloco = [p for p in pilotos_lista if p in lista_extras_turno]
    titulares_no_bloco = [
        p for p in pilotos_lista if p not in lista_extras_turno
    ]

    if extras_no_bloco and titulares_no_bloco:
      pct_efetivo = max(0.0, min(100.0, float(pct_extra_dia)))
      min_total_extras = int(
          round(total_minutos_disponiveis * (pct_efetivo / 100.0))
      )
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
      fatias_tit = [
          base_tit + (1 if idx < sobra_tit else 0) for idx in range(n_tit)
      ]
      if n_tit >= 2 and base_tit > 20:
        for i in range(0, n_tit - 1, 2):
          jitter = random.randint(-2, 2)
          fatias_tit[i] += jitter
          fatias_tit[i + 1] -= jitter
      for idx, p in enumerate(titulares_no_bloco):
        metas[p] = fatias_tit[idx]

      return metas

  # Caso 2: Porcentagem customizada fixa por piloto (Sidebar)
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

  # Caso 3: Divisão igualitária padrão entre todos
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
  extras = [p for idx, p in enumerate(lista_pilotos) if eh_piloto_extra(p, idx)]

  pilotos_ativos_dia = lista_pilotos.copy()
  if pct_extra_dia is not None and float(pct_extra_dia) <= 0.0 and extras:
    pilotos_ativos_dia = [p for p in lista_pilotos if p not in extras]

  registros = []
  horas_dia = {p: 0.0 for p in [s_68, s_69] + lista_pilotos}

  if modo_unico:

    def calcular_tempo_total_rov(mergulhos_str):
      total_minutos = 0
      primeiros_inicios = []
      for i, f in mergulhos_str:
        dt_i = parse_horario_turno(i, data_base_str, shift_noite)
        dt_f = parse_horario_turno(f, data_base_str, shift_noite)
        if dt_f <= dt_i:

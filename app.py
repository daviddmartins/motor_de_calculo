from __future__ import annotations

import html
import unicodedata
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from io import BytesIO

import pandas as pd
import streamlit as st

from core.engines import calculate_contract_evolution
from core.exports import build_evolution_workbook
from core.indices import load_configuration
from core.models import (
    ContractSettings, ExtraordinaryAmortization, RoundingSettings,
    MODALITY_LABELS, SUPPORTED_EVOLUTION_MODALITIES, NO_CORRECTION_MODALITIES,
)
from core.monetary_update import (
    Abatement, Penalty, UpdateItem, UpdateRule, calculate_monetary_update,
    build_monetary_update_workbook, build_import_template,
    CORRECTION_NONE, CORRECTION_DAILY_BUSINESS, CORRECTION_DAILY_CALENDAR, CORRECTION_PRORATA_CALENDAR, CORRECTION_MONTHLY_CLOSE, CORRECTION_FACTOR_TABLE,
    INTEREST_NONE, INTEREST_SIMPLE_30, INTEREST_SIMPLE_COMPETENCE, INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC, INTEREST_COMPOUND_EQUIV, INTEREST_LEGAL_BCB,
    INTEREST_APPLICATION_DAILY, INTEREST_APPLICATION_END,
    INTEREST_BASE_PRINCIPAL, INTEREST_BASE_CORRECTED,
    ORDER_CORRECTION_INTEREST, ORDER_INTEREST_CORRECTION,
    PENALTY_BASE_PRINCIPAL, PENALTY_BASE_CORRECTED, PENALTY_BASE_BALANCE,
    parse_ui_date, excel_yearfrac,
)
from core.pdf_reports import ReportIdentity, build_opinion_pdf, pdf_filename
from core.pdf_funcef_template import DEFAULT_AUTHORS, DEFAULT_OPERATION, ManifestationHeader
from core.docx_export import build_opinion_docx, docx_filename
from core.monetary_pdf import build_monetary_update_pdf, build_update_copy_text, monetary_pdf_filename
from core.court_profiles import get_profiles, build_profile_rules
from core.majs import (
    MAJSSettings, MAJSPayment, MAJSExtraAmortization, MAJSSuspension, MAJSEarlySettlement, calculate_majs, build_majs_workbook,
)
from core.majs_pdf import MAJSReportIdentity, build_majs_opinion_pdf, majs_opinion_pdf_filename
from core.engine_versions import EVOLUTION_ENGINE_VERSION, MAJS_ENGINE_VERSION, MONETARY_UPDATE_ENGINE_VERSION
from services.configuration_service import load_runtime_configuration, factor_series_for_runtime
from services.index_service import discover_index_workbook, load_index_workbook
from services.access_service import build_access_context
from services.health_service import build_health_snapshot
from ui.screens.home import render_home
from ui.screens.indices import render_indices_center
from ui.screens.admin import render_holidays, render_configuration, render_access
from ui.components import (
    apply_design, sidebar_brand, brand_header as ui_brand_header, page_intro as ui_page_intro,
    step_header, workflow_strip, rule_timeline, result_hero, copy_block_title, callout,
)


APP_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = APP_DIR / "config" / "configuracoes_validacao.xlsx"
ACCESS_CONFIG = APP_DIR / "config" / "parametros_acesso_motor_calculos.xlsx"
INDEX_CONFIG = APP_DIR / "config" / "indices_motor_calculos.xlsx"
MIN_CREDIT_DATE = date(1994, 1, 1)
LOGO_PATH = APP_DIR / "assets" / "funcef_logo.png"
LOGO_WHITE_PATH = APP_DIR / "assets" / "logo_funcef_branca.png"
VERSION = "0.10.10"
WEEKDAYS_PT = [
    "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira",
    "Sexta-feira", "Sábado", "Domingo",
]
NAVIGATION = [
    "Início",
    "Evolução de contrato",
    "Recálculo de diferenças",
    "Atualização do saldo devedor",
    "Central de índices",
    "Feriados",
    "Configuração",
    "Acesso ao sistema",
]
NAV_LABELS = {
    "Início": "Página inicial",
    "Evolução de contrato": "Evolução de contrato e parecer",
    "Recálculo de diferenças": "Recálculo MAJS",
    "Atualização do saldo devedor": "Atualização do saldo devedor",
    "Central de índices": "Central de índices",
    "Feriados": "Feriados",
    "Configuração": "Configuração",
    "Acesso ao sistema": "Acesso ao sistema",
}


st.set_page_config(
    page_title="Motor de Cálculos · FUNCEF",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


apply_design()


def number_br(value, places: int = 2) -> str:
    number = Decimal(str(value))
    text = f"{number:,.{places}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def money_br(value, places: int = 2) -> str:
    number = Decimal(str(value))
    prefix = "-R$ " if number < 0 else "R$ "
    return prefix + number_br(abs(number), places)


def percent_br(rate, places: int = 4) -> str:
    return f"{number_br(Decimal(str(rate)) * Decimal('100'), places)}%"


def date_br(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else "-"


def parse_money_br(value) -> Decimal:
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = str(value or "").strip().replace("R$", "").replace(" ", "")
    if not text:
        return Decimal("0")
    negative = text.startswith("-")
    text = text.lstrip("-")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    result = Decimal(text)
    return -result if negative else result


@st.cache_data(show_spinner=False)
def load_default_config():
    return load_configuration(DEFAULT_CONFIG)


def brand_header():
    ui_brand_header(VERSION)


def page_intro(eyebrow: str, title: str, description: str):
    ui_page_intro(eyebrow, title, description)



def set_navigation(target: str, open_report: bool = False):
    """Callback de navegação. Executa antes da nova renderização do Streamlit."""
    st.session_state["navigation"] = target
    if open_report:
        st.session_state["show_report_builder"] = True


def _active_configuration_source():
    return st.session_state.get("active_config_bytes") or DEFAULT_CONFIG


def get_runtime_configuration():
    return load_runtime_configuration(
        app_dir=APP_DIR,
        configuration_source=_active_configuration_source(),
        preferred_index_path=INDEX_CONFIG,
        active_index_bytes=st.session_state.get("active_indices_bytes"),
    )


def get_active_config():
    return get_runtime_configuration().configuration


def get_index_repository():
    return get_runtime_configuration().indices


def export_active_configuration_bytes() -> bytes:
    """Exporta somente parâmetros gerais/feriados. Índices possuem base própria."""
    data = st.session_state.get("active_config_bytes")
    if data:
        return data
    return DEFAULT_CONFIG.read_bytes()


def _index_status_rows(cfg=None):
    repo = get_index_repository()
    return repo.status_rows() if repo is not None else []


def get_factor_series():
    runtime = get_runtime_configuration()
    return factor_series_for_runtime(app_dir=APP_DIR, repository=runtime.indices)


def _judicial_status_rows():
    factors = get_factor_series()
    profiles = get_profiles()
    rows, seen = [], set()
    for code, factor in sorted(factors.items()):
        profile = profiles.get(code)
        rows.append({
            "Código": code,
            "Tribunal": profile.tribunal if profile else code.split("_")[0],
            "Critério / tabela": profile.title if profile else getattr(factor, "name", code),
            "Cobertura inicial": getattr(factor, "coverage_start", "-"),
            "Última competência": getattr(factor, "coverage_end", "-"),
            "Fonte": (profile.source_label if profile else getattr(factor, "source", "")) or "Base única de índices",
            "Situação": profile.status if profile else "TABELA DISPONÍVEL",
        })
        seen.add(code)
    for code, profile in profiles.items():
        if code == "PERSONALIZADO" or code in seen:
            continue
        rows.append({
            "Código": code,
            "Tribunal": profile.tribunal,
            "Critério / tabela": profile.title,
            "Cobertura inicial": "-",
            "Última competência": "Critério composto",
            "Fonte": profile.source_label,
            "Situação": profile.status,
        })
    return rows


def _request_headers() -> dict:
    try:
        return dict(st.context.headers)
    except Exception:
        return {}


def _access_context():
    ctx = build_access_context(
        app_dir=APP_DIR,
        preferred_path=ACCESS_CONFIG,
        headers=_request_headers(),
    )
    return (
        ctx.registry, ctx.identity, ctx.entry, ctx.allowed, ctx.user_name,
        ctx.profile, ctx.is_admin, list(ctx.validator_options),
    )


def initialize_state():
    defaults = {
        "navigation": "Início",
        "show_report_builder": False,
        "evolution_inputs": {},
        "evolution_revision": 0,
        "report_pdf": None,
        "report_pdf_name": None,
        "report_docx": None,
        "report_docx_name": None,
        "report_adjustments": {},
        "report_selection_mode": "Duas primeiras",
        "report_identity": None,
        "report_selected_numbers": [],
        "report_validation_messages": [],
        "result_view": "Resumo por prestação",
        "update_result": None,
        "update_excel": None,
        "update_view": "Resumo por valor",
        "update_mode": "Saldo consolidado",
        "update_profile": "PERSONALIZADO",
        "update_rules_revision": 0,
        "update_pdf": None,
        "update_pdf_name": None,
        "update_profile_label": "Critério personalizado",
        "active_indices_bytes": None,
    }
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def saved_input(key: str, default):
    return st.session_state.get("evolution_inputs", {}).get(key, default)


def clear_calculation():
    for key in (
        "evolution_state", "report_pdf", "report_pdf_name", "report_docx", "report_docx_name", "report_adjustments",
        "show_report_builder", "report_values_editor", "report_selected_custom",
        "report_identity", "report_selected_numbers", "report_validation_messages",
        "update_result", "update_excel",
    ):
        st.session_state.pop(key, None)
    st.session_state["evolution_inputs"] = {}
    st.session_state["evolution_revision"] = st.session_state.get("evolution_revision", 0) + 1


def build_installments_dataframe(settings, result) -> pd.DataFrame:
    places = settings.rounding.money_places
    return pd.DataFrame([
        {
            "Nº": row.installment_number,
            "Referência": date_br(row.reference_due_date),
            "Vencimento operacional": date_br(row.operational_due_date),
            "Competência": row.competence,
            "Índice referência": row.index_reference or "-",
            "Saldo inicial": money_br(row.opening_balance, places),
            "Saldo antes da prestação": money_br(row.balance_before_installment, places),
            "Base da amortização": money_br(row.amortization_base, places),
            "Correção": money_br(row.correction_amount, places),
            "Juros": money_br(row.interest_amount, places),
            "Amortização": money_br(row.regular_amortization, places),
            "Amort. extraordinária": money_br(row.extraordinary_amortization, places),
            "Prestação": money_br(row.installment_amount, places),
            "Saldo final": money_br(row.closing_balance, places),
        }
        for row in result.installments
    ])


def build_daily_dataframe(settings, result) -> pd.DataFrame:
    places = settings.rounding.money_places
    pct_places = settings.rounding.percentage_display_places
    return pd.DataFrame([
        {
            "Prestação": row.installment_number,
            "Dia da semana": WEEKDAYS_PT[row.day.weekday()],
            "Feriado": "Sim" if row.is_holiday else "",
            "Data": date_br(row.day),
            "Juros (%)": percent_br(row.daily_interest_rate, pct_places),
            "Juros (R$)": money_br(row.interest_amount, places),
            "S.D. com juros": money_br(row.balance_after_interest, places),
            "C.M. (%)": percent_br(row.daily_correction_rate, pct_places),
            "C.M. (R$)": money_br(row.correction_amount, places),
            "Juros + C.M. (R$)": money_br(row.interest_amount + row.correction_amount, places),
            "Saldo devedor (J+CM)": money_br(row.balance_after_correction, places),
            "Amortização extra": money_br(row.extraordinary_amortization, places),
            "Juros acumulados": money_br(row.accumulated_interest, places),
            "Amortização": money_br(row.regular_amortization, places),
            "Prestação do dia": money_br(row.installment_amount, places),
            "Saldo finalizado": money_br(
                row.closing_balance_after_installment
                if row.closing_balance_after_installment is not None
                else row.balance_before_installment,
                places,
            ),
        }
        for row in result.daily_rows
    ])


initialize_state()
config = load_default_config()

ACCESS_REGISTRY, REQUEST_IDENTITY, CURRENT_ACCESS, ACCESS_ALLOWED, CURRENT_USER_NAME, CURRENT_PROFILE, CURRENT_IS_ADMIN, VALIDATOR_OPTIONS = _access_context()

USER_NAVIGATION = [
    "Início", "Evolução de contrato", "Recálculo de diferenças",
    "Atualização do saldo devedor",
]
ADMIN_NAVIGATION = USER_NAVIGATION + ["Central de índices", "Feriados", "Configuração", "Acesso ao sistema"]
ACTIVE_NAVIGATION = ADMIN_NAVIGATION if CURRENT_IS_ADMIN else USER_NAVIGATION

if st.session_state.get("navigation") not in ACTIVE_NAVIGATION:
    st.session_state["navigation"] = "Início"

sidebar_brand(VERSION)
module = st.sidebar.radio("Navegação", ACTIVE_NAVIGATION, key="navigation", label_visibility="collapsed", format_func=lambda x: NAV_LABELS.get(x, x))
st.sidebar.markdown("---")


def start_new_calculation():
    current = st.session_state.get("navigation", "Evolução de contrato")
    if current == "Atualização do saldo devedor":
        for key in (
            "update_result", "update_excel", "update_values_editor_seed", "update_values_editor",
            "update_rules_seed", "update_rules_editor", "update_abatements_seed", "update_abatements_editor",
            "update_values_upload", "update_consolidated_description", "update_pdf", "update_pdf_name",
        ):
            st.session_state.pop(key, None)
        st.session_state["navigation"] = "Atualização do saldo devedor"
    else:
        clear_calculation()
        st.session_state["navigation"] = "Evolução de contrato"


st.sidebar.button(
    "Novo cálculo", use_container_width=True,
    on_click=start_new_calculation,
)

brand_header()

if ACCESS_REGISTRY.enabled and not ACCESS_ALLOWED:
    if not REQUEST_IDENTITY.email:
        st.error("Não foi possível identificar o e-mail do usuário logado no Databricks App. O acesso foi bloqueado porque a planilha oficial de controle está ativa.")
    elif CURRENT_ACCESS is None:
        st.error(f"O usuário {REQUEST_IDENTITY.email} não está cadastrado na lista de acesso do Motor de Cálculos.")
    else:
        st.error(f"O acesso de {REQUEST_IDENTITY.email} está inativo na planilha de controle.")
    st.caption("Solicite ao administrador a atualização do arquivo config/parametros_acesso_motor_calculos.xlsx.")
    st.stop()


if module == "Início":
    runtime = get_runtime_configuration()
    active_cfg = runtime.configuration
    repo = runtime.indices
    profiles = get_profiles()
    judicial_profile_count = len([code for code in profiles if code != "PERSONALIZADO"])
    active_modalities = len(SUPPORTED_EVOLUTION_MODALITIES)
    index_count = len(repo.active_definitions) if repo is not None else len(active_cfg.indices)

    health = None
    if CURRENT_IS_ADMIN:
        health = build_health_snapshot(
            configuration_path=DEFAULT_CONFIG,
            configuration_loaded=True,
            repository=repo,
            index_error=runtime.index_error,
            access_registry=ACCESS_REGISTRY,
            holidays_count=len(active_cfg.holidays),
            engine_versions={
                "Evolução": EVOLUTION_ENGINE_VERSION,
                "MAJS": MAJS_ENGINE_VERSION,
                "Atualização": MONETARY_UPDATE_ENGINE_VERSION,
            },
        )

    render_home(
        set_navigation=set_navigation,
        modality_count=active_modalities,
        index_count=index_count,
        judicial_profile_count=judicial_profile_count,
        health=health,
        is_admin=CURRENT_IS_ADMIN,
    )

elif module == "Evolução de contrato":
    page_intro(
        "Módulo 01",
        "Evolução de contrato",
        "A data inicial é excluída e a data de referência final é incluída. O cálculo fecha no dia 20, mesmo quando o vencimento operacional ocorre no dia útil seguinte.",
    )

    with st.expander("Configuração e índices temporários", expanded=False):
        uploaded_config = st.file_uploader(
            "Importar configuração/índices nesta sessão (opcional)",
            type=["xlsx"],
            key="uploaded_config",
            help="O arquivo é usado apenas na sessão atual e não altera a configuração do deploy.",
        )
    if uploaded_config:
        st.session_state["active_config_bytes"] = uploaded_config.getvalue()
    active_config = get_active_config()

    implemented = list(SUPPORTED_EVOLUTION_MODALITIES)
    if not implemented:
        st.error("A configuração não contém modalidades implementadas nesta versão.")
        st.stop()

    # Reidratação dos controles após navegação entre módulos.
    widget_defaults = {
        "ev_modality": saved_input("modality_code", implemented[0]),
        "ev_initial_balance": float(saved_input("initial_balance", 10000.0)),
        "ev_credit_date": max(saved_input("credit_date", date(2023, 8, 22)), MIN_CREDIT_DATE),
        "ev_term": int(saved_input("term", 48)),
        "ev_annual_rate": float(saved_input("annual_rate_pct", 13.53)),
        "ev_competence_start": int(saved_input("competence_start", 21)),
        "ev_payment_timing": int(saved_input("payment_timing", 0)),
        "ev_lag": int(saved_input("lag", 2)),
        "ev_money_places": int(saved_input("money_places", 2)),
        "ev_interest_places": int(saved_input("interest_rate_places", 12)),
        "ev_correction_places": int(saved_input("correction_rate_places", 12)),
        "ev_percentage_places": int(saved_input("percentage_display_places", 4)),
        "ev_round_daily": bool(saved_input("round_daily_money", False)),
    }
    for key, value in widget_defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

    st.markdown('<div class="mini-heading">Dados e parâmetros do contrato</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    with c1:
        code = st.selectbox(
            "Modalidade", implemented, key="ev_modality",
            format_func=lambda item: active_config.modality_names.get(item, MODALITY_LABELS.get(item, item)),
        )
        initial_balance = st.number_input(
            "Saldo inicial (R$)", min_value=0.01, step=100.0, key="ev_initial_balance"
        )
        credit_date = st.date_input(
            "Data do crédito",
            min_value=MIN_CREDIT_DATE,
            max_value=date.today(),
            format="DD/MM/YYYY",
            key="ev_credit_date",
            help="São aceitas datas de crédito a partir de 01/01/1994.",
        )
    with c2:
        term = st.number_input("Prazo (prestações)", min_value=1, max_value=600, step=1, key="ev_term")
        annual_rate_pct = st.number_input(
            "Taxa de juros anual (%)", step=0.01, format="%.6f", key="ev_annual_rate"
        )
        lock_competence_21 = code in {"MOD_004", "MOD_005", "MOD_006"}
        if lock_competence_21:
            st.session_state["ev_competence_start"] = 21
        competence_start = st.selectbox(
            "Início da competência", [21, 1], key="ev_competence_start",
            format_func=lambda x: "Dia 21 (ciclo 21-20)" if x == 21 else "Dia 1 (mês civil)",
            disabled=lock_competence_21,
            help=(
                "As modalidades das abas 4, 5 e 6 utilizam a referência 21-20 da planilha metodológica."
                if lock_competence_21 else None
            ),
        )
    with c3:
        index_options = ["SEM_CORRECAO"] + sorted(active_config.indices.keys())
        is_no_correction = code in NO_CORRECTION_MODALITIES
        is_new_variable = code == "MOD_006"
        is_standard_fixed = code == "MOD_002"

        if is_no_correction:
            st.session_state["ev_index"] = "SEM_CORRECAO"
        elif is_new_variable:
            if "INPC" not in index_options:
                st.error("A modalidade Novo Credinâmico Variável requer a série INPC na configuração ativa.")
                st.stop()
            st.session_state["ev_index"] = "INPC"
        elif "ev_index" not in st.session_state:
            saved_index = saved_input("index_code", "INPC")
            st.session_state["ev_index"] = saved_index if saved_index in index_options else "INPC"

        index_code = st.selectbox(
            "Índice de correção",
            index_options,
            key="ev_index",
            disabled=is_no_correction or is_new_variable,
            help=(
                "Esta modalidade não utiliza correção monetária." if is_no_correction
                else "O Novo Credinâmico Variável utiliza INPC conforme a planilha metodológica." if is_new_variable
                else "Selecione o índice aplicável à evolução do saldo."
            ),
        )

        if is_new_variable:
            st.session_state["ev_lag"] = 2
        lag = st.number_input(
            "Defasagem do índice (meses)", min_value=0, max_value=24, step=1,
            disabled=is_no_correction or is_new_variable or index_code == "SEM_CORRECAO", key="ev_lag",
        )

        if is_no_correction:
            st.caption("Correção monetária não aplicável à modalidade selecionada.")
        elif is_new_variable:
            st.caption("INPC com defasagem de 2 meses, conforme a metodologia da Aba 6.")

        if is_standard_fixed:
            payment_timing = st.radio(
                "Tipo da fórmula Price (PGTO)", [0, 1], key="ev_payment_timing",
                format_func=lambda x: "0 - pagamento no fim do período" if x == 0 else "1 - pagamento no início do período",
                help="Corresponde ao argumento tipo da fórmula PGTO do Excel.",
            )
        elif code in {"MOD_004", "MOD_006"}:
            st.session_state["ev_payment_timing"] = 1
            payment_timing = st.radio(
                "Tipo da fórmula Price (PGTO)", [1], key="ev_payment_timing",
                format_func=lambda x: "1 - pagamento no início do período", disabled=True,
            )
        elif code == "MOD_005":
            st.session_state["ev_payment_timing"] = 0
            payment_timing = st.radio(
                "Tipo da fórmula Price (PGTO)", [0], key="ev_payment_timing",
                format_func=lambda x: "0 - pagamento no fim do período", disabled=True,
            )
        else:
            payment_timing = 0
            st.caption("A fórmula Price não se aplica à modalidade selecionada.")


    with st.expander("Precisão e arredondamentos", expanded=False):
        r1, r2, r3, r4, r5 = st.columns(5)
        with r1:
            interest_rate_places = st.number_input(
                "Precisão dos juros", min_value=4, max_value=18, step=1, key="ev_interest_places"
            )
        with r2:
            correction_rate_places = st.number_input(
                "Precisão da correção", min_value=4, max_value=18, step=1, key="ev_correction_places"
            )
        with r3:
            percentage_display_places = st.number_input(
                "Casas dos percentuais", min_value=0, max_value=10, step=1, key="ev_percentage_places"
            )
        with r4:
            money_places = st.number_input(
                "Casas dos valores", min_value=0, max_value=8, step=1, key="ev_money_places"
            )
        with r5:
            round_daily_money = st.checkbox(
                "Arredondar diariamente", key="ev_round_daily",
                help="Os juros e saldos diários passam a usar o mesmo valor monetário exibido na memória.",
            )

    modality_rule_notes = {
        "MOD_001": (
            "Regra diária",
            "juros do dia → correção monetária em dia útil → eventos financeiros. A amortização regular é calculada no fechamento de cada prestação.",
        ),
        "MOD_002": (
            "Credplan Fixo",
            "juros apropriados diariamente, sem correção monetária; a prestação fixa é formada pela Tabela Price e a amortização corresponde à prestação menos os juros do período.",
        ),
        "MOD_004": (
            "Novo Credinâmico Fixo",
            "juros apropriados diariamente, sem correção monetária; em cada vencimento a prestação é apurada pela Price tipo 1 sobre o saldo e o prazo remanescentes.",
        ),
        "MOD_005": (
            "Credinâmico Fixo",
            "sem correção monetária. O encargo de juros é incorporado ao saldo no aniversário anual da primeira prestação; a prestação mensal é recalculada pela Price tipo 0 e reduz integralmente o saldo já atualizado.",
        ),
        "MOD_006": (
            "Novo Credinâmico Variável",
            "juros do dia → correção monetária pelo INPC em dia útil → eventos financeiros → Price tipo 1 recalculada no fechamento de cada prestação.",
        ),
    }
    if code in modality_rule_notes:
        rule_title, rule_text = modality_rule_notes[code]
        st.markdown(
            f'<div class="success-note"><b>{rule_title}:</b> {rule_text}</div>',
            unsafe_allow_html=True,
        )

    st.markdown('<div class="mini-heading">Amortizações extraordinárias</div>', unsafe_allow_html=True)
    if "extra_editor" not in st.session_state:
        saved_extra = st.session_state.get("evolution_inputs", {}).get("extra_df")
        st.session_state["extra_editor"] = saved_extra if isinstance(saved_extra, pd.DataFrame) else pd.DataFrame({
            "Data": pd.Series(dtype="datetime64[ns]"),
            "Valor": pd.Series(dtype="float"),
            "Observação": pd.Series(dtype="str"),
        })
    extra_df = st.data_editor(
        st.session_state["extra_editor"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Valor": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.01),
        },
        key="extra_editor_widget",
    )

    b1, b2 = st.columns([3, 1])
    with b1:
        calculate = st.button("Calcular evolução", type="primary", use_container_width=True)
    with b2:
        if st.button("Limpar resultado", use_container_width=True):
            clear_calculation()
            st.rerun()

    if calculate:
        try:
            extras = []
            for _, row in extra_df.dropna(subset=["Data", "Valor"]).iterrows():
                extras.append(
                    ExtraordinaryAmortization(
                        pd.to_datetime(row["Data"]).date(),
                        Decimal(str(row["Valor"])),
                        str(row.get("Observação") or ""),
                    )
                )

            rounding = RoundingSettings(
                rate_places=int(interest_rate_places),
                factor_places=int(correction_rate_places),
                money_places=int(money_places),
                percentage_display_places=int(percentage_display_places),
                round_daily_money=bool(round_daily_money),
            )
            settings = ContractSettings(
                modality_code=code,
                initial_balance=Decimal(str(initial_balance)),
                credit_date=credit_date,
                annual_interest_rate=Decimal(str(annual_rate_pct)) / Decimal("100"),
                term=int(term),
                competence_start_day=int(competence_start),
                index_code=None if index_code == "SEM_CORRECAO" else index_code,
                index_lag_months=int(lag),
                payment_timing=int(payment_timing),
                rounding=rounding,
            )
            index_series = None if settings.index_code is None else active_config.indices.get(settings.index_code)
            modality_name = active_config.modality_names.get(code, MODALITY_LABELS.get(code, code))
            with st.spinner("Processando a evolução diária..."):
                result = calculate_contract_evolution(
                    settings,
                    index_series,
                    active_config.holidays,
                    extras,
                )

            st.session_state["evolution_inputs"] = {
                "modality_code": code,
                "initial_balance": initial_balance,
                "credit_date": credit_date,
                "term": int(term),
                "annual_rate_pct": annual_rate_pct,
                "competence_start": int(competence_start),
                "payment_timing": int(payment_timing),
                "index_code": index_code,
                "lag": int(lag),
                "money_places": int(money_places),
                "interest_rate_places": int(interest_rate_places),
                "correction_rate_places": int(correction_rate_places),
                "percentage_display_places": int(percentage_display_places),
                "round_daily_money": bool(round_daily_money),
                "extra_df": extra_df.copy(),
            }
            st.session_state["evolution_state"] = {
                "settings": settings,
                "result": result,
                "modality_name": modality_name,
            }
            st.session_state["extra_editor"] = extra_df.copy()
            st.session_state["evolution_revision"] += 1
            st.session_state["report_pdf"] = None
            st.session_state["report_docx"] = None
            st.session_state["report_adjustments"] = {}
            st.session_state["report_identity"] = None
            st.session_state["report_selected_numbers"] = []
            st.session_state["report_validation_messages"] = []
            st.session_state["result_view"] = "Resumo por prestação"
            st.success("Evolução processada e preservada nesta sessão.")
        except (ValueError, KeyError, InvalidOperation) as exc:
            st.error(str(exc))

    evolution_state = st.session_state.get("evolution_state")
    if evolution_state:
        settings = evolution_state["settings"]
        result = evolution_state["result"]
        modality_name = evolution_state["modality_name"]

        st.markdown('<div class="mini-heading">Resultado da evolução</div>', unsafe_allow_html=True)
        status_class = "status-complete" if result.is_complete else "status-partial"
        status_label = "Cálculo completo" if result.is_complete else "Cálculo parcial"
        st.markdown(f'<span class="{status_class}">{status_label}</span>', unsafe_allow_html=True)
        if not result.is_complete:
            st.markdown(
                f'<div class="technical-note"><b>Nota/crítica:</b> {html.escape(result.status_note)}</div>',
                unsafe_allow_html=True,
            )

        value_places = settings.rounding.money_places
        pct_places = settings.rounding.percentage_display_places
        last_installment = result.installments[-1] if result.installments else None
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Taxa mensal equivalente", percent_br(result.monthly_interest_rate, pct_places))
        m2.metric(
            "Primeira prestação",
            money_br(result.installments[0].installment_amount, value_places) if result.installments else "Não concluída",
        )
        m3.metric(
            "Último saldo concluído",
            money_br(last_installment.closing_balance, value_places) if last_installment else "Sem fechamento",
        )
        m4.metric("Prestações concluídas", f"{len(result.installments)} de {result.requested_installments or settings.term}")

        st.caption("Nas tabelas, use a barra nativa no canto superior direito para tela cheia, ocultar colunas, pesquisar e baixar os dados.")
        view_options = ["Resumo por prestação", "Memória diária", "Avisos e crítica"]
        if st.session_state.get("result_view") not in view_options:
            st.session_state["result_view"] = view_options[0]
        result_view = st.radio(
            "Visualização do resultado",
            view_options,
            horizontal=True,
            key="result_view",
            label_visibility="collapsed",
        )

        if result_view == "Resumo por prestação":
            if result.installments:
                if "installments_df" not in evolution_state:
                    evolution_state["installments_df"] = build_installments_dataframe(settings, result)
                st.dataframe(
                    evolution_state["installments_df"],
                    use_container_width=True,
                    hide_index=True,
                    height=420,
                )
            else:
                st.warning("Nenhuma prestação foi concluída antes da ausência do índice necessário.")
        elif result_view == "Memória diária":
            if "daily_df" not in evolution_state:
                with st.spinner("Preparando a visualização da memória diária..."):
                    evolution_state["daily_df"] = build_daily_dataframe(settings, result)
            st.dataframe(
                evolution_state["daily_df"],
                use_container_width=True,
                hide_index=True,
                height=520,
            )
        else:
            if result.warnings:
                for warning in result.warnings:
                    st.warning(warning)
            else:
                st.success("Nenhum aviso ou crítica foi gerado.")

        if "excel_bytes" not in evolution_state:
            with st.spinner("Preparando a memória de cálculo em Excel..."):
                evolution_state["excel_bytes"] = build_evolution_workbook(settings, modality_name, result)
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "Baixar memória de cálculo em Excel",
                data=evolution_state["excel_bytes"],
                file_name=f"memoria_evolucao_{settings.modality_code}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            if result.installments and settings.modality_code in SUPPORTED_EVOLUTION_MODALITIES:
                if st.button("Confeccionar parecer", type="primary", use_container_width=True):
                    st.session_state["show_report_builder"] = True
                    st.rerun()
            else:
                st.button("Confeccionar parecer", disabled=True, use_container_width=True)

        if st.session_state.get("show_report_builder"):
            st.markdown("---")
            page_intro(
                "Etapa 06",
                "Confeccionar parecer em PDF",
                "Selecione as prestações, preencha os dados e aplique todas as alterações de uma só vez. O cálculo original permanece preservado.",
            )

            if not result.installments:
                st.warning("Não há prestações concluídas para compor o parecer.")
            else:
                total_available = len(result.installments)
                selection_options = ["Primeira", "Duas primeiras", "Primeiras 3", "Primeiras 6", "Primeiras 12", "Personalizado"]
                selection_mode = st.selectbox(
                    "Prestações a apresentar", selection_options, key="report_selection_mode"
                )
                counts = {
                    "Primeira": 1,
                    "Duas primeiras": 2,
                    "Primeiras 3": 3,
                    "Primeiras 6": 6,
                    "Primeiras 12": 12,
                }
                if selection_mode == "Personalizado":
                    default_custom = st.session_state.get("report_selected_custom", [1])
                    st.session_state.setdefault(
                        "report_selected_custom_widget",
                        [n for n in default_custom if n <= total_available] or [1],
                    )
                    selected_numbers = st.multiselect(
                        "Escolha as prestações",
                        options=[row.installment_number for row in result.installments],
                        key="report_selected_custom_widget",
                    )
                    st.session_state["report_selected_custom"] = selected_numbers
                else:
                    selected_numbers = [row.installment_number for row in result.installments[:counts[selection_mode]]]

                if not selected_numbers:
                    st.warning("Selecione ao menos uma prestação.")
                else:
                    rows_by_number = {row.installment_number: row for row in result.installments}
                    existing_adjustments = st.session_state.get("report_adjustments", {})
                    field_labels = [
                        ("balance_before_installment", "Saldo antes"),
                        ("correction_amount", "Correção"),
                        ("interest_amount", "Juros"),
                        ("regular_amortization", "Amortização"),
                        ("installment_amount", "Prestação"),
                        ("closing_balance", "Saldo final"),
                    ]
                    editor_rows = []
                    for number in selected_numbers:
                        row = rows_by_number[number]
                        item = {"Nº": number, "Referência": date_br(row.reference_due_date)}
                        for field, label in field_labels:
                            calculated = getattr(row, field)
                            presented = existing_adjustments.get(number, {}).get(field, calculated)
                            item[f"{label} calculado"] = money_br(calculated, settings.rounding.money_places)
                            item[f"{label} apresentado"] = money_br(presented, settings.rounding.money_places)
                        fgqc_shown = existing_adjustments.get(number, {}).get("fgqc_amount", Decimal("0"))
                        item["FGQC apresentado"] = money_br(fgqc_shown, settings.rounding.money_places)
                        editor_rows.append(item)
                    editor_df = pd.DataFrame(editor_rows)
                    calculated_columns = [column for column in editor_df.columns if column.endswith("calculado")] + ["Nº", "Referência"]

                    st.markdown(
                        '<div class="success-note"><b>Preenchimento otimizado:</b> os campos abaixo estão em um formulário. '
                        'Digite IOF, FGQC, taxa e demais informações livremente; a página só será processada quando você clicar em “Aplicar dados e ajustes”.</div>',
                        unsafe_allow_html=True,
                    )

                    with st.form(
                        key=f"report_form_{st.session_state['evolution_revision']}_{'_'.join(map(str, selected_numbers))}",
                        clear_on_submit=False,
                    ):
                        st.markdown('<div class="mini-heading">Identificação do documento</div>', unsafe_allow_html=True)
                        st.caption(
                            "Modo web/LGPD: dados pessoais do participante (nome, matrícula, CPF) e validador não são coletados nesta versão. "
                            "O documento traz apenas o responsável pela informação, conforme o padrão FUNCEF."
                        )
                        i1, i2, i3 = st.columns(3)
                        with i1:
                            contract_number = st.text_input("Número do contrato", key="report_contract_number")
                        with i2:
                            st.session_state.setdefault("report_request_date", settings.credit_date)
                            request_date = st.date_input(
                                "Data da solicitação",
                                format="DD/MM/YYYY",
                                key="report_request_date",
                            )
                        with i3:
                            st.session_state.setdefault("report_net_amount", float(settings.initial_balance))
                            net_amount = st.number_input(
                                "Valor líquido (R$)", min_value=0.0, key="report_net_amount"
                            )

                        participant_name = ""
                        registration = ""
                        cpf = ""
                        elaborator = ""
                        validator = ""
                        additional_note = st.text_area("Observação adicional", key="report_additional_note", height=76)

                        with st.expander("Cabeçalho padrão FUNCEF (Manifestação de Subsídios)", expanded=True):
                            st.caption(
                                "Campos da abertura do documento. Contrato(s) e Modalidade são preenchidos automaticamente; "
                                "Mutuário(s) e Matrícula(s) saem em branco no modo web/LGPD."
                            )
                            st.session_state.setdefault("report_hdr_date", date.today())
                            st.session_state.setdefault("report_hdr_authors", DEFAULT_AUTHORS)
                            st.session_state.setdefault("report_hdr_operation", DEFAULT_OPERATION)
                            h1, h2, h3 = st.columns(3)
                            with h1:
                                hdr_date = st.date_input("Data do documento", format="DD/MM/YYYY", key="report_hdr_date")
                                hdr_court = st.text_input("Vara", key="report_hdr_court")
                                hdr_authors = st.text_input("Autor(es)", key="report_hdr_authors")
                            with h2:
                                hdr_process = st.text_input("Processo nº", key="report_hdr_process")
                                hdr_state = st.text_input("UF", key="report_hdr_state")
                                hdr_operation = st.text_input("Operação com participante", key="report_hdr_operation")
                            with h3:
                                hdr_district = st.text_input("Comarca", key="report_hdr_district")
                                hdr_lawyer = st.text_input("Advogado responsável", key="report_hdr_lawyer")
                                hdr_area = st.text_input("Área de destino", key="report_hdr_area")
                            s1, s2 = st.columns(2)
                            with s1:
                                hdr_subject = st.text_input("Assunto", key="report_hdr_subject")
                            with s2:
                                hdr_reference = st.text_input("Referência", key="report_hdr_reference")
                            hdr_responsible = st.text_input("Responsável pela informação", key="report_hdr_responsible")

                        with st.expander("Dados contratuais complementares", expanded=True):
                            x1, x2, x3 = st.columns(3)
                            with x1:
                                margin_amount = st.number_input("Margem consignável (R$)", min_value=0.0, key="report_margin")
                                fgqc_concession = st.number_input("FGQC - concessão (R$)", min_value=0.0, key="report_fgqc")
                            with x2:
                                iof_amount = st.number_input("IOF (R$)", min_value=0.0, key="report_iof")
                                administrative_fee = st.number_input("Taxa administrativa (R$)", min_value=0.0, key="report_admin_fee")
                            with x3:
                                settled_loan_balance = st.number_input(
                                    "Saldo quitado na operação (R$)", min_value=0.0, key="report_settled_balance"
                                )

                        st.markdown('<div class="mini-heading">Revisão dos valores apresentados</div>', unsafe_allow_html=True)
                        st.caption(
                            "As colunas calculadas são somente referência. Edite as colunas 'Apresentado' no padrão brasileiro. "
                            "O FGQC é informado por prestação e permanece separado do valor da prestação. "
                            "Os ajustes não alteram o motor nem a memória de cálculo."
                        )
                        edited_df = st.data_editor(
                            editor_df,
                            use_container_width=True,
                            hide_index=True,
                            disabled=calculated_columns,
                            key=f"report_values_editor_{st.session_state['evolution_revision']}_{'_'.join(map(str, selected_numbers))}",
                            column_config={
                                "Nº": st.column_config.NumberColumn(width="small"),
                                "Referência": st.column_config.TextColumn(width="small"),
                            },
                        )
                        apply_report = st.form_submit_button(
                            "Aplicar dados e ajustes",
                            type="primary",
                            use_container_width=True,
                        )

                    if apply_report:
                        missing = ["Número do contrato"] if not str(contract_number).strip() else []
                        if missing:
                            st.error("Preencha os campos obrigatórios: " + ", ".join(missing) + ".")
                        else:
                            try:
                                overrides = {}
                                validation_messages = []
                                for _, item in edited_df.iterrows():
                                    number = int(item["Nº"])
                                    per_row = {}
                                    for field, label in field_labels:
                                        per_row[field] = parse_money_br(item[f"{label} apresentado"])
                                    per_row["fgqc_amount"] = parse_money_br(item["FGQC apresentado"])
                                    overrides[number] = per_row
                                    if settings.modality_code == "MOD_005":
                                        if abs(per_row["installment_amount"] - per_row["regular_amortization"]) > Decimal("0.01"):
                                            validation_messages.append(
                                                f"Prestação {number}: no Credinâmico Fixo, o valor apresentado da prestação deve corresponder à amortização do ciclo."
                                            )
                                    elif abs(per_row["installment_amount"] - (per_row["interest_amount"] + per_row["regular_amortization"])) > Decimal("0.01"):
                                        validation_messages.append(
                                            f"Prestação {number}: o valor apresentado da prestação não corresponde à soma dos juros e da amortização."
                                        )
                                identity = ReportIdentity(
                                    contract_number=contract_number.strip(),
                                    participant_name=participant_name.strip(),
                                    registration=registration.strip(),
                                    cpf=cpf.strip(),
                                    request_date=request_date,
                                    margin_amount=Decimal(str(margin_amount)),
                                    fgqc_concession=Decimal(str(fgqc_concession)),
                                    iof_amount=Decimal(str(iof_amount)),
                                    administrative_fee=Decimal(str(administrative_fee)),
                                    settled_loan_balance=Decimal(str(settled_loan_balance)),
                                    net_amount=Decimal(str(net_amount)),
                                    elaborator=elaborator.strip(),
                                    validator=validator.strip(),
                                    additional_note=additional_note.strip(),
                                    header=ManifestationHeader(
                                        issue_date=hdr_date,
                                        process_number=hdr_process.strip(),
                                        district=hdr_district.strip(),
                                        court=hdr_court.strip(),
                                        state=hdr_state.strip(),
                                        authors=hdr_authors.strip(),
                                        lawyer=hdr_lawyer.strip(),
                                        operation=hdr_operation.strip(),
                                        contracts=contract_number.strip(),
                                        modality=modality_name,
                                        destination_area=hdr_area.strip(),
                                        subject=hdr_subject.strip(),
                                        reference=hdr_reference.strip(),
                                        responsible=hdr_responsible.strip(),
                                    ),
                                )
                                st.session_state["report_identity"] = identity
                                st.session_state["report_adjustments"] = overrides
                                st.session_state["report_selected_numbers"] = list(selected_numbers)
                                st.session_state["report_validation_messages"] = validation_messages
                                st.session_state["report_pdf"] = None
                                st.session_state["report_pdf_name"] = None
                                st.session_state["report_docx"] = None
                                st.session_state["report_docx_name"] = None
                                st.success("Dados e ajustes aplicados. O parecer está pronto para ser gerado.")
                            except (InvalidOperation, ValueError) as exc:
                                st.error(f"Não foi possível interpretar um valor apresentado: {exc}")

                    for message in st.session_state.get("report_validation_messages", []):
                        st.warning(message)

                    applied_identity = st.session_state.get("report_identity")
                    applied_numbers = st.session_state.get("report_selected_numbers", [])
                    ready_to_generate = (
                        applied_identity is not None
                        and list(applied_numbers) == list(selected_numbers)
                        and bool(st.session_state.get("report_adjustments"))
                    )
                    if applied_identity is not None and list(applied_numbers) != list(selected_numbers):
                        st.warning("A seleção de prestações mudou. Clique novamente em “Aplicar dados e ajustes” antes de gerar o PDF.")

                    g1, g2 = st.columns([3, 1])
                    with g1:
                        generate_pdf = st.button(
                            "Gerar parecer (PDF e Word)",
                            type="primary",
                            use_container_width=True,
                            disabled=not ready_to_generate,
                        )
                    with g2:
                        if st.button("Fechar confecção", use_container_width=True):
                            st.session_state["show_report_builder"] = False
                            st.rerun()

                    if generate_pdf and ready_to_generate:
                        try:
                            pdf_bytes = build_opinion_pdf(
                                settings=settings,
                                result=result,
                                modality_name=modality_name,
                                identity=applied_identity,
                                selected_installment_numbers=applied_numbers,
                                overrides=st.session_state["report_adjustments"],
                                logo_path=LOGO_WHITE_PATH,
                            )
                            st.session_state["report_pdf"] = pdf_bytes
                            st.session_state["report_pdf_name"] = pdf_filename(
                                modality_name, applied_identity.contract_number
                            )
                            st.session_state["report_docx"] = build_opinion_docx(
                                settings=settings,
                                result=result,
                                modality_name=modality_name,
                                identity=applied_identity,
                                selected_installment_numbers=applied_numbers,
                                overrides=st.session_state["report_adjustments"],
                            )
                            st.session_state["report_docx_name"] = docx_filename(
                                modality_name, applied_identity.contract_number
                            )
                            st.success("Parecer gerado em PDF e Word. O cálculo original e os campos da evolução continuam preservados.")
                        except (ValueError, KeyError, InvalidOperation) as exc:
                            st.error(str(exc))

                    if st.session_state.get("report_pdf"):
                        d1, d2 = st.columns(2)
                        with d1:
                            st.download_button(
                                "Baixar parecer em PDF",
                                data=st.session_state["report_pdf"],
                                file_name=st.session_state["report_pdf_name"],
                                mime="application/pdf",
                                use_container_width=True,
                            )
                        with d2:
                            if st.session_state.get("report_docx"):
                                st.download_button(
                                    "Baixar parecer em Word (.docx)",
                                    data=st.session_state["report_docx"],
                                    file_name=st.session_state["report_docx_name"],
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                    use_container_width=True,
                                )
                        st.caption(
                            "O arquivo Word traz o mesmo conteúdo do PDF no modelo corporativo e pode ser editado, "
                            "por exemplo para preencher Mutuário(s) e Matrícula(s), que o sistema não coleta por causa da LGPD."
                        )


elif module == "Recálculo de diferenças":
    majs_config = get_active_config()
    page_intro(
        "Módulo 02",
        "Recálculo por MAJS e diferenças",
        "Reconstrua o fluxo teórico pela metodologia MAJS, confronte pagamentos realizados com as prestações devidas e atualize as diferenças até a data-base.",
    )
    workflow_strip(["Contrato", "Pagamentos", "Eventos", "Diferenças", "Calcular", "Resultado"], active=0)

    callout(
        "Como o módulo trata os eventos",
        "Pagamentos ordinários são comparados com o valor teoricamente devido. Amortizações extraordinárias reduzem o saldo e reiniciam o fator; suspensões mantêm apenas a correção monetária, sem prestação ou juros. Se houver quitação antecipada, o fluxo MAJS é encerrado na data informada.",
        "info",
    )

    step_header(1, "Parâmetros do recálculo MAJS", "Informe os dados contratuais usados para reconstruir o fluxo teórico.")
    with st.container(border=True):
        m1, m2, m3, m4 = st.columns(4)
        with m1:
            majs_modality = st.selectbox(
                "Modalidade do contrato",
                ["Variável — com correção monetária", "Fixa — sem correção monetária"],
                key="majs_modality",
                help="Na modalidade fixa o saldo MAJS não recebe correção monetária. Os juros contratuais continuam sendo calculados pela taxa informada.",
            )
            majs_original = st.number_input(
                "Valor original", min_value=0.01,
                value=float(st.session_state.get("majs_original", 80000.0)),
                step=100.0, format="%.2f", key="majs_original"
            )
        with m2:
            majs_credit = st.date_input(
                "Data do crédito", value=st.session_state.get("majs_credit", date(2015, 7, 8)),
                format="DD/MM/YYYY", key="majs_credit"
            )
            majs_annual_pct = st.number_input(
                "Taxa nominal (% a.a.)", min_value=0.0,
                value=float(st.session_state.get("majs_annual_pct", 7.9)),
                step=0.01, format="%.6f", key="majs_annual_pct",
                help="A metodologia MAJS converte a taxa nominal anual em taxa mensal por divisão por 12."
            )
        with m3:
            majs_term = st.number_input(
                "Prazo (prestações)", min_value=1, max_value=600,
                value=int(st.session_state.get("majs_term", 96)), step=1, key="majs_term",
                help="Competências em suspensão não consomem uma prestação do prazo."
            )
            majs_base = st.date_input(
                "Data-base", value=st.session_state.get("majs_base", date.today()),
                min_value=majs_credit + timedelta(days=1), format="DD/MM/YYYY", key="majs_base"
            )
        with m4:
            is_fixed_majs = majs_modality.startswith("Fixa")
            majs_index_options = [c for c in sorted(majs_config.indices.keys()) if c != "TAXA_LEGAL"]
            if is_fixed_majs:
                majs_index = "SEM_CORRECAO"
                st.text_input("Índice do saldo MAJS", value="Sem correção monetária", disabled=True)
                majs_lag = 0
                st.number_input("Defasagem do índice (meses)", value=0, min_value=0, max_value=12, disabled=True)
            else:
                default_majs_idx = "INPC" if "INPC" in majs_index_options else majs_index_options[0]
                saved_majs_idx = st.session_state.get("majs_index", default_majs_idx)
                if saved_majs_idx not in majs_index_options:
                    saved_majs_idx = default_majs_idx
                majs_index = st.selectbox(
                    "Índice do saldo MAJS", majs_index_options,
                    index=majs_index_options.index(saved_majs_idx), key="majs_index"
                )
                majs_lag = st.number_input(
                    "Defasagem do índice (meses)", min_value=0, max_value=12,
                    value=int(st.session_state.get("majs_lag", 2)), step=1, key="majs_lag"
                )

    monthly_preview = Decimal(str(majs_annual_pct)) / Decimal("12") / Decimal("100")
    modality_caption = "modalidade fixa, sem correção monetária" if is_fixed_majs else f"correção por {majs_index} com defasagem de {majs_lag} mês(es)"
    st.caption(
        f"Taxa mensal usada no MAJS: {percent_br(monthly_preview, 6)} a.m. · "
        f"1º vencimento: dia 20 do mês seguinte ao crédito · {modality_caption}."
    )

    with st.container(border=True):
        majs_early_settlement_enabled = st.toggle(
            "Contrato quitado antecipadamente",
            value=bool(st.session_state.get("majs_early_settlement_enabled", False)),
            key="majs_early_settlement_enabled",
            help="Use quando o contrato deixou de evoluir até o prazo originalmente contratado e foi encerrado em uma data anterior.",
        )
        majs_settlement_date = None
        majs_settlement_is_novation = False
        majs_settlement_paid = Decimal("0")
        majs_settlement_note = ""
        if majs_early_settlement_enabled:
            q1, q2, q3 = st.columns([1.0, 1.15, 1.15])
            with q1:
                saved_settlement_date = st.session_state.get("majs_settlement_date", majs_base)
                if not isinstance(saved_settlement_date, date) or saved_settlement_date <= majs_credit or saved_settlement_date > majs_base:
                    saved_settlement_date = majs_base
                majs_settlement_date = st.date_input(
                    "Data da quitação",
                    value=saved_settlement_date,
                    min_value=majs_credit + timedelta(days=1),
                    max_value=majs_base,
                    format="DD/MM/YYYY",
                    key="majs_settlement_date",
                )
            with q2:
                settlement_type = st.selectbox(
                    "Forma da quitação",
                    ["Novação", "Pagamento pelo(a) participante"],
                    key="majs_settlement_type",
                    help="Na novação não há desembolso do participante a considerar como pagamento na tabela de diferenças.",
                )
                majs_settlement_is_novation = settlement_type == "Novação"
            with q3:
                if majs_settlement_is_novation:
                    st.text_input(
                        "Valor pago pelo(a) participante", value="Não se aplica - novação", disabled=True,
                        key="majs_settlement_paid_disabled",
                    )
                    majs_settlement_paid = Decimal("0")
                else:
                    raw_settlement_paid = st.number_input(
                        "Valor efetivamente pago", min_value=0.0, step=100.0, format="%.2f",
                        value=float(st.session_state.get("majs_settlement_paid", 0.0)),
                        key="majs_settlement_paid",
                    )
                    majs_settlement_paid = Decimal(str(raw_settlement_paid))
            majs_settlement_note = st.text_input(
                "Observação da quitação (opcional)",
                key="majs_settlement_note",
                placeholder="Ex.: contrato encerrado por acordo / instrumento de novação...",
            )
            if majs_settlement_is_novation:
                callout(
                    "Quitação por novação",
                    "O Motor apurará o saldo MAJS teórico na data da novação e encerrará a evolução ali. Como não houve desembolso do(a) participante, esse saldo não será lançado como pagamento nem gerará uma linha de diferença de quitação.",
                    "info",
                )
            else:
                callout(
                    "Quitação por pagamento",
                    "O Motor apurará o saldo MAJS teórico na data da quitação, confrontará esse saldo com o valor efetivamente pago e levará a diferença para a tabela de diferenças. Prestações posteriores não serão geradas.",
                    "info",
                )

    step_header(2, "Pagamentos realizados", "Importe a planilha de pagamentos ou faça ajustes diretamente na grade.")

    def _normalize_majs_column(name: object) -> str:
        text_col = unicodedata.normalize("NFKD", str(name or "")).encode("ascii", "ignore").decode("ascii")
        return " ".join(text_col.strip().lower().replace("_", " ").split())

    def _prepare_majs_payments(df: pd.DataFrame) -> pd.DataFrame:
        aliases = {
            "prestacao": "Prestação",
            "data pagamento": "Data pagamento",
            "data de pagamento": "Data pagamento",
            "valor pago": "Valor pago",
            "competencia": "Competência",
            "observacao": "Observação",
        }
        rename = {}
        for col in df.columns:
            normalized = _normalize_majs_column(col)
            if normalized in aliases:
                rename[col] = aliases[normalized]
        df = df.rename(columns=rename).copy()
        required = ["Prestação", "Data pagamento", "Valor pago", "Competência", "Observação"]
        for col in required:
            if col not in df.columns:
                df[col] = None
        df = df[required]
        df["Prestação"] = pd.to_numeric(df["Prestação"], errors="coerce")
        df["Data pagamento"] = pd.to_datetime(df["Data pagamento"], errors="coerce", dayfirst=True).dt.date
        df["Valor pago"] = df["Valor pago"].apply(
            lambda x: (float(parse_money_br(x)) if not pd.isna(x) and str(x).strip() else None)
        )
        df["Competência"] = df["Competência"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        df["Observação"] = df["Observação"].apply(lambda x: "" if pd.isna(x) else str(x).strip())
        return df

    pay_model = APP_DIR / "exemplos" / "modelo_majs_pagamentos.csv"
    if pay_model.exists():
        st.download_button(
            "Baixar modelo CSV de pagamentos", pay_model.read_bytes(),
            file_name="modelo_majs_pagamentos.csv", mime="text/csv"
        )
    pay_upload = st.file_uploader(
        "Importar pagamentos (CSV ou XLSX)", type=["csv", "xlsx"], key="majs_pay_upload",
        help="Aceita o modelo Pagamentos MAJS com separador ;, vírgula decimal e cabeçalhos com ou sem acentos."
    )
    if pay_upload is not None:
        try:
            if pay_upload.name.lower().endswith(".csv"):
                imported_pay = pd.read_csv(pay_upload, sep=None, engine="python", dtype=str)
            else:
                imported_pay = pd.read_excel(pay_upload, dtype=object)
            imported_pay = _prepare_majs_payments(imported_pay)
            st.session_state["majs_pay_seed"] = imported_pay
            st.success(f"Pagamentos importados: {len(imported_pay)} linha(s).")
        except Exception as exc:
            st.error(f"Não foi possível ler os pagamentos: {exc}")
    if "majs_pay_seed" not in st.session_state:
        st.session_state["majs_pay_seed"] = pd.DataFrame(
            columns=["Prestação", "Data pagamento", "Valor pago", "Competência", "Observação"]
        )
    majs_pay_df = st.data_editor(
        st.session_state["majs_pay_seed"], num_rows="dynamic", hide_index=True,
        use_container_width=True, key="majs_pay_editor",
        column_config={
            "Prestação": st.column_config.NumberColumn(min_value=1, step=1),
            "Data pagamento": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Valor pago": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0),
            "Competência": st.column_config.TextColumn(help="Opcional. Aceita AAAA-MM, MM/AAAA ou DD/MM/AAAA."),
            "Observação": st.column_config.TextColumn(),
        },
    )

    def _prepare_majs_event_seed(df: pd.DataFrame, *, columns: list[str], date_columns: list[str]) -> pd.DataFrame:
        """Mantém tipos compatíveis com st.data_editor/DateColumn, inclusive em grades vazias."""
        out = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
        for col in columns:
            if col not in out.columns:
                out[col] = None
        out = out[columns].copy()
        for col in date_columns:
            out[col] = pd.to_datetime(out[col], errors="coerce")
        if "Valor" in out.columns:
            out["Valor"] = pd.to_numeric(out["Valor"], errors="coerce")
        for col in [c for c in columns if c not in date_columns and c != "Valor"]:
            out[col] = out[col].apply(lambda x: "" if pd.isna(x) else str(x))
        return out

    step_header(3, "Amortizações extraordinárias e suspensões", "Cadastre eventos que alteram o fluxo teórico do contrato.")
    tab_extra, tab_suspend = st.tabs(["Amortizações extraordinárias", "Suspensões"])
    with tab_extra:
        st.caption(
            "Na data da amortização extraordinária, o saldo recebe primeiro somente a correção proporcional aos dias transcorridos no ciclo; "
            "no vencimento seguinte, aplica-se apenas a parcela remanescente do índice. A amortização é aplicada depois dessa correção proporcional e reinicia a sequência do fator. "
            "Se o valor informado superar o saldo já corrigido, "
            "somente o necessário quita a dívida e o excedente entra como diferença credora."
        )
        extra_model = APP_DIR / "exemplos" / "modelo_majs_amortizacoes.csv"
        if extra_model.exists():
            st.download_button(
                "Baixar modelo CSV de amortizações", extra_model.read_bytes(),
                file_name="modelo_majs_amortizacoes.csv", mime="text/csv"
            )
        if "majs_extra_seed" not in st.session_state:
            st.session_state["majs_extra_seed"] = pd.DataFrame({
                "Data": pd.Series(dtype="datetime64[ns]"),
                "Valor": pd.Series(dtype="float64"),
                "Observação": pd.Series(dtype="object"),
            })
        st.session_state["majs_extra_seed"] = _prepare_majs_event_seed(
            st.session_state["majs_extra_seed"],
            columns=["Data", "Valor", "Observação"], date_columns=["Data"],
        )
        majs_extra_df = st.data_editor(
            st.session_state["majs_extra_seed"], num_rows="dynamic", hide_index=True,
            use_container_width=True, key="majs_extra_editor",
            column_config={
                "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Valor": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.0),
                "Observação": st.column_config.TextColumn(),
            },
        )
    with tab_suspend:
        st.caption(
            "Durante a suspensão, cada competência apenas corrige o saldo. Não há prestação, juros nem amortização contratual, "
            "e a quantidade de prestações do prazo é preservada."
        )
        if "majs_suspend_seed" not in st.session_state:
            st.session_state["majs_suspend_seed"] = pd.DataFrame({
                "Data inicial": pd.Series(dtype="datetime64[ns]"),
                "Data final": pd.Series(dtype="datetime64[ns]"),
                "Observação": pd.Series(dtype="object"),
            })
        st.session_state["majs_suspend_seed"] = _prepare_majs_event_seed(
            st.session_state["majs_suspend_seed"],
            columns=["Data inicial", "Data final", "Observação"],
            date_columns=["Data inicial", "Data final"],
        )
        majs_suspend_df = st.data_editor(
            st.session_state["majs_suspend_seed"], num_rows="dynamic", hide_index=True,
            use_container_width=True, key="majs_suspend_editor",
            column_config={
                "Data inicial": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Data final": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Observação": st.column_config.TextColumn(),
            },
        )

    step_header(4, "Atualização das diferenças", "Defina como cada diferença entre pago e devido será levada até a data-base.")
    with st.container(border=True):
        d1, d2 = st.columns(2)
        corr_diff_labels = {
            "INPC até 08/2024 e IPCA a partir de 09/2024": "INPC_IPCA_14905",
            "Sem correção monetária": "SEM_CORRECAO",
        }
        for code in sorted(majs_config.indices.keys()):
            if code != "TAXA_LEGAL":
                corr_diff_labels[f"Usar {code} em todo o período"] = code
        with d1:
            majs_diff_corr_label = st.selectbox(
                "Correção das diferenças", list(corr_diff_labels.keys()), index=0, key="majs_diff_corr"
            )
        int_diff_labels = {
            "1% a.m. — pró-rata dias/30": "UM_PCT",
            "1% a.m. — fração de ano (convenção de dias)": "UM_PCT_YEARFRAC",
            "1% a.m. dias/30 até 29/08/2024 + Taxa Legal a partir de 30/08/2024": "UM_PCT_ATE_TL",
            "1% a.m. por fração de ano até 29/08/2024 + Taxa Legal a partir de 30/08/2024": "UM_PCT_YEARFRAC_ATE_TL",
            "Somente Taxa Legal a partir de 30/08/2024": "TAXA_LEGAL",
            "Sem juros": "SEM_JUROS",
        }
        with d2:
            majs_diff_int_label = st.selectbox(
                "Juros sobre as diferenças", list(int_diff_labels.keys()), index=0, key="majs_diff_int"
            )
        st.markdown("**Início dos juros de mora**")
        start_mode = st.radio(
            "Critério para o termo inicial", ["Data de cada diferença", "Data fixa"],
            horizontal=True, key="majs_diff_start_mode",
            format_func=lambda value: (
                "Vencimento/data de origem de cada diferença"
                if value == "Data de cada diferença"
                else "Data específica (ex.: citação)"
            ),
            label_visibility="collapsed",
        )
        majs_fixed_start = None
        if start_mode == "Data fixa":
            majs_fixed_start = st.date_input(
                "Data de início dos juros de mora",
                value=st.session_state.get("majs_fixed_start", majs_credit),
                min_value=MIN_CREDIT_DATE, max_value=majs_base,
                format="DD/MM/YYYY", key="majs_fixed_start"
            )
            st.caption(
                "Ex.: data da citação. Se uma diferença surgir depois dessa data, os juros começam na própria "
                "data de origem da diferença, pois não há mora sobre valor ainda não constituído."
            )
        selected_interest_mode = int_diff_labels[majs_diff_int_label]
        yearfrac_basis_labels = {
            "US/NASD 30/360": 0,
            "Real/Real": 1,
            "Real/360": 2,
            "Real/365": 3,
            "Europeu 30/360": 4,
        }
        majs_yearfrac_basis = 0
        if "YEARFRAC" in selected_interest_mode:
            basis_label = st.selectbox(
                "Convenção de contagem de dias", list(yearfrac_basis_labels.keys()),
                index=0, key="majs_yearfrac_basis_label",
                help="A taxa de 1% a.m. é anualizada nominalmente para 12% a.a. e aplicada proporcionalmente à fração de ano apurada pela convenção de contagem selecionada.",
            )
            majs_yearfrac_basis = yearfrac_basis_labels[basis_label]
            st.info(
                "Critério aplicado aos juros de 1%: a taxa mensal é anualizada nominalmente para 12% a.a. e multiplicada pela "
                f"fração de ano entre as datas, segundo a convenção selecionada (código {majs_yearfrac_basis}). "
                "Os juros correspondem à diferença corrigida multiplicada pelo percentual acumulado."
            )
        elif selected_interest_mode == "UM_PCT":
            st.caption(
                "Fórmula: taxa acumulada = 1% a.m. × (dias corridos / 30). "
                "Juros = diferença corrigida × taxa acumulada."
            )
        elif selected_interest_mode == "UM_PCT_ATE_TL":
            st.caption(
                "Até 29/08/2024: 1% a.m. × (dias/30). A partir de 30/08/2024: Taxa Legal, em juros simples."
            )
        elif selected_interest_mode == "TAXA_LEGAL":
            st.caption(
                "A Taxa Legal usa a série oficial mensal do Banco Central e juros simples. "
                "Frações de mês são apropriadas pro rata por dias corridos."
            )
        st.caption(
            "Se algum índice necessário ainda não estiver disponível, o cálculo será parcial: "
            "o sistema evolui somente até a última competência disponível e informa o corte, sem projetar índice futuro."
        )

    step_header(5, "Executar recálculo", "O cálculo gera evolução teórica, conciliação pago × devido e memória das diferenças.")
    if st.button("Calcular MAJS e diferenças", type="primary", use_container_width=True):
        try:
            payments = []
            for _, r in majs_pay_df.iterrows():
                raw_amount = r.get("Valor pago")
                raw_date = r.get("Data pagamento")
                amount = parse_money_br(raw_amount) if not pd.isna(raw_amount) else Decimal("0")
                if amount == 0 or pd.isna(raw_date):
                    continue
                inst_raw = r.get("Prestação")
                inst = None if pd.isna(inst_raw) else int(inst_raw)
                comp_raw = r.get("Competência")
                comp = None if pd.isna(comp_raw) else str(comp_raw).strip() or None
                payments.append(MAJSPayment(
                    payment_date=parse_ui_date(raw_date), amount=amount,
                    installment_number=inst, competence=comp,
                    note="" if pd.isna(r.get("Observação")) else str(r.get("Observação") or "")
                ))

            extras = []
            for _, r in majs_extra_df.iterrows():
                raw_value = r.get("Valor")
                value = parse_money_br(raw_value) if not pd.isna(raw_value) else Decimal("0")
                if pd.isna(r.get("Data")) or value == 0:
                    continue
                extras.append(MAJSExtraAmortization(
                    event_date=parse_ui_date(r.get("Data")), amount=value,
                    note="" if pd.isna(r.get("Observação")) else str(r.get("Observação") or "")
                ))

            suspensions = []
            for _, r in majs_suspend_df.iterrows():
                raw_start, raw_end = r.get("Data inicial"), r.get("Data final")
                if pd.isna(raw_start) and pd.isna(raw_end):
                    continue
                if pd.isna(raw_start) or pd.isna(raw_end):
                    raise ValueError("Toda suspensão deve ter data inicial e data final.")
                suspensions.append(MAJSSuspension(
                    start_date=parse_ui_date(raw_start), end_date=parse_ui_date(raw_end),
                    note="" if pd.isna(r.get("Observação")) else str(r.get("Observação") or "")
                ))

            settings_majs = MAJSSettings(
                original_amount=Decimal(str(majs_original)), credit_date=majs_credit,
                monthly_interest_rate=monthly_preview, term_months=int(majs_term),
                base_date=majs_base, correction_index_code=majs_index,
                correction_lag_months=int(majs_lag),
            )
            majs_early_settlement = None
            if majs_early_settlement_enabled:
                if majs_settlement_date is None:
                    raise ValueError("Informe a data da quitação antecipada.")
                majs_early_settlement = MAJSEarlySettlement(
                    settlement_date=majs_settlement_date,
                    novation=majs_settlement_is_novation,
                    amount_paid=(Decimal("0") if majs_settlement_is_novation else majs_settlement_paid),
                    note=majs_settlement_note,
                )
            majs_result = calculate_majs(
                settings_majs, majs_config.indices, payments, extras, suspensions,
                early_settlement=majs_early_settlement,
                difference_correction_mode=corr_diff_labels[majs_diff_corr_label],
                difference_interest_mode=int_diff_labels[majs_diff_int_label],
                fixed_interest_start=majs_fixed_start,
                yearfrac_basis=int(majs_yearfrac_basis),
            )
            st.session_state["majs_result"] = majs_result
            st.session_state["majs_excel"] = build_majs_workbook(majs_result)
            st.session_state["majs_pay_seed"] = majs_pay_df.copy()
            st.session_state["majs_extra_seed"] = majs_extra_df.copy()
            st.session_state["majs_suspend_seed"] = majs_suspend_df.copy()
            st.session_state.pop("majs_report_pdf", None)
            st.session_state.pop("majs_report_pdf_name", None)
            st.success("Recálculo MAJS concluído.")
        except Exception as exc:
            st.error(str(exc))

    majs_result = st.session_state.get("majs_result")
    if majs_result:
        step_header(6, "Resultado", "Confira o fluxo teórico, as suspensões e as diferenças reconciliadas.")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Devido até a data-base", money_br(majs_result.due_total))
        c2.metric("Pago/crédito conciliado", money_br(majs_result.paid_total))
        c3.metric("Diferença pago - devido", money_br(majs_result.difference_total))
        c4.metric("Diferença atualizada", money_br(majs_result.updated_difference_total))
        c5.metric("Saldo devedor a vencer", money_br(majs_result.balance_to_mature))
        if majs_result.was_settled_early and majs_result.early_settlement:
            settlement = majs_result.early_settlement
            kind = "Novação" if settlement.novation else "Pagamento pelo(a) participante"
            paid_text = "não compõe pagamentos/diferenças" if settlement.novation else f"valor pago: {money_br(settlement.amount_paid)}"
            callout(
                "Quitação antecipada processada",
                f"Data: {settlement.settlement_date.strftime('%d/%m/%Y')} · forma: {kind} · saldo MAJS teórico na quitação: {money_br(majs_result.early_settlement_balance)} · {paid_text}. O fluxo contratual não evolui após essa data.",
                "success",
            )
        for warning in majs_result.warnings:
            st.warning(warning)

        tab_evo, tab_diff = st.tabs(["Evolução MAJS", "Diferenças pago × devido"])
        with tab_evo:
            event_labels = {
                "PRESTACAO": "Prestação",
                "SUSPENSAO": "Suspensão",
                "AMORTIZACAO_EXTRA": "Amortização extraordinária",
                "QUITACAO_ANTECIPADA": "Quitação antecipada",
            }
            evo_df = pd.DataFrame([{
                "Evento": event_labels.get(r.event_type, r.event_type),
                "Prestação": r.installment_number,
                "Data": r.due_date,
                "Saldo inicial": float(r.opening_balance),
                "Posição fator": r.factor_position,
                "Fator financeiro": float(r.financial_factor),
                "Soma fatores": float(r.remaining_factor_sum),
                "Prestação devida": float(r.installment_due),
                "Juros": float(r.interest_component),
                "Amortização": float(r.amortization_component),
                "Ref. índice": r.correction_reference or "Sem correção",
                "Correção aplicada (%)": float(r.correction_rate * Decimal('100')),
                "Amortização extra aplicada": float(r.extra_amortization),
                "Valor pago na quitação": float(r.settlement_payment),
                "Excesso da amortização": float(r.excess_extra_amortization),
                "Saldo final": float(r.closing_balance),
                "Observação": r.event_note,
            } for r in majs_result.schedule])
            st.dataframe(evo_df, use_container_width=True, hide_index=True)
        with tab_diff:
            dif_df = pd.DataFrame([{
                "Origem": r.origin,
                "Prestação": r.installment_number,
                "Data de origem": r.due_date,
                "Devido": float(r.amount_due),
                "Pago/crédito": float(r.amount_paid),
                "Diferença (pago - devido)": float(r.difference_paid_minus_due),
                "Fator correção": float(r.correction_factor),
                "Diferença corrigida": float(r.corrected_difference),
                "Início dos juros de mora": r.interest_start_date,
                "Juros acumulados (%)": float(r.accumulated_interest_rate * Decimal('100')),
                "Juros R$": float(r.interest_amount),
                "Diferença atualizada": float(r.updated_difference),
                "Atualizado até": r.updated_through_date,
                "Observação": r.note,
            } for r in majs_result.differences])
            st.dataframe(dif_df, use_container_width=True, hide_index=True)
        if st.session_state.get("majs_excel"):
            st.download_button(
                "Baixar memória MAJS em Excel", st.session_state["majs_excel"],
                file_name=f"memoria_majs_{majs_result.settings.base_date.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.markdown("---")
        st.markdown("### Parecer técnico MAJS")
        st.caption(
            "Gere um parecer premium em PDF com resumo executivo, evolução integral do MAJS em página paisagem "
            "e, quando houver pagamentos/créditos, tabela completa das diferenças com correção e juros detalhados."
        )
        with st.container(border=True):
            rp1, rp2 = st.columns(2)
            with rp1:
                majs_report_contract = st.text_input(
                    "Contrato", key="majs_report_contract", placeholder="Ex.: 123456"
                )
            with rp2:
                majs_report_document = st.text_input(
                    "Documento / processo", key="majs_report_document", placeholder="Ex.: Processo 0000000-00.0000.0.00.0000"
                )
            majs_report_participant = ""
            majs_report_elaborator = ""
            st.caption(
                "Modo web/LGPD: o parecer MAJS não coleta nem exibe nome do participante, matrícula, CPF, elaborador ou validador nesta versão."
            )
            majs_report_observation = st.text_area(
                "Observações do analista", key="majs_report_observation", height=80,
                placeholder="Observações que devam constar no parecer (opcional)."
            )

            generate_majs_pdf = st.button(
                "Gerar parecer premium MAJS em PDF", type="primary",
                use_container_width=True, key="generate_majs_pdf"
            )
            if generate_majs_pdf:
                try:
                    majs_identity = MAJSReportIdentity(
                        contract_number=majs_report_contract.strip(),
                        participant_name=majs_report_participant.strip(),
                        document_reference=majs_report_document.strip(),
                        elaborator=majs_report_elaborator.strip(),
                        observation=majs_report_observation.strip(),
                    )
                    majs_pdf_bytes = build_majs_opinion_pdf(
                        majs_result, identity=majs_identity, logo_path=LOGO_WHITE_PATH
                    )
                    st.session_state["majs_report_pdf"] = majs_pdf_bytes
                    st.session_state["majs_report_pdf_name"] = majs_opinion_pdf_filename(majs_result, majs_identity)
                    st.success("Parecer premium MAJS gerado.")
                except Exception as exc:
                    st.error(f"Não foi possível gerar o parecer MAJS: {exc}")

            if st.session_state.get("majs_report_pdf"):
                st.download_button(
                    "Baixar parecer premium MAJS em PDF",
                    data=st.session_state["majs_report_pdf"],
                    file_name=st.session_state.get("majs_report_pdf_name", "parecer_majs.pdf"),
                    mime="application/pdf",
                    use_container_width=True,
                    key="download_majs_report_pdf",
                )


elif module == "Atualização do saldo devedor":
    page_intro(
        "Módulo 03",
        "Atualização monetária do saldo devedor",
        "Atualize um saldo consolidado ou várias prestações, combinando perfis judiciais, regras por vigência, juros e abatimentos. O resultado gera memória auditável, texto pronto para copiar e PDF técnico.",
    )
    workflow_strip([
        "Valores", "Metodologia", "Correção", "Juros", "Extras", "Calcular", "Resultado",
    ], active=0)

    with st.expander("Fontes, índices e configuração temporária", expanded=False):
        update_config_upload = st.file_uploader(
            "Importar configuração/índices nesta sessão (opcional)", type=["xlsx"], key="update_config_upload",
            help="Permite usar séries adicionais sem alterar os arquivos do deploy. O upload vale somente para esta sessão.",
        )
        if update_config_upload:
            st.session_state["active_config_bytes"] = update_config_upload.getvalue()
            st.success("Configuração temporária carregada para esta sessão.")

    update_config = get_active_config()
    index_codes = sorted(update_config.indices.keys())
    factor_series = get_factor_series()
    factor_codes = sorted(factor_series.keys())
    profiles = get_profiles()

    correction_labels = {
        "Sem correção": CORRECTION_NONE,
        "Pró-rata mensal por dias corridos": CORRECTION_PRORATA_CALENDAR,
        "Diária equivalente — dias úteis": CORRECTION_DAILY_BUSINESS,
        "Diária equivalente — dias corridos da competência": CORRECTION_DAILY_CALENDAR,
        "Mensal integral — fechamento da competência": CORRECTION_MONTHLY_CLOSE,
        "Tabela prática — razão entre fatores mensais": CORRECTION_FACTOR_TABLE,
    }
    correction_labels_rev = {v: k for k, v in correction_labels.items()}
    interest_labels = {
        "Sem juros": INTEREST_NONE,
        "Simples diário — taxa mensal ÷ 30": INTEREST_SIMPLE_30,
        "Simples diário — dias da competência": INTEREST_SIMPLE_COMPETENCE,
        "TJRJ — juros simples, ano de 360 dias": INTEREST_SIMPLE_360_TJRJ,
        "Simples — fração de ano (convenção de dias)": INTEREST_SIMPLE_YEARFRAC,
        "Taxa Legal — BCB / art. 406 do Código Civil": INTEREST_LEGAL_BCB,
        "Composto diário equivalente": INTEREST_COMPOUND_EQUIV,
    }
    interest_labels_rev = {v: k for k, v in interest_labels.items()}
    interest_application_labels = {
        "Acumular percentual e aplicar no final": INTEREST_APPLICATION_END,
        "Aplicar juros diariamente": INTEREST_APPLICATION_DAILY,
    }
    interest_application_rev = {v: k for k, v in interest_application_labels.items()}
    yearfrac_basis_labels = {
        "US/NASD 30/360": 0,
        "Real/Real": 1,
        "Real/360": 2,
        "Real/365": 3,
        "Europeu 30/360": 4,
    }
    yearfrac_basis_rev = {v: k for k, v in yearfrac_basis_labels.items()}
    interest_base_labels = {"Principal nominal": INTEREST_BASE_PRINCIPAL, "Principal corrigido": INTEREST_BASE_CORRECTED}
    interest_base_rev = {v: k for k, v in interest_base_labels.items()}
    order_labels = {"Correção → juros": ORDER_CORRECTION_INTEREST, "Juros → correção": ORDER_INTEREST_CORRECTION}
    order_labels_rev = {v: k for k, v in order_labels.items()}
    competence_labels = {"Dia 1 (mês civil)": 1, "Dia 21 (ciclo 21–20)": 21}
    competence_rev = {v: k for k, v in competence_labels.items()}
    penalty_base_labels = {
        "Principal nominal": PENALTY_BASE_PRINCIPAL,
        "Principal corrigido (principal + correção)": PENALTY_BASE_CORRECTED,
        "Saldo antes da multa (principal + correção + juros)": PENALTY_BASE_BALANCE,
    }
    penalty_base_rev = {v: k for k, v in penalty_base_labels.items()}
    correction_index_codes = [c for c in index_codes if c != "TAXA_LEGAL"]
    available_index_labels = ["SEM_CORRECAO"] + correction_index_codes + factor_codes

    def _to_date(value):
        return parse_ui_date(value)

    def _normalize_rule_seed(seed: pd.DataFrame, earliest_origin: date) -> pd.DataFrame:
        df = seed.copy()
        if "Vigência a partir de" not in df.columns and "Início" in df.columns:
            df = df.rename(columns={"Início": "Vigência a partir de"})
        if "Fim" in df.columns:
            df = df.drop(columns=["Fim"])
        if df.empty:
            return df
        if "Vigência a partir de" not in df.columns:
            df["Vigência a partir de"] = earliest_origin
        if "Início dos juros" not in df.columns:
            df["Início dos juros"] = None
        if "Aplicação dos juros" not in df.columns:
            df["Aplicação dos juros"] = "Acumular percentual e aplicar no final"
        df["Vigência a partir de"] = df["Vigência a partir de"].map(
            lambda x: None if x is None or pd.isna(x) else _to_date(x)
        )
        df["Início dos juros"] = df["Início dos juros"].map(
            lambda x: None if x is None or pd.isna(x) else _to_date(x)
        )
        valid = df["Vigência a partir de"].dropna()
        if not valid.empty:
            first_idx = valid.index[0]
            df.loc[first_idx, "Vigência a partir de"] = earliest_origin
        return df

    def _effective_rule_rows(df: pd.DataFrame, earliest_origin: date, final_date: date):
        clean = df.dropna(subset=["Ordem", "Vigência a partir de"]).copy()
        if clean.empty:
            return []
        clean["Vigência a partir de"] = clean["Vigência a partir de"].map(_to_date)
        clean = clean.sort_values(["Vigência a partir de", "Ordem"], kind="stable").reset_index(drop=True)
        clean.loc[0, "Vigência a partir de"] = earliest_origin
        starts = clean["Vigência a partir de"].tolist()
        if len(starts) != len(set(starts)):
            raise ValueError("Não pode haver duas regras iniciando na mesma data. Use uma única regra por mudança de critério.")
        rows = []
        for idx, row in clean.iterrows():
            start_date = max(_to_date(row["Vigência a partir de"]), earliest_origin)
            if start_date > final_date:
                continue
            end_date = final_date
            if idx + 1 < len(clean):
                end_date = min(final_date, _to_date(clean.loc[idx + 1, "Vigência a partir de"]) - timedelta(days=1))
            if end_date < start_date:
                continue
            rows.append((row, start_date, end_date))
        return rows

    step_header(1, "Valores e período da atualização", "A data de origem define quando cada valor começa a evoluir. A data-base é o destino final do cálculo.")
    callout(
        "Como as datas funcionam",
        "Você informa a data de origem apenas junto ao valor. Na tabela de regras, a data significa somente 'a partir de quando este critério passa a valer'. O término de cada regra é calculado automaticamente pela próxima mudança ou pela data-base.",
        "info",
    )
    mode = st.radio("Forma de entrada", ["Saldo consolidado", "Por prestação"], horizontal=True, key="update_mode")
    base_date = st.date_input(
        "Data-base da atualização", value=st.session_state.get("update_base_date", date.today()),
        min_value=MIN_CREDIT_DATE, max_value=date.today(), format="DD/MM/YYYY", key="update_base_date",
    )

    if mode == "Saldo consolidado":
        a1, a2, a3 = st.columns([1.1, 1, 1.9])
        with a1:
            consolidated_amount = st.number_input(
                "Saldo/valor de origem (R$)", min_value=0.01, step=100.0,
                value=float(st.session_state.get("update_consolidated_amount", 10000.0)), key="update_consolidated_amount",
            )
        with a2:
            consolidated_date = st.date_input(
                "Data de origem", value=st.session_state.get("update_consolidated_date", date(2024, 1, 1)),
                min_value=MIN_CREDIT_DATE, max_value=base_date, format="DD/MM/YYYY", key="update_consolidated_date",
                help="O primeiro dia de evolução é o dia seguinte à data de origem.",
            )
        with a3:
            consolidated_description = st.text_input(
                "Descrição (opcional)", key="update_consolidated_description", placeholder="Ex.: saldo devedor apurado na sentença",
            )
        value_df = pd.DataFrame([{
            "Prestação": 1, "Data de origem": consolidated_date, "Valor": float(consolidated_amount), "Descrição": consolidated_description,
        }])
    else:
        up1, up2 = st.columns([2, 1])
        with up1:
            values_upload = st.file_uploader(
                "Importar valores/prestações (opcional)", type=["xlsx", "csv"], key="update_values_upload",
                help="Colunas esperadas: Prestação, Data de origem, Valor e Descrição (opcional).",
            )
        with up2:
            st.download_button(
                "Baixar modelo de importação", data=build_import_template(), file_name="modelo_atualizacao_monetaria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
            )
        if values_upload is not None:
            try:
                imported = pd.read_csv(values_upload) if values_upload.name.lower().endswith(".csv") else pd.read_excel(values_upload, sheet_name="Valores")
                rename_map = {}
                for col in imported.columns:
                    norm = str(col).strip().lower()
                    if norm in {"prestação", "prestacao", "nº", "numero", "número"}: rename_map[col] = "Prestação"
                    elif norm in {"data", "data de origem", "vencimento", "data origem"}: rename_map[col] = "Data de origem"
                    elif norm in {"valor", "valor original", "prestação (r$)", "prestacao (r$)"}: rename_map[col] = "Valor"
                    elif norm in {"descrição", "descricao", "item"}: rename_map[col] = "Descrição"
                imported = imported.rename(columns=rename_map)
                if not {"Prestação", "Data de origem", "Valor"}.issubset(imported.columns):
                    raise ValueError("A planilha precisa conter as colunas Prestação, Data de origem e Valor.")
                if "Descrição" not in imported.columns: imported["Descrição"] = ""
                imported = imported[["Prestação", "Data de origem", "Valor", "Descrição"]]
                imported["Data de origem"] = imported["Data de origem"].map(
                    lambda x: None if x is None or pd.isna(x) else _to_date(x)
                )
                st.session_state["update_values_editor_seed"] = imported
            except Exception as exc:
                st.error(f"Não foi possível ler os valores importados: {exc}")
        if "update_values_editor_seed" not in st.session_state:
            st.session_state["update_values_editor_seed"] = pd.DataFrame([
                {"Prestação": 1, "Data de origem": date(2024, 1, 20), "Valor": 1000.0, "Descrição": ""},
                {"Prestação": 2, "Data de origem": date(2024, 2, 20), "Valor": 1000.0, "Descrição": ""},
            ])
        value_df = st.data_editor(
            st.session_state["update_values_editor_seed"], num_rows="dynamic", use_container_width=True, hide_index=True,
            key="update_values_editor",
            column_config={
                "Prestação": st.column_config.NumberColumn(min_value=1, step=1, required=True),
                "Data de origem": st.column_config.DateColumn(format="DD/MM/YYYY", required=True),
                "Valor": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.01, required=True),
                "Descrição": st.column_config.TextColumn(),
            },
        )

    origin_dates = [
        _to_date(x) for x in value_df.get("Data de origem", [])
        if x is not None and not pd.isna(x)
    ]
    earliest_origin = min(origin_dates or [MIN_CREDIT_DATE])

    # ------------------------------------------------------------------
    # Metodologia: perfil judicial como PRESET, nunca como modo bloqueante.
    # ------------------------------------------------------------------
    def _default_update_rule_seed() -> pd.DataFrame:
        return pd.DataFrame([{
            "Ordem": 1,
            "Vigência a partir de": earliest_origin,
            "Descrição": "Regra principal",
            "Índice": "INPC" if "INPC" in index_codes else "SEM_CORRECAO",
            "Correção": "Pró-rata mensal por dias corridos" if "INPC" in index_codes else "Sem correção",
            "Defasagem (meses)": 0,
            "Juros (% a.m.)": 0.0,
            "Início dos juros": None,
            "Método dos juros": "Sem juros",
            "Aplicação dos juros": "Acumular percentual e aplicar no final",
            "Convenção de contagem": yearfrac_basis_rev[0],
            "Base dos juros": "Principal corrigido",
            "Ordem de incidência": "Correção → juros",
            "Competência": "Dia 1 (mês civil)",
        }])

    def _profile_rules_dataframe(profile_code: str) -> pd.DataFrame:
        seeded_rules = build_profile_rules(profile_code, earliest_origin, base_date)
        df = pd.DataFrame([{
            "Ordem": r.order,
            "Vigência a partir de": max(r.start_date, earliest_origin),
            "Descrição": r.description,
            "Índice": r.correction_index_code,
            "Correção": correction_labels_rev[r.correction_method],
            "Defasagem (meses)": r.index_lag_months,
            "Juros (% a.m.)": float(r.monthly_interest_rate * Decimal("100")),
            "Início dos juros": r.interest_start_date,
            "Método dos juros": interest_labels_rev[r.interest_method],
            "Aplicação dos juros": (
                "Acumular percentual e aplicar no final"
                if r.interest_method == INTEREST_NONE
                else interest_application_rev.get(r.interest_application, "Acumular percentual e aplicar no final")
            ),
            "Convenção de contagem": yearfrac_basis_rev.get(r.yearfrac_basis, yearfrac_basis_rev[0]),
            "Base dos juros": interest_base_rev[r.interest_base],
            "Ordem de incidência": order_labels_rev[r.event_order],
            "Competência": competence_rev[r.competence_start_day],
        } for r in seeded_rules])
        if not df.empty:
            df.loc[df.index[0], "Vigência a partir de"] = earliest_origin
        return df

    if "update_rules_seed" not in st.session_state:
        st.session_state["update_rules_seed"] = _default_update_rule_seed()
    st.session_state["update_rules_seed"] = _normalize_rule_seed(st.session_state["update_rules_seed"], earliest_origin)
    for col, default in {
        "Convenção de contagem": yearfrac_basis_rev[0],
        "Aplicação dos juros": "Acumular percentual e aplicar no final",
    }.items():
        if col not in st.session_state["update_rules_seed"].columns:
            st.session_state["update_rules_seed"][col] = default

    st.session_state.setdefault("update_profile", "PERSONALIZADO")
    st.session_state.setdefault("update_profile_origin", None)
    st.session_state.setdefault(
        "update_setup_mode",
        "Perfil judicial" if st.session_state.get("update_profile") != "PERSONALIZADO" else "Configuração livre",
    )
    if st.session_state.get("update_pending_setup_mode"):
        st.session_state["update_setup_mode"] = st.session_state.pop("update_pending_setup_mode")

    step_header(
        2,
        "Metodologia do cálculo",
        "Comece livremente ou use um perfil judicial como ponto de partida. O perfil apenas preenche parâmetros: você pode sair dele sem perder o que já foi configurado.",
    )
    setup_mode = st.radio(
        "Como deseja definir os critérios?",
        ["Configuração livre", "Perfil judicial"],
        horizontal=True,
        key="update_setup_mode",
    )

    active_code = st.session_state.get("update_profile", "PERSONALIZADO")
    # Ao sair de um perfil para a configuração livre, preservamos tudo e apenas
    # removemos o vínculo com o preset judicial.
    if setup_mode == "Configuração livre" and active_code != "PERSONALIZADO":
        st.session_state["update_profile_origin"] = active_code
        st.session_state["update_profile"] = "PERSONALIZADO"
        active_code = "PERSONALIZADO"

    judicial_profiles = {code: p for code, p in profiles.items() if code != "PERSONALIZADO"}
    if setup_mode == "Perfil judicial":
        tribunals = sorted({p.tribunal for p in judicial_profiles.values()})
        current_profile = profiles.get(st.session_state.get("update_profile", ""))
        default_tribunal = current_profile.tribunal if current_profile and current_profile.code != "PERSONALIZADO" else tribunals[0]
        if "update_tribunal_choice" not in st.session_state or st.session_state["update_tribunal_choice"] not in tribunals:
            st.session_state["update_tribunal_choice"] = default_tribunal

        pcol1, pcol2 = st.columns([1, 2])
        with pcol1:
            selected_tribunal = st.selectbox("Tribunal", tribunals, key="update_tribunal_choice")
        criteria_codes = [code for code, p in judicial_profiles.items() if p.tribunal == selected_tribunal]
        with pcol2:
            selected_profile = st.selectbox(
                "Critério",
                criteria_codes,
                format_func=lambda code: judicial_profiles[code].title,
                key=f"update_profile_choice_v2_{selected_tribunal}",
            )

        selected_def = judicial_profiles[selected_profile]
        coverage = ""
        if selected_def.factor_code and selected_def.factor_code in factor_series:
            fs = factor_series[selected_def.factor_code]
            coverage = f" · cobertura {fs.coverage_start} a {fs.coverage_end}"
        with st.container(border=True):
            st.markdown(f"**{selected_def.tribunal} · {selected_def.title}**")
            st.caption(selected_def.description)
            st.caption(f"{selected_def.status}{coverage} · Fonte: {selected_def.source_label}")
            if selected_def.note:
                st.caption(selected_def.note)
            b1, b2 = st.columns([1.1, 2.9])
            with b1:
                apply_profile = st.button("Usar este perfil", type="primary", use_container_width=True)
            with b2:
                if selected_def.source_url.startswith("http"):
                    st.markdown(f"[Abrir fonte oficial do {selected_def.tribunal}]({selected_def.source_url})")

        if apply_profile:
            try:
                seeded = _profile_rules_dataframe(selected_profile)
                st.session_state["update_rules_seed"] = seeded
                st.session_state["update_profile"] = selected_profile
                st.session_state["update_profile_origin"] = selected_profile
                st.session_state["update_rule_change_mode"] = (
                    "Sim — há mudança de critério no período" if len(seeded) > 1 else "Não — uma regra para todo o período"
                )
                st.session_state["update_rules_revision"] = st.session_state.get("update_rules_revision", 0) + 1
                for key in ("update_result", "update_excel", "update_pdf", "update_pdf_name"):
                    st.session_state.pop(key, None)
                st.rerun()
            except Exception as exc:
                st.error(str(exc))

    active_code = st.session_state.get("update_profile", "PERSONALIZADO")
    active_profile = profiles.get(active_code, profiles["PERSONALIZADO"])
    origin_code = st.session_state.get("update_profile_origin")
    origin_profile = profiles.get(origin_code) if origin_code else None

    if active_code != "PERSONALIZADO":
        callout(
            "Perfil judicial aplicado",
            f"{active_profile.tribunal} — {active_profile.title}. Ele é apenas um preset. Se você alterar qualquer parâmetro, o cálculo passa automaticamente a ser uma configuração personalizada baseada neste perfil.",
            "success",
        )
        ac1, ac2 = st.columns([1, 1])
        with ac1:
            if st.button("Editar livremente sem perder os parâmetros", use_container_width=True):
                st.session_state["update_profile_origin"] = active_code
                st.session_state["update_profile"] = "PERSONALIZADO"
                st.session_state["update_pending_setup_mode"] = "Configuração livre"
                st.rerun()
        with ac2:
            if st.button("Restaurar padrão deste perfil", use_container_width=True):
                st.session_state["update_rules_seed"] = _profile_rules_dataframe(active_code)
                st.session_state["update_rules_revision"] = st.session_state.get("update_rules_revision", 0) + 1
                st.rerun()
    elif origin_profile is not None:
        callout(
            "Configuração personalizada",
            f"Os parâmetros atuais foram originados de {origin_profile.tribunal} — {origin_profile.title}, mas agora estão livres para edição.",
            "info",
        )
    else:
        callout("Configuração livre", "Nenhum perfil judicial está vinculando o cálculo. Você define diretamente correção, juros e eventuais mudanças de critério.", "info")

    if setup_mode == "Configuração livre":
        if st.button("Começar uma configuração em branco", use_container_width=False):
            st.session_state["update_rules_seed"] = _default_update_rule_seed()
            st.session_state["update_profile"] = "PERSONALIZADO"
            st.session_state["update_profile_origin"] = None
            st.session_state["update_rule_change_mode"] = "Não — uma regra para todo o período"
            st.session_state["update_rules_revision"] = st.session_state.get("update_rules_revision", 0) + 1
            for key in ("update_result", "update_excel", "update_pdf", "update_pdf_name"):
                st.session_state.pop(key, None)
            st.rerun()

    seed_df = _normalize_rule_seed(st.session_state["update_rules_seed"], earliest_origin)
    if len(seed_df) > 1 and "update_rule_change_mode" not in st.session_state:
        st.session_state["update_rule_change_mode"] = "Sim — há mudança de critério no período"
    st.session_state.setdefault("update_rule_change_mode", "Não — uma regra para todo o período")

    change_mode = st.radio(
        "O critério muda durante o período?",
        ["Não — uma regra para todo o período", "Sim — há mudança de critério no período"],
        horizontal=True,
        key="update_rule_change_mode",
        help="Na maioria dos cálculos, escolha Não. Use Sim somente quando índice, juros ou outra regra mudarem em uma data específica.",
    )
    multi_rule_mode = change_mode.startswith("Sim")
    if not multi_rule_mode and len(seed_df) > 1:
        st.warning("Os parâmetros atuais contêm mais de uma regra. Para trabalhar em modo simples, primeiro consolide o cálculo em uma única regra.")
        if st.button("Usar apenas a primeira regra e continuar no modo simples"):
            st.session_state["update_rules_seed"] = seed_df.iloc[[0]].copy()
            st.session_state["update_rules_seed"].loc[st.session_state["update_rules_seed"].index[0], "Vigência a partir de"] = earliest_origin
            st.session_state["update_rules_revision"] = st.session_state.get("update_rules_revision", 0) + 1
            st.rerun()
        multi_rule_mode = True

    rev = st.session_state.get("update_rules_revision", 0)

    if not multi_rule_mode:
        row = seed_df.iloc[0].copy()
        original_row = row.copy()

        step_header(3, "Correção monetária", "Defina como o valor de origem será atualizado até a data-base. Nesta etapa não há incidência monetária de juros.")
        with st.container(border=True):
            c1, c2 = st.columns([1.45, 1.15])
            with c1:
                current_corr = str(row.get("Correção") or "Sem correção")
                correction_choice = st.selectbox(
                    "Forma de aplicar a correção monetária",
                    list(correction_labels.keys()),
                    index=list(correction_labels.keys()).index(current_corr) if current_corr in correction_labels else 0,
                    key=f"update_simple_correction_{rev}",
                    help="O índice define QUAL série será usada; esta opção define COMO a variação daquela série será reconhecida no cálculo.",
                )
            with c2:
                current_index = str(row.get("Índice") or "SEM_CORRECAO")
                index_choice = st.selectbox(
                    "Índice / tabela",
                    available_index_labels,
                    index=available_index_labels.index(current_index) if current_index in available_index_labels else 0,
                    key=f"update_simple_index_{rev}",
                )

            c3, c4 = st.columns(2)
            with c3:
                lag_months = st.number_input(
                    "Defasagem do índice (meses)", min_value=0, max_value=120, step=1,
                    value=int(row.get("Defasagem (meses)") or 0), key=f"update_simple_lag_{rev}",
                    disabled=(correction_choice in {"Sem correção", "Tabela prática — razão entre fatores mensais"}),
                )
            with c4:
                current_comp = str(row.get("Competência") or "Dia 1 (mês civil)")
                competence_choice = st.selectbox(
                    "Competência",
                    list(competence_labels.keys()),
                    index=list(competence_labels.keys()).index(current_comp) if current_comp in competence_labels else 0,
                    key=f"update_simple_competence_{rev}",
                    disabled=(correction_choice == "Pró-rata mensal por dias corridos"),
                    help="No pró-rata mensal por dias corridos, o Motor usa obrigatoriamente o mês civil (1º ao último dia do mês).",
                )
                if correction_choice == "Pró-rata mensal por dias corridos":
                    competence_choice = "Dia 1 (mês civil)"

            if correction_choice == "Tabela prática — razão entre fatores mensais":
                st.caption("O sistema utiliza a razão entre os fatores oficiais dos meses inicial e final. Use esta opção somente para tabelas judiciais de fatores, como TJSP, TJRJ e TJMG.")
            elif correction_choice == "Sem correção":
                st.caption("Nenhuma correção monetária será gerada neste período.")
            elif correction_choice == "Pró-rata mensal por dias corridos":
                st.caption("Critério da planilha operacional: para cada mês, a cotação mensal é transformada em fator equivalente pelo número de dias corridos efetivamente abrangidos. Em uma fração de mês, fator = (1 + índice mensal)^(dias utilizados / dias do mês). A data inicial é incluída no período.")
            elif correction_choice == "Mensal integral — fechamento da competência":
                st.caption("A variação mensal publicada é reconhecida integralmente no fechamento de cada competência. É a forma mais simples para índices mensais como INPC/IPCA quando não houver determinação de pró-rata diário.")
            else:
                st.caption("A variação mensal é convertida em taxa diária equivalente e distribuída ao longo da competência, conforme dias úteis ou dias corridos.")

            with st.expander("Entenda a forma de aplicar a correção monetária", expanded=False):
                st.markdown("""
**Índice / tabela** responde *qual série será usada*. **Forma de aplicar a correção** responde *como essa série entra no cálculo*.

| Forma | Quando usar |
|---|---|
| **Pró-rata mensal por dias corridos** | Reproduz o critério da planilha operacional: em cada mês usa `(1 + índice mensal)^(dias abrangidos / dias do mês)`. Meses completos recebem a cotação integral; meses parciais recebem somente a fração equivalente. |
| **Mensal integral — fechamento da competência** | Aplica o percentual mensal completo quando a competência fecha. É a opção mais intuitiva para INPC/IPCA mensais quando não há regra de pró-rata. |
| **Diária equivalente — dias corridos** | Converte a taxa mensal em equivalente diário e distribui por todos os dias da competência. |
| **Diária equivalente — dias úteis** | Igual ao anterior, mas a incidência ocorre somente nos dias úteis. |
| **Tabela prática — razão entre fatores mensais** | Para tabelas judiciais de fatores. O motor usa a razão entre o fator inicial e o fator final. |
| **Sem correção** | Mantém o principal sem atualização monetária. |
""")
                st.caption("Ao escolher um perfil judicial, o Motor já pré-seleciona a forma de correção adequada ao perfil. Só altere se o título/decisão exigir outro critério.")

        step_header(4, "Juros", "Escolha como o tempo vira percentual de juros e, separadamente, quando esse percentual será convertido em valor monetário.")
        with st.container(border=True):
            current_interest_method = str(row.get("Método dos juros") or "Sem juros")
            interest_enabled_default = current_interest_method != "Sem juros"
            apply_interest = st.radio(
                "Aplicar juros?", ["Não", "Sim"], horizontal=True,
                index=1 if interest_enabled_default else 0,
                key=f"update_simple_apply_interest_{rev}",
            )

            interest_method_choice = "Sem juros"
            application_choice = "Acumular percentual e aplicar no final"
            interest_rate_pct = 0.0
            interest_start_value = None
            yearfrac_basis_choice = yearfrac_basis_rev[0]
            interest_base_choice = str(row.get("Base dos juros") or "Principal corrigido")
            event_order_choice = "Correção → juros"

            if apply_interest == "Sim":
                interest_options = [x for x in interest_labels.keys() if x != "Sem juros"]
                default_method = current_interest_method if current_interest_method in interest_options else "Simples — fração de ano (convenção de dias)"
                i1, i2 = st.columns([1.7, 1])
                with i1:
                    interest_method_choice = st.selectbox(
                        "Método dos juros", interest_options,
                        index=interest_options.index(default_method), key=f"update_simple_interest_method_{rev}",
                    )
                with i2:
                    if interest_labels[interest_method_choice] == INTEREST_LEGAL_BCB:
                        interest_rate_pct = 0.0
                        st.text_input("Taxa de juros", value="Automática · série BCB", disabled=True, key=f"update_simple_interest_rate_legal_{rev}")
                    else:
                        interest_rate_pct = st.number_input(
                            "Taxa de juros (% a.m.)", min_value=0.0, step=0.01, format="%.6f",
                            value=float(row.get("Juros (% a.m.)") or 0), key=f"update_simple_interest_rate_{rev}",
                        )

                i3, i4 = st.columns(2)
                with i3:
                    raw_start = row.get("Início dos juros")
                    default_start = earliest_origin if raw_start is None or pd.isna(raw_start) else _to_date(raw_start)
                    interest_start_value = st.date_input(
                        "Início dos juros", value=default_start,
                        min_value=earliest_origin, max_value=base_date, format="DD/MM/YYYY",
                        key=f"update_simple_interest_start_{rev}",
                    )
                with i4:
                    interest_base_choice = st.selectbox(
                        "Base de incidência", list(interest_base_labels.keys()),
                        index=list(interest_base_labels.keys()).index(interest_base_choice) if interest_base_choice in interest_base_labels else 1,
                        key=f"update_simple_interest_base_{rev}",
                    )

                forced_end = interest_labels[interest_method_choice] in {INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC, INTEREST_LEGAL_BCB}
                forced_daily = interest_labels[interest_method_choice] == INTEREST_COMPOUND_EQUIV
                if forced_end:
                    application_choice = "Acumular percentual e aplicar no final"
                    callout(
                        "Aplicação dos juros",
                        "Este método calcula primeiro o percentual total do período e aplica os juros uma única vez no fechamento, sobre a base selecionada já apurada.",
                        "info",
                    )
                elif forced_daily:
                    application_choice = "Aplicar juros diariamente"
                    callout(
                        "Aplicação dos juros",
                        "Juros compostos exigem aplicação diária: os juros acumulados passam a integrar a base do dia seguinte.",
                        "info",
                    )
                else:
                    current_application = str(row.get("Aplicação dos juros") or "Acumular percentual e aplicar no final")
                    application_choice = st.radio(
                        "Forma de aplicação dos juros",
                        ["Acumular percentual e aplicar no final", "Aplicar juros diariamente"],
                        horizontal=True,
                        index=0 if current_application != "Aplicar juros diariamente" else 1,
                        key=f"update_simple_interest_application_{rev}",
                        help="No fechamento, o motor soma as taxas simples do período e aplica o percentual uma única vez. Na aplicação diária, calcula o valor monetário dos juros em cada dia sobre a base daquele dia.",
                    )
                    if application_choice == "Acumular percentual e aplicar no final":
                        callout(
                            "Modo recomendado para a maioria dos cálculos simples",
                            "O motor acumula o percentual de juros durante o período, termina primeiro a correção monetária e só então aplica o percentual total sobre o saldo corrigido. Assim, correção e juros permanecem matematicamente separados.",
                            "success",
                        )
                    else:
                        callout(
                            "Aplicação diária",
                            "O valor dos juros é reconhecido dia a dia. Se a base corrigida variar durante o período, cada dia pode ter uma base monetária diferente.",
                            "warning",
                        )

                if interest_labels[interest_method_choice] == INTEREST_LEGAL_BCB:
                    callout(
                        "Taxa Legal do Código Civil",
                        "O percentual é obtido automaticamente da série oficial do Banco Central (SGS 29543). O regime é de juros simples: as taxas mensais são somadas e frações de mês são apropriadas por dias corridos. Quando houver correção monetária, os juros legais são aplicados sobre o valor já corrigido.",
                        "info",
                    )
                    if interest_start_value and interest_start_value < date(2024, 8, 30):
                        st.warning("A Taxa Legal passou a vigorar em 30/08/2024. Para um período anterior, use uma mudança de critério e defina a regra precedente separadamente.")
                elif interest_labels[interest_method_choice] == INTEREST_SIMPLE_YEARFRAC:
                    current_basis = str(row.get("Convenção de contagem") or yearfrac_basis_rev[0])
                    yearfrac_basis_choice = st.selectbox(
                        "Convenção de contagem",
                        list(yearfrac_basis_labels.keys()),
                        index=list(yearfrac_basis_labels.keys()).index(current_basis) if current_basis in yearfrac_basis_labels else 0,
                        key=f"update_simple_yearfrac_basis_{rev}",
                    )
                    if interest_start_value:
                        fraction = excel_yearfrac(interest_start_value, base_date, yearfrac_basis_labels[yearfrac_basis_choice])
                        accumulated = fraction * Decimal(str(interest_rate_pct)) / Decimal("100") * Decimal("12")
                        st.caption(
                            f"Prévia: Fração do ano = {number_br(fraction, 6)} · meses equivalentes = {number_br(fraction * Decimal('12'), 4)} · percentual acumulado = {percent_br(accumulated, 4)}."
                        )
                elif interest_labels[interest_method_choice] == INTEREST_SIMPLE_360_TJRJ:
                    st.caption("Este método usa a convenção comercial validada no calculador cível do TJRJ e sempre fecha os juros no fim do período.")
                elif interest_labels[interest_method_choice] == INTEREST_SIMPLE_30:
                    st.caption("Taxa mensal ÷ 30. Você pode apenas acumular as taxas diárias e aplicá-las no final, ou reconhecer o valor monetário diariamente.")
                elif interest_labels[interest_method_choice] == INTEREST_SIMPLE_COMPETENCE:
                    st.caption("A taxa mensal é distribuída pelos dias reais de cada competência. O percentual pode ser acumulado para aplicação final ou convertido em juros diariamente.")

                if application_choice == "Aplicar juros diariamente" and interest_labels[interest_method_choice] != INTEREST_COMPOUND_EQUIV:
                    current_order = str(row.get("Ordem de incidência") or "Correção → juros")
                    event_order_choice = st.selectbox(
                        "Ordem no dia", list(order_labels.keys()),
                        index=list(order_labels.keys()).index(current_order) if current_order in order_labels else 0,
                        key=f"update_simple_event_order_{rev}",
                        help="Só interfere quando os juros são efetivamente aplicados dia a dia.",
                    )

            with st.expander("Entenda os métodos de juros", expanded=False):
                st.markdown(
                    "| Método | Como o percentual é formado | Aplicação |\n"
                    "|---|---|---|\n"
                    "| **Simples diário — taxa mensal ÷ 30** | Taxa mensal / 30 por dia elegível. | Pode acumular o percentual e aplicar no final **ou** aplicar juros diariamente. |\n"
                    "| **Simples diário — dias da competência** | Divide a taxa mensal por 28/29/30/31, conforme a competência. | Pode acumular e aplicar no final **ou** aplicar diariamente. |\n"
                    "| **TJRJ — 360 dias** | Taxa anual × dias comerciais / 360. | Sempre no fechamento. |\n"
                    "| **Fração de ano** | Taxa anual nominal × fração de ano segundo a convenção de contagem escolhida. | Sempre no fechamento. |\n"
                    "| **Taxa Legal — BCB** | Série oficial mensal; juros simples; soma das taxas e pro rata por dias corridos nas frações. | Sempre no fechamento, sobre a base corrigida quando houver correção. |\n"
                    "| **Composto diário equivalente** | Converte a taxa mensal em equivalente diário. | Sempre diariamente, com capitalização. |"
                )
                st.caption(
                    "Regra prática: se você só precisa encontrar o percentual de correção, atualizar o saldo e depois encontrar o percentual de juros para aplicá-lo sobre esse saldo corrigido, use 'Acumular percentual e aplicar no final'."
                )

        simple_rule = {
            "Ordem": 1,
            "Vigência a partir de": earliest_origin,
            "Descrição": str(row.get("Descrição") or "Regra principal"),
            "Índice": index_choice,
            "Correção": correction_choice,
            "Defasagem (meses)": int(lag_months),
            "Juros (% a.m.)": float(interest_rate_pct if apply_interest == "Sim" else 0.0),
            "Início dos juros": interest_start_value if apply_interest == "Sim" else None,
            "Método dos juros": interest_method_choice if apply_interest == "Sim" else "Sem juros",
            "Aplicação dos juros": application_choice,
            "Convenção de contagem": yearfrac_basis_choice,
            "Base dos juros": interest_base_choice,
            "Ordem de incidência": event_order_choice,
            "Competência": competence_choice,
        }
        rules_df = pd.DataFrame([simple_rule])
        st.session_state["update_rules_seed"] = rules_df.copy()

        # Se um perfil estava ativo e qualquer parâmetro foi alterado, o perfil
        # vira apenas a origem da configuração; nenhuma informação é perdida.
        active_now = st.session_state.get("update_profile", "PERSONALIZADO")
        if active_now != "PERSONALIZADO":
            compare_cols = [c for c in simple_rule if c in original_row.index]
            changed = any(str(simple_rule[c]) != str(original_row.get(c)) for c in compare_cols)
            if changed:
                st.session_state["update_profile_origin"] = active_now
                st.session_state["update_profile"] = "PERSONALIZADO"

        rule_timeline([{
            "period": f"{date_br(earliest_origin)} → {date_br(base_date)}",
            "title": "Regra única",
            "meta": f"{correction_choice} · {index_choice}" + (
                f" · {interest_method_choice} · {application_choice.lower()} · desde {date_br(interest_start_value)}"
                if apply_interest == "Sim" else " · sem juros"
            ),
        }])

    else:
        step_header(3, "Mudanças de critério — modo avançado", "Use este modo somente quando correção, juros ou outro parâmetro mudarem ao longo do período. O fim de cada regra é automático.")
        st.caption("Cada nova linha informa apenas a data em que o novo critério começa. A regra anterior termina no dia imediatamente anterior.")

        profile_for_interest = active_profile if active_code != "PERSONALIZADO" else origin_profile
        if profile_for_interest is not None and profile_for_interest.tribunal == "TJDFT":
            with st.container(border=True):
                st.markdown("**Juros do TJDFT — opcionais**")
                st.caption(
                    "O perfil TJDFT pré-configura a correção monetária, mas não impõe juros. "
                    "Se o título judicial determinar juros, configure-os aqui; o Motor aplica a mesma regra de juros ao longo das mudanças INPC/IPCA sem alterar a correção do perfil."
                )
                existing_interest = any(str(v or "Sem juros") != "Sem juros" for v in seed_df.get("Método dos juros", pd.Series(dtype=object)).tolist())
                tjdft_apply_interest = st.radio(
                    "Aplicar juros no cálculo?",
                    ["Não", "Sim"],
                    horizontal=True,
                    index=1 if existing_interest else 0,
                    key=f"tjdft_apply_interest_{rev}",
                )

                tjdft_method = "Sem juros"
                tjdft_rate = 0.0
                tjdft_start = earliest_origin
                tjdft_base = "Principal corrigido"
                tjdft_application = "Acumular percentual e aplicar no final"
                tjdft_yearfrac = yearfrac_basis_rev[0]

                if tjdft_apply_interest == "Sim":
                    existing_methods = [str(v) for v in seed_df.get("Método dos juros", pd.Series(dtype=object)).tolist() if str(v or "Sem juros") != "Sem juros"]
                    existing_method = existing_methods[0] if existing_methods else "Simples — fração de ano (convenção de dias)"
                    options = [x for x in interest_labels.keys() if x != "Sem juros"]
                    if existing_method not in options:
                        existing_method = "Simples — fração de ano (convenção de dias)"
                    cju1, cju2 = st.columns([1.7, 1.0])
                    with cju1:
                        tjdft_method = st.selectbox(
                            "Método dos juros",
                            options,
                            index=options.index(existing_method),
                            key=f"tjdft_interest_method_{rev}",
                        )
                    with cju2:
                        if interest_labels[tjdft_method] == INTEREST_LEGAL_BCB:
                            st.text_input("Taxa", value="Automática · série BCB", disabled=True, key=f"tjdft_interest_legal_{rev}")
                            tjdft_rate = 0.0
                        else:
                            existing_rates = [float(v or 0) for v in seed_df.get("Juros (% a.m.)", pd.Series(dtype=float)).tolist() if float(v or 0) != 0]
                            default_rate = existing_rates[0] if existing_rates else 1.0
                            tjdft_rate = st.number_input(
                                "Taxa de juros (% a.m.)", min_value=0.0, step=0.01, format="%.6f",
                                value=float(default_rate), key=f"tjdft_interest_rate_{rev}",
                            )
                    cju3, cju4 = st.columns(2)
                    with cju3:
                        tjdft_start = st.date_input(
                            "Início dos juros", value=earliest_origin, min_value=earliest_origin, max_value=base_date,
                            format="DD/MM/YYYY", key=f"tjdft_interest_start_{rev}",
                            help="Ex.: data da citação. Os juros não retroagem para antes desta data.",
                        )
                    with cju4:
                        tjdft_base = st.selectbox(
                            "Base de incidência", list(interest_base_labels.keys()), index=1,
                            key=f"tjdft_interest_base_{rev}",
                        )

                    if interest_labels[tjdft_method] in {INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC, INTEREST_LEGAL_BCB}:
                        tjdft_application = "Acumular percentual e aplicar no final"
                    elif interest_labels[tjdft_method] == INTEREST_COMPOUND_EQUIV:
                        tjdft_application = "Aplicar juros diariamente"
                    else:
                        tjdft_application = st.radio(
                            "Forma de aplicação dos juros",
                            ["Acumular percentual e aplicar no final", "Aplicar juros diariamente"],
                            horizontal=True, key=f"tjdft_interest_application_{rev}",
                        )
                    if interest_labels[tjdft_method] == INTEREST_SIMPLE_YEARFRAC:
                        tjdft_yearfrac = st.selectbox(
                            "Convenção de contagem", list(yearfrac_basis_labels.keys()),
                            index=0, key=f"tjdft_yearfrac_{rev}",
                        )

                if st.button("Aplicar configuração de juros ao perfil TJDFT", use_container_width=True, key=f"tjdft_apply_interest_btn_{rev}"):
                    updated = seed_df.copy()
                    for idx in updated.index:
                        row_start_raw = updated.at[idx, "Vigência a partir de"] if "Vigência a partir de" in updated.columns else earliest_origin
                        row_start = earliest_origin if row_start_raw is None or pd.isna(row_start_raw) else _to_date(row_start_raw)
                        if tjdft_apply_interest == "Não":
                            updated.at[idx, "Juros (% a.m.)"] = 0.0
                            updated.at[idx, "Início dos juros"] = None
                            updated.at[idx, "Método dos juros"] = "Sem juros"
                            updated.at[idx, "Aplicação dos juros"] = "Acumular percentual e aplicar no final"
                        else:
                            updated.at[idx, "Juros (% a.m.)"] = float(tjdft_rate)
                            updated.at[idx, "Início dos juros"] = max(tjdft_start, row_start)
                            updated.at[idx, "Método dos juros"] = tjdft_method
                            updated.at[idx, "Aplicação dos juros"] = tjdft_application
                            updated.at[idx, "Base dos juros"] = tjdft_base
                            updated.at[idx, "Convenção de contagem"] = tjdft_yearfrac
                    st.session_state["update_rules_seed"] = updated
                    st.session_state["update_rules_revision"] = st.session_state.get("update_rules_revision", 0) + 1
                    for key in ("update_result", "update_excel", "update_pdf", "update_pdf_name"):
                        st.session_state.pop(key, None)
                    st.rerun()

        rules_df = st.data_editor(
            seed_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key=f"update_rules_editor_{rev}",
            column_config={
                "Ordem": st.column_config.NumberColumn(min_value=1, step=1, width="small", required=True),
                "Vigência a partir de": st.column_config.DateColumn(format="DD/MM/YYYY", required=True, help="Na primeira regra, o sistema usa automaticamente a data de origem mais antiga."),
                "Descrição": st.column_config.TextColumn(width="medium"),
                "Índice": st.column_config.SelectboxColumn(options=available_index_labels, required=True),
                "Correção": st.column_config.SelectboxColumn("Forma de aplicar a correção", options=list(correction_labels.keys()), required=True, help="Define como a série mensal/fator é reconhecida no cálculo; o índice/tabela é escolhido na coluna ao lado."),
                "Defasagem (meses)": st.column_config.NumberColumn(min_value=0, max_value=120, step=1, width="small"),
                "Juros (% a.m.)": st.column_config.NumberColumn(format="%.6f", step=0.01),
                "Início dos juros": st.column_config.DateColumn(format="DD/MM/YYYY", help="Opcional. Se vazio, os juros começam junto com a vigência da regra."),
                "Método dos juros": st.column_config.SelectboxColumn(options=list(interest_labels.keys()), required=True),
                "Aplicação dos juros": st.column_config.SelectboxColumn(
                    options=list(interest_application_labels.keys()), required=True,
                    help="Para fração de ano e TJRJ o fechamento é obrigatório; para juros compostos a aplicação diária é obrigatória."
                ),
                "Convenção de contagem": st.column_config.SelectboxColumn(
                    options=list(yearfrac_basis_labels.keys()), required=True,
                    help="Usada somente no método fração de ano. Base 0 corresponde à convenção US/NASD 30/360 adotada como padrão quando a base não é especificada."
                ),
                "Base dos juros": st.column_config.SelectboxColumn(options=list(interest_base_labels.keys()), required=True),
                "Ordem de incidência": st.column_config.SelectboxColumn(options=list(order_labels.keys()), required=True),
                "Competência": st.column_config.SelectboxColumn(options=list(competence_labels.keys()), required=True),
            },
        )
        st.session_state["update_rules_seed"] = rules_df.copy()
        if active_code != "PERSONALIZADO" and not rules_df.equals(seed_df):
            st.session_state["update_profile_origin"] = active_code
            st.session_state["update_profile"] = "PERSONALIZADO"

        try:
            preview_rows = _effective_rule_rows(rules_df, earliest_origin, base_date)
            timeline_rows = []
            for row_adv, start_date, end_date in preview_rows:
                corr = str(row_adv.get("Correção") or "Sem correção")
                idx_code = str(row_adv.get("Índice") or "SEM_CORRECAO")
                juros = str(row_adv.get("Método dos juros") or "Sem juros")
                rate = float(row_adv.get("Juros (% a.m.)") or 0)
                application = str(row_adv.get("Aplicação dos juros") or "Acumular percentual e aplicar no final")
                interest_start_raw = row_adv.get("Início dos juros")
                interest_start_text = ""
                if juros != "Sem juros" and interest_start_raw is not None and not pd.isna(interest_start_raw):
                    interest_start_text = f" a partir de {date_br(_to_date(interest_start_raw))}"
                timeline_rows.append({
                    "period": f"{date_br(start_date)} → {date_br(end_date)}",
                    "title": str(row_adv.get("Descrição") or f"Regra {int(row_adv['Ordem'])}"),
                    "meta": f"{corr} · {idx_code}" + (
                        f" · juros {rate:.4f}% a.m. · {application.lower()}{interest_start_text}"
                        if juros != "Sem juros" else " · sem juros"
                    ),
                })
            rule_timeline(timeline_rows)
        except Exception as exc:
            st.warning(str(exc))

        with st.expander("Como funciona o modo avançado?", expanded=False):
            st.markdown(
                "**Correção monetária:** a coluna *Índice* informa qual série será usada; a coluna *Forma de aplicar a correção* informa como essa série entra no cálculo. No pró-rata mensal por dias corridos, meses parciais usam fator equivalente proporcional aos dias efetivamente abrangidos.  \n"
                "**Aplicação final dos juros:** acumula o percentual simples dentro de cada regra e o converte em valor somente no encerramento daquela regra, após a correção monetária.  \n"
                "**Aplicação diária dos juros:** converte a taxa em valor em cada dia, usando a base existente naquele momento.  \n"
                "**Fração de ano, TJRJ/360 e Taxa Legal:** são sempre métodos de fechamento.  \n"
                "**Composto diário equivalente:** é sempre diário, pois os juros precisam integrar a base subsequente."
            )

    step_header(5, "Multas, abatimentos e amortizações", "Cadastre acréscimos pontuais e pagamentos. Multas são aplicadas na data indicada; abatimentos permanecem proporcionais entre os componentes da obrigação.")


    st.markdown("**Multas / acréscimos percentuais**")
    st.caption("A multa é aplicada uma única vez na data informada. Prestação alvo = 0 aplica a todas as obrigações abertas.")
    if "update_penalties_seed" not in st.session_state:
        st.session_state["update_penalties_seed"] = pd.DataFrame(columns=["Data", "Multa (%)", "Base da multa", "Prestação alvo", "Observação"])
    penalties_df = st.data_editor(
        st.session_state["update_penalties_seed"], num_rows="dynamic", use_container_width=True, hide_index=True, key="update_penalties_editor",
        column_config={
            "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
            "Multa (%)": st.column_config.NumberColumn(format="%.4f", min_value=0.0, step=0.1),
            "Base da multa": st.column_config.SelectboxColumn(options=list(penalty_base_labels.keys()), default="Saldo antes da multa (principal + correção + juros)"),
            "Prestação alvo": st.column_config.NumberColumn(min_value=0, step=1, help="0 = aplicar a todas as prestações abertas"),
            "Observação": st.column_config.TextColumn(),
        },
    )

    st.markdown("**Abatimentos / amortizações**")
    st.caption("Informe a data e o valor do recurso/pagamento. O Motor corrige e calcula os encargos até a data do evento e somente depois reduz a obrigação.")

    # O antigo campo numérico "Prestação alvo" era tecnicamente correto, porém pouco
    # intuitivo. A interface passa a trabalhar com um destino legível e converte a
    # seleção para o número interno do item apenas no momento do cálculo.
    clean_targets = value_df.dropna(subset=["Prestação", "Data de origem", "Valor"]).copy()
    abatement_target_map: dict[str, int] = {}
    if mode == "Saldo consolidado":
        abatement_target_map["Saldo consolidado"] = 1
    else:
        abatement_target_map["Todas as prestações abertas — rateio proporcional"] = 0
        for _, target_row in clean_targets.iterrows():
            item_no = int(target_row["Prestação"])
            label = f"Prestação {item_no} — {date_br(_to_date(target_row['Data de origem']))} — {money_br(target_row['Valor'], 2)}"
            abatement_target_map[label] = item_no

    default_abatement_target = next(iter(abatement_target_map.keys()))
    if mode == "Saldo consolidado":
        callout(
            "Destino automático",
            "Como o cálculo possui um único saldo consolidado, todo abatimento será aplicado a esse saldo. Você não precisa informar número de prestação.",
            "info",
        )
    else:
        callout(
            "Escolha o destino do recurso",
            "Use 'Todas as prestações abertas — rateio proporcional' quando o recurso deve reduzir o conjunto da dívida. Se o pagamento estiver vinculado a uma prestação específica, selecione essa prestação na lista.",
            "info",
        )

    with st.expander("Como o Motor aplica o abatimento?", expanded=False):
        st.markdown(
            "1. O saldo é evoluído **até a data do abatimento** conforme correção, juros e multa configurados.  \n"
            "2. O abatimento é aplicado **depois dos encargos daquela data**.  \n"
            "3. Se o destino for todas as prestações, o valor é rateado proporcionalmente pelos saldos abertos de cada obrigação.  \n"
            "4. Dentro de cada obrigação, a redução é proporcional entre principal, correção, juros e multa já existentes.  \n"
            "5. Se o abatimento superar o saldo elegível, o excedente não é transformado em crédito: ele é informado como valor não aplicado."
        )

    if "update_abatements_seed" not in st.session_state:
        st.session_state["update_abatements_seed"] = pd.DataFrame(columns=["Data", "Valor", "Destino", "Observação"])
    else:
        legacy_ab = st.session_state["update_abatements_seed"].copy()
        if "Destino" not in legacy_ab.columns:
            legacy_ab["Destino"] = default_abatement_target
            if "Prestação alvo" in legacy_ab.columns:
                reverse_targets = {v: k for k, v in abatement_target_map.items()}
                legacy_ab["Destino"] = legacy_ab["Prestação alvo"].map(
                    lambda x: reverse_targets.get(int(x or 0), default_abatement_target) if not pd.isna(x) else default_abatement_target
                )
                legacy_ab = legacy_ab.drop(columns=["Prestação alvo"])
            st.session_state["update_abatements_seed"] = legacy_ab

    abatements_df = st.data_editor(
        st.session_state["update_abatements_seed"], num_rows="dynamic", use_container_width=True, hide_index=True, key="update_abatements_editor",
        column_config={
            "Data": st.column_config.DateColumn(format="DD/MM/YYYY", help="Data em que o recurso/pagamento deve reduzir a dívida."),
            "Valor": st.column_config.NumberColumn(format="R$ %.2f", min_value=0.01),
            "Destino": st.column_config.SelectboxColumn(
                options=list(abatement_target_map.keys()), required=True,
                help="Saldo consolidado, todas as prestações abertas com rateio proporcional ou uma prestação específica."
            ),
            "Observação": st.column_config.TextColumn(),
        },
    )

    step_header(6, "Precisão e execução", "Revise as opções e execute. O motor registrará a regra aplicada em cada dia da memória.")
    with st.expander("Precisão do cálculo", expanded=False):
        uc1, uc2 = st.columns(2)
        with uc1:
            update_round_daily = st.checkbox("Arredondar encargos diariamente", value=True, key="update_round_daily")
        with uc2:
            update_money_places = st.number_input("Casas dos valores", min_value=0, max_value=8, value=2, step=1, key="update_money_places")

    calculate_update = st.button("Calcular atualização", type="primary", use_container_width=True)
    if calculate_update:
        try:
            items = []
            clean_values = value_df.dropna(subset=["Prestação", "Data de origem", "Valor"]).copy()
            if clean_values.empty:
                raise ValueError("Informe ao menos um valor para atualização.")
            for _, row in clean_values.iterrows():
                origin = _to_date(row["Data de origem"])
                if origin >= base_date:
                    raise ValueError(f"A data de origem da prestação {int(row['Prestação'])} deve ser anterior à data-base.")
                items.append(UpdateItem(item_number=int(row["Prestação"]), origin_date=origin, amount=Decimal(str(row["Valor"])), description=str(row.get("Descrição", "") or "")))
            if len({x.item_number for x in items}) != len(items):
                raise ValueError("Não pode haver números de prestação/itens duplicados.")

            rules = []
            effective_rows = _effective_rule_rows(rules_df, earliest_origin, base_date)
            if not effective_rows:
                raise ValueError("Cadastre ao menos uma regra de atualização.")
            for row, start_date, end_date in effective_rows:
                index_code = str(row.get("Índice") or "SEM_CORRECAO")
                correction_label = str(row.get("Correção") or "Sem correção")
                correction_method = correction_labels[correction_label]
                if correction_method != CORRECTION_NONE and index_code == "SEM_CORRECAO":
                    raise ValueError(f"Regra {int(row['Ordem'])}: selecione um índice/fator para o método de correção informado.")
                if correction_method == CORRECTION_FACTOR_TABLE and index_code not in factor_codes:
                    raise ValueError(f"Regra {int(row['Ordem'])}: o método Tabela Prática exige uma tabela de fatores judiciais.")
                if index_code in factor_codes and correction_method != CORRECTION_FACTOR_TABLE:
                    raise ValueError(f"Regra {int(row['Ordem'])}: a tabela {index_code} deve usar o método Tabela Prática.")
                interest_method_code = interest_labels[str(row.get("Método dos juros") or "Sem juros")]
                interest_application_code = interest_application_labels[str(row.get("Aplicação dos juros") or "Acumular percentual e aplicar no final")]
                if interest_method_code in {INTEREST_SIMPLE_360_TJRJ, INTEREST_SIMPLE_YEARFRAC, INTEREST_LEGAL_BCB}:
                    interest_application_code = INTEREST_APPLICATION_END
                elif interest_method_code == INTEREST_COMPOUND_EQUIV:
                    interest_application_code = INTEREST_APPLICATION_DAILY
                rules.append(UpdateRule(
                    order=int(row["Ordem"]), start_date=start_date, end_date=end_date, description=str(row.get("Descrição", "") or ""),
                    correction_index_code=index_code, correction_method=correction_method,
                    index_lag_months=int(row.get("Defasagem (meses)") or 0),
                    monthly_interest_rate=Decimal(str(row.get("Juros (% a.m.)") or 0)) / Decimal("100"),
                    interest_start_date=(None if row.get("Início dos juros") is None or pd.isna(row.get("Início dos juros")) else _to_date(row.get("Início dos juros"))),
                    interest_method=interest_method_code,
                    interest_application=interest_application_code,
                    yearfrac_basis=yearfrac_basis_labels[str(row.get("Convenção de contagem") or yearfrac_basis_rev[0])],
                    interest_base=interest_base_labels[str(row.get("Base dos juros") or "Principal corrigido")],
                    event_order=order_labels[str(row.get("Ordem de incidência") or "Correção → juros")],
                    competence_start_day=(
                        1 if correction_method == CORRECTION_PRORATA_CALENDAR
                        else competence_labels[str(row.get("Competência") or "Dia 1 (mês civil)")]
                    ),
                ))

            penalties = []
            if not penalties_df.empty:
                for _, row in penalties_df.dropna(subset=["Data", "Multa (%)"]).iterrows():
                    rate = Decimal(str(row.get("Multa (%)") or 0)) / Decimal("100")
                    if rate <= 0:
                        continue
                    penalties.append(Penalty(
                        event_date=_to_date(row["Data"]), rate=rate,
                        base=penalty_base_labels[str(row.get("Base da multa") or "Saldo antes da multa (principal + correção + juros)")],
                        target_item=int(row.get("Prestação alvo") or 0), note=str(row.get("Observação", "") or ""),
                    ))

            abatements = []
            if not abatements_df.empty:
                for _, row in abatements_df.dropna(subset=["Data", "Valor"]).iterrows():
                    destination = str(row.get("Destino") or default_abatement_target)
                    target_item = abatement_target_map.get(destination)
                    if target_item is None:
                        raise ValueError(f"Destino de abatimento inválido: {destination}.")
                    abatements.append(Abatement(
                        event_date=_to_date(row["Data"]), amount=Decimal(str(row["Valor"])),
                        target_item=int(target_item), note=str(row.get("Observação", "") or ""),
                    ))

            result_update = calculate_monetary_update(
                items=items, rules=rules, base_date=base_date, indices=update_config.indices,
                holidays=update_config.holidays, abatements=abatements, penalties=penalties, factor_series=factor_series,
                round_daily_money=bool(update_round_daily), money_places=int(update_money_places),
            )
            active_label = "Critério personalizado"
            active_code = st.session_state.get("update_profile", "PERSONALIZADO")
            active_def = profiles.get(active_code)
            if active_def and active_code != "PERSONALIZADO":
                active_label = f"{active_def.tribunal} — {active_def.title}"
            elif st.session_state.get("update_profile_origin") in profiles:
                origin_def = profiles[st.session_state["update_profile_origin"]]
                if origin_def.code != "PERSONALIZADO":
                    active_label = f"Critério personalizado · origem {origin_def.tribunal} — {origin_def.title}"
            st.session_state["update_profile_label"] = active_label
            st.session_state["update_result"] = result_update
            st.session_state["update_excel"] = build_monetary_update_workbook(result_update, money_places=int(update_money_places), percent_places=4)
            st.session_state["update_money_places_used"] = int(update_money_places)
            st.session_state.pop("update_pdf", None); st.session_state.pop("update_pdf_name", None)
            st.success("Atualização concluída." if result_update.is_complete else "Atualização calculada até o último fator/índice disponível.")
        except Exception as exc:
            st.error(str(exc))

    result_update = st.session_state.get("update_result")
    if result_update is not None:
        places = int(st.session_state.get("update_money_places_used", 2))
        profile_label = st.session_state.get("update_profile_label", "Critério personalizado")
        step_header(7, "Resultado e documentos", "Consulte o resultado, copie um texto pronto para Word e gere a memória em Excel ou PDF.")
        correction_limit = (result_update.stopped_at - timedelta(days=1)) if (not result_update.is_complete and result_update.stopped_at) else result_update.base_date
        result_hero(
            label="Saldo atualizado na data-base" if result_update.is_complete else "Saldo apurado na data-base · correção parcial",
            value=money_br(result_update.updated_total, places),
            subtitle=(
                f"Data-base {date_br(result_update.base_date)} · {profile_label}"
                if result_update.is_complete
                else f"Correção até {date_br(correction_limit)} · juros, se aplicados, até {date_br(result_update.base_date)} · {profile_label}"
            ),
        )
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Valor original", money_br(result_update.original_total, places))
        m2.metric("Correção monetária", money_br(result_update.correction_total, places))
        m3.metric("Juros", money_br(result_update.interest_total, places))
        m4.metric("Multa", money_br(result_update.penalty_total, places))
        m5.metric("Abatimentos", money_br(result_update.abatements_total, places))

        if not result_update.is_complete:
            correction_limit = (result_update.stopped_at - timedelta(days=1)) if result_update.stopped_at else None
            st.warning(
                "Correção monetária parcial: "
                + (f"{result_update.missing_index_code} não possui índice/fator para {result_update.missing_index_reference}. " if result_update.missing_index_code else "há índice/fator ausente. ")
                + (f"A correção foi mantida somente até {date_br(correction_limit)}. " if correction_limit else "")
                + f"O Motor não projeta índice futuro; o saldo corrigido fica congelado a partir daí. "
                + f"Se juros foram selecionados, eles continuam sendo apurados normalmente até a data-base de {date_br(result_update.base_date)}, sobre a base disponível conforme a regra escolhida."
            )
        for warning in result_update.warnings:
            st.info(warning)

        # Demonstrativo visual inspirado na tela operacional atualmente utilizada.
        # Ele existe apenas na navegação do aplicativo e não é incorporado ao PDF.
        with st.container(border=True):
            st.markdown(f"### Atualização do débito em {date_br(result_update.base_date)}")
            st.caption("Visão de conferência durante a elaboração, inspirada no demonstrativo operacional atual. Este quadro não é incluído no PDF.")
            visual_rows = []
            for item in result_update.items:
                corrected_value = item.original_amount + item.correction_amount
                gross_before_abatements = corrected_value + item.interest_amount + item.penalty_amount
                correction_factor = (corrected_value / item.original_amount) if item.original_amount else Decimal("1")
                interest_pct = (item.interest_amount / corrected_value) if corrected_value else Decimal("0")
                visual_rows.append({
                    "Data do valor devido": date_br(item.origin_date),
                    "Valor devido": money_br(item.original_amount, places),
                    "Fator efetivo C.M.": number_br(correction_factor, 6),
                    "Correção monetária": money_br(item.correction_amount, places),
                    "Valor corrigido": money_br(corrected_value, places),
                    "Juros / valor corrigido": percent_br(interest_pct, 2),
                    "Juros de mora": money_br(item.interest_amount, places),
                    "Multa": money_br(item.penalty_amount, places),
                    "Valor antes dos abatimentos": money_br(gross_before_abatements, places),
                    "Abatimentos": money_br(item.abatements, places),
                    "Valor atualizado": money_br(item.updated_total, places),
                })
            st.dataframe(pd.DataFrame(visual_rows), use_container_width=True, hide_index=True)

        abatement_rows = [x for x in result_update.daily_rows if x.abatement_amount != 0]
        if abatement_rows:
            with st.expander("Ver como os abatimentos foram distribuídos", expanded=True):
                st.caption("O saldo antes do abatimento já contém a correção, os juros e a multa apurados até a data do evento. Em seguida o recurso é reduzido dos componentes indicados abaixo.")
                abatement_view = pd.DataFrame([{
                    "Data": date_br(x.day),
                    "Prestação / item": x.item_number,
                    "Saldo antes": money_br(x.closing_total + x.abatement_amount, places),
                    "Abatimento aplicado": money_br(x.abatement_amount, places),
                    "Redução do principal": money_br(x.abatement_principal, places),
                    "Redução da correção": money_br(x.abatement_correction, places),
                    "Redução dos juros": money_br(x.abatement_interest, places),
                    "Redução da multa": money_br(x.abatement_penalty, places),
                    "Saldo após": money_br(x.closing_total, places),
                } for x in abatement_rows])
                st.dataframe(abatement_view, use_container_width=True, hide_index=True)

        with st.form("update_document_identification_form", border=True):
            st.markdown("**Identificação do demonstrativo (opcional)**")
            st.caption(
                "Modo web/LGPD: apenas o número do contrato pode ser informado. Dados pessoais do participante e identificação de responsáveis não são coletados nesta versão."
            )
            contract_input = st.text_input(
                "Contrato",
                value=st.session_state.get("update_contract_number", ""),
                placeholder="Ex.: 123456-7",
            )
            save_identification = st.form_submit_button("Salvar identificação", use_container_width=True)
            if save_identification:
                st.session_state["update_contract_number"] = contract_input.strip()
                st.session_state.pop("update_participant_name", None)
                st.session_state.pop("update_pdf", None)
                st.session_state.pop("update_pdf_name", None)
                st.success("Identificação salva para o demonstrativo.")

        contract_number = st.session_state.get("update_contract_number", "")
        participant_name = ""

        copy_block_title("Texto pronto para copiar", "Use o ícone de cópia do bloco e cole diretamente no Word, e-mail ou manifestação.")
        copy_mode = st.radio("Conteúdo para copiar", ["Resumo objetivo", "Memória técnica"], horizontal=True, key="update_copy_mode")
        copy_text = build_update_copy_text(
            result_update, profile_label=profile_label, detailed=(copy_mode == "Memória técnica"), factor_series=factor_series,
            contract_number=contract_number, participant_name=participant_name,
        )
        st.code(copy_text, language=None)

        d1, d2, d3 = st.columns(3)
        with d1:
            if st.session_state.get("update_excel"):
                st.download_button(
                    "Baixar memória em Excel", data=st.session_state["update_excel"],
                    file_name=f"atualizacao_monetaria_{result_update.base_date.strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True,
                )
        with d2:
            if not result_update.is_complete:
                st.caption("O PDF será emitido como **cálculo parcial de correção**, indicando o último índice disponível e deixando claro que os juros, quando selecionados e com taxa disponível/determinável, foram apurados até a data-base.")
            if st.button("Preparar PDF premium", type="primary", use_container_width=True):
                try:
                    st.session_state["update_pdf"] = build_monetary_update_pdf(
                        result_update, logo_path=LOGO_WHITE_PATH, profile_label=profile_label, factor_series=factor_series,
                        contract_number=contract_number, participant_name=participant_name, elaborator="",
                    )
                    st.session_state["update_pdf_name"] = monetary_pdf_filename(result_update.base_date)
                except Exception as exc:
                    st.error(f"Não foi possível gerar o PDF: {exc}")
        with d3:
            if st.session_state.get("update_pdf"):
                st.download_button(
                    "Baixar PDF premium", data=st.session_state["update_pdf"],
                    file_name=st.session_state["update_pdf_name"], mime="application/pdf", use_container_width=True,
                )

        rv1, rv2 = st.columns([2, 1])
        with rv1:
            view = st.radio("Visualização", ["Resumo por valor", "Memória diária", "Regras aplicadas"], horizontal=True, key="update_view")
        with rv2:
            st.caption("As tabelas mantêm tela cheia, pesquisa, ocultação de colunas e download nativo.")

        if view == "Resumo por valor":
            summary_df = pd.DataFrame([{
                "Item": x.item_number, "Descrição": x.description, "Data de origem": date_br(x.origin_date), "Valor original": money_br(x.original_amount, places),
                "Correção": money_br(x.correction_amount, places), "Juros": money_br(x.interest_amount, places), "Multa": money_br(x.penalty_amount, places), "Abatimentos": money_br(x.abatements, places), "Saldo atualizado": money_br(x.updated_total, places),
            } for x in result_update.items])
            st.dataframe(summary_df, use_container_width=True, hide_index=True, height=min(620, 76 + 35 * max(len(summary_df), 1)))
        elif view == "Memória diária":
            daily_df = pd.DataFrame([{
                "Data": date_br(x.day), "Item": x.item_number, "Regra": x.rule_order if x.rule_order is not None else "-", "Competência": x.competence or "-", "Índice/fator ref.": x.index_reference or "-",
                "Principal inicial": money_br(x.opening_principal, places), "Correção inicial": money_br(x.opening_correction, places), "Juros inicial": money_br(x.opening_interest, places), "Multa inicial": money_br(x.opening_penalty, places),
                "Taxa C.M.": percent_br(x.correction_rate, 4), "C.M. do dia": money_br(x.correction_amount, places), "Taxa juros": percent_br(x.interest_rate, 4), "Juros do dia": money_br(x.interest_amount, places), "Taxa multa": percent_br(x.penalty_rate, 4), "Multa do dia": money_br(x.penalty_amount, places),
                "Abatimento": money_br(x.abatement_amount, places), "Principal final": money_br(x.closing_principal, places), "Correção final": money_br(x.closing_correction, places), "Juros final": money_br(x.closing_interest, places), "Multa final": money_br(x.closing_penalty, places), "Saldo final": money_br(x.closing_total, places),
            } for x in result_update.daily_rows])
            st.dataframe(daily_df, use_container_width=True, hide_index=True, height=620)
        else:
            rules_view = pd.DataFrame([{
                "Regra": x.order, "Período": f"{date_br(x.start_date)} a {date_br(x.end_date)}", "Descrição": x.description,
                "Índice/tabela": x.correction_index_code, "Correção": correction_labels_rev.get(x.correction_method, x.correction_method), "Defasagem": x.index_lag_months,
                "Juros a.m.": percent_br(x.monthly_interest_rate, 4), "Início dos juros": date_br(x.interest_start_date) if x.interest_start_date else date_br(x.start_date), "Método juros": interest_labels_rev.get(x.interest_method, x.interest_method), "Aplicação juros": interest_application_rev.get(x.interest_application, x.interest_application), "Base juros": interest_base_rev.get(x.interest_base, x.interest_base),
                "Ordem incidência": order_labels_rev.get(x.event_order, x.event_order), "Competência": competence_rev.get(x.competence_start_day, str(x.competence_start_day)),
            } for x in result_update.rules])
            st.dataframe(rules_view, use_container_width=True, hide_index=True)


elif module == "Central de índices":
    runtime = get_runtime_configuration()
    render_indices_center(
        repository=runtime.indices,
        source_label=runtime.index_source,
        set_navigation=set_navigation,
        page_intro=page_intro,
        callout=callout,
    )


elif module == "Feriados":
    render_holidays(
        active_cfg=get_active_config(),
        weekdays_pt=WEEKDAYS_PT,
        page_intro=page_intro,
        callout=callout,
    )


elif module == "Configuração":
    render_configuration(
        active_cfg=get_active_config(),
        export_configuration_bytes=export_active_configuration_bytes,
        page_intro=page_intro,
        callout=callout,
    )


elif module == "Acesso ao sistema":
    render_access(
        app_dir=APP_DIR,
        access_config=ACCESS_CONFIG,
        registry=ACCESS_REGISTRY,
        identity=REQUEST_IDENTITY,
        current_user_name=CURRENT_USER_NAME,
        current_profile=CURRENT_PROFILE,
        page_intro=page_intro,
        callout=callout,
    )

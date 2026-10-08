from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum


class ModalityCode(str, Enum):
    VARIABLE = "MOD_001"
    FIXED = "MOD_002"
    NEW_CREDINAMICO_FIXED = "MOD_004"
    CREDINAMICO_FIXED = "MOD_005"
    NEW_CREDINAMICO_VARIABLE = "MOD_006"


MODALITY_LABELS: dict[str, str] = {
    ModalityCode.VARIABLE.value: "C Variável",
    ModalityCode.FIXED.value: "C Fixo",
    ModalityCode.NEW_CREDINAMICO_FIXED.value: "Novo Credinâmico Fixo",
    ModalityCode.CREDINAMICO_FIXED.value: "Credinâmico Fixo",
    ModalityCode.NEW_CREDINAMICO_VARIABLE.value: "Novo Credinâmico Variável",
}

SUPPORTED_EVOLUTION_MODALITIES: tuple[str, ...] = tuple(MODALITY_LABELS)
NO_CORRECTION_MODALITIES: frozenset[str] = frozenset({
    ModalityCode.FIXED.value,
    ModalityCode.NEW_CREDINAMICO_FIXED.value,
    ModalityCode.CREDINAMICO_FIXED.value,
})
PRICE_MODALITIES: frozenset[str] = frozenset({
    ModalityCode.FIXED.value,
    ModalityCode.NEW_CREDINAMICO_FIXED.value,
    ModalityCode.CREDINAMICO_FIXED.value,
    ModalityCode.NEW_CREDINAMICO_VARIABLE.value,
})
CORRECTION_MODALITIES: frozenset[str] = frozenset({
    ModalityCode.VARIABLE.value,
    ModalityCode.NEW_CREDINAMICO_VARIABLE.value,
})


class PaymentTiming(int, Enum):
    END_OF_PERIOD = 0
    BEGINNING_OF_PERIOD = 1


class ExtraEffect(str, Enum):
    REDUCE_BALANCE = "REDUZIR_SALDO"
    REDUCE_INSTALLMENT = "REDUZIR_PRESTACAO"
    REDUCE_TERM = "REDUZIR_PRAZO"
    MODALITY_RULE = "REGRA_MODALIDADE"


@dataclass(frozen=True)
class RoundingSettings:
    rate_places: int = 12
    factor_places: int = 12
    money_places: int = 2
    percentage_display_places: int = 4
    mode: str = ROUND_HALF_UP
    round_daily_money: bool = False

    def _quantum(self, places: int) -> Decimal:
        return Decimal(1).scaleb(-places)

    def rate(self, value: Decimal) -> Decimal:
        return value.quantize(self._quantum(self.rate_places), rounding=self.mode)

    def factor(self, value: Decimal) -> Decimal:
        return value.quantize(self._quantum(self.factor_places), rounding=self.mode)

    def money(self, value: Decimal) -> Decimal:
        return value.quantize(self._quantum(self.money_places), rounding=self.mode)


@dataclass(frozen=True)
class ContractSettings:
    modality_code: str
    initial_balance: Decimal
    credit_date: date
    annual_interest_rate: Decimal
    term: int
    competence_start_day: int = 21
    due_day: int = 20
    index_code: str | None = "INPC"
    index_lag_months: int = 2
    payment_timing: int = 0
    interest_day_divisor: str = "COMPETENCIA"
    extra_effect: str = ExtraEffect.REDUCE_BALANCE.value
    rounding: RoundingSettings = field(default_factory=RoundingSettings)

    def __post_init__(self) -> None:
        if self.initial_balance <= 0:
            raise ValueError("O saldo inicial deve ser positivo.")
        if self.annual_interest_rate <= Decimal("-1"):
            raise ValueError("A taxa anual deve ser superior a -100%.")
        if self.term <= 0:
            raise ValueError("O prazo deve ser positivo.")
        if self.competence_start_day not in (1, 21):
            raise ValueError("O início da competência deve ser 1 ou 21.")
        if not 1 <= self.due_day <= 28:
            raise ValueError("Nesta versão, o dia de vencimento deve estar entre 1 e 28.")
        if self.payment_timing not in (0, 1):
            raise ValueError("O tipo da fórmula Price deve ser 0 ou 1.")


@dataclass(frozen=True)
class ExtraordinaryAmortization:
    event_date: date
    amount: Decimal
    note: str = ""

    def __post_init__(self) -> None:
        if self.amount <= 0:
            raise ValueError("A amortização extraordinária deve ser positiva.")


@dataclass(frozen=True)
class DailyRow:
    day: date
    installment_number: int
    competence: str
    index_reference: str | None
    is_business_day: bool
    is_holiday: bool
    opening_balance: Decimal
    daily_interest_rate: Decimal
    interest_amount: Decimal
    balance_after_interest: Decimal
    daily_correction_rate: Decimal
    correction_amount: Decimal
    balance_after_correction: Decimal
    extraordinary_amortization: Decimal
    accumulated_interest: Decimal
    balance_before_installment: Decimal
    regular_amortization: Decimal = Decimal("0")
    installment_amount: Decimal = Decimal("0")
    closing_balance_after_installment: Decimal | None = None


@dataclass(frozen=True)
class InstallmentRow:
    installment_number: int
    reference_due_date: date
    operational_due_date: date
    competence: str
    index_reference: str | None
    opening_balance: Decimal
    balance_before_installment: Decimal
    amortization_base: Decimal
    correction_amount: Decimal
    interest_amount: Decimal
    regular_amortization: Decimal
    extraordinary_amortization: Decimal
    installment_amount: Decimal
    closing_balance: Decimal
    remaining_installments: int

    @property
    def opening_principal(self) -> Decimal:
        return self.opening_balance

    @property
    def closing_principal(self) -> Decimal:
        return self.closing_balance


@dataclass
class EvolutionResult:
    modality_code: str
    fixed_payment: Decimal | None
    monthly_interest_rate: Decimal
    installments: list[InstallmentRow]
    daily_rows: list[DailyRow]
    warnings: list[str]
    is_complete: bool = True
    requested_installments: int = 0
    stopped_at: date | None = None
    missing_index_reference: str | None = None
    status_note: str = ""

    @property
    def processed_installments(self) -> int:
        return len(self.installments)

    @property
    def last_calculated_date(self) -> date | None:
        if self.daily_rows:
            return self.daily_rows[-1].day
        if self.installments:
            return self.installments[-1].reference_due_date
        return None

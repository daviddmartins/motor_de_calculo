from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date
from decimal import Decimal, localcontext

from .calendar_rules import (
    business_days_in_day_competence,
    calendar_days_in_day_competence,
    competence_label_for_day,
    first_reference_due_date,
    is_business_day,
    iter_dates_excluding_start,
    lagged_competence,
    next_business_day,
    next_reference_due_date,
)
from .indices import IndexSeries
from .models import (
    ContractSettings,
    DailyRow,
    EvolutionResult,
    ExtraordinaryAmortization,
    InstallmentRow,
    ModalityCode,
)
from .rates import (
    annual_to_monthly_equivalent,
    monthly_index_to_daily_equivalent,
    monthly_to_daily_equivalent,
    price_payment,
)


@dataclass(frozen=True)
class _DailyEvolutionOutcome:
    balance: Decimal
    accumulated_interest: Decimal
    accumulated_correction: Decimal
    accumulated_extra: Decimal
    rows: list[DailyRow]
    complete: bool = True
    stopped_at: date | None = None
    missing_index_reference: str | None = None


def _group_events(events: list[ExtraordinaryAmortization]) -> dict[date, Decimal]:
    grouped: dict[date, Decimal] = defaultdict(lambda: Decimal("0"))
    for event in events:
        grouped[event.event_date] += event.amount
    return dict(grouped)


def _period_summary(rows: list[DailyRow]) -> tuple[str, str | None]:
    if not rows:
        return "", None
    competences = list(dict.fromkeys(row.competence for row in rows))
    refs = list(dict.fromkeys(row.index_reference for row in rows if row.index_reference))
    comp = competences[0] if len(competences) == 1 else f"{competences[0]} a {competences[-1]}"
    if not refs:
        index_ref = None
    else:
        index_ref = refs[0] if len(refs) == 1 else f"{refs[0]} a {refs[-1]}"
    return comp, index_ref


def _partial_note(
    *,
    requested_installments: int,
    processed_installments: int,
    stopped_at: date | None,
    missing_index_reference: str | None,
) -> str:
    last_complete = (
        f"Foram concluídas {processed_installments} de {requested_installments} prestações solicitadas."
        if requested_installments
        else ""
    )
    stop_text = f" A evolução foi interrompida antes de {stopped_at:%d/%m/%Y}." if stopped_at else ""
    index_text = (
        f" O índice necessário para a competência {missing_index_reference} não está disponível."
        if missing_index_reference
        else ""
    )
    return ("Cálculo parcial. " + last_complete + stop_text + index_text).strip()


def _apply_daily_evolution(
    *,
    settings: ContractSettings,
    installment_number: int,
    period_start: date,
    due: date,
    opening_balance: Decimal,
    monthly_rate: Decimal,
    index_series: IndexSeries | None,
    holidays: set[date],
    events: dict[date, Decimal],
    warnings: list[str],
) -> _DailyEvolutionOutcome:
    """Evolui um único saldo, recalculando taxas quando muda a competência diária.

    Quando o índice de correção necessário não estiver disponível, retorna o período
    já calculado até o dia anterior, sem lançar erro. A prestação incompleta não é
    fechada, mas a memória diária parcial é preservada.
    """
    rounding = settings.rounding
    balance = opening_balance
    accumulated_interest = Decimal("0")
    accumulated_correction = Decimal("0")
    accumulated_extra = Decimal("0")
    rows: list[DailyRow] = []
    interest_rate_cache: dict[str, Decimal] = {}
    correction_rate_cache: dict[str, Decimal] = {}

    for current_day in iter_dates_excluding_start(period_start, due):
        comp = competence_label_for_day(current_day, settings.competence_start_day, settings.due_day)
        index_ref = lagged_competence(comp, settings.index_lag_months) if settings.index_code else None

        if comp not in interest_rate_cache:
            calendar_days = calendar_days_in_day_competence(
                current_day, settings.competence_start_day, settings.due_day
            )
            interest_rate_cache[comp] = rounding.rate(
                monthly_to_daily_equivalent(monthly_rate, calendar_days)
            )
        daily_interest = interest_rate_cache[comp]

        business = is_business_day(current_day, holidays)
        holiday = current_day in holidays

        # A correção só existe em dia útil. Portanto, a série é exigida somente
        # quando o primeiro dia útil daquela competência for efetivamente processado.
        if settings.index_code and business and comp not in correction_rate_cache:
            if not index_series or index_ref not in index_series.values:
                return _DailyEvolutionOutcome(
                    balance=balance,
                    accumulated_interest=accumulated_interest,
                    accumulated_correction=accumulated_correction,
                    accumulated_extra=accumulated_extra,
                    rows=rows,
                    complete=False,
                    stopped_at=current_day,
                    missing_index_reference=index_ref,
                )
            business_days = business_days_in_day_competence(
                current_day, settings.competence_start_day, holidays, settings.due_day
            )
            monthly_index = index_series.values[index_ref]
            correction_rate_cache[comp] = rounding.factor(
                monthly_index_to_daily_equivalent(monthly_index, business_days)
            )

        opening = balance
        interest_amount = opening * daily_interest
        if rounding.round_daily_money:
            interest_amount = rounding.money(interest_amount)
        after_interest = opening + interest_amount
        if rounding.round_daily_money:
            after_interest = rounding.money(after_interest)
        accumulated_interest += interest_amount

        correction_amount = Decimal("0")
        correction_rate = Decimal("0")
        if settings.index_code and business:
            correction_rate = correction_rate_cache[comp]
            correction_amount = after_interest * correction_rate
            if rounding.round_daily_money:
                correction_amount = rounding.money(correction_amount)
            accumulated_correction += correction_amount
        after_correction_before_extra = after_interest + correction_amount
        if rounding.round_daily_money:
            after_correction_before_extra = rounding.money(after_correction_before_extra)

        extra_requested = events.get(current_day, Decimal("0"))
        extra_applied = Decimal("0")
        balance_after_events = after_correction_before_extra
        if extra_requested:
            extra_applied = min(extra_requested, balance_after_events)
            balance_after_events -= extra_applied
            accumulated_extra += extra_applied
            if extra_requested > extra_applied:
                warnings.append(
                    f"Amortização extraordinária de {current_day:%d/%m/%Y} excedeu o saldo; "
                    f"R$ {extra_requested - extra_applied:.2f} não foram aplicados."
                )

        if rounding.round_daily_money:
            balance_after_events = rounding.money(balance_after_events)

        balance = balance_after_events
        rows.append(
            DailyRow(
                day=current_day,
                installment_number=installment_number,
                competence=comp,
                index_reference=index_ref,
                is_business_day=business,
                is_holiday=holiday,
                opening_balance=opening,
                daily_interest_rate=daily_interest,
                interest_amount=interest_amount,
                balance_after_interest=after_interest,
                daily_correction_rate=correction_rate,
                correction_amount=correction_amount,
                balance_after_correction=after_correction_before_extra,
                extraordinary_amortization=extra_applied,
                accumulated_interest=accumulated_interest,
                balance_before_installment=balance,
            )
        )

    return _DailyEvolutionOutcome(
        balance=balance,
        accumulated_interest=accumulated_interest,
        accumulated_correction=accumulated_correction,
        accumulated_extra=accumulated_extra,
        rows=rows,
    )


def _calculate_variable(
    settings: ContractSettings,
    index_series: IndexSeries | None,
    holidays: set[date],
    events: dict[date, Decimal],
) -> EvolutionResult:
    rounding = settings.rounding
    monthly_rate = rounding.rate(annual_to_monthly_equivalent(settings.annual_interest_rate))
    balance = settings.initial_balance
    period_start = settings.credit_date
    due = first_reference_due_date(settings.credit_date, settings.due_day)
    installments: list[InstallmentRow] = []
    daily_rows: list[DailyRow] = []
    warnings: list[str] = []
    stopped_at: date | None = None
    missing_index_reference: str | None = None

    with localcontext() as ctx:
        ctx.prec = 40
        for installment_number in range(1, settings.term + 1):
            if balance <= Decimal("0.0000001"):
                warnings.append(f"O saldo foi encerrado antes da prestação {installment_number}.")
                break

            opening_balance = balance
            outcome = _apply_daily_evolution(
                settings=settings,
                installment_number=installment_number,
                period_start=period_start,
                due=due,
                opening_balance=balance,
                monthly_rate=monthly_rate,
                index_series=index_series,
                holidays=holidays,
                events=events,
                warnings=warnings,
            )

            if not outcome.complete:
                daily_rows.extend(outcome.rows)
                stopped_at = outcome.stopped_at
                missing_index_reference = outcome.missing_index_reference
                break

            balance_before_installment = outcome.balance
            interest_period = outcome.accumulated_interest
            correction_period = outcome.accumulated_correction
            extra_period = outcome.accumulated_extra
            period_rows = outcome.rows

            remaining = settings.term - installment_number + 1
            amortization_base = balance_before_installment - interest_period
            if amortization_base < Decimal("-0.01"):
                raise ValueError(
                    f"A base de amortização ficou negativa na prestação {installment_number}. "
                    "Revise taxa, índice, prazo e eventos."
                )
            amortization_base = max(amortization_base, Decimal("0"))
            regular_amortization = (
                amortization_base if remaining == 1 else amortization_base / Decimal(remaining)
            )
            regular_amortization = min(rounding.money(regular_amortization), balance_before_installment)
            installment_amount = rounding.money(interest_period + regular_amortization)
            closing_balance = balance_before_installment - installment_amount
            if abs(closing_balance) < Decimal("0.01"):
                closing_balance = Decimal("0")

            if period_rows:
                period_rows[-1] = replace(
                    period_rows[-1],
                    regular_amortization=regular_amortization,
                    installment_amount=installment_amount,
                    closing_balance_after_installment=closing_balance,
                )
            daily_rows.extend(period_rows)

            comp, index_ref = _period_summary(period_rows)
            installments.append(
                InstallmentRow(
                    installment_number=installment_number,
                    reference_due_date=due,
                    operational_due_date=next_business_day(due, holidays),
                    competence=comp,
                    index_reference=index_ref,
                    opening_balance=rounding.money(opening_balance),
                    balance_before_installment=rounding.money(balance_before_installment),
                    amortization_base=rounding.money(amortization_base),
                    correction_amount=rounding.money(correction_period),
                    interest_amount=rounding.money(interest_period),
                    regular_amortization=rounding.money(regular_amortization),
                    extraordinary_amortization=rounding.money(extra_period),
                    installment_amount=rounding.money(installment_amount),
                    closing_balance=rounding.money(max(closing_balance, Decimal("0"))),
                    remaining_installments=max(settings.term - installment_number, 0),
                )
            )

            balance = max(closing_balance, Decimal("0"))
            period_start = due
            due = next_reference_due_date(due, settings.due_day)

    is_complete = stopped_at is None
    status_note = ""
    if not is_complete:
        status_note = _partial_note(
            requested_installments=settings.term,
            processed_installments=len(installments),
            stopped_at=stopped_at,
            missing_index_reference=missing_index_reference,
        )
        warnings.append(status_note)

    return EvolutionResult(
        modality_code=settings.modality_code,
        fixed_payment=None,
        monthly_interest_rate=monthly_rate,
        installments=installments,
        daily_rows=daily_rows,
        warnings=warnings,
        is_complete=is_complete,
        requested_installments=settings.term,
        stopped_at=stopped_at,
        missing_index_reference=missing_index_reference,
        status_note=status_note,
    )


def _calculate_fixed(
    settings: ContractSettings,
    index_series: IndexSeries | None,
    holidays: set[date],
    events: dict[date, Decimal],
) -> EvolutionResult:
    rounding = settings.rounding
    monthly_rate = rounding.rate(annual_to_monthly_equivalent(settings.annual_interest_rate))
    fixed_payment: Decimal | None = None
    balance = settings.initial_balance
    period_start = settings.credit_date
    due = first_reference_due_date(settings.credit_date, settings.due_day)
    installments: list[InstallmentRow] = []
    daily_rows: list[DailyRow] = []
    warnings: list[str] = []
    stopped_at: date | None = None
    missing_index_reference: str | None = None

    with localcontext() as ctx:
        ctx.prec = 40
        for installment_number in range(1, settings.term + 1):
            if balance <= Decimal("0.0000001"):
                warnings.append(f"O saldo foi encerrado antes da prestação {installment_number}.")
                break

            opening_balance = balance
            outcome = _apply_daily_evolution(
                settings=settings,
                installment_number=installment_number,
                period_start=period_start,
                due=due,
                opening_balance=balance,
                monthly_rate=monthly_rate,
                index_series=index_series,
                holidays=holidays,
                events=events,
                warnings=warnings,
            )

            if not outcome.complete:
                daily_rows.extend(outcome.rows)
                stopped_at = outcome.stopped_at
                missing_index_reference = outcome.missing_index_reference
                break

            balance_before_installment = outcome.balance
            interest_period = outcome.accumulated_interest
            correction_period = outcome.accumulated_correction
            extra_period = outcome.accumulated_extra
            period_rows = outcome.rows

            remaining = settings.term - installment_number + 1
            if fixed_payment is None:
                fixed_payment = rounding.money(
                    price_payment(
                        balance_before_installment,
                        monthly_rate,
                        settings.term,
                        settings.payment_timing,
                    )
                )
            installment_amount = fixed_payment
            if remaining == 1 or installment_amount >= balance_before_installment:
                installment_amount = rounding.money(balance_before_installment)
            regular_amortization = installment_amount - interest_period
            if regular_amortization <= 0:
                raise ValueError(
                    f"Na prestação {installment_number}, os juros superaram a prestação Price. "
                    "Revise taxa, prazo, tipo da fórmula e regra de carência."
                )
            closing_balance = balance_before_installment - installment_amount
            if abs(closing_balance) < Decimal("0.01"):
                closing_balance = Decimal("0")
            amortization_base = balance_before_installment - interest_period

            if period_rows:
                period_rows[-1] = replace(
                    period_rows[-1],
                    regular_amortization=regular_amortization,
                    installment_amount=installment_amount,
                    closing_balance_after_installment=closing_balance,
                )
            daily_rows.extend(period_rows)

            comp, index_ref = _period_summary(period_rows)
            installments.append(
                InstallmentRow(
                    installment_number=installment_number,
                    reference_due_date=due,
                    operational_due_date=next_business_day(due, holidays),
                    competence=comp,
                    index_reference=index_ref,
                    opening_balance=rounding.money(opening_balance),
                    balance_before_installment=rounding.money(balance_before_installment),
                    amortization_base=rounding.money(amortization_base),
                    correction_amount=rounding.money(correction_period),
                    interest_amount=rounding.money(interest_period),
                    regular_amortization=rounding.money(regular_amortization),
                    extraordinary_amortization=rounding.money(extra_period),
                    installment_amount=rounding.money(installment_amount),
                    closing_balance=rounding.money(max(closing_balance, Decimal("0"))),
                    remaining_installments=max(settings.term - installment_number, 0),
                )
            )

            balance = max(closing_balance, Decimal("0"))
            period_start = due
            due = next_reference_due_date(due, settings.due_day)

    is_complete = stopped_at is None
    status_note = ""
    if not is_complete:
        status_note = _partial_note(
            requested_installments=settings.term,
            processed_installments=len(installments),
            stopped_at=stopped_at,
            missing_index_reference=missing_index_reference,
        )
        warnings.append(status_note)

    return EvolutionResult(
        modality_code=settings.modality_code,
        fixed_payment=fixed_payment,
        monthly_interest_rate=monthly_rate,
        installments=installments,
        daily_rows=daily_rows,
        warnings=warnings,
        is_complete=is_complete,
        requested_installments=settings.term,
        stopped_at=stopped_at,
        missing_index_reference=missing_index_reference,
        status_note=status_note,
    )



def _calculate_repriced_price(
    settings: ContractSettings,
    index_series: IndexSeries | None,
    holidays: set[date],
    events: dict[date, Decimal],
) -> EvolutionResult:
    """Price recalculada em cada vencimento sobre saldo e prazo remanescentes.

    Esta é a regra das modalidades Novo Credinâmico Fixo e Novo Credinâmico
    Variável das abas 4 e 6 da planilha de referência. O tipo Price é 1
    (pagamento no início do período). Na modalidade variável, juros e correção
    monetária são apropriados diariamente antes do fechamento da prestação.
    """
    rounding = settings.rounding
    monthly_rate = rounding.rate(annual_to_monthly_equivalent(settings.annual_interest_rate))
    balance = settings.initial_balance
    period_start = settings.credit_date
    due = first_reference_due_date(settings.credit_date, settings.due_day)
    installments: list[InstallmentRow] = []
    daily_rows: list[DailyRow] = []
    warnings: list[str] = []
    stopped_at: date | None = None
    missing_index_reference: str | None = None

    with localcontext() as ctx:
        ctx.prec = 40
        for installment_number in range(1, settings.term + 1):
            if balance <= Decimal("0.0000001"):
                warnings.append(f"O saldo foi encerrado antes da prestação {installment_number}.")
                break

            opening_balance = balance
            outcome = _apply_daily_evolution(
                settings=settings,
                installment_number=installment_number,
                period_start=period_start,
                due=due,
                opening_balance=balance,
                monthly_rate=monthly_rate,
                index_series=index_series,
                holidays=holidays,
                events=events,
                warnings=warnings,
            )
            if not outcome.complete:
                daily_rows.extend(outcome.rows)
                stopped_at = outcome.stopped_at
                missing_index_reference = outcome.missing_index_reference
                break

            balance_before_installment = outcome.balance
            interest_period = outcome.accumulated_interest
            correction_period = outcome.accumulated_correction
            extra_period = outcome.accumulated_extra
            period_rows = outcome.rows
            remaining = settings.term - installment_number + 1

            installment_amount = rounding.money(
                price_payment(
                    balance_before_installment,
                    monthly_rate,
                    remaining,
                    1,
                )
            )
            if remaining == 1 or installment_amount >= balance_before_installment:
                installment_amount = rounding.money(balance_before_installment)
            regular_amortization = installment_amount - interest_period
            if regular_amortization <= 0:
                raise ValueError(
                    f"Na prestação {installment_number}, os juros superaram a prestação Price recalculada. "
                    "Revise taxa, prazo e eventos."
                )
            closing_balance = balance_before_installment - installment_amount
            if abs(closing_balance) < Decimal("0.01"):
                closing_balance = Decimal("0")
            amortization_base = balance_before_installment - interest_period

            if period_rows:
                period_rows[-1] = replace(
                    period_rows[-1],
                    regular_amortization=regular_amortization,
                    installment_amount=installment_amount,
                    closing_balance_after_installment=closing_balance,
                )
            daily_rows.extend(period_rows)
            comp, index_ref = _period_summary(period_rows)
            installments.append(
                InstallmentRow(
                    installment_number=installment_number,
                    reference_due_date=due,
                    operational_due_date=next_business_day(due, holidays),
                    competence=comp,
                    index_reference=index_ref,
                    opening_balance=rounding.money(opening_balance),
                    balance_before_installment=rounding.money(balance_before_installment),
                    amortization_base=rounding.money(amortization_base),
                    correction_amount=rounding.money(correction_period),
                    interest_amount=rounding.money(interest_period),
                    regular_amortization=rounding.money(regular_amortization),
                    extraordinary_amortization=rounding.money(extra_period),
                    installment_amount=rounding.money(installment_amount),
                    closing_balance=rounding.money(max(closing_balance, Decimal("0"))),
                    remaining_installments=max(settings.term - installment_number, 0),
                )
            )
            balance = max(closing_balance, Decimal("0"))
            period_start = due
            due = next_reference_due_date(due, settings.due_day)

    is_complete = stopped_at is None
    status_note = ""
    if not is_complete:
        status_note = _partial_note(
            requested_installments=settings.term,
            processed_installments=len(installments),
            stopped_at=stopped_at,
            missing_index_reference=missing_index_reference,
        )
        warnings.append(status_note)

    fixed_payment = (
        installments[0].installment_amount
        if installments and settings.modality_code == ModalityCode.NEW_CREDINAMICO_FIXED.value
        else None
    )
    return EvolutionResult(
        modality_code=settings.modality_code,
        fixed_payment=fixed_payment,
        monthly_interest_rate=monthly_rate,
        installments=installments,
        daily_rows=daily_rows,
        warnings=warnings,
        is_complete=is_complete,
        requested_installments=settings.term,
        stopped_at=stopped_at,
        missing_index_reference=missing_index_reference,
        status_note=status_note,
    )


def _calculate_credinamico_fixed_annual(
    settings: ContractSettings,
    holidays: set[date],
    events: dict[date, Decimal],
) -> EvolutionResult:
    """Credinâmico Fixo da aba 5: juros no aniversário anual e Price tipo 0.

    Não há correção monetária. Fora da data anual de aniversário da primeira
    prestação, a taxa de juros diária é zero. No aniversário, a taxa anual
    contratual é incorporada ao saldo. A prestação mensal é recalculada em cada
    vencimento pela taxa mensal equivalente e pelo prazo remanescente; na memória
    da planilha, a prestação é tratada integralmente como amortização do saldo já
    acrescido do eventual encargo anual.
    """
    rounding = settings.rounding
    monthly_rate = rounding.rate(annual_to_monthly_equivalent(settings.annual_interest_rate))
    balance = settings.initial_balance
    first_due = first_reference_due_date(settings.credit_date, settings.due_day)
    period_start = settings.credit_date
    due = first_due
    installments: list[InstallmentRow] = []
    daily_rows: list[DailyRow] = []
    warnings: list[str] = []

    with localcontext() as ctx:
        ctx.prec = 40
        for installment_number in range(1, settings.term + 1):
            if balance <= Decimal("0.0000001"):
                warnings.append(f"O saldo foi encerrado antes da prestação {installment_number}.")
                break

            opening_balance = balance
            accumulated_interest = Decimal("0")
            accumulated_extra = Decimal("0")
            period_rows: list[DailyRow] = []

            for current_day in iter_dates_excluding_start(period_start, due):
                opening = balance
                years_after_first_due = current_day.year - first_due.year
                anniversary = (
                    years_after_first_due >= 1
                    and current_day.month == first_due.month
                    and current_day.day == first_due.day
                )
                daily_interest = rounding.rate(settings.annual_interest_rate) if anniversary else Decimal("0")
                interest_amount = opening * daily_interest if anniversary else Decimal("0")
                if rounding.round_daily_money:
                    interest_amount = rounding.money(interest_amount)
                after_interest = opening + interest_amount
                if rounding.round_daily_money:
                    after_interest = rounding.money(after_interest)
                accumulated_interest += interest_amount

                extra_requested = events.get(current_day, Decimal("0"))
                extra_applied = Decimal("0")
                balance_after_events = after_interest
                if extra_requested:
                    extra_applied = min(extra_requested, balance_after_events)
                    balance_after_events -= extra_applied
                    accumulated_extra += extra_applied
                    if extra_requested > extra_applied:
                        warnings.append(
                            f"Amortização extraordinária de {current_day:%d/%m/%Y} excedeu o saldo; "
                            f"R$ {extra_requested - extra_applied:.2f} não foram aplicados."
                        )
                if rounding.round_daily_money:
                    balance_after_events = rounding.money(balance_after_events)
                balance = balance_after_events

                period_rows.append(
                    DailyRow(
                        day=current_day,
                        installment_number=installment_number,
                        competence=competence_label_for_day(current_day, settings.competence_start_day, settings.due_day),
                        index_reference=None,
                        is_business_day=is_business_day(current_day, holidays),
                        is_holiday=current_day in holidays,
                        opening_balance=opening,
                        daily_interest_rate=daily_interest,
                        interest_amount=interest_amount,
                        balance_after_interest=after_interest,
                        daily_correction_rate=Decimal("0"),
                        correction_amount=Decimal("0"),
                        balance_after_correction=after_interest,
                        extraordinary_amortization=extra_applied,
                        accumulated_interest=accumulated_interest,
                        balance_before_installment=balance,
                    )
                )

            balance_before_installment = balance
            remaining = settings.term - installment_number + 1
            installment_amount = rounding.money(
                price_payment(balance_before_installment, monthly_rate, remaining, 0)
            )
            if remaining == 1 or installment_amount >= balance_before_installment:
                installment_amount = rounding.money(balance_before_installment)
            regular_amortization = installment_amount
            closing_balance = balance_before_installment - installment_amount
            if abs(closing_balance) < Decimal("0.01"):
                closing_balance = Decimal("0")

            if period_rows:
                period_rows[-1] = replace(
                    period_rows[-1],
                    regular_amortization=regular_amortization,
                    installment_amount=installment_amount,
                    closing_balance_after_installment=closing_balance,
                )
            daily_rows.extend(period_rows)
            comp, _ = _period_summary(period_rows)
            installments.append(
                InstallmentRow(
                    installment_number=installment_number,
                    reference_due_date=due,
                    operational_due_date=next_business_day(due, holidays),
                    competence=comp,
                    index_reference=None,
                    opening_balance=rounding.money(opening_balance),
                    balance_before_installment=rounding.money(balance_before_installment),
                    amortization_base=rounding.money(balance_before_installment),
                    correction_amount=Decimal("0"),
                    interest_amount=rounding.money(accumulated_interest),
                    regular_amortization=rounding.money(regular_amortization),
                    extraordinary_amortization=rounding.money(accumulated_extra),
                    installment_amount=rounding.money(installment_amount),
                    closing_balance=rounding.money(max(closing_balance, Decimal("0"))),
                    remaining_installments=max(settings.term - installment_number, 0),
                )
            )
            balance = max(closing_balance, Decimal("0"))
            period_start = due
            due = next_reference_due_date(due, settings.due_day)

    return EvolutionResult(
        modality_code=settings.modality_code,
        fixed_payment=None,
        monthly_interest_rate=monthly_rate,
        installments=installments,
        daily_rows=daily_rows,
        warnings=warnings,
        is_complete=True,
        requested_installments=settings.term,
        status_note="",
    )

def calculate_contract_evolution(
    settings: ContractSettings,
    index_series: IndexSeries | None,
    holidays: set[date] | None = None,
    extraordinary_amortizations: list[ExtraordinaryAmortization] | None = None,
) -> EvolutionResult:
    holidays = holidays or set()
    events = _group_events(extraordinary_amortizations or [])
    if settings.modality_code == ModalityCode.VARIABLE.value:
        return _calculate_variable(settings, index_series, holidays, events)
    if settings.modality_code == ModalityCode.FIXED.value:
        return _calculate_fixed(settings, index_series, holidays, events)
    if settings.modality_code == ModalityCode.NEW_CREDINAMICO_FIXED.value:
        normalized = replace(settings, index_code=None, payment_timing=1)
        return _calculate_repriced_price(normalized, None, holidays, events)
    if settings.modality_code == ModalityCode.CREDINAMICO_FIXED.value:
        normalized = replace(settings, index_code=None, payment_timing=0)
        return _calculate_credinamico_fixed_annual(normalized, holidays, events)
    if settings.modality_code == ModalityCode.NEW_CREDINAMICO_VARIABLE.value:
        if settings.index_code is None:
            raise ValueError("Novo Credinâmico Variável requer índice de correção monetária.")
        normalized = replace(settings, payment_timing=1)
        return _calculate_repriced_price(normalized, index_series, holidays, events)
    raise ValueError("Modalidade ainda não implementada nesta versão.")

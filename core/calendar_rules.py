from __future__ import annotations

import calendar
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta


def first_reference_due_date(credit_date: date, due_day: int = 20) -> date:
    """Primeira referência sempre no mês seguinte ao crédito."""
    next_month = credit_date + relativedelta(months=1, day=1)
    return next_month.replace(day=due_day)


def next_reference_due_date(current_due: date, due_day: int = 20) -> date:
    next_month = current_due + relativedelta(months=1, day=1)
    return next_month.replace(day=due_day)


def is_business_day(day: date, holidays: set[date]) -> bool:
    return day.weekday() < 5 and day not in holidays


def next_business_day(day: date, holidays: set[date]) -> date:
    current = day
    while not is_business_day(current, holidays):
        current += timedelta(days=1)
    return current


def iter_dates_excluding_start(start: date, end: date):
    if end < start:
        raise ValueError("A data final não pode ser anterior à inicial.")
    current = start + timedelta(days=1)
    while current <= end:
        yield current
        current += timedelta(days=1)


def competence_bounds_for_due(reference_due: date, start_day: int, due_day: int = 20) -> tuple[date, date]:
    if start_day == 21:
        start = (reference_due - relativedelta(months=1)).replace(day=21)
        end = reference_due.replace(day=due_day)
        return start, end
    if start_day == 1:
        start = reference_due.replace(day=1)
        last = calendar.monthrange(reference_due.year, reference_due.month)[1]
        return start, reference_due.replace(day=last)
    raise ValueError("start_day deve ser 1 ou 21.")


def competence_bounds_for_day(day: date, start_day: int, due_day: int = 20) -> tuple[date, date]:
    """Retorna a competência efetivamente aplicável a uma data diária."""
    if start_day == 21:
        if day.day >= 21:
            start = day.replace(day=21)
            end = (start + relativedelta(months=1)).replace(day=due_day)
        else:
            end = day.replace(day=due_day)
            start = (end - relativedelta(months=1)).replace(day=21)
        return start, end
    if start_day == 1:
        start = day.replace(day=1)
        last = calendar.monthrange(day.year, day.month)[1]
        return start, day.replace(day=last)
    raise ValueError("start_day deve ser 1 ou 21.")


def competence_label_for_day(day: date, start_day: int, due_day: int = 20) -> str:
    start, _ = competence_bounds_for_day(day, start_day, due_day)
    return f"{start.year:04d}-{start.month:02d}"


def competence_label(reference_due: date, start_day: int) -> str:
    ref = reference_due - relativedelta(months=1) if start_day == 21 else reference_due
    return f"{ref.year:04d}-{ref.month:02d}"


def lagged_competence(label: str, lag_months: int) -> str:
    year, month = [int(x) for x in label.split("-")]
    ref = date(year, month, 1) - relativedelta(months=lag_months)
    return f"{ref.year:04d}-{ref.month:02d}"


def calendar_days_in_competence(reference_due: date, start_day: int, due_day: int = 20) -> int:
    start, end = competence_bounds_for_due(reference_due, start_day, due_day)
    return (end - start).days + 1


def calendar_days_in_day_competence(day: date, start_day: int, due_day: int = 20) -> int:
    start, end = competence_bounds_for_day(day, start_day, due_day)
    return (end - start).days + 1


def business_days_in_competence(reference_due: date, start_day: int, holidays: set[date], due_day: int = 20) -> int:
    start, end = competence_bounds_for_due(reference_due, start_day, due_day)
    return _count_business_days(start, end, holidays)


def business_days_in_day_competence(day: date, start_day: int, holidays: set[date], due_day: int = 20) -> int:
    start, end = competence_bounds_for_day(day, start_day, due_day)
    return _count_business_days(start, end, holidays)


def _count_business_days(start: date, end: date, holidays: set[date]) -> int:
    total = 0
    current = start
    while current <= end:
        if is_business_day(current, holidays):
            total += 1
        current += timedelta(days=1)
    return total

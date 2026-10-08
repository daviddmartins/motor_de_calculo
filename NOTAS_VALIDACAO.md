# Notas de validação — v0.9.3

## Evolução de contrato

O baseline matemático preservado na migração para o GitHub é `2026.08-baseline`. A migração não altera as regras matemáticas já existentes para as modalidades variável e fixa.

## Atualização monetária

Motor `2026.08.2`. A base migrada inclui a metodologia livre com pró-rata mensal por dias corridos, tratamento de índice ausente sem projeção e possibilidade de continuidade dos juros configurados até a data-base sobre a base disponível.

## MAJS

Motor `2026.08.1`. Além das regras já existentes de pagamentos, suspensões e amortizações extraordinárias, a v0.9.3 inclui quitação antecipada.

Cenário de regressão da entrega:

- valor original: R$ 80.000,00;
- crédito: 08/01/2020;
- prazo: 96 prestações;
- taxa mensal: 0,600000%;
- índice MAJS de teste: 0,3000% ao mês;
- data-base: 15/08/2026;
- quitação: 10/05/2022;
- saldo MAJS teórico esperado na quitação: R$ 58.737,45.

Os testes automatizados cobrem tanto novação quanto pagamento pelo participante de R$ 50.000,00, cuja diferença específica de quitação esperada é -R$ 8.737,45.

## Regra de uso

Resultados destinados a uso oficial devem continuar passando pelo processo de homologação e conferência definido pela área responsável. O GitHub passa a controlar versões e regressões do código, mas não substitui a validação de negócio.

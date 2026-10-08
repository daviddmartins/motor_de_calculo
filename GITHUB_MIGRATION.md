# Migração para GitHub

## Base utilizada

A pasta exportada do Databricks correspondia à aplicação `0.9.2`. Antes da preparação do primeiro versionamento GitHub foi aplicado o delta cumulativo `0.9.3`.

## Correções preservadas da base 0.9.2

- metodologia livre de atualização monetária com pró-rata mensal por dias corridos;
- correção parcial quando falta índice, sem projeção de competência futura;
- continuidade opcional de juros até a data-base quando a correção fica congelada;
- juros opcionais em perfil TJDFT;
- correção visual do componente de linha do tempo de regras;
- navegação de abatimentos/amortizações e conferência em tela.

## Delta aplicado para 0.9.3

- quitação antecipada no MAJS;
- distinção entre novação e pagamento pelo participante;
- encerramento do fluxo na data de quitação;
- saldo teórico proporcional na data do evento;
- diferença específica da quitação apenas quando houver pagamento pelo participante;
- atualização das saídas Excel/PDF;
- motor MAJS identificado como `2026.08.1`.

## Preparação GitHub

- `.gitignore` com proteção da planilha real de acesso e arquivos secretos;
- `requirements-dev.txt` e `pytest.ini`;
- workflow de testes em `.github/workflows/tests.yml`;
- `Dockerfile` para execução portável;
- documentação atualizada;
- planilha real de acesso removida do conjunto versionável.

## Validação de regressão

A suíte automatizada deve permanecer verde antes da promoção de mudanças para `main`.

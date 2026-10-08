# Testes de arquitetura e regressão

Esta pasta inicia a camada automatizada de testes sem alterar os motores matemáticos.

- `test_index_service.py`: valida estrutura, tipos (percentual/fator), séries obrigatórias e round-trip da nova base única.
- `test_health_service.py`: garante que a Central de Saúde reflita o status real do catálogo, sem inventar atualização/frescor.

Os casos matemáticos homologados dos motores devem ser acrescentados progressivamente como testes de regressão em arquivos separados (`test_evolution_regression.py`, `test_majs_regression.py` e `test_monetary_regression.py`) a partir das memórias reais já homologadas pela área. A recomendação é congelar cada caso somente depois da conferência de negócio, para evitar transformar um valor ainda não homologado em “verdade” automatizada.

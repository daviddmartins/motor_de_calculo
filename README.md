# Motor de Cálculos

Aplicação Streamlit para cálculos financeiros e contratuais, organizada para execução no Databricks Apps ou em qualquer ambiente Python/contêiner compatível.

**Versão do aplicativo:** `0.10.11`

## Módulos

- **Evolução de contrato e parecer** — cinco modalidades, com memória por prestação, memória diária quando aplicável e parecer técnico específico:
  - C Variável (`MOD_001`);
  - C Fixo (`MOD_002`);
  - Novo Credinâmico Fixo (`MOD_004`);
  - Credinâmico Fixo (`MOD_005`);
  - Novo Credinâmico Variável (`MOD_006`).
- **Recálculo de diferenças (MAJS)** — reconstrução do fluxo teórico, pagamentos, suspensões, amortizações extraordinárias, diferenças e quitação antecipada.
- **Atualização do saldo devedor** — regras livres e perfis judiciais, correção monetária, juros, multas e abatimentos, com tratamento explícito de índice ausente.
- **Administração** — índices, feriados, configurações, saúde do Motor e controle de acesso quando disponibilizado pelo ambiente.

## Arquitetura

```text
app.py                     interface principal Streamlit
core/                      motores matemáticos, relatórios e modelos
services/                  acesso a configurações e infraestrutura
ui/                        componentes, telas e identidade visual
data/                       séries e tabelas versionáveis
config/                     índices/configurações não sensíveis
assets/                     identidade visual
templates/                  modelos de parecer
exemplos/                   arquivos de apoio/teste
tests/                      testes automatizados
.github/workflows/          integração contínua no GitHub
```

O núcleo matemático permanece separado da interface. Alterações exclusivamente visuais não devem mudar os identificadores dos motores em `core/engine_versions.py`.

## Versões dos motores

- Evolução de contrato: `2026.08.1`
- MAJS: `2026.08.1`
- Atualização monetária: `2026.08.2`

## MAJS — quitação antecipada

Na v0.9.3 o MAJS permite encerrar o contrato antes do prazo originalmente contratado:

- **Novação:** o Motor apura o saldo teórico na data da novação e encerra o fluxo. O saldo não é tratado como desembolso do participante e não cria diferença de quitação.
- **Pagamento pelo(a) participante:** o Motor apura o saldo teórico na data da quitação, confronta com o valor efetivamente pago e inclui a diferença na tabela de diferenças.
- Em ambos os casos não são geradas prestações posteriores à data da quitação.

## Privacidade na versão web

O Motor não acessa a base de participantes e não coleta CPF. As seções técnicas dos pareceres não exibem dados pessoais nem identificação de elaborador ou validador.

O parecer de evolução segue o padrão documental FUNCEF (“Manifestação de Subsídios”): faixa institucional, quadros 01 a 03 (Dados do Processo, Participante e Operação, Demanda) e, ao final, o quadro “Responsável pela informação”. Todo esse cabeçalho é preenchido manualmente pelo usuário no formulário do parecer, inclusive Mutuário(s) e Matrícula(s); os dados digitados são usados somente na geração do PDF. Campos deixados em branco aparecem vazios no documento, e o formulário avisa quais ficaram em branco e se o número do processo foge do padrão CNJ.

Nos métodos de juros por fração de ano, a apresentação utiliza a terminologia financeira de convenção de contagem de dias (US/NASD 30/360, Real/Real, Real/360, Real/365 ou Europeu 30/360), sem referência a funções de planilha eletrônica.

## Execução local

Requer Python 3.11+.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/macOS: source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Testes

```bash
pip install -r requirements-dev.txt
pytest -q
```

O workflow `.github/workflows/tests.yml` executa compilação e testes a cada push ou pull request para `main`.

## Docker

```bash
docker build -t motor-de-calculos .
docker run --rm -p 8501:8501 motor-de-calculos
```

A aplicação ficará disponível na porta `8501`.

## Databricks Apps

O `app.yaml` original foi preservado. Assim, o repositório também pode continuar sendo usado como fonte do Databricks Apps quando esse ambiente estiver disponível.

## Dados sensíveis e controle de acesso

A planilha real de controle de acesso **não deve ser versionada no GitHub**. Os padrões de nome usados por essa planilha estão bloqueados no `.gitignore`.

Quando necessário, disponibilize o arquivo de acesso diretamente no ambiente de execução em `config/`. Veja `config/README.md`.

Nunca inclua no repositório senhas, tokens, chaves, dados pessoais de participantes ou arquivos de produção contendo informações individuais.

## Histórico

Consulte `CHANGELOG.md` para as alterações relevantes de cada entrega.

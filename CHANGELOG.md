# Changelog

## 0.10.11 — 2026-10-07

- cabeçalho do parecer (quadros 01 a 03, data e responsável) passa a ser preenchido integralmente à mão no formulário, na mesma ordem do PDF, incluindo Mutuário(s) e Matrícula(s);
- UF em lista com as 27 unidades da federação; Contrato(s) e Modalidade, quando em branco, usam o número do contrato e a modalidade do cálculo;
- avisos não bloqueantes ao aplicar os dados: campos do cabeçalho em branco e número do processo fora do padrão CNJ;
- removida a exportação do parecer em Word (.docx); o parecer volta a ser emitido somente em PDF;
- nenhuma regra matemática foi alterada.

## 0.10.10 — 2026-10-07

- tipografia corporativa única no parecer: somente Helvetica (equivalente métrica do Arial do modelo), com corpo de texto 9 pt, tabelas 7,5 pt e nenhum texto abaixo de 7 pt;
- fórmulas passam a ser texto na fonte do documento (antes eram imagens em outra fonte); símbolos Σ e ∏ substituídos por notação explícita;
- títulos de seção no padrão do modelo (selo numerado + título) e numeração contínua após os quadros 01 a 03 em todas as modalidades;
- paginação revisada: fim do bloco indivisível que deixava páginas quase vazias nas modalidades 004 a 006, títulos nunca isolados no pé da página, tabela de parâmetros sem quebra e quadro do responsável sempre acompanhado do último parágrafo;
- novo arquivo Word (.docx) do parecer, gerado a partir do modelo corporativo com o mesmo conteúdo do PDF, para edição e preenchimento manual dos campos bloqueados pela LGPD;
- nenhuma regra matemática foi alterada.

## 0.10.9 — 2026-10-07

- parecer de evolução contratual passa a seguir o padrão documental FUNCEF do modelo “Manifestação de Subsídios”: faixa institucional com logo, áreas (DIBEN, GERAT, COPART) e data; título; quadros 01 Dados do Processo, 02 Participante e Operação e 03 Demanda;
- rótulo “#10 Corporativo - FUNCEF” no topo e rodapé “COPART · GERAT · DIBEN” em todas as páginas, com numeração “Página X de Y”;
- quadro “Responsável pela informação” ao final do documento (sem “Conferido por”);
- o conteúdo técnico existente vem após a abertura; nas modalidades 004 a 006 as seções foram renumeradas de 04 a 09;
- formulário do parecer ganhou o bloco “Cabeçalho padrão FUNCEF”; Mutuário(s) e Matrícula(s) permanecem em branco no modo web/LGPD;
- nenhuma regra matemática foi alterada.

## 0.10.2 — 2026-08-16

- reforço de LGPD na versão web: PDFs ignoram por completo nome, matrícula, CPF, elaborador e validador, mesmo quando valores legados persistem em sessão;
- parecer de evolução contratual passa a mencionar somente o número do contrato na identificação;
- parecer MAJS deixa de renderizar participante e elaboração;
- atualização monetária deixa de renderizar participante;
- nenhuma regra matemática foi alterada.

## 0.10.1 — 2026-08-15

- aplicada política temporária de LGPD para a versão web: nome, matrícula e CPF do participante deixam de ser coletados nos formulários de documentos;
- elaborador/usuário logado e validador deixam de ser coletados e exibidos nos PDFs enquanto o Motor estiver fora do Databricks;
- campos de identificação vazios passam a ser omitidos integralmente dos PDFs, sem rótulos vazios ou “Não informado”;
- parecer de evolução contratual, parecer MAJS e demonstrativo de atualização monetária foram alinhados à mesma regra de privacidade;
- a metodologia antes apresentada como “FRAÇÃOANO do Excel” passa a ser descrita tecnicamente como juros simples por fração de ano segundo convenção de contagem de dias;
- mantida a matemática interna e as convenções US/NASD 30/360, Real/Real, Real/360, Real/365 e Europeu 30/360;
- nenhuma versão dos motores matemáticos foi alterada, pois a mudança é de apresentação, privacidade e nomenclatura metodológica.

## 0.10.0 — 2026-08-15

- adicionadas três modalidades ao módulo **Evolução de contrato e parecer**, com base nas abas 4, 5 e 6 do arquivo de referência `Modalidades.xlsx`;
- **Novo Credinâmico Fixo (MOD_004):** evolução diária sem correção monetária, juros equivalentes e prestação Price tipo 1 recalculada pelo saldo e prazo remanescentes;
- **Credinâmico Fixo (MOD_005):** sem correção monetária, encargo de juros incorporado apenas no aniversário anual da primeira prestação e prestação mensal Price tipo 0 recalculada;
- **Novo Credinâmico Variável (MOD_006):** juros diários, correção monetária pelo INPC em dias úteis, ciclo 21–20 com defasagem de dois meses e prestação Price tipo 1 recalculada;
- criadas fundamentação, metodologia, fórmulas, demonstração e conclusão específicas no parecer técnico de cada nova modalidade;
- adicionados testes de regressão com os valores de referência das abas 4, 5 e 6;
- versão do motor de evolução atualizada para `2026.08.1`.

## 0.9.3 — 2026-08-15

- migração da base do Motor para estrutura versionável no GitHub;
- quitação antecipada no MAJS, por novação ou pagamento pelo(a) participante;
- encerramento do fluxo MAJS na data de quitação, sem geração de prestações posteriores;
- apuração proporcional do saldo teórico na data de quitação;
- novação sem lançamento do saldo como desembolso/diferença;
- quitação sem novação com confronto entre saldo teórico e valor efetivamente pago;
- atualização das memórias Excel e PDF do MAJS;
- preservadas as correções cumulativas da v0.9.2 para atualização monetária;
- adicionados testes automatizados, workflow GitHub Actions e regras de proteção de arquivos locais.

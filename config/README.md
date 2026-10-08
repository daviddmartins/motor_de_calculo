# Configuração local

O repositório contém as bases de índices e configurações necessárias ao Motor, mas **não deve conter a planilha real de controle de acesso**.

Para ambientes que utilizam o cadastro local de usuários, disponibilize o arquivo de acesso diretamente no ambiente de execução, dentro de `config/`, sem adicioná-lo ao Git. O `.gitignore` bloqueia os nomes usuais dessa planilha.

Em produção, prefira identidade/autorização provida pela plataforma de hospedagem sempre que possível.

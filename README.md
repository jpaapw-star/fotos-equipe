# Fotos da equipe

App da Web (Google Apps Script) que coloca o link da foto de cada pessoa na coluna **FOTOS** da planilha
"NUMERO DE IDENTIFICAÇÃO DE PRODUÇÃO FUNCIONARIOS EMANX", em todas as abas onde a pessoa aparece
(Página1 e FUNÇÕES).

- Você escolhe as fotos (várias ou um .zip). O app acha a pessoa pelo nome do arquivo (nome inteiro, parte
  do nome ou número identificador; sem diferença de acento ou maiúscula).
- Mostra a lista antes de gravar. Se tiver dúvida (ex.: "PEDRO.jpg"), você escolhe a pessoa na lista.
- Cada foto vai para a pasta **Fotos da equipe** no Drive, com "qualquer pessoa com o link pode ver",
  e o link vai para a coluna FOTOS. Foto nova da mesma pessoa troca a antiga (a antiga vai para a lixeira).
- Procura a pessoa pelo nome em cada aba, porque as abas não têm as mesmas pessoas nas mesmas linhas.

Instalação: [docs/INSTALACAO.md](docs/INSTALACAO.md).

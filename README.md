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

- Também aceita uma **pasta inteira**: do computador (botão "Pasta do PC" ou arrastando a pasta) ou do Google Drive (cole o link; as fotos ficam onde estão, o app só libera o link de leitura e grava na planilha).

- **Continua sozinho**: fotos da pasta do Drive entram numa fila que o Google processa a cada minuto, mesmo com o app
  ou a planilha fechados. Fotos do PC precisam do app aberto enquanto são enviadas.
- **Não repete**: foto que já foi colocada não aparece de novo (fica em "já na planilha"). Foto nova da mesma pessoa
  troca a antiga (a antiga vai para a lixeira do Drive).
- **Fotos guardadas**: toda foto fica guardada na pasta do app ("Fotos da equipe") e aparece em "Fotos guardadas".
- **Busca**: em cada foto, digite parte do nome ou o número para escolher a pessoa.

Instalação: [docs/INSTALACAO.md](docs/INSTALACAO.md).

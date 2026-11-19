# Fotos da equipe

App da Web (Google Apps Script) que coloca o link da foto de cada pessoa na coluna **FOTOS** da planilha
"IDENTIFICAÇÃO FUNCIONARIOS EMANX", em todas as abas que têm as colunas NOME e FOTOS
(hoje as abas NUMERO PRODUCAO e FUNCIONÁRIOS; o título NOME COMPLETO também vale). Abas sem essas duas colunas são ignoradas.

- Você escolhe as fotos (várias ou um .zip). O app acha a pessoa pelo nome do arquivo (nome inteiro, parte
  do nome ou número identificador; sem diferença de acento ou maiúscula).
- Mostra a lista antes de gravar. Se tiver dúvida (ex.: "PEDRO.jpg"), você escolhe a pessoa na lista.
- Cada foto vai para a pasta **Fotos da equipe** no Drive, com "qualquer pessoa com o link pode ver",
  e o link vai para a coluna FOTOS. Foto nova da mesma pessoa troca a antiga (a antiga vai para a lixeira).
- Procura a pessoa pelo nome em cada aba, porque as abas não têm as mesmas pessoas nas mesmas linhas.
- A planilha manda: ao abrir, as fotos guardadas acompanham o nome que está na planilha (se o nome ou o
  número mudar, a foto segue o link). Foto sem ninguém na planilha continua guardada e aparece marcada.

- Também aceita uma **pasta inteira**: do computador (botão "Pasta do PC" ou arrastando a pasta) ou do Google Drive (cole o link; as fotos ficam onde estão, o app só libera o link de leitura e grava na planilha).

- **Continua sozinho**: fotos da pasta do Drive entram numa fila que o Google processa a cada minuto, mesmo com o app
  ou a planilha fechados. Fotos do PC precisam do app aberto enquanto são enviadas.
- **Não repete**: foto que já foi colocada não aparece de novo (fica em "já na planilha"). Foto nova da mesma pessoa
  troca a antiga (a antiga vai para a lixeira do Drive).
- **Fotos guardadas**: toda foto fica guardada na pasta do app ("Fotos da equipe") e aparece em "Fotos guardadas", onde dá para **Trocar** a pessoa (nome errado) ou **Excluir** a foto (sai da planilha e vai para a lixeira).
- **Busca**: em cada foto, digite parte do nome ou o número para escolher a pessoa.

Instalação: [docs/INSTALACAO.md](docs/INSTALACAO.md).

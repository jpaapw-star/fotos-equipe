# Instalação (uma vez só)

Depois disso, todo código novo que entrar na branch `main` é enviado e publicado sozinho.

## 1. Um segredo no GitHub

No repositório: **Settings › Secrets and variables › Actions › New repository secret**.

| Nome | O que colar |
| --- | --- |
| `CLASPRC_JSON` | O mesmo conteúdo usado no Produção por voz: o arquivo `.clasprc.json` da sua pasta de usuário no Windows (`C:\Users\<você>\.clasprc.json`), aberto no Bloco de Notas. |

Nunca cole esse valor no chat.

## 2. Primeira publicação

Em **Actions › Apps Script › Run workflow**. Na primeira vez, o projeto do Apps Script é criado na sua conta
Google e o link do app aparece no resumo da execução.

## 3. Autorizar uma vez

Abra o link do app logado na sua conta Google e aceite o pedido de acesso (planilhas e Drive). Se aparecer
"O Google não verificou este app", toque em **Avançado › Acessar**. É normal em app próprio.

## Bom saber

- Qualquer pessoa com o link e uma conta Google (qualquer e-mail) abre o app. Ele grava como você, então os outros não precisam ter acesso à planilha nem à pasta. Você precisa poder editar a planilha.
- Ele só escreve na coluna FOTOS. As outras colunas não mudam.
- Para usar outra planilha: no projeto do Apps Script › Configurações do projeto › Propriedades do script,
  crie `PLANILHA_ID` com o ID da planilha.

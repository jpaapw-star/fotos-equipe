/*
 * Fotos da equipe — app da Web que recebe fotos, guarda no Drive e escreve o link na coluna FOTOS
 * de cada aba da planilha onde a pessoa aparece (procura pelo nome em cada aba, nunca pela linha).
 * Só escreve na coluna FOTOS; não mexe nas outras colunas.
 *
 * Planilha: PLANILHA_PADRAO, ou a propriedade do script PLANILHA_ID se existir.
 */

var VERSAO = '1.1.0';
var PLANILHA_PADRAO = '1sCB4bhqsPXJqm8gaMJNAoZb2EmfGLaevwBvT_QLaQBA';
var PASTA_NOME = 'Fotos da equipe';

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Fotos da equipe')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function planilhaId_() {
  return PropertiesService.getScriptProperties().getProperty('PLANILHA_ID') || PLANILHA_PADRAO;
}

/* ------------------------------------------------------------------ planilha */

// Em cada aba, acha a linha de títulos com NOME e FOTOS (e NÚMERO, se houver).
function colunas_(valores) {
  for (var r = 0; r < Math.min(5, valores.length); r++) {
    var col = {};
    valores[r].forEach(function (v, i) {
      var t = normalizar_(v);
      if (t === 'NOME' && col.nome === undefined) col.nome = i;
      else if ((t === 'FOTO' || t === 'FOTOS') && col.foto === undefined) col.foto = i;
      else if (t.indexOf('NUMERO') === 0 && col.numero === undefined) col.numero = i;
    });
    if (col.nome !== undefined && col.foto !== undefined) { col.titulo = r; return col; }
  }
  return null;
}

// Lê todas as abas que têm NOME e FOTOS. Junta a mesma pessoa (mesmo nome) das várias abas.
function lerPessoas_(ss) {
  var mapa = {}, lista = [], abas = [];
  ss.getSheets().forEach(function (aba) {
    var ult = aba.getLastRow(), larg = aba.getLastColumn();
    if (ult < 2 || larg < 2) return;
    var valores = aba.getRange(1, 1, ult, larg).getValues();
    var col = colunas_(valores);
    if (!col) return;
    abas.push(aba.getName());
    for (var r = col.titulo + 1; r < valores.length; r++) {
      var nome = String(valores[r][col.nome] || '').replace(/\s+/g, ' ').trim();
      var chave = normalizar_(nome);
      if (!chave) continue;
      var p = mapa[chave];
      if (!p) { p = mapa[chave] = { chave: chave, nome: nome, numeros: [], lugares: [], temFoto: false }; lista.push(p); }
      var num = col.numero !== undefined ? String(valores[r][col.numero]).trim() : '';
      if (num !== '' && p.numeros.indexOf(num) < 0) p.numeros.push(num);
      if (String(valores[r][col.foto] || '').trim()) p.temFoto = true;
      p.lugares.push({ aba: aba.getName(), linha: r + 1, coluna: col.foto + 1 });
    }
  });
  if (!abas.length) throw new Error('Nenhuma aba da planilha tem as colunas NOME e FOTOS.');
  lista.sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return { pessoas: lista, abas: abas };
}

/* ------------------------------------------------------------------ chamadas do app */

// Recebe só os nomes dos arquivos e devolve quem é cada um, mais a lista de pessoas para escolher.
function apiPreparar(arquivos) {
  var ss = SpreadsheetApp.openById(planilhaId_());
  var dados = lerPessoas_(ss);
  return {
    versao: VERSAO,
    planilha: ss.getName(),
    planilhaUrl: ss.getUrl(),
    abas: dados.abas,
    pessoas: dados.pessoas.map(function (p) {
      return { chave: p.chave, nome: p.nome, numeros: p.numeros, temFoto: p.temFoto,
        abas: p.lugares.map(function (l) { return l.aba; }) };
    }),
    resultado: combinar_(arquivos || [], dados.pessoas)
  };
}

// Guarda uma foto no Drive (qualquer pessoa com o link vê) e escreve o link em todas as abas da pessoa.
function apiEnviar(envio) {
  if (!envio || !envio.chave || !envio.base64) throw new Error('Foto ou pessoa faltando.');
  var ss = SpreadsheetApp.openById(planilhaId_());
  var pessoa = lerPessoas_(ss).pessoas.filter(function (p) { return p.chave === envio.chave; })[0];
  if (!pessoa) throw new Error('A pessoa não está mais na planilha.');

  var ext = (String(envio.arquivo || '').match(/\.([A-Za-z0-9]{2,5})$/) || [, 'jpg'])[1].toLowerCase();
  if (envio.mime === 'image/jpeg') ext = ext === 'jpeg' ? 'jpeg' : 'jpg';
  var nomeArquivo = (pessoa.numeros.length ? pessoa.numeros[0] + ' - ' : '') + pessoa.nome + '.' + ext;
  var blob = Utilities.newBlob(Utilities.base64Decode(envio.base64), envio.mime || 'image/jpeg', nomeArquivo);

  var pasta = pasta_(ss);
  // Foto antiga da mesma pessoa vai para a lixeira do Drive (dá para recuperar por 30 dias).
  var antigas = pasta.getFilesByName(nomeArquivo);
  while (antigas.hasNext()) antigas.next().setTrashed(true);
  var arquivo = pasta.createFile(blob);
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';

  return gravarLink_(ss, pessoa, link);
}

// Escreve o link na coluna FOTOS de todas as abas onde a pessoa aparece.
function gravarLink_(ss, pessoa, link) {
  var trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    pessoa.lugares.forEach(function (l) {
      ss.getSheetByName(l.aba).getRange(l.linha, l.coluna).setValue(link);
    });
    SpreadsheetApp.flush();
  } finally {
    trava.releaseLock();
  }
  return {
    link: link,
    substituiu: pessoa.temFoto,
    lugares: pessoa.lugares.map(function (l) { return l.aba + ' linha ' + l.linha; })
  };
}

/* ------------------------------------------------------------------ pasta do Drive */

function idDaPasta_(link) {
  var t = String(link || '').trim();
  var m = t.match(/\/folders\/([A-Za-z0-9_-]{10,})/) || t.match(/[?&]id=([A-Za-z0-9_-]{10,})/) || t.match(/^([A-Za-z0-9_-]{10,})$/);
  if (!m) throw new Error('Isso não parece um link de pasta do Drive.');
  return m[1];
}

// Lista as fotos da pasta do Drive (e das subpastas). As fotos não são movidas nem renomeadas.
function apiListarPasta(link) {
  var pasta;
  try {
    pasta = DriveApp.getFolderById(idDaPasta_(link));
  } catch (e) {
    throw new Error('Não consegui abrir essa pasta. Confira o link e se ela está no seu Drive ou compartilhada com você.');
  }
  var arquivos = [], pilha = [pasta];
  while (pilha.length && arquivos.length < 2000) {
    var atual = pilha.shift();
    var fs = atual.getFiles();
    while (fs.hasNext()) {
      var f = fs.next();
      var tipo = f.getMimeType() || '';
      if (tipo.indexOf('image/') === 0 || /\.(jpe?g|png|webp|heic|heif)$/i.test(f.getName())) {
        arquivos.push({ id: f.getId(), nome: f.getName() });
      }
    }
    var subs = atual.getFolders();
    while (subs.hasNext()) pilha.push(subs.next());
  }
  return { pasta: pasta.getName(), arquivos: arquivos };
}

// Usa a foto que já está no Drive: libera "qualquer pessoa com o link vê" e grava o link.
// Se a foto for de outra pessoa (não dá para mudar o compartilhamento), faz uma cópia na pasta "Fotos da equipe".
function apiUsarDoDrive(envio) {
  if (!envio || !envio.chave || !envio.driveId) throw new Error('Foto ou pessoa faltando.');
  var ss = SpreadsheetApp.openById(planilhaId_());
  var pessoa = lerPessoas_(ss).pessoas.filter(function (p) { return p.chave === envio.chave; })[0];
  if (!pessoa) throw new Error('A pessoa não está mais na planilha.');
  var arquivo = DriveApp.getFileById(envio.driveId);
  try {
    if (arquivo.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK &&
        arquivo.getSharingAccess() !== DriveApp.Access.ANYONE) {
      arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (e) {
    arquivo = arquivo.makeCopy(arquivo.getName(), pasta_(ss));
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }
  return gravarLink_(ss, pessoa, 'https://drive.google.com/file/d/' + arquivo.getId() + '/view');
}

// Pasta "Fotos da equipe" ao lado da planilha (criada na primeira vez).
function pasta_(ss) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('PASTA_ID');
  if (id) {
    try {
      var salva = DriveApp.getFolderById(id);
      if (!salva.isTrashed()) return salva;
    } catch (e) { /* pasta apagada: cria outra */ }
  }
  var pai = DriveApp.getRootFolder(), pasta;
  try {
    var pais = DriveApp.getFileById(ss.getId()).getParents();
    if (pais.hasNext()) pai = pais.next();
  } catch (e) { /* planilha de outra pessoa sem acesso à pasta: usa o Meu Drive */ }
  try {
    var existentes = pai.getFoldersByName(PASTA_NOME);
    pasta = existentes.hasNext() ? existentes.next() : pai.createFolder(PASTA_NOME);
  } catch (e) {
    pasta = DriveApp.getRootFolder().createFolder(PASTA_NOME);
  }
  props.setProperty('PASTA_ID', pasta.getId());
  return pasta;
}

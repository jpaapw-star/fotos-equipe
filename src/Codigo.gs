/*
 * Fotos da equipe — app da Web que recebe fotos, guarda no Drive e escreve o link na coluna FOTOS
 * de cada aba da planilha onde a pessoa aparece (procura pelo nome em cada aba, nunca pela linha).
 * Só escreve na coluna FOTOS; não mexe nas outras colunas.
 *
 * Planilha: PLANILHA_PADRAO, ou a propriedade do script PLANILHA_ID se existir.
 */

var VERSAO = '1.3.0';
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

// Recebe os arquivos ({nome, origem}) e devolve quem é cada um, mais a lista de pessoas para escolher.
// "origem" identifica o arquivo (id do Drive, ou nome|tamanho do PC): foto já colocada não volta para a lista.
function apiPreparar(arquivos) {
  arquivos = (arquivos || []).map(function (a) { return typeof a === 'string' ? { nome: a, origem: '' } : a; });
  var ss = SpreadsheetApp.openById(planilhaId_());
  var dados = lerPessoas_(ss);
  var reg = lerDados_().registro;
  var resultado = combinar_(arquivos.map(function (a) { return a.nome; }), dados.pessoas);
  marcarJaColocadas_(resultado, arquivos, dados.pessoas, reg);
  return {
    versao: VERSAO,
    planilha: ss.getName(),
    planilhaUrl: ss.getUrl(),
    abas: dados.abas,
    pessoas: dados.pessoas.map(function (p) {
      return { chave: p.chave, nome: p.nome, numeros: p.numeros, temFoto: p.temFoto,
        abas: p.lugares.map(function (l) { return l.aba; }) };
    }),
    resultado: resultado
  };
}

// Foto que já está na planilha não aparece de novo. Pessoa que já tem foto só recebe outra se o arquivo for diferente.
function marcarJaColocadas_(resultado, arquivos, pessoas, reg) {
  var usadas = {};
  Object.keys(reg).forEach(function (c) { if (reg[c].origem) usadas[reg[c].origem] = c; });
  var porChave = {};
  pessoas.forEach(function (p) { porChave[p.chave] = p; });
  resultado.forEach(function (r, i) {
    var origem = arquivos[i].origem;
    if (origem && usadas[origem]) {
      r.ja = true;
      r.chave = usadas[origem];
      r.motivo = 'Esta foto já está na planilha.';
      return;
    }
    var p = r.chave && porChave[r.chave];
    if (!p || !p.temFoto) return;
    if (reg[r.chave]) {
      r.motivo += ' Foto nova: vai trocar a atual.';
    } else {
      // Foto colocada antes (sem registro no app): fica junto das já colocadas; dá para escolher e trocar.
      r.ja = true;
      r.motivo = 'Essa pessoa já tem foto na planilha. Escolha aqui se quiser trocar.';
    }
  });
}

// Foto vinda do PC: guarda na pasta do app e grava o link.
function apiEnviar(envio) {
  if (!envio || !envio.chave || !envio.base64) throw new Error('Foto ou pessoa faltando.');
  var ss = SpreadsheetApp.openById(planilhaId_());
  var pessoa = acharPessoa_(ss, envio.chave);
  var blob = Utilities.newBlob(Utilities.base64Decode(envio.base64), envio.mime || 'image/jpeg', envio.arquivo || 'foto.jpg');
  return guardarFoto_(ss, pessoa, { blob: blob, arquivo: envio.arquivo, mime: envio.mime }, envio.origem || '');
}

function acharPessoa_(ss, chave, pessoas) {
  var lista = pessoas || lerPessoas_(ss).pessoas;
  var p = lista.filter(function (x) { return x.chave === chave; })[0];
  if (!p) throw new Error('A pessoa não está mais na planilha.');
  return p;
}

// Cria a foto na pasta "Fotos da equipe" (link de leitura), apaga a foto antiga da pessoa,
// escreve o link em todas as abas e guarda no registro do app.
function guardarFoto_(ss, pessoa, fonte, origem) {
  var pasta = pasta_(ss);
  var ext = (String(fonte.arquivo || '').match(/\.([A-Za-z0-9]{2,5})$/) || [, 'jpg'])[1].toLowerCase();
  if (fonte.mime === 'image/jpeg' && ext !== 'jpeg') ext = 'jpg';
  var nomeArquivo = (pessoa.numeros.length ? pessoa.numeros[0] + ' - ' : '') + pessoa.nome + '.' + ext;
  var arquivo;
  if (fonte.blob) {
    fonte.blob.setName(nomeArquivo);
    arquivo = pasta.createFile(fonte.blob);
  } else {
    arquivo = DriveApp.getFileById(fonte.driveId).makeCopy(nomeArquivo, pasta);
  }
  arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var link = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';

  var trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    var dados = lerDados_();
    var antiga = dados.registro[pessoa.chave];
    // A foto antiga vai para a lixeira do Drive (dá para recuperar por 30 dias).
    var apagar = [];
    if (antiga && antiga.fileId) apagar.push(antiga.fileId);
    var mesmas = pasta.getFiles();
    var prefixo = nomeArquivo.replace(/\.[^.]+$/, '.');
    while (mesmas.hasNext()) {
      var f = mesmas.next();
      if (f.getName().indexOf(prefixo) === 0 && f.getId() !== arquivo.getId()) apagar.push(f.getId());
    }
    apagar.forEach(function (id) {
      if (id === arquivo.getId()) return;
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* já apagada */ }
    });
    pessoa.lugares.forEach(function (l) {
      ss.getSheetByName(l.aba).getRange(l.linha, l.coluna).setValue(link);
    });
    SpreadsheetApp.flush();
    dados.registro[pessoa.chave] = { nome: pessoa.nome, fileId: arquivo.getId(), link: link, origem: origem,
      quando: new Date().toISOString() };
    salvarDados_(dados);
  } finally {
    trava.releaseLock();
  }
  return {
    link: link,
    substituiu: pessoa.temFoto,
    lugares: pessoa.lugares.map(function (l) { return l.aba + ' linha ' + l.linha; })
  };
}

/* ------------------------------------------------------------------ registro e fila (continuam sem o app aberto) */

// Arquivo dados-do-app.json na pasta do app: registro das fotos guardadas e fila de fotos do Drive a colocar.
function arquivoDados_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('DADOS_ID');
  if (id) {
    try {
      var f = DriveApp.getFileById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* apagado: cria outro */ }
  }
  var novo = pasta_(SpreadsheetApp.openById(planilhaId_()))
    .createFile('dados-do-app.json', JSON.stringify(dadosVazios_()), 'application/json');
  props.setProperty('DADOS_ID', novo.getId());
  return novo;
}

function dadosVazios_() {
  return { registro: {}, fila: [], total: 0, feitas: 0, erros: [] };
}

function lerDados_() {
  try {
    var d = JSON.parse(arquivoDados_().getBlob().getDataAsString());
    var v = dadosVazios_();
    Object.keys(v).forEach(function (k) { if (d[k] === undefined) d[k] = v[k]; });
    return d;
  } catch (e) {
    return dadosVazios_();
  }
}

function salvarDados_(d) {
  arquivoDados_().setContent(JSON.stringify(d));
}

// Fotos do Drive: entram numa fila que o Google processa sozinho, mesmo com o app fechado.
function apiEnfileirar(lista) {
  lista = (lista || []).filter(function (x) { return x && x.driveId && x.chave; });
  if (!lista.length) return apiEstado();
  var trava = LockService.getScriptLock();
  trava.waitLock(30000);
  try {
    var d = lerDados_();
    var naFila = {};
    d.fila.forEach(function (x) { naFila[x.driveId] = true; });
    if (!d.fila.length) { d.total = 0; d.feitas = 0; d.erros = []; }
    lista.forEach(function (x) {
      if (naFila[x.driveId]) return;
      d.fila.push({ driveId: x.driveId, chave: x.chave, arquivo: x.arquivo || '' });
      d.total++;
    });
    salvarDados_(d);
  } finally {
    trava.releaseLock();
  }
  ligarRotina_();
  return apiEstado();
}

function ligarRotina_() {
  var existe = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'processarFila'; });
  if (!existe) ScriptApp.newTrigger('processarFila').timeBased().everyMinutes(1).create();
}

function desligarRotina_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'processarFila') ScriptApp.deleteTrigger(t);
  });
}

// Roda a cada minuto enquanto houver fila. Trabalha até ~4 minutos por vez e continua na próxima rodada.
function processarFila() {
  var props = PropertiesService.getScriptProperties();
  var agora = Date.now();
  if (Number(props.getProperty('FILA_RODANDO_ATE') || 0) > agora) return;
  props.setProperty('FILA_RODANDO_ATE', String(agora + 5 * 60 * 1000));
  try {
    var ss = SpreadsheetApp.openById(planilhaId_());
    var pessoas = lerPessoas_(ss).pessoas;
    while (Date.now() - agora < 4 * 60 * 1000) {
      var item = lerDados_().fila[0];
      if (!item) break;
      var erro = '';
      try {
        var pessoa = acharPessoa_(ss, item.chave, pessoas);
        guardarFoto_(ss, pessoa, { driveId: item.driveId, arquivo: item.arquivo }, 'drive:' + item.driveId);
        pessoa.temFoto = true;
      } catch (e) {
        erro = item.arquivo + ': ' + (e && e.message || e);
      }
      var trava = LockService.getScriptLock();
      trava.waitLock(30000);
      try {
        var d = lerDados_();
        d.fila = d.fila.filter(function (x) { return x.driveId !== item.driveId; });
        if (erro) d.erros.push(erro); else d.feitas++;
        salvarDados_(d);
      } finally {
        trava.releaseLock();
      }
    }
    if (!lerDados_().fila.length) desligarRotina_();
  } finally {
    props.deleteProperty('FILA_RODANDO_ATE');
  }
}

// O que o app mostra ao abrir: andamento da fila e as fotos guardadas.
function apiEstado() {
  var d = lerDados_();
  var guardadas = Object.keys(d.registro).map(function (c) {
    var r = d.registro[c];
    return { chave: c, nome: r.nome, link: r.link, fileId: r.fileId, quando: r.quando };
  }).sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return {
    fila: { total: d.total, feitas: d.feitas, faltam: d.fila.length, erros: d.erros.slice(-20),
      naFila: d.fila.map(function (x) { return x.driveId; }) },
    guardadas: guardadas
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

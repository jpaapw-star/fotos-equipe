/*
 * Fotos da equipe — app da Web que recebe fotos, guarda no Drive e escreve o link na coluna FOTOS
 * de cada aba da planilha onde a pessoa aparece (procura pelo nome em cada aba, nunca pela linha).
 * Só escreve na coluna FOTOS; não mexe nas outras colunas.
 *
 * Planilha: PLANILHA_PADRAO, ou a propriedade do script PLANILHA_ID se existir.
 */

var VERSAO = '1.7.0';
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
      if ((t === 'NOME' || t === 'NOME COMPLETO') && col.nome === undefined) col.nome = i;
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
      if (!p) { p = mapa[chave] = { chave: chave, nome: nome, numeros: [], lugares: [], links: [], temFoto: false }; lista.push(p); }
      var num = col.numero !== undefined ? String(valores[r][col.numero]).trim() : '';
      if (num !== '' && p.numeros.indexOf(num) < 0) p.numeros.push(num);
      var link = String(valores[r][col.foto] || '').trim();
      if (link) { p.temFoto = true; if (p.links.indexOf(link) < 0) p.links.push(link); }
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
  sincronizarRegistro_(ss, dados.pessoas);
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
  var link = 'https://drive.google.com/file/d/' + arquivo.getId() + '/view';
  var gravou = false;
  try {
    arquivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Fotos antigas com o mesmo nome: procura antes de pegar a vez, para não segurar os outros.
    var apagar = [];
    var mesmas = pasta.getFiles();
    var prefixo = nomeArquivo.replace(/\.[^.]+$/, '.');
    while (mesmas.hasNext()) {
      var f = mesmas.next();
      if (f.getName().indexOf(prefixo) === 0 && f.getId() !== arquivo.getId()) apagar.push(f.getId());
    }

    var trava = esperarVez_();
    try {
      var dados = lerDados_();
      var antiga = dados.registro[pessoa.chave];
      if (antiga && antiga.fileId) apagar.push(antiga.fileId);
      // Duas fotos da mesma pessoa ao mesmo tempo: a outra pode ter mandado esta para a lixeira. A última vale.
      if (arquivo.isTrashed()) arquivo.setTrashed(false);
      gravou = true;
      pessoa.lugares.forEach(function (l) {
        ss.getSheetByName(l.aba).getRange(l.linha, l.coluna).setValue(link);
      });
      SpreadsheetApp.flush();
      dados.registro[pessoa.chave] = { nome: pessoa.nome, fileId: arquivo.getId(), link: link, origem: origem,
        quando: new Date().toISOString() };
      salvarDados_(dados);
      // Só com o link novo já na planilha a foto antiga vai para a lixeira (dá para recuperar por 30 dias):
      // se o Drive falhar no meio, nenhum link fica apontando para foto apagada.
      // A lista foi feita antes da vez: foto que neste meio-tempo passou para outra pessoa (Trocar) fica.
      var deOutros = {};
      Object.keys(dados.registro).forEach(function (c) { if (c !== pessoa.chave) deOutros[dados.registro[c].fileId] = true; });
      apagar.forEach(function (id) {
        if (id === arquivo.getId() || deOutros[id]) return;
        try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* já apagada; sobra na pasta */ }
      });
    } finally {
      trava.releaseLock();
    }
  } catch (e) {
    // Não chegou na planilha: a foto recém-criada não fica sobrando na pasta.
    if (!gravou) { try { arquivo.setTrashed(true); } catch (x) { /* tanto faz */ } }
    throw e;
  }
  return {
    link: link,
    substituiu: pessoa.temFoto,
    lugares: pessoa.lugares.map(function (l) { return l.aba + ' linha ' + l.linha; })
  };
}

/* ------------------------------------------------------------------ registro e fila (continuam sem o app aberto) */

// Pega a vez de mexer na planilha e no registro. Muita gente ao mesmo tempo: espera até 2 minutos.
function esperarVez_() {
  var trava = LockService.getScriptLock();
  if (trava.tryLock(120000)) return trava;
  throw new Error('Muita gente usando o app agora. Tente de novo em 1 minuto.');
}

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

// Se o Drive falhar na leitura, dá erro (tente de novo) em vez de seguir com um registro vazio:
// seguir vazio apagaria a lista de fotos guardadas na próxima gravação.
function lerDados_() {
  var texto = arquivoDados_().getBlob().getDataAsString();
  var v = dadosVazios_();
  if (!String(texto || '').trim()) return v;
  var d = JSON.parse(texto);
  if (!d || typeof d !== 'object') throw new Error('O registro do app (dados-do-app.json) está estragado.');
  Object.keys(v).forEach(function (k) { if (d[k] == null) d[k] = v[k]; });
  return d;
}

function salvarDados_(d) {
  arquivoDados_().setContent(JSON.stringify(d));
}

// Fotos do Drive: entram numa fila que o Google processa sozinho, mesmo com o app fechado.
function apiEnfileirar(lista) {
  lista = (lista || []).filter(function (x) { return x && x.driveId && x.chave; });
  if (!lista.length) return apiEstado();
  var trava = esperarVez_();
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
  // Conferir e marcar "rodando" de uma vez só: duas rodadas juntas fariam a mesma foto duas vezes.
  var vez = LockService.getScriptLock();
  if (!vez.tryLock(10000)) return;
  try {
    if (Number(props.getProperty('FILA_RODANDO_ATE') || 0) > agora) return;
    props.setProperty('FILA_RODANDO_ATE', String(agora + 5 * 60 * 1000));
  } finally {
    vez.releaseLock();
  }
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
      var trava = esperarVez_();
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

function idDoLink_(link) {
  var t = String(link || '');
  var m = t.match(/\/d\/([A-Za-z0-9_-]+)/) || t.match(/[?&]id=([A-Za-z0-9_-]+)/);
  return m ? m[1] : '';
}

// A planilha manda: cada foto guardada fica com a pessoa que tem o link dela na planilha
// (nome e número podem ter mudado). Foto sem ninguém na planilha continua guardada, marcada.
function sincronizarRegistro_(ss, pessoas) {
  // Quase sempre nada mudou: confere sem pegar a vez, e só pega a vez quando há o que arrumar.
  if (!registroSincronizado_(lerDados_(), pessoas).mudou) return;
  var trava = esperarVez_();
  try {
    // Relê a planilha já com a vez: alguém pode ter trocado uma foto de pessoa neste meio-tempo.
    var d = lerDados_(), r = registroSincronizado_(d, lerPessoas_(ss).pessoas);
    if (r.mudou) { d.registro = r.registro; salvarDados_(d); }
  } finally {
    trava.releaseLock();
  }
}

function registroSincronizado_(d, pessoas) {
  var dono = {};
  pessoas.forEach(function (p) {
    p.links.forEach(function (l) { var id = idDoLink_(l); if (id && !dono[id]) dono[id] = p; });
  });
  var novo = {}, mudou = false;
  Object.keys(d.registro).forEach(function (c) {
    var r = d.registro[c], p = dono[r.fileId];
    var chave = p && !novo[p.chave] ? p.chave : c;
    if (novo[chave]) chave = c + ' (' + r.fileId + ')';
    if (chave !== c) mudou = true;
    if (p && r.nome !== p.nome) { r.nome = p.nome; mudou = true; }
    if (!!r.semPessoa === !!p) { r.semPessoa = !p; mudou = true; }
    novo[chave] = r;
  });
  return { mudou: mudou, registro: novo };
}

// O que o app mostra ao abrir: andamento da fila e as fotos guardadas.
// Com sincronizar = true, antes confere as fotos guardadas com a planilha.
function apiEstado(sincronizar) {
  if (sincronizar) {
    var ss = SpreadsheetApp.openById(planilhaId_());
    sincronizarRegistro_(ss, lerPessoas_(ss).pessoas);
  }
  var d = lerDados_();
  var guardadas = Object.keys(d.registro).map(function (c) {
    var r = d.registro[c];
    return { chave: c, nome: r.nome, link: r.link, fileId: r.fileId, quando: r.quando, semPessoa: !!r.semPessoa };
  }).sort(function (a, b) { return a.nome.localeCompare(b.nome, 'pt-BR'); });
  return {
    fila: { total: d.total, feitas: d.feitas, faltam: d.fila.length, erros: d.erros.slice(-20),
      naFila: d.fila.map(function (x) { return x.driveId; }) },
    guardadas: guardadas
  };
}

/* ------------------------------------------------------------------ editar fotos guardadas */

// Lista de pessoas para o campo de busca das fotos guardadas.
function apiPessoas() {
  return lerPessoas_(SpreadsheetApp.openById(planilhaId_())).pessoas.map(function (p) {
    return { chave: p.chave, nome: p.nome, numeros: p.numeros, temFoto: p.temFoto };
  });
}

// Apaga o link da pessoa nas abas, só onde a célula ainda tem o link desta foto.
function limparLink_(ss, pessoa, link) {
  pessoa.lugares.forEach(function (l) {
    var c = ss.getSheetByName(l.aba).getRange(l.linha, l.coluna);
    if (String(c.getValue()).trim() === link) c.setValue('');
  });
}

// A foto guardada era de outra pessoa: passa para a pessoa certa (nome do arquivo e links na planilha).
function apiTrocarPessoa(pedido) {
  if (!pedido || !pedido.de || !pedido.para) throw new Error('Pessoa faltando.');
  if (pedido.de === pedido.para) return apiEstado();
  var ss = SpreadsheetApp.openById(planilhaId_());
  var pessoas = lerPessoas_(ss).pessoas;
  var nova = acharPessoa_(ss, pedido.para, pessoas);
  var trava = esperarVez_();
  try {
    var d = lerDados_();
    var foto = d.registro[pedido.de];
    if (!foto) throw new Error('Essa foto não está mais guardada.');
    // Primeiro o link vai para a pessoa certa, depois sai da errada e por último a foto antiga dela vai para a
    // lixeira: se o Drive falhar no meio, nenhum link fica apontando para foto apagada.
    nova.lugares.forEach(function (l) {
      ss.getSheetByName(l.aba).getRange(l.linha, l.coluna).setValue(foto.link);
    });
    var antiga = pessoas.filter(function (p) { return p.chave === pedido.de; })[0];
    if (antiga) limparLink_(ss, antiga, foto.link);
    SpreadsheetApp.flush();
    var dela = d.registro[nova.chave];
    delete d.registro[pedido.de];
    foto.nome = nova.nome;
    d.registro[nova.chave] = foto;
    salvarDados_(d);
    // Se a pessoa certa já tinha foto, a dela vai para a lixeira (fica só a foto trocada).
    if (dela && dela.fileId && dela.fileId !== foto.fileId) {
      try { DriveApp.getFileById(dela.fileId).setTrashed(true); } catch (e) { /* já apagada; sobra na pasta */ }
    }
    // Renomear o arquivo é só para organizar a pasta: se o Drive recusar, a troca continua.
    try {
      var arquivo = DriveApp.getFileById(foto.fileId);
      var ext = (arquivo.getName().match(/\.([A-Za-z0-9]{2,5})$/) || [, 'jpg'])[1];
      arquivo.setName((nova.numeros.length ? nova.numeros[0] + ' - ' : '') + nova.nome + '.' + ext);
    } catch (e) { console.warn('Não renomeou ' + foto.fileId + ': ' + (e && e.message || e)); }
  } finally {
    trava.releaseLock();
  }
  return apiEstado();
}

// Exclui a foto guardada: vai para a lixeira do Drive e o link sai da planilha.
function apiExcluirFoto(chave) {
  var ss = SpreadsheetApp.openById(planilhaId_());
  var pessoas = lerPessoas_(ss).pessoas;
  var trava = esperarVez_();
  try {
    var d = lerDados_();
    var foto = d.registro[chave];
    if (!foto) return apiEstado();
    var pessoa = pessoas.filter(function (p) { return p.chave === chave; })[0];
    if (pessoa) limparLink_(ss, pessoa, foto.link);
    SpreadsheetApp.flush();
    delete d.registro[chave];
    salvarDados_(d);
    try { DriveApp.getFileById(foto.fileId).setTrashed(true); } catch (e) { /* já apagada; sobra na pasta */ }
  } finally {
    trava.releaseLock();
  }
  return apiEstado();
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

// Fila do Drive, registro das fotos e troca da foto antiga, com Google Apps Script imitado.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function montar() {
  const arquivos = {};
  let seq = 0;
  const pasta = {
    getId: () => 'PASTA', isTrashed: () => false,
    createFile: (a, b, c) => {
      const nome = typeof a === 'string' ? a : a.getName();
      const conteudo = typeof a === 'string' ? b : a.dados;
      return novoArquivo(nome, conteudo);
    },
    getFiles: () => iter(Object.values(arquivos).filter((f) => !f.lixo)),
    getFilesByName: (n) => iter(Object.values(arquivos).filter((f) => !f.lixo && f.nome === n)),
  };
  function novoArquivo(nome, conteudo) {
    const id = 'F' + (++seq);
    const f = arquivos[id] = {
      id, nome, conteudo, lixo: false, compartilhado: false,
      getId: () => id, getName: () => f.nome, isTrashed: () => f.lixo,
      setTrashed: (v) => { f.lixo = v; },
      setSharing: () => { f.compartilhado = true; },
      setContent: (t) => { f.conteudo = t; },
      getBlob: () => ({ getDataAsString: () => f.conteudo }),
      makeCopy: (n) => novoArquivo(n, f.conteudo),
      setName: (n) => { f.nome = n; },
    };
    return f;
  }
  function iter(l) { let i = 0; return { hasNext: () => i < l.length, next: () => l[i++] }; }
  const linhas = {
    'Página1': [['NOME', 'NÚMERO', 'FOTOS'], ['BRUNO VITOR', 1, ''], ['KAREN', 76, '']],
    'FUNÇÕES': [['NOME', 'NÚMERO', 'FOTOS'], ['KAREN', 76, ''], ['BRUNO VITOR', 1, '']],
  };
  const aba = (nome) => ({
    getName: () => nome,
    getLastRow: () => linhas[nome].length,
    getLastColumn: () => 3,
    getRange: (r, c, nr = 1, nc = 1) => ({
      getValues: () => linhas[nome].slice(r - 1, r - 1 + nr).map((l) => l.slice(c - 1, c - 1 + nc)),
      setValue: (v) => { linhas[nome][r - 1][c - 1] = v; },
      getValue: () => linhas[nome][r - 1][c - 1],
    }),
  });
  const ss = {
    getId: () => 'PLAN', getName: () => 'Plan', getUrl: () => 'u',
    getSheets: () => Object.keys(linhas).map(aba), getSheetByName: aba,
  };
  const props = {};
  let gatilhos = [];
  const ctx = {
    SpreadsheetApp: { openById: () => ss, flush: () => {} },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = v; },
      deleteProperty: (k) => { delete props[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
    ScriptApp: {
      getProjectTriggers: () => gatilhos,
      newTrigger: (h) => ({ timeBased: () => ({ everyMinutes: () => ({ create: () => {
        gatilhos.push({ getHandlerFunction: () => h }); } }) }) }),
      deleteTrigger: (t) => { gatilhos = gatilhos.filter((x) => x !== t); },
    },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 1 }, Permission: { VIEW: 1 },
      getFolderById: () => pasta,
      getFileById: (id) => { if (id === 'PLAN') return { getParents: () => iter([]) }; if (!arquivos[id]) throw new Error('não existe'); return arquivos[id]; },
      getRootFolder: () => ({ getFoldersByName: () => iter([pasta]), createFolder: () => pasta }),
    },
    Utilities: { newBlob: (d, t, n) => ({ dados: d, nome: n, getName() { return this.nome; }, setName(x) { this.nome = x; } }),
      base64Decode: (b) => b },
    Date, JSON, Math, Object, String, Number,
  };
  vm.createContext(ctx);
  for (const f of ['Combinar.gs', 'Codigo.gs']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), ctx);
  return { ctx, arquivos, linhas, gatilhos: () => gatilhos, novoArquivo };
}

test('fila do Drive continua sozinha, grava nas duas abas e desliga a rotina no fim', () => {
  const m = montar();
  const original = m.novoArquivo('bruno.jpg', 'foto1');
  m.ctx.apiEnfileirar([{ driveId: original.id, chave: 'BRUNO VITOR', arquivo: 'bruno.jpg' }]);
  assert.strictEqual(m.gatilhos().length, 1, 'rotina ligada');
  m.ctx.processarFila();
  const link = m.linhas['Página1'][1][2];
  assert.match(link, /^https:\/\/drive\.google\.com\/file\/d\/F\d+\/view$/);
  assert.strictEqual(m.linhas['FUNÇÕES'][2][2], link);
  assert.strictEqual(original.lixo, false, 'foto original do usuário não é mexida');
  assert.strictEqual(m.gatilhos().length, 0, 'rotina desligada');
  const est = m.ctx.apiEstado();
  assert.strictEqual(est.fila.faltam, 0);
  assert.strictEqual(est.fila.feitas, 1);
  assert.strictEqual(est.guardadas.length, 1);
  assert.strictEqual(est.guardadas[0].nome, 'BRUNO VITOR');
});

test('foto nova da mesma pessoa apaga a antiga e a mesma foto não volta para a lista', () => {
  const m = montar();
  m.ctx.apiEnviar({ chave: 'KAREN', arquivo: 'k1.jpg', origem: 'k1.jpg|10', mime: 'image/jpeg', base64: 'a' });
  const primeira = m.ctx.apiEstado().guardadas[0].fileId;
  m.ctx.apiEnviar({ chave: 'KAREN', arquivo: 'k2.png', origem: 'k2.png|20', mime: 'image/png', base64: 'b' });
  const segunda = m.ctx.apiEstado().guardadas[0].fileId;
  assert.notStrictEqual(primeira, segunda);
  assert.strictEqual(m.arquivos[primeira].lixo, true, 'antiga foi para a lixeira');
  assert.strictEqual(m.arquivos[segunda].lixo, false);
  assert.strictEqual(m.arquivos[segunda].nome, '76 - KAREN.png');
  const prep = m.ctx.apiPreparar([{ nome: 'k2.png', origem: 'k2.png|20' }, { nome: 'KAREN.jpg', origem: 'k3.jpg|30' }]);
  assert.strictEqual(prep.resultado[0].ja, true);
  assert.strictEqual(prep.resultado[1].ja, undefined);
  assert.match(prep.resultado[1].motivo, /trocar a atual/);
});

test('trocar a pessoa de uma foto guardada e excluir', () => {
  const m = montar();
  m.ctx.apiEnviar({ chave: 'BRUNO VITOR', arquivo: 'x.jpg', origem: 'x|1', mime: 'image/jpeg', base64: 'a' });
  m.ctx.apiEnviar({ chave: 'KAREN', arquivo: 'k.jpg', origem: 'k|1', mime: 'image/jpeg', base64: 'b' });
  const foto = m.ctx.apiEstado().guardadas.find((g) => g.chave === 'BRUNO VITOR');
  const daKaren = m.ctx.apiEstado().guardadas.find((g) => g.chave === 'KAREN');
  // A foto "do Bruno" era na verdade da Karen.
  const est = m.ctx.apiTrocarPessoa({ de: 'BRUNO VITOR', para: 'KAREN' });
  assert.deepStrictEqual(est.guardadas.map((g) => g.chave), ['KAREN']);
  assert.strictEqual(est.guardadas[0].fileId, foto.fileId);
  assert.strictEqual(m.arquivos[foto.fileId].nome, '76 - KAREN.jpg');
  assert.strictEqual(m.arquivos[daKaren.fileId].lixo, true, 'foto antiga da Karen foi para a lixeira');
  assert.strictEqual(m.linhas['Página1'][1][2], '', 'link saiu do Bruno');
  assert.strictEqual(m.linhas['FUNÇÕES'][2][2], '');
  assert.strictEqual(m.linhas['Página1'][2][2], foto.link, 'link foi para a Karen');
  assert.strictEqual(m.linhas['FUNÇÕES'][1][2], foto.link);
  const depois = m.ctx.apiExcluirFoto('KAREN');
  assert.strictEqual(depois.guardadas.length, 0);
  assert.strictEqual(m.arquivos[foto.fileId].lixo, true);
  assert.strictEqual(m.linhas['Página1'][2][2], '');
  assert.strictEqual(m.linhas['FUNÇÕES'][1][2], '');
});

// Testa a procura da pessoa pelo nome do arquivo e a leitura das abas, com nomes da planilha real.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const ctx = {};
vm.createContext(ctx);
for (const f of ['Combinar.gs', 'Codigo.gs']) vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8'), ctx);

const titulo = ['NOME', 'CARGO', 'SETOR', 'NÚMERO IDENTIFICADOR ', 'COR', 'FOTOS'];
const pagina1 = [titulo,
  ['BRUNO VITOR', 'SUPERVISOR', 'INFORMÁTICA', 1, '', ''],
  ['PEDRO DIAS', 'ESTAGIÁRIO', 'INFORMÁTICA', 23, '', ''],
  ['PEDRO HENRIQUE', 'ESTAGIÁRIO', 'INFORMÁTICA', 24, '', ''],
  ['PEDRO HENRIQUE BISPO DOS SANTOS', 'ESTAGIÁRIO', 'SMARTPHONES', 27, '', ''],
  ['ALICE VITORIA DE OLIVEIRA SILVA', 'ESTAGIÁRIO', 'SMARTPHONES', 21, 'ROSA', ''],
  ['KAREN KATLIN', 'FUNCIONÁRIO', 'SMARTPHONES', 52, '', ''],
  ['KAREN', 'ESTAGIÁRIO', 'SMARTPHONES', 76, 'AZUL', ''],
  ['KEVILIN', 'FUNCIONÁRIO', 'ELETROELETRÔNICOS', 46, '', ''],
  ['ARTHUR LUÍS', 'FUNCIONÁRIO', 'ELETROELETRÔNICOS', 45, '', ''],
  ['JOÃO PEDRO', 'SUPERVISOR', 'INFORMÁTICA', 8, '', 'https://antigo'],
  ['', '', '', 67, '', ''],
  ['MICHELE', 'GESTÃO', '', 0, '', ''],
];
const funcoes = [titulo,
  ['BRUNO VITOR', 'SUPERVISOR', 'INFORMÁTICA', 1, '', ''],
  ['JOAO VITOR SENNA', 'ESTAGIÁRIO', 'SMARTPHONES', 23, '', ''],
  ['ARTHUR LUIZ', 'ESTAGIÁRIO', 'SMARTPHONES', 24, '', ''],
  ['DAVI  FERNANDO ', 'FUNCIONÁRIO', 'ELETROELETRÔNICOS', 33, '', ''],
  ['KAREN', 'ESTAGIÁRIO', 'SMARTPHONES', 76, 'AZUL', ''],
];

function aba(nome, linhas) {
  return {
    getName: () => nome,
    getLastRow: () => linhas.length,
    getLastColumn: () => linhas[0].length,
    getRange: (r, c, nr, nc) => ({ getValues: () => linhas.slice(r - 1, r - 1 + nr).map((l) => l.slice(c - 1, c - 1 + nc)) }),
  };
}
const ss = { getSheets: () => [aba('Página1', pagina1), aba('FUNÇÕES ', funcoes), aba('Vazia', [['x']])] };
const { pessoas, abas } = ctx.lerPessoas_(ss);
const quem = (arquivo) => ctx.combinar_([arquivo], pessoas)[0];

test('lê as abas com NOME e FOTOS e junta a mesma pessoa', () => {
  assert.deepStrictEqual([...abas], ['Página1', 'FUNÇÕES ']);
  const bruno = pessoas.find((p) => p.chave === 'BRUNO VITOR');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(bruno.lugares.map((l) => [l.aba, l.linha, l.coluna]))), [['Página1', 2, 6], ['FUNÇÕES ', 2, 6]]);
  assert.ok(pessoas.find((p) => p.chave === 'JOAO PEDRO').temFoto);
  assert.ok(pessoas.find((p) => p.chave === 'DAVI FERNANDO'), 'espaço duplo no nome');
  assert.ok(!pessoas.find((p) => p.chave === ''), 'linha sem nome fica fora');
});

test('nome igual, com acento, minúscula ou extensão diferente', () => {
  assert.strictEqual(quem('Bruno Vitor.jpg').chave, 'BRUNO VITOR');
  assert.strictEqual(quem('joao pedro.HEIC').chave, 'JOAO PEDRO');
  assert.strictEqual(quem('ARTHUR_LUÍS.png').chave, 'ARTHUR LUIS');
  assert.strictEqual(quem('ARTHUR LUIZ.png').chave, 'ARTHUR LUIZ');
  assert.strictEqual(quem('KAREN.jpg').chave, 'KAREN');
  assert.strictEqual(quem('foto - KAREN KATLIN (1).jpeg').chave, 'KAREN KATLIN');
});

test('nome mais longo ganha e parte única do nome basta', () => {
  assert.strictEqual(quem('PEDRO HENRIQUE BISPO DOS SANTOS.jpg').chave, 'PEDRO HENRIQUE BISPO DOS SANTOS');
  assert.strictEqual(quem('PEDRO HENRIQUE.jpg').chave, 'PEDRO HENRIQUE');
  assert.strictEqual(quem('alice.jpg').chave, 'ALICE VITORIA DE OLIVEIRA SILVA');
  assert.strictEqual(quem('Michele.jpg').nivel, 'certo');
});

test('parte do nome de várias pessoas vira escolha', () => {
  const r = quem('PEDRO.jpg');
  assert.strictEqual(r.nivel, 'duvida');
  assert.strictEqual(r.chave, '');
  assert.ok(r.opcoes.includes('PEDRO DIAS') && r.opcoes.includes('PEDRO HENRIQUE'));
});

test('número no arquivo desempata ou identifica', () => {
  assert.strictEqual(quem('PEDRO 23.jpg').chave, 'PEDRO DIAS');
  assert.strictEqual(quem('1.jpg').chave, 'BRUNO VITOR');
  assert.strictEqual(quem('076.jpg').chave, 'KAREN');
  assert.strictEqual(quem('0.jpg').chave, 'MICHELE');
  const dois = quem('23.jpg');
  assert.strictEqual(dois.nivel, 'duvida', 'nº 23 é PEDRO DIAS numa aba e JOAO VITOR SENNA na outra');
  assert.deepStrictEqual([...dois.opcoes].sort(), ['JOAO VITOR SENNA', 'PEDRO DIAS']);
  assert.strictEqual(quem('67.jpg').nivel, 'nada');
});

test('letra trocada fica para conferir; nome desconhecido não acha', () => {
  const r = quem('KEVELIN.jpg');
  assert.strictEqual(r.nivel, 'conferir');
  assert.strictEqual(r.chave, 'KEVILIN');
  assert.strictEqual(quem('IMG_4021.jpg').nivel, 'nada');
  assert.strictEqual(quem('FULANO DE TAL.jpg').nivel, 'nada');
});

test('duas fotos para a mesma pessoa: a segunda fica para escolher', () => {
  const r = ctx.combinar_(['BRUNO VITOR.jpg', 'bruno vitor 2.jpg'], pessoas);
  assert.strictEqual(r[0].chave, 'BRUNO VITOR');
  assert.strictEqual(r[1].nivel, 'duvida');
  assert.deepStrictEqual([...r[1].opcoes], ['BRUNO VITOR']);
});

test('link da pasta do Drive vira o ID da pasta', () => {
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz012345';
  assert.strictEqual(ctx.idDaPasta_('https://drive.google.com/drive/folders/' + id + '?usp=sharing'), id);
  assert.strictEqual(ctx.idDaPasta_('https://drive.google.com/drive/u/0/folders/' + id), id);
  assert.strictEqual(ctx.idDaPasta_('https://drive.google.com/open?id=' + id), id);
  assert.strictEqual(ctx.idDaPasta_(id), id);
  assert.throws(() => ctx.idDaPasta_('fotos da equipe'));
});

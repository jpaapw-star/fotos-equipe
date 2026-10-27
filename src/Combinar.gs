/*
 * Acha, pelo nome do arquivo, de qual pessoa é cada foto.
 * Funções puras (sem planilha), testadas em tests/combinar.test.js.
 *
 * pessoas: [{ chave, nome, numeros: ['1', ...] }]  (chave = nome normalizado)
 * Resultado de cada arquivo: { arquivo, chave, nivel, motivo, opcoes }
 *   nivel 'certo'    — achou uma pessoa só, vai direto
 *   nivel 'conferir' — parecido (letra trocada), vem marcado mas é bom olhar
 *   nivel 'duvida'   — mais de uma pessoa possível, você escolhe na lista
 *   nivel 'nada'     — não achou ninguém
 */

var PALAVRAS_IGNORADAS = ['FOTO', 'FOTOS', 'COPIA', 'COPY', 'PERFIL', 'IMG', 'IMAGEM', 'WHATSAPP', 'IMAGE'];

function normalizar_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
}

function semExtensao_(arquivo) {
  var nome = String(arquivo || '').split('/').pop();
  return nome.replace(/\.[A-Za-z0-9]{2,5}$/, '');
}

function distancia_(a, b) {
  if (a === b) return 0;
  var ant = [], i, j;
  for (j = 0; j <= b.length; j++) ant[j] = j;
  for (i = 1; i <= a.length; i++) {
    var atual = [i];
    for (j = 1; j <= b.length; j++) {
      atual[j] = Math.min(ant[j] + 1, atual[j - 1] + 1, ant[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    ant = atual;
  }
  return ant[b.length];
}

function contemTodas_(maior, menor) {
  for (var i = 0; i < menor.length; i++) if (maior.indexOf(menor[i]) < 0) return false;
  return true;
}

function combinarUm_(arquivo, pessoas) {
  var todas = normalizar_(semExtensao_(arquivo)).split(' ').filter(function (t) {
    return t && PALAVRAS_IGNORADAS.indexOf(t) < 0;
  });
  var palavras = todas.filter(function (t) { return !/^\d+$/.test(t); });
  var numeros = todas.filter(function (t) { return /^\d+$/.test(t); }).map(function (n) { return String(Number(n)); });
  var r = { arquivo: arquivo, chave: '', nivel: 'nada', motivo: 'Nenhum nome ou número da planilha no nome do arquivo.', opcoes: [] };
  var chaves = function (lista) { return lista.map(function (p) { return p.chave; }); };

  if (palavras.length) {
    var texto = palavras.join(' ');
    var iguais = pessoas.filter(function (p) { return p.chave === texto; });
    if (iguais.length) return certo_(r, iguais[0].chave, 'Nome igual ao da planilha.');

    // O nome inteiro da pessoa está no arquivo (ex.: "BRUNO VITOR SUPERVISOR.jpg"). Fica o nome mais longo.
    var dentro = pessoas.filter(function (p) { return contemTodas_(palavras, p.chave.split(' ')); });
    if (dentro.length) {
      var maior = Math.max.apply(null, dentro.map(function (p) { return p.chave.split(' ').length; }));
      dentro = dentro.filter(function (p) { return p.chave.split(' ').length === maior; });
      dentro = desempatarPorNumero_(dentro, numeros);
      if (dentro.length === 1) return certo_(r, dentro[0].chave, 'Nome da planilha achado no arquivo.');
      return duvida_(r, chaves(dentro));
    }

    // O arquivo tem só parte do nome (ex.: "ALICE.jpg").
    var parte = pessoas.filter(function (p) { return contemTodas_(p.chave.split(' '), palavras); });
    parte = desempatarPorNumero_(parte, numeros);
    if (parte.length === 1) return certo_(r, parte[0].chave, 'Parte do nome bate com uma pessoa só.');
    if (parte.length > 1) return duvida_(r, chaves(parte));

    // Letra trocada ou faltando (ex.: "KEVELIN" → "KEVILIN"). Vem marcado para conferir.
    var melhores = [], menor = Infinity;
    pessoas.forEach(function (p) {
      var d = distancia_(texto, p.chave);
      if (d < menor) { menor = d; melhores = [p]; } else if (d === menor) melhores.push(p);
    });
    var limite = Math.max(1, Math.floor(texto.length * 0.15));
    if (melhores.length && menor <= limite) {
      melhores = desempatarPorNumero_(melhores, numeros);
      if (melhores.length === 1) {
        r.chave = melhores[0].chave;
        r.nivel = 'conferir';
        r.motivo = 'Nome parecido (' + menor + (menor === 1 ? ' letra diferente' : ' letras diferentes') + '). Confira.';
        r.opcoes = [r.chave];
        return r;
      }
      return duvida_(r, chaves(melhores));
    }
  }

  if (numeros.length) {
    var porNumero = pessoas.filter(function (p) { return p.numeros.indexOf(numeros[0]) >= 0; });
    if (porNumero.length === 1) return certo_(r, porNumero[0].chave, 'Número identificador ' + numeros[0] + '.');
    if (porNumero.length > 1) {
      duvida_(r, chaves(porNumero));
      r.motivo = 'O número ' + numeros[0] + ' é de pessoas diferentes nas abas. Escolha quem é.';
      return r;
    }
    r.motivo = numeros[0].length > 4
      ? 'O nome do arquivo é só um número longo, sem nome de pessoa. Escolha quem é.'
      : 'O número ' + numeros[0] + ' não tem pessoa na planilha.';
  }
  return r;
}

function desempatarPorNumero_(lista, numeros) {
  if (lista.length < 2 || !numeros.length) return lista;
  var comNumero = lista.filter(function (p) { return p.numeros.indexOf(numeros[0]) >= 0; });
  return comNumero.length ? comNumero : lista;
}

function certo_(r, chave, motivo) {
  r.chave = chave; r.nivel = 'certo'; r.motivo = motivo; r.opcoes = [chave];
  return r;
}

function duvida_(r, opcoes) {
  r.chave = ''; r.nivel = 'duvida'; r.opcoes = opcoes;
  r.motivo = 'Pode ser mais de uma pessoa. Escolha quem é.';
  return r;
}

function combinar_(arquivos, pessoas) {
  var lista = arquivos.map(function (a) { return combinarUm_(a, pessoas); });
  // Duas fotos para a mesma pessoa: a segunda fica para você decidir.
  var vistos = {};
  lista.forEach(function (r) {
    if (!r.chave) return;
    if (vistos[r.chave]) {
      r.nivel = 'duvida';
      r.motivo = 'Outra foto (' + vistos[r.chave] + ') já é dessa pessoa. Escolha se esta substitui.';
      r.chave = '';
    } else {
      vistos[r.chave] = r.arquivo;
    }
  });
  return lista;
}

// =====================================================================
//  Leitura dos .tsv que o UNLOAD gerou + utilidades comuns.
//
//  Peculiaridades do arquivo do UNLOAD:
//    * o separador é a sequência literal \ + t, NÃO um TAB — é por isso
//      que o split é em '\\t' e não em '\t'. Isso é até bom: TAB dentro
//      de um campo de texto não quebra a linha.
//    * encoding cp1252 / latin1.
//    * sem cabeçalho — os nomes de coluna ficam num .cols ao lado.
// =====================================================================
const fs = require('fs');
const path = require('path');

// argv[2] = arquivo de saída, argv[3] = pasta dos .tsv
const OUT = (process.argv[3] || 'C:/dbtmp/out').replace(/\\/g, '/');

// ---------------------------------------------------------------------
//  Filtro "só empresas ativas" (DOM_SO_ATIVAS=1)
//
//  Aplicado aqui na leitura, e não no SQL, para valer igual em todos os
//  relatórios: qualquer .tsv com coluna codi_emp perde as linhas de
//  empresa com geempre.stat_emp <> 'A'. Sem a variável, nada muda.
// ---------------------------------------------------------------------
let ATIVAS = null;
function ativas() {
  if (ATIVAS) return ATIVAS;
  ATIVAS = new Set();
  const { cols, rows } = readBruto('01_empresas');
  const iCod = cols.indexOf('codi_emp'), iSt = cols.indexOf('stat_emp');
  for (const r of rows) if (r[iSt] === 'A') ATIVAS.add(r[iCod]);
  return ATIVAS;
}

function read(nome) {
  const t = readBruto(nome);
  if (process.env.DOM_SO_ATIVAS !== '1') return t;
  const i = t.cols.indexOf('codi_emp');
  if (i < 0) return t;
  const set = ativas();
  return { cols: t.cols, rows: t.rows.filter(r => set.has(r[i])) };
}

function readBruto(nome) {
  const file = path.join(OUT, nome + '.tsv');
  if (!fs.existsSync(file)) return { cols: [], rows: [] };
  const colsFile = path.join(OUT, nome + '.cols');
  let cols = [];
  if (fs.existsSync(colsFile)) {
    cols = fs.readFileSync(colsFile, 'utf8').replace(/^\uFEFF/, '').trim().split('\t');
  }
  const txt = fs.readFileSync(file).toString('latin1');
  const rows = [];
  for (const line of txt.split(/\r?\n/)) {
    if (!line) continue;
    rows.push(line.split('\\t').map(v => v.trim()));
  }
  return { cols, rows };
}

function readObj(nome) {
  const { cols, rows } = read(nome);
  return rows.map(r => {
    const o = {};
    cols.forEach((c, i) => { o[c] = r[i] === undefined ? '' : r[i]; });
    return o;
  });
}

const existe = nome => fs.existsSync(path.join(OUT, nome + '.tsv'));

const num = v => {
  if (v === '' || v === null || v === undefined) return 0;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? 0 : n;
};
const int = v => Math.round(num(v));

// 'AAAA-MM-DD' -> 'AAAA-MM'
const comp = d => (d && d.length >= 7) ? d.slice(0, 7) : '';
const mm = (ano, mes) => ano + '-' + String(mes).padStart(2, '0');

const cnpjFmt = c => {
  const d = String(c || '').replace(/\D/g, '');
  if (d.length === 14) return d.replace(/^(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})$/, '$1.$2.$3/$4-$5');
  if (d.length === 11) return d.replace(/^(\d{3})(\d{3})(\d{3})(\d{2})$/, '$1.$2.$3-$4');
  return String(c || '').trim();
};

// geloguser.sist_log -> GEMODULOS.NOME
const MODULOS = {
  '1': 'Contabilidade', '2': 'Escrita Fiscal', '3': 'Honorários', '4': 'Patrimônio',
  '5': 'Escrita Fiscal', '6': 'Lalur', '7': 'Atualizar', '8': 'Protocolos',
  '9': 'Administrar', '11': 'Outros', '12': 'Folha', '13': 'Ponto Eletrônico',
  '14': 'Auditoria', '15': 'Registro', '19': 'Processos'
};
const MODS = ['Escrita Fiscal', 'Contabilidade', 'Folha', 'Honorários',
              'Patrimônio', 'Lalur', 'Protocolos', 'Registro'];

// ---------------------------------------------------------------------
//  Departamento do escritório
//
//  O Domínio não tem cadastro de departamento — o que ele grava é o MÓDULO
//  em que o usuário estava (geloguser.sist_log). O mapa abaixo é a convenção
//  do escritório: Lalur é apuração de IRPJ/CSLL, então conta como Fiscal;
//  Patrimônio é imobilizado, conta como Contábil; Ponto é Pessoal.
//  Se a divisão do escritório for outra, edite AQUI e todos os relatórios
//  acompanham.
// ---------------------------------------------------------------------
const DEPARTAMENTO = {
  '1': 'Contábil',        // Contabilidade
  '2': 'Fiscal',          // Escrita Fiscal (código alternativo)
  '3': 'Honorários',      // Honorários
  '4': 'Contábil',        // Patrimônio
  '5': 'Fiscal',          // Escrita Fiscal
  '6': 'Fiscal',          // Lalur
  '7': 'Administrativo',  // Atualizar
  '8': 'Administrativo',  // Protocolos
  '9': 'Administrativo',  // Administrar
  '11': 'Administrativo',
  '12': 'Pessoal',        // Folha
  '13': 'Pessoal',        // Ponto Eletrônico
  '14': 'Administrativo', // Auditoria
  '15': 'Societário',     // Registro
  '19': 'Administrativo'  // Processos
};
const DEPTS = ['Fiscal', 'Contábil', 'Pessoal', 'Honorários', 'Societário', 'Administrativo'];
const dept = sistLog => DEPARTAMENTO[String(sistLog)] || 'Administrativo';

// geimposto.codi_sis
const SISTEMA = { '5': 'Fiscal', '6': 'Contábil/Lalur', '12': 'Folha' };

// 'hh:mm:ss' -> minutos
function hms(t) {
  if (!t) return null;
  const m = String(t).match(/(\d{1,2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  return (+m[1]) * 60 + (+m[2]) + (+m[3]) / 60;
}

// ---------------------------------------------------------------------
//  Dimensão empresa, usada por todos os geradores
// ---------------------------------------------------------------------
function carregaEmpresas() {
  const cnae = {};
  for (const c of readObj('01b_cnae')) {
    cnae[c.codi_emp] = { cod: (c.i_cnae20 || '').trim(), ramo: (c.ramo_emp || '').trim() };
  }
  const cnaeDesc = {};
  for (const c of readObj('24_cnae20')) cnaeDesc[(c.codigo || '').trim()] = (c.descricao || '').trim();

  const emp = {};
  for (const e of readObj('01_empresas')) {
    const doc = String(e.cgce_emp || '').replace(/\D/g, '');
    const cn = cnae[e.codi_emp] || {};
    emp[e.codi_emp] = {
      cod: int(e.codi_emp),
      nome: (e.nome_emp || '').trim(),
      razao: (e.razao_emp || '').trim(),
      doc, pf: doc.length === 11, cnpj: cnpjFmt(e.cgce_emp),
      ie: (e.iest_emp || '').trim(),
      im: (e.imun_emp || '').trim(),
      cidade: (e.cida_emp || '').trim(),
      uf: (e.esta_emp || '').trim(),
      status: e.stat_emp === 'A' ? 'Ativa' : (e.stat_emp === 'I' ? 'Inativa' : (e.stat_emp || '')),
      simples: (e.simples_emp === '1' || e.simples_emp === 'S') ? 'Sim' : '',
      cnae: cn.cod || '',
      atividade: cn.ramo || cnaeDesc[cn.cod] || ''
    };
  }
  const vazio = c => ({
    cod: int(c), nome: '(empresa ' + c + ')', razao: '', cnpj: '', ie: '', im: '',
    cidade: '', uf: '', status: '', simples: '', cnae: '', atividade: '', pf: false
  });
  return { emp, cnaeDesc, E: c => emp[c] || vazio(c) };
}

// ---------------------------------------------------------------------
//  Planilha: cria aba já com cabeçalho formatado e filtro
// ---------------------------------------------------------------------
const MOEDA = { numFmt: '#,##0.00' };

function fazAba(wb, nome, cols, linhas, cor) {
  const ws = wb.addWorksheet(nome, { views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }] });
  ws.columns = cols;
  ws.addRows(linhas);
  const r1 = ws.getRow(1);
  r1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: cor || 'FF1F4E79' } };
  r1.alignment = { vertical: 'middle', wrapText: true };
  r1.height = 32;
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  return ws;
}

const destino = () => process.argv[2];

module.exports = {
  read, readObj, existe, num, int, comp, mm, cnpjFmt,
  MODULOS, MODS, SISTEMA, DEPARTAMENTO, DEPTS, dept, hms, MOEDA,
  carregaEmpresas, fazAba, destino, OUT
};

// 03-CONTABIL-BALANCOS.xlsx — balanço patrimonial, DRE e balancete
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, num, int, MOEDA, fazAba, destino } = L;

const { E } = L.carregaEmpresas();
const DEST = destino();

// 1º dígito da classificação da conta define grupo e natureza
const GRUPOS = {
  '1': { nome: 'ATIVO',                          nat: 'D', patrimonial: true },
  '2': { nome: 'PASSIVO E PATRIMÔNIO LÍQUIDO',   nat: 'C', patrimonial: true },
  '3': { nome: 'CUSTOS E DESPESAS',              nat: 'D', patrimonial: false },
  '4': { nome: 'RECEITAS',                       nat: 'C', patrimonial: false },
  '5': { nome: 'APURAÇÃO DO RESULTADO',          nat: 'C', patrimonial: false },
  '6': { nome: 'ENCERRAMENTO DO EXERCÍCIO',      nat: 'C', patrimonial: false },
};

// plano de contas
const plano = {};
const nomeClas = {};
for (const p of readObj('20_plano_contas')) {
  const clas = (p.clas_cta || '').trim();
  plano[p.codi_emp + '|' + p.codi_cta] = {
    nome: (p.nome_cta || '').trim(), clas,
    tipo: (p.tipo_cta || '').trim() === 'A' ? 'Analítica' : 'Sintética'
  };
  const k = p.codi_emp + '|' + clas;
  if (clas && !nomeClas[k]) nomeClas[k] = (p.nome_cta || '').trim();
}

// balancete
const raw = readObj('22_balancete').map(b => {
  const p = plano[b.codi_emp + '|' + b.codi_cta] || { nome: '(conta ' + b.codi_cta + ')', clas: '', tipo: '' };
  return {
    cod: b.codi_emp, ano: int(b.ano), conta: b.codi_cta,
    clas: p.clas, nome: p.nome, tipo: p.tipo, grupo: (p.clas || '')[0] || '',
    debito: num(b.debito), credito: num(b.credito), n: int(b.lancamentos)
  };
}).filter(r => r.ano >= 1995 && r.ano <= 2100);

const saldoNat = r => {
  const nat = (GRUPOS[r.grupo] || {}).nat || 'D';
  return nat === 'D' ? r.debito - r.credito : r.credito - r.debito;
};

// saldo acumulado ao longo dos anos, por empresa+conta
raw.sort((a, b) => a.cod.localeCompare(b.cod) || a.conta.localeCompare(b.conta) || a.ano - b.ano);
let chave = '', acumulado = 0;
for (const r of raw) {
  const k = r.cod + '|' + r.conta;
  if (k !== chave) { chave = k; acumulado = 0; }
  r.saldoAno = saldoNat(r);
  acumulado += r.saldoAno;
  r.saldoAcum = acumulado;
}

const wb = new ExcelJS.Workbook();
wb.creator = 'Extração Domínio'; wb.created = new Date();

// ---------- resumo por empresa e ano ----------
const res = {};
for (const r of raw) {
  const k = r.cod + '|' + r.ano;
  const o = res[k] || (res[k] = {
    cod: r.cod, ano: r.ano, ativo: 0, passivo: 0, receita: 0, despesa: 0, n: 0, contas: new Set()
  });
  o.n += r.n; o.contas.add(r.conta);
  // patrimonial usa saldo acumulado; resultado usa só o movimento do ano
  if (r.grupo === '1') o.ativo += r.saldoAcum;
  else if (r.grupo === '2') o.passivo += r.saldoAcum;
  else if (r.grupo === '4') o.receita += r.saldoAno;
  else if (r.grupo === '3') o.despesa += r.saldoAno;
}

fazAba(wb, 'Resumo por empresa e ano', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Ano', key: 'ano', width: 8 },
  { header: 'Ativo (saldo acumulado)', key: 'ativo', width: 20, style: MOEDA },
  { header: 'Passivo + PL (saldo acumulado)', key: 'passivo', width: 22, style: MOEDA },
  { header: 'Diferença Ativo − Passivo', key: 'dif', width: 19, style: MOEDA },
  { header: 'Receitas do ano', key: 'receita', width: 18, style: MOEDA },
  { header: 'Custos e despesas do ano', key: 'despesa', width: 20, style: MOEDA },
  { header: 'Resultado do exercício', key: 'result', width: 19, style: MOEDA },
  { header: 'Contas movimentadas', key: 'ncontas', width: 15 },
  { header: 'Lançamentos', key: 'n', width: 12 },
], Object.values(res).sort((a, b) => int(a.cod) - int(b.cod) || a.ano - b.ano).map(o => {
  const e = E(o.cod);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ano: o.ano,
    ativo: o.ativo, passivo: o.passivo, dif: o.ativo - o.passivo,
    receita: o.receita, despesa: o.despesa, result: o.receita - o.despesa,
    ncontas: o.contas.size, n: o.n
  };
}), 'FF7F3300');

// ---------- balanço patrimonial, por grupo de 2º nível ----------
const bp = {};
for (const r of raw) {
  const g = GRUPOS[r.grupo];
  if (!g || !g.patrimonial) continue;
  const n2 = r.clas.slice(0, 2);
  const k = r.cod + '|' + r.ano + '|' + n2;
  const o = bp[k] || (bp[k] = { cod: r.cod, ano: r.ano, grupo: r.grupo, n2, saldo: 0, n: 0 });
  o.saldo += r.saldoAcum; o.n += r.n;
}

fazAba(wb, 'Balanço Patrimonial', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Ano', key: 'ano', width: 8 },
  { header: 'Grupo', key: 'grupo', width: 30 },
  { header: 'Classificação', key: 'n2', width: 13 },
  { header: 'Descrição', key: 'desc', width: 42 },
  { header: 'Saldo em 31/12', key: 'saldo', width: 20, style: MOEDA },
  { header: 'Lançamentos', key: 'n', width: 12 },
], Object.values(bp).sort((a, b) => int(a.cod) - int(b.cod) || a.ano - b.ano || a.n2.localeCompare(b.n2)).map(o => {
  const e = E(o.cod);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ano: o.ano,
    grupo: GRUPOS[o.grupo].nome, n2: o.n2,
    desc: nomeClas[o.cod + '|' + o.n2] || '', saldo: o.saldo, n: o.n
  };
}), 'FF1F4E79');

// ---------- DRE ----------
const dre = {};
for (const r of raw) {
  if (r.grupo !== '3' && r.grupo !== '4') continue;
  const n2 = r.clas.slice(0, 2);
  const k = r.cod + '|' + r.ano + '|' + n2;
  const o = dre[k] || (dre[k] = { cod: r.cod, ano: r.ano, grupo: r.grupo, n2, valor: 0, n: 0 });
  o.valor += r.saldoAno; o.n += r.n;
}

fazAba(wb, 'DRE - Resultado', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Ano', key: 'ano', width: 8 },
  { header: 'Natureza', key: 'grupo', width: 22 },
  { header: 'Classificação', key: 'n2', width: 13 },
  { header: 'Descrição', key: 'desc', width: 42 },
  { header: 'Valor do ano', key: 'valor', width: 20, style: MOEDA },
  { header: 'Lançamentos', key: 'n', width: 12 },
], Object.values(dre).sort((a, b) => int(a.cod) - int(b.cod) || a.ano - b.ano || a.n2.localeCompare(b.n2)).map(o => {
  const e = E(o.cod);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ano: o.ano,
    grupo: GRUPOS[o.grupo].nome, n2: o.n2,
    desc: nomeClas[o.cod + '|' + o.n2] || '', valor: o.valor, n: o.n
  };
}), 'FF375623');

// ---------- balancete analítico ----------
fazAba(wb, 'Balancete analítico', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 40 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Ano', key: 'ano', width: 8 },
  { header: 'Conta', key: 'conta', width: 10 },
  { header: 'Classificação', key: 'clas', width: 16 },
  { header: 'Descrição da conta', key: 'desc', width: 42 },
  { header: 'Tipo', key: 'tipo', width: 11 },
  { header: 'Grupo', key: 'grupo', width: 28 },
  { header: 'Débito', key: 'debito', width: 17, style: MOEDA },
  { header: 'Crédito', key: 'credito', width: 17, style: MOEDA },
  { header: 'Saldo do ano', key: 'saldoAno', width: 17, style: MOEDA },
  { header: 'Saldo acumulado', key: 'saldoAcum', width: 17, style: MOEDA },
  { header: 'Lançamentos', key: 'n', width: 12 },
], raw.slice().sort((a, b) => int(a.cod) - int(b.cod) || a.ano - b.ano || a.clas.localeCompare(b.clas)).map(r => {
  const e = E(r.cod);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ano: r.ano, conta: int(r.conta),
    clas: r.clas, desc: r.nome, tipo: r.tipo,
    grupo: (GRUPOS[r.grupo] || {}).nome || '',
    debito: r.debito, credito: r.credito, saldoAno: r.saldoAno, saldoAcum: r.saldoAcum, n: r.n
  };
}), 'FF833C00');

wb.xlsx.writeFile(DEST).then(() => {
  const anos = raw.map(r => r.ano);
  console.log('   empresas: %d | linhas de balancete: %d | anos: %d a %d | lançamentos: %s',
    new Set(raw.map(r => r.cod)).size, raw.length,
    Math.min(...anos), Math.max(...anos),
    raw.reduce((a, r) => a + r.n, 0).toLocaleString('pt-BR'));
});

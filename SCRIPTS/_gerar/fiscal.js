// 01-FISCAL.xlsx  — faturamento e impostos apurados
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, num, int, comp, mm, MOEDA, SISTEMA, fazAba, destino } = L;

const { E } = L.carregaEmpresas();
const DEST = destino();

// ---------- movimento por empresa x competência ----------
const mov = {};
const M = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!mov[k]) mov[k] = { cod, mes, fat: 0, prod: 0, nsai: 0, compras: 0, nent: 0, imp: {} };
  return mov[k];
};

for (const r of readObj('02_faturamento')) {
  const o = M(r.codi_emp, mm(r.ano, r.mes));
  o.fat += num(r.faturamento); o.prod += num(r.valor_produtos); o.nsai += int(r.notas);
}
for (const r of readObj('03_compras')) {
  const o = M(r.codi_emp, mm(r.ano, r.mes));
  o.compras += num(r.compras); o.nent += int(r.notas);
}

// ---------- impostos: só os do fiscal (codi_sis 5 e 6) ----------
const impRows = [];
const siglasTotais = {};
for (const r of readObj('04_impostos')) {
  const sis = (r.codi_sis || '').trim();
  if (sis !== '5' && sis !== '6') continue;      // 12 = folha, sai em 02-FOLHA
  const sig = (r.sigla || '').trim();
  const mes = comp(r.competencia);
  const v = num(r.saldo_devedor);
  const o = M(r.codi_emp, mes);
  o.imp[sig] = (o.imp[sig] || 0) + v;
  siglasTotais[sig] = (siglasTotais[sig] || 0) + v;
  const e = E(r.codi_emp);
  impRows.push({
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes,
    sigla: sig, imposto: (r.nome_imposto || '').trim(), sistema: SISTEMA[sis] || sis,
    devedor: v, credor: num(r.saldo_credor), base: num(r.base_calculo),
    pago: num(r.valor_pago), apuracoes: int(r.apuracoes)
  });
}

// vira coluna só o tributo que realmente teve valor, do maior para o menor
const SIGLAS = Object.entries(siglasTotais).filter(([, v]) => v > 0)
  .sort((a, b) => b[1] - a[1]).map(([s]) => s);

// ---------- totais por empresa ----------
const tot = {};
for (const o of Object.values(mov)) {
  const t = tot[o.cod] || (tot[o.cod] = { cod: o.cod, fat: 0, nsai: 0, compras: 0, nent: 0, imp: {} });
  t.fat += o.fat; t.nsai += o.nsai; t.compras += o.compras; t.nent += o.nent;
  for (const [s, v] of Object.entries(o.imp)) t.imp[s] = (t.imp[s] || 0) + v;
}

const wb = new ExcelJS.Workbook();
wb.creator = 'Extração Domínio'; wb.created = new Date();

fazAba(wb, 'Resumo por empresa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'Razão social', key: 'razao', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Inscrição estadual', key: 'ie', width: 20 },
  { header: 'Inscrição municipal', key: 'im', width: 20 },
  { header: 'Cidade', key: 'cidade', width: 20 },
  { header: 'UF', key: 'uf', width: 6 },
  { header: 'Situação', key: 'status', width: 10 },
  { header: 'Simples', key: 'simples', width: 9 },
  { header: 'CNAE', key: 'cnae', width: 12 },
  { header: 'Atividade', key: 'atividade', width: 44 },
  { header: 'Faturamento', key: 'fat', width: 17, style: MOEDA },
  { header: 'Notas de saída', key: 'nsai', width: 13 },
  { header: 'Compras', key: 'compras', width: 17, style: MOEDA },
  { header: 'Notas de entrada', key: 'nent', width: 14 },
  { header: 'Total impostos apurados', key: 'impTot', width: 20, style: MOEDA },
  ...SIGLAS.map(s => ({ header: s, key: 'i_' + s, width: 14, style: MOEDA })),
], Object.values(tot).sort((a, b) => b.fat - a.fat).map(t => {
  const e = E(String(t.cod));
  const o = {
    cod: t.cod, nome: e.nome, razao: e.razao, cnpj: e.cnpj, ie: e.ie, im: e.im,
    cidade: e.cidade, uf: e.uf, status: e.status, simples: e.simples,
    cnae: e.cnae, atividade: e.atividade,
    fat: t.fat, nsai: t.nsai, compras: t.compras, nent: t.nent,
    impTot: SIGLAS.reduce((a, s) => a + (t.imp[s] || 0), 0)
  };
  SIGLAS.forEach(s => o['i_' + s] = t.imp[s] || 0);
  return o;
}));

fazAba(wb, 'Mensal por empresa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Inscrição estadual', key: 'ie', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Faturamento', key: 'fat', width: 16, style: MOEDA },
  { header: 'Notas de saída', key: 'nsai', width: 13 },
  { header: 'Compras', key: 'compras', width: 16, style: MOEDA },
  { header: 'Notas de entrada', key: 'nent', width: 14 },
  { header: 'Total impostos', key: 'impTot', width: 16, style: MOEDA },
  ...SIGLAS.map(s => ({ header: s, key: 'i_' + s, width: 13, style: MOEDA })),
], Object.values(mov).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes)).map(o => {
  const e = E(String(o.cod));
  const x = {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ie: e.ie, mes: o.mes,
    fat: o.fat, nsai: o.nsai, compras: o.compras, nent: o.nent,
    impTot: SIGLAS.reduce((a, s) => a + (o.imp[s] || 0), 0)
  };
  SIGLAS.forEach(s => x['i_' + s] = o.imp[s] || 0);
  return x;
}));

fazAba(wb, 'Impostos detalhado', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Sigla', key: 'sigla', width: 12 },
  { header: 'Imposto', key: 'imposto', width: 32 },
  { header: 'Módulo', key: 'sistema', width: 14 },
  { header: 'Base de cálculo', key: 'base', width: 17, style: MOEDA },
  { header: 'Saldo devedor (apurado)', key: 'devedor', width: 20, style: MOEDA },
  { header: 'Saldo credor', key: 'credor', width: 15, style: MOEDA },
  { header: 'Valor pago', key: 'pago', width: 15, style: MOEDA },
  { header: 'Apurações', key: 'apuracoes', width: 11 },
], impRows.sort((a, b) => a.cod - b.cod || a.mes.localeCompare(b.mes) || a.sigla.localeCompare(b.sigla)));

fazAba(wb, 'Totais por tributo', [
  { header: 'Sigla', key: 'sigla', width: 14 },
  { header: 'Total apurado', key: 'v', width: 20, style: MOEDA },
  { header: 'Empresas', key: 'n', width: 12 },
], SIGLAS.map(s => ({
  sigla: s, v: siglasTotais[s],
  n: Object.values(tot).filter(t => (t.imp[s] || 0) !== 0).length
})));

wb.xlsx.writeFile(DEST).then(() => {
  const f = Object.values(tot).reduce((a, t) => a + t.fat, 0);
  const i = Object.values(siglasTotais).reduce((a, v) => a + v, 0);
  console.log('   empresas: %d | faturamento: R$ %s | impostos: R$ %s | tributos: %d',
    Object.keys(tot).length,
    f.toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
    i.toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
    SIGLAS.length);
});

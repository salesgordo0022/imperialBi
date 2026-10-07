// 05-PRODUTOR-RURAL.xlsx — recorte dos produtores rurais
//
// O banco NÃO tem marca de produtor rural (CNPJ_PRODUTOR_RURAL_EMP vazio,
// ucxa_emp = 0, LCDPR_CTA = 0 em todas as contas). A identificação é por
// CNAE 2.0 agropecuário e por FUNRURAL apurado — ver cabeçalho do
// 05-PRODUTOR-RURAL.ps1.
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, num, int, comp, mm, MOEDA, fazAba, destino } = L;

const { emp, cnaeDesc, E } = L.carregaEmpresas();
const DEST = destino();
const VERDE = 'FF548235';
const AGRO = /^0[123]/;      // divisões 01 agricultura/pecuária, 02 florestal, 03 pesca

// ---------- FUNRURAL ----------
const funrural = {};
for (const r of readObj('04_impostos')) {
  if ((r.sigla || '').trim() !== 'FUNRURAL') continue;
  funrural[r.codi_emp] = (funrural[r.codi_emp] || 0) + num(r.saldo_devedor);
}

// ---------- seleção ----------
const criterio = {};
for (const [cod, e] of Object.entries(emp)) {
  const c = [];
  if (AGRO.test(e.cnae)) c.push('CNAE agropecuário');
  if ((funrural[cod] || 0) > 0) c.push('FUNRURAL apurado');
  if (e.pf && AGRO.test(e.cnae)) c.push('Pessoa física');
  if (c.length) criterio[cod] = c.join(' + ');
}
const CODS = new Set(Object.keys(criterio));

// ---------- movimento ----------
const mov = {};
const M = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!mov[k]) mov[k] = { cod, mes, fat: 0, nsai: 0, compras: 0, nent: 0, imp: {} };
  return mov[k];
};
for (const r of readObj('02_faturamento')) {
  if (!CODS.has(r.codi_emp)) continue;
  const o = M(r.codi_emp, mm(r.ano, r.mes)); o.fat += num(r.faturamento); o.nsai += int(r.notas);
}
for (const r of readObj('03_compras')) {
  if (!CODS.has(r.codi_emp)) continue;
  const o = M(r.codi_emp, mm(r.ano, r.mes)); o.compras += num(r.compras); o.nent += int(r.notas);
}
const siglas = {};
for (const r of readObj('04_impostos')) {
  if (!CODS.has(r.codi_emp)) continue;
  const sis = (r.codi_sis || '').trim();
  if (sis !== '5' && sis !== '6') continue;
  const s = (r.sigla || '').trim(), v = num(r.saldo_devedor);
  const o = M(r.codi_emp, comp(r.competencia));
  o.imp[s] = (o.imp[s] || 0) + v;
  siglas[s] = (siglas[s] || 0) + v;
}
const SIGLAS = Object.entries(siglas).filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]).map(([s]) => s);

const tot = {};
for (const o of Object.values(mov)) {
  const t = tot[o.cod] || (tot[o.cod] = { cod: o.cod, fat: 0, compras: 0, imp: 0 });
  t.fat += o.fat; t.compras += o.compras;
  t.imp += Object.values(o.imp).reduce((a, v) => a + v, 0);
}

const wb = new ExcelJS.Workbook();
wb.creator = 'Extração Domínio'; wb.created = new Date();

fazAba(wb, 'Produtores identificados', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Produtor / Empresa', key: 'nome', width: 44 },
  { header: 'Razão social', key: 'razao', width: 44 },
  { header: 'CPF/CNPJ', key: 'cnpj', width: 20 },
  { header: 'Tipo', key: 'tipo', width: 14 },
  { header: 'Inscrição estadual', key: 'ie', width: 18 },
  { header: 'Cidade', key: 'cidade', width: 20 },
  { header: 'UF', key: 'uf', width: 6 },
  { header: 'Situação', key: 'status', width: 10 },
  { header: 'CNAE 2.0', key: 'cnae', width: 12 },
  { header: 'Atividade', key: 'atividade', width: 48 },
  { header: 'Critério de identificação', key: 'crit', width: 34 },
  { header: 'Faturamento no período', key: 'fat', width: 19, style: MOEDA },
  { header: 'Compras no período', key: 'compras', width: 18, style: MOEDA },
  { header: 'Impostos apurados', key: 'imp', width: 17, style: MOEDA },
  { header: 'FUNRURAL apurado', key: 'fun', width: 17, style: MOEDA },
], [...CODS].map(c => {
  const e = E(c), t = tot[c] || {};
  return {
    cod: e.cod, nome: e.nome, razao: e.razao, cnpj: e.cnpj,
    tipo: e.pf ? 'Pessoa física' : 'Pessoa jurídica',
    ie: e.ie, cidade: e.cidade, uf: e.uf, status: e.status,
    cnae: e.cnae, atividade: e.atividade || cnaeDesc[e.cnae] || '', crit: criterio[c],
    fat: t.fat || 0, compras: t.compras || 0, imp: t.imp || 0, fun: funrural[c] || 0
  };
}).sort((a, b) => b.fat - a.fat), VERDE);

fazAba(wb, 'Faturamento e impostos', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Produtor / Empresa', key: 'nome', width: 44 },
  { header: 'CPF/CNPJ', key: 'cnpj', width: 20 },
  { header: 'Inscrição estadual', key: 'ie', width: 18 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Faturamento', key: 'fat', width: 17, style: MOEDA },
  { header: 'Notas de saída', key: 'nsai', width: 13 },
  { header: 'Compras', key: 'compras', width: 17, style: MOEDA },
  { header: 'Notas de entrada', key: 'nent', width: 14 },
  ...SIGLAS.map(s => ({ header: s, key: 'i_' + s, width: 14, style: MOEDA })),
], Object.values(mov).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes)).map(o => {
  const e = E(o.cod);
  const x = { cod: e.cod, nome: e.nome, cnpj: e.cnpj, ie: e.ie, mes: o.mes, fat: o.fat, nsai: o.nsai, compras: o.compras, nent: o.nent };
  SIGLAS.forEach(s => x['i_' + s] = o.imp[s] || 0);
  return x;
}), VERDE);

// ---------- livro caixa ----------
const lc = {};
for (const r of readObj('30_livro_caixa')) {
  if (!CODS.has(r.codi_emp)) continue;
  const mes = r.ano + '-' + String(r.mes).padStart(2, '0');
  const k = r.codi_emp + '|' + mes;
  const o = lc[k] || (lc[k] = { cod: r.codi_emp, mes, ent: 0, sai: 0, n: 0 });
  o.ent += num(r.entradas); o.sai += num(r.saidas); o.n += int(r.lancamentos);
}
const lcLista = Object.values(lc).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes));
let atual = '', saldo = 0;
for (const o of lcLista) { if (o.cod !== atual) { atual = o.cod; saldo = 0; } saldo += o.ent - o.sai; o.acum = saldo; }

fazAba(wb, 'Livro caixa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Produtor / Empresa', key: 'nome', width: 44 },
  { header: 'CPF/CNPJ', key: 'cnpj', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Entradas (receitas)', key: 'ent', width: 18, style: MOEDA },
  { header: 'Saídas (despesas)', key: 'sai', width: 18, style: MOEDA },
  { header: 'Saldo do mês', key: 'saldo', width: 17, style: MOEDA },
  { header: 'Saldo acumulado', key: 'acum', width: 17, style: MOEDA },
  { header: 'Lançamentos', key: 'n', width: 12 },
], lcLista.map(o => {
  const e = E(o.cod);
  return { cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes: o.mes, ent: o.ent, sai: o.sai, saldo: o.ent - o.sai, acum: o.acum, n: o.n };
}), VERDE);

// ---------- folha ----------
const folha = {};
const F = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!folha[k]) folha[k] = { cod, mes, prov: 0, nemp: 0, pat: 0, rat: 0, ter: 0, seg: 0, guia: 0 };
  return folha[k];
};
for (const r of readObj('05_folha_encargos')) {
  if (!CODS.has(r.codi_emp)) continue;
  const o = F(r.codi_emp, comp(r.competencia));
  o.prov += num(r.proventos); o.nemp += int(r.empregados);
}
for (const r of readObj('05b_guia_inss')) {
  if (!CODS.has(r.codi_emp)) continue;
  const o = F(r.codi_emp, comp(r.competencia));
  o.pat += num(r.inss_patronal_guia); o.rat += num(r.rat_guia); o.ter += num(r.terceiros_guia);
  o.seg += num(r.segurados_guia); o.guia += num(r.total_guia);
}

fazAba(wb, 'Folha e encargos', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Produtor / Empresa', key: 'nome', width: 44 },
  { header: 'CPF/CNPJ', key: 'cnpj', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Empregados', key: 'nemp', width: 11 },
  { header: 'Proventos', key: 'prov', width: 16, style: MOEDA },
  { header: 'INSS patronal', key: 'pat', width: 15, style: MOEDA },
  { header: 'RAT/SAT', key: 'rat', width: 13, style: MOEDA },
  { header: 'Terceiros', key: 'ter', width: 14, style: MOEDA },
  { header: 'INSS segurados', key: 'seg', width: 15, style: MOEDA },
  { header: 'Total da guia', key: 'guia', width: 16, style: MOEDA },
], Object.values(folha).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes)).map(o => {
  const e = E(o.cod);
  return { cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes: o.mes, nemp: o.nemp, prov: o.prov, pat: o.pat, rat: o.rat, ter: o.ter, seg: o.seg, guia: o.guia };
}), VERDE);

wb.xlsx.writeFile(DEST).then(() => {
  console.log('   produtores: %d | faturamento: R$ %s | FUNRURAL: R$ %s',
    CODS.size,
    Object.values(tot).reduce((a, t) => a + t.fat, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 }),
    Object.values(funrural).reduce((a, v) => a + v, 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 }));
});

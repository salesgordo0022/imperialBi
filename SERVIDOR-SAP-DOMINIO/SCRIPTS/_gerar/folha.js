// 02-FOLHA-E-ENCARGOS.xlsx — encargos patronais, headcount, admissões e demissões
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, num, int, comp, MOEDA, fazAba, destino } = L;

const { E } = L.carregaEmpresas();
const DEST = destino();
const VERDE = 'FF375623';

// ---------- série de competências: do primeiro ao último dado ----------
const mesesVistos = new Set();
for (const r of readObj('05_folha_encargos')) mesesVistos.add(comp(r.competencia));
for (const r of readObj('05b_guia_inss'))     mesesVistos.add(comp(r.competencia));
const ordenados = [...mesesVistos].filter(Boolean).sort();
const INI = ordenados[0] || '2024-01';
const FIM = ordenados[ordenados.length - 1] || '2026-12';

const meses = [];
{
  let [y, m] = INI.split('-').map(Number);
  const [fy, fm] = FIM.split('-').map(Number);
  while (y < fy || (y === fy && m <= fm)) {
    meses.push(y + '-' + String(m).padStart(2, '0'));
    m++; if (m > 12) { m = 1; y++; }
  }
}

// ---------- empregados e desligamentos ----------
const empregados = readObj('06_empregados').map(r => ({
  cod: r.codi_emp, id: r.i_empregados, nome: (r.nome || '').trim(), cpf: (r.cpf || '').trim(),
  admissao: r.admissao || '', salario: num(r.salario)
}));

// Complemento de rescisão gera outra linha para o mesmo empregado;
// vale a data mais antiga.
const demis = {};
const rescRows = [];
for (const r of readObj('07_rescisoes')) {
  const d = r.demissao || '';
  if (!d) continue;
  const k = r.codi_emp + '|' + r.i_empregados;
  if (!demis[k] || d < demis[k].data) demis[k] = { data: d, motivo: (r.motivo || '').trim() };
  const e = E(r.codi_emp);
  rescRows.push({
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, id: int(r.i_empregados),
    demissao: d, mes: comp(d), motivo: (r.motivo || '').trim(),
    proventos: num(r.proventos), descontos: num(r.descontos),
    _k: k
  });
}
const nomeFunc = {};
empregados.forEach(e => nomeFunc[e.cod + '|' + e.id] = e.nome);

// ---------- headcount ----------
const hc = {};
const H = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!hc[k]) hc[k] = { cod, mes, ativos: 0, adm: 0, dem: 0 };
  return hc[k];
};
const codsComFolha = new Set(empregados.map(e => e.cod));
for (const cod of codsComFolha) meses.forEach(m => H(cod, m));

for (const e of empregados) {
  if (!e.admissao || e.admissao.length < 7) continue;
  const admMes = e.admissao.slice(0, 7);
  const dm = demis[e.cod + '|' + e.id];
  const demMes = dm ? dm.data.slice(0, 7) : null;
  if (admMes >= INI && admMes <= FIM) H(e.cod, admMes).adm++;
  if (demMes && demMes >= INI && demMes <= FIM) H(e.cod, demMes).dem++;
  // ativo no mês: admitido até o fim dele e sem demissão antes do início
  for (const m of meses) {
    if (m < admMes) continue;
    if (demMes && m > demMes) break;
    H(e.cod, m).ativos++;
  }
}

// ---------- encargos ----------
const enc = {};
const G = (cod, mes) => {
  const k = cod + '|' + mes;
  if (!enc[k]) enc[k] = {
    cod, mes, proventos: 0, descontos: 0, empregadosCalc: 0,
    patronal: 0, rat: 0, terceiros: 0, segurados: 0, totalGuia: 0, multa: 0, guias: 0, pisFolha: 0
  };
  return enc[k];
};
for (const r of readObj('05_folha_encargos')) {
  const o = G(r.codi_emp, comp(r.competencia));
  o.proventos += num(r.proventos); o.descontos += num(r.descontos);
  o.empregadosCalc += int(r.empregados);
}
for (const r of readObj('05b_guia_inss')) {
  const o = G(r.codi_emp, comp(r.competencia));
  o.patronal += num(r.inss_patronal_guia); o.rat += num(r.rat_guia);
  o.terceiros += num(r.terceiros_guia);    o.segurados += num(r.segurados_guia);
  o.totalGuia += num(r.total_guia);        o.multa += num(r.multa);
  o.guias += int(r.guias);
}
for (const r of readObj('05c_guia_pis')) {
  G(r.codi_emp, comp(r.competencia)).pisFolha += num(r.pis_folha);
}

// ---------- totais ----------
const totEnc = {};
for (const o of Object.values(enc)) {
  const t = totEnc[o.cod] || (totEnc[o.cod] = {
    cod: o.cod, proventos: 0, descontos: 0, patronal: 0, rat: 0,
    terceiros: 0, segurados: 0, totalGuia: 0, pisFolha: 0
  });
  ['proventos', 'descontos', 'patronal', 'rat', 'terceiros', 'segurados', 'totalGuia', 'pisFolha']
    .forEach(k => t[k] += o[k]);
}
const hcTot = {};
for (const o of Object.values(hc)) {
  const t = hcTot[o.cod] || (hcTot[o.cod] = { cod: o.cod, adm: 0, dem: 0, ultAtivos: 0, pico: 0 });
  t.adm += o.adm; t.dem += o.dem;
  if (o.ativos > t.pico) t.pico = o.ativos;
  if (o.mes === FIM) t.ultAtivos = o.ativos;
}

const wb = new ExcelJS.Workbook();
wb.creator = 'Extração Domínio'; wb.created = new Date();

fazAba(wb, 'Resumo por empresa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Inscrição estadual', key: 'ie', width: 18 },
  { header: 'Cidade', key: 'cidade', width: 18 },
  { header: 'UF', key: 'uf', width: 6 },
  { header: 'Situação', key: 'status', width: 10 },
  { header: 'Funcionários ativos (' + FIM + ')', key: 'ativos', width: 16 },
  { header: 'Pico de ativos no período', key: 'pico', width: 15 },
  { header: 'Admissões no período', key: 'adm', width: 14 },
  { header: 'Demissões no período', key: 'dem', width: 14 },
  { header: 'Proventos (folha bruta)', key: 'proventos', width: 19, style: MOEDA },
  { header: 'INSS patronal', key: 'patronal', width: 15, style: MOEDA },
  { header: 'RAT/SAT', key: 'rat', width: 13, style: MOEDA },
  { header: 'Terceiros', key: 'terceiros', width: 14, style: MOEDA },
  { header: 'PIS sobre folha', key: 'pisFolha', width: 14, style: MOEDA },
  { header: 'TOTAL ENCARGO PATRONAL', key: 'patTot', width: 21, style: MOEDA },
  { header: 'INSS segurados (desc. empregado)', key: 'segurados', width: 18, style: MOEDA },
  { header: 'Total da guia INSS', key: 'totalGuia', width: 17, style: MOEDA },
], [...new Set([...Object.keys(totEnc), ...Object.keys(hcTot)])].map(c => {
  const e = E(c), t = totEnc[c] || {}, h = hcTot[c] || {};
  const patTot = (t.patronal || 0) + (t.rat || 0) + (t.terceiros || 0) + (t.pisFolha || 0);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, ie: e.ie, cidade: e.cidade, uf: e.uf, status: e.status,
    ativos: h.ultAtivos || 0, pico: h.pico || 0, adm: h.adm || 0, dem: h.dem || 0,
    proventos: t.proventos || 0, patronal: t.patronal || 0, rat: t.rat || 0,
    terceiros: t.terceiros || 0, pisFolha: t.pisFolha || 0, patTot,
    segurados: t.segurados || 0, totalGuia: t.totalGuia || 0
  };
}).sort((a, b) => b.patTot - a.patTot || b.proventos - a.proventos), VERDE);

fazAba(wb, 'Encargos mensais', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Empregados na folha', key: 'nemp', width: 14 },
  { header: 'Proventos', key: 'proventos', width: 16, style: MOEDA },
  { header: 'Descontos', key: 'descontos', width: 15, style: MOEDA },
  { header: 'INSS patronal', key: 'patronal', width: 15, style: MOEDA },
  { header: 'RAT/SAT', key: 'rat', width: 13, style: MOEDA },
  { header: 'Terceiros', key: 'terceiros', width: 14, style: MOEDA },
  { header: 'PIS sobre folha', key: 'pisFolha', width: 14, style: MOEDA },
  { header: 'TOTAL ENCARGO PATRONAL', key: 'patTot', width: 21, style: MOEDA },
  { header: 'INSS segurados', key: 'segurados', width: 15, style: MOEDA },
  { header: 'Multa/juros', key: 'multa', width: 13, style: MOEDA },
  { header: 'Total da guia INSS', key: 'totalGuia', width: 17, style: MOEDA },
  { header: 'Guias emitidas', key: 'guias', width: 12 },
], Object.values(enc).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes)).map(o => {
  const e = E(o.cod);
  return {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes: o.mes, nemp: o.empregadosCalc,
    proventos: o.proventos, descontos: o.descontos, patronal: o.patronal, rat: o.rat,
    terceiros: o.terceiros, pisFolha: o.pisFolha,
    patTot: o.patronal + o.rat + o.terceiros + o.pisFolha,
    segurados: o.segurados, multa: o.multa, totalGuia: o.totalGuia, guias: o.guias
  };
}), VERDE);

fazAba(wb, 'Funcionários por mês', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Funcionários ativos', key: 'ativos', width: 15 },
  { header: 'Admissões', key: 'adm', width: 11 },
  { header: 'Demissões', key: 'dem', width: 11 },
  { header: 'Saldo do mês', key: 'saldo', width: 12 },
], Object.values(hc).filter(o => o.ativos || o.adm || o.dem)
  .sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes))
  .map(o => {
    const e = E(o.cod);
    return { cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes: o.mes, ativos: o.ativos, adm: o.adm, dem: o.dem, saldo: o.adm - o.dem };
  }), VERDE);

fazAba(wb, 'Admissões', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Matrícula', key: 'id', width: 10 },
  { header: 'Funcionário', key: 'func', width: 38 },
  { header: 'CPF', key: 'cpf', width: 15 },
  { header: 'Admissão', key: 'admissao', width: 12 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Salário', key: 'salario', width: 14, style: MOEDA },
  { header: 'Demissão', key: 'demissao', width: 12 },
  { header: 'Situação', key: 'sit', width: 12 },
], empregados
  .filter(e => e.admissao && e.admissao.slice(0, 7) >= INI && e.admissao.slice(0, 7) <= FIM)
  .map(e => {
    const d = demis[e.cod + '|' + e.id], x = E(e.cod);
    return {
      cod: x.cod, nome: x.nome, cnpj: x.cnpj, id: int(e.id), func: e.nome, cpf: e.cpf,
      admissao: e.admissao, mes: e.admissao.slice(0, 7), salario: e.salario,
      demissao: d ? d.data : '', sit: d ? 'Desligado' : 'Ativo'
    };
  }).sort((a, b) => a.cod - b.cod || a.admissao.localeCompare(b.admissao)), VERDE);

// O código do motivo não tem legenda no banco — ela está no executável
// do Domínio (Folha > Rescisões).
fazAba(wb, 'Demissões', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 44 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Matrícula', key: 'id', width: 10 },
  { header: 'Funcionário', key: 'func', width: 38 },
  { header: 'Demissão', key: 'demissao', width: 12 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Cód. motivo', key: 'motivo', width: 11 },
  { header: 'Proventos da rescisão', key: 'proventos', width: 17, style: MOEDA },
  { header: 'Descontos da rescisão', key: 'descontos', width: 17, style: MOEDA },
], rescRows.filter(r => r.mes >= INI && r.mes <= FIM)
  .map(r => ({ ...r, func: nomeFunc[r._k] || '' }))
  .sort((a, b) => a.cod - b.cod || a.demissao.localeCompare(b.demissao)), VERDE);

wb.xlsx.writeFile(DEST).then(() => {
  const tp = Object.values(enc).reduce((a, o) => a + o.patronal + o.rat + o.terceiros + o.pisFolha, 0);
  console.log('   período %s a %s | empresas com folha: %d | empregados: %d | encargo patronal: R$ %s',
    INI, FIM, codsComFolha.size, empregados.length,
    tp.toLocaleString('pt-BR', { minimumFractionDigits: 2 }));
});

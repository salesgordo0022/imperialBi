// 06-PRODUTIVIDADE-E-TEMPO.xlsx — tempo por empresa e por usuário
//
// Fonte: geloguser, uma linha por sessão (entrada e saída com hora).
//
// ATENÇÃO AO RATEIO POR CONCORRÊNCIA (função rateiaConcorrencia, abaixo).
// O mesmo usuário pode ter várias sessões abertas ao mesmo tempo — dois
// módulos, ou duas empresas. Somar as durações conta o mesmo minuto mais de
// uma vez: na base inteira a soma bruta dá 41.081 h para 37.675 h de relógio,
// 8,3% a mais. Por isso cada minuto é dividido entre as sessões que estavam
// abertas naquele instante.
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, num, int, comp, mm, MODULOS, MODS, DEPTS, dept, hms, MOEDA, fazAba, destino } = L;

const { E } = L.carregaEmpresas();
const DEST = destino();
const ROXO = 'FF4B2E83';
const h = m => Math.round(m / 60 * 100) / 100;

// Minutos absolutos desde uma origem fixa. Date.UTC não depende de fuso nem
// de horário de verão — só as diferenças importam.
const diaAbs = d => {
  const [y, m, dd] = String(d).split('-').map(Number);
  return Date.UTC(y, m - 1, dd) / 60000;
};

// ---------- sessões ----------
const sess = [];
for (const r of readObj('10_loguser')) {
  const a = hms(r.hora_ini), b = hms(r.hora_fim);
  if (a == null || b == null) continue;
  const ini = diaAbs(r.data_ini) + a;
  let fim = diaAbs(r.data_fim || r.data_ini) + b;
  if (fim < ini) fim = ini;                 // relógio virado para trás: descarta a diferença
  sess.push({
    cod: r.codi_emp, usuario: (r.usuario || '').trim(),
    mod: MODULOS[r.modulo] || ('cód ' + r.modulo),
    dep: dept(r.modulo),
    data: r.data_ini, mes: comp(r.data_ini),
    hIni: r.hora_ini, hFim: r.hora_fim,
    ini, fim,
    min: fim - ini,     // duração bruta da sessão
    minAjust: 0         // fatia depois do rateio por concorrência
  });
}

// ---------- rateio por concorrência ----------
// Varredura por usuário: percorre a linha do tempo dele, e em cada trecho
// entre dois eventos divide a duração igualmente entre as sessões abertas.
// Resultado: a soma dos minAjust de um usuário é exatamente o tempo de
// relógio dele, e cada empresa fica com uma fatia justa do tempo dividido.
function rateiaConcorrencia(sessoes) {
  const porUsu = {};
  sessoes.forEach((s, i) => { (porUsu[s.usuario] = porUsu[s.usuario] || []).push(i); });

  for (const idxs of Object.values(porUsu)) {
    const pontos = [...new Set(idxs.flatMap(i => [sessoes[i].ini, sessoes[i].fim]))]
      .sort((a, b) => a - b);
    const ordem = idxs.slice().sort((a, b) => sessoes[a].ini - sessoes[b].ini);

    let p = 0;
    const ativos = new Set();
    for (let k = 0; k < pontos.length - 1; k++) {
      const ini = pontos[k], fim = pontos[k + 1], dur = fim - ini;
      while (p < ordem.length && sessoes[ordem[p]].ini <= ini) { ativos.add(ordem[p]); p++; }
      for (const i of ativos) if (sessoes[i].fim <= ini) ativos.delete(i);
      if (dur <= 0 || ativos.size === 0) continue;
      const fatia = dur / ativos.size;
      for (const i of ativos) sessoes[i].minAjust += fatia;
    }
  }
}
rateiaConcorrencia(sess);

// ---------- volume de lançamentos, por empresa e mês ----------
const lanc = {};
const bump = (cod, mes, campo, v) => {
  if (!mes) return;
  const k = cod + '|' + mes;
  if (!lanc[k]) lanc[k] = { cod, mes, fiscalSai: 0, fiscalEnt: 0, contabil: 0, folha: 0, apuracoes: 0 };
  lanc[k][campo] += v;
};
for (const r of readObj('02_faturamento'))    bump(r.codi_emp, mm(r.ano, r.mes), 'fiscalSai', int(r.notas));
for (const r of readObj('03_compras'))        bump(r.codi_emp, mm(r.ano, r.mes), 'fiscalEnt', int(r.notas));
for (const r of readObj('11_lanc_contabeis')) bump(r.codi_emp, mm(r.ano, r.mes), 'contabil', int(r.lancamentos));
for (const r of readObj('05_folha_encargos')) bump(r.codi_emp, comp(r.competencia), 'folha', int(r.empregados));
for (const r of readObj('04_impostos'))       bump(r.codi_emp, comp(r.competencia), 'apuracoes', int(r.apuracoes));

// ---------- agregações ----------
// Em toda parte: 'min' é bruto (soma das sessões) e 'aj' é atribuído
// (depois do rateio). As horas de referência são as atribuídas.
const porEmp = {}, porUsu = {}, porEmpUsu = {}, porEmpMes = {};
for (const s of sess) {
  const e = porEmp[s.cod] || (porEmp[s.cod] = {
    cod: s.cod, sess: 0, min: 0, aj: 0, mods: {}, usus: new Set(), dias: new Set(),
    ini: s.data, fim: s.data
  });
  e.sess++; e.min += s.min; e.aj += s.minAjust;
  e.mods[s.mod] = (e.mods[s.mod] || 0) + s.minAjust;
  e.usus.add(s.usuario); e.dias.add(s.data);
  if (s.data < e.ini) e.ini = s.data;
  if (s.data > e.fim) e.fim = s.data;

  const u = porUsu[s.usuario] || (porUsu[s.usuario] = {
    usuario: s.usuario, sess: 0, min: 0, aj: 0, mods: {}, emps: new Set(), dias: new Set()
  });
  u.sess++; u.min += s.min; u.aj += s.minAjust;
  u.mods[s.mod] = (u.mods[s.mod] || 0) + s.minAjust;
  u.emps.add(s.cod); u.dias.add(s.data);

  const ku = s.cod + '|' + s.usuario;
  const eu = porEmpUsu[ku] || (porEmpUsu[ku] = {
    cod: s.cod, usuario: s.usuario, sess: 0, min: 0, aj: 0, dias: new Set()
  });
  eu.sess++; eu.min += s.min; eu.aj += s.minAjust; eu.dias.add(s.data);

  const km = s.cod + '|' + s.mes;
  const em = porEmpMes[km] || (porEmpMes[km] = {
    cod: s.cod, mes: s.mes, sess: 0, min: 0, aj: 0, usus: new Set(), dias: new Set()
  });
  em.sess++; em.min += s.min; em.aj += s.minAjust; em.usus.add(s.usuario); em.dias.add(s.data);
}

// ---------- agregações por DEPARTAMENTO ----------
// Cada lançamento é roteado ao departamento que o produz:
//   Fiscal   notas de saída + notas de entrada + apurações de imposto
//   Contábil lançamentos contábeis
//   Pessoal  empregados processados na folha
// Os demais departamentos não têm contador de volume no banco.
const lancDoDept = (l, d) => {
  if (!l) return 0;
  if (d === 'Fiscal')   return l.fiscalSai + l.fiscalEnt + l.apuracoes;
  if (d === 'Contábil') return l.contabil;
  if (d === 'Pessoal')  return l.folha;
  return 0;
};

const porDept = {}, porDeptMes = {}, porDeptEmp = {}, porDeptUsu = {};
for (const s of sess) {
  const d = porDept[s.dep] || (porDept[s.dep] = {
    dep: s.dep, sess: 0, min: 0, aj: 0, usus: new Set(), emps: new Set(), dias: new Set()
  });
  d.sess++; d.min += s.min; d.aj += s.minAjust;
  d.usus.add(s.usuario); d.emps.add(s.cod); d.dias.add(s.data);

  const km = s.dep + '|' + s.mes;
  const dm = porDeptMes[km] || (porDeptMes[km] = {
    dep: s.dep, mes: s.mes, sess: 0, aj: 0, usus: new Set(), emps: new Set()
  });
  dm.sess++; dm.aj += s.minAjust; dm.usus.add(s.usuario); dm.emps.add(s.cod);

  const ke = s.dep + '|' + s.cod;
  const de = porDeptEmp[ke] || (porDeptEmp[ke] = {
    dep: s.dep, cod: s.cod, sess: 0, aj: 0, usus: new Set(), dias: new Set()
  });
  de.sess++; de.aj += s.minAjust; de.usus.add(s.usuario); de.dias.add(s.data);

  const ku = s.dep + '|' + s.usuario;
  const du = porDeptUsu[ku] || (porDeptUsu[ku] = {
    dep: s.dep, usuario: s.usuario, sess: 0, aj: 0, emps: new Set(), dias: new Set()
  });
  du.sess++; du.aj += s.minAjust; du.emps.add(s.cod); du.dias.add(s.data);
}

// volume por departamento, somando o que caiu em cada empresa/mês
const volDept = {}, volDeptMes = {}, volDeptEmp = {};
for (const [k, l] of Object.entries(lanc)) {
  const [cod, mes] = k.split('|');
  for (const d of DEPTS) {
    const v = lancDoDept(l, d);
    if (!v) continue;
    volDept[d] = (volDept[d] || 0) + v;
    volDeptMes[d + '|' + mes] = (volDeptMes[d + '|' + mes] || 0) + v;
    volDeptEmp[d + '|' + cod] = (volDeptEmp[d + '|' + cod] || 0) + v;
  }
}

const totalAj = sess.reduce((a, s) => a + s.minAjust, 0);
const ordemDept = (a, b) => DEPTS.indexOf(a.dep) - DEPTS.indexOf(b.dep);

const wb = new ExcelJS.Workbook();
wb.creator = 'Extração Domínio'; wb.created = new Date();

// ---------- 0. departamento: o resumo que responde "tempo e lançamento por departamento"
fazAba(wb, 'Por departamento', [
  { header: 'Departamento', key: 'dep', width: 18 },
  { header: 'Horas', key: 'horas', width: 13, style: MOEDA },
  { header: '% do tempo total', key: 'pct', width: 14, style: { numFmt: '0.0%' } },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Usuários', key: 'nusu', width: 10 },
  { header: 'Empresas atendidas', key: 'nemp', width: 16 },
  { header: 'Dias com atividade', key: 'dias', width: 15 },
  { header: 'Lançamentos', key: 'lanc', width: 14 },
  { header: 'Lanç. por hora', key: 'lph', width: 13, style: { numFmt: '#,##0.0' } },
  { header: 'Média h/dia', key: 'mdia', width: 12, style: MOEDA },
], Object.values(porDept).sort(ordemDept).map(d => {
  const horas = h(d.aj), vol = volDept[d.dep] || 0;
  return {
    dep: d.dep, horas, pct: totalAj > 0 ? d.aj / totalAj : 0,
    sess: d.sess, nusu: d.usus.size, nemp: d.emps.size, dias: d.dias.size,
    lanc: vol, lph: horas > 0 && vol ? Math.round(vol / horas * 10) / 10 : '',
    mdia: Math.round(d.aj / d.dias.size / 60 * 100) / 100
  };
}), ROXO);

// ---------- 0b. departamento mês a mês ----------
fazAba(wb, 'Departamento por mês', [
  { header: 'Departamento', key: 'dep', width: 18 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Horas', key: 'horas', width: 13, style: MOEDA },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Usuários', key: 'nusu', width: 10 },
  { header: 'Empresas atendidas', key: 'nemp', width: 16 },
  { header: 'Lançamentos', key: 'lanc', width: 14 },
  { header: 'Lanç. por hora', key: 'lph', width: 13, style: { numFmt: '#,##0.0' } },
], Object.values(porDeptMes)
  .sort((a, b) => ordemDept(a, b) || a.mes.localeCompare(b.mes))
  .map(d => {
    const horas = h(d.aj), vol = volDeptMes[d.dep + '|' + d.mes] || 0;
    return {
      dep: d.dep, mes: d.mes, horas, sess: d.sess, nusu: d.usus.size, nemp: d.emps.size,
      lanc: vol, lph: horas > 0 && vol ? Math.round(vol / horas * 10) / 10 : ''
    };
  }), ROXO);

// ---------- 0c. departamento por empresa ----------
fazAba(wb, 'Departamento por empresa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Departamento', key: 'dep', width: 18 },
  { header: 'Horas', key: 'horas', width: 13, style: MOEDA },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Dias com acesso', key: 'dias', width: 13 },
  { header: 'Usuários', key: 'nusu', width: 10 },
  { header: 'Lançamentos', key: 'lanc', width: 14 },
  { header: 'Lanç. por hora', key: 'lph', width: 13, style: { numFmt: '#,##0.0' } },
], Object.values(porDeptEmp)
  .sort((a, b) => int(a.cod) - int(b.cod) || ordemDept(a, b))
  .map(d => {
    const e = E(d.cod), horas = h(d.aj), vol = volDeptEmp[d.dep + '|' + d.cod] || 0;
    return {
      cod: e.cod, nome: e.nome, cnpj: e.cnpj, dep: d.dep, horas,
      sess: d.sess, dias: d.dias.size, nusu: d.usus.size,
      lanc: vol, lph: horas > 0 && vol ? Math.round(vol / horas * 10) / 10 : ''
    };
  }), ROXO);

// ---------- 0d. quem trabalha em qual departamento ----------
const totalUsu = {};
for (const d of Object.values(porDeptUsu)) totalUsu[d.usuario] = (totalUsu[d.usuario] || 0) + d.aj;

fazAba(wb, 'Departamento por usuário', [
  { header: 'Usuário', key: 'usuario', width: 22 },
  { header: 'Departamento', key: 'dep', width: 18 },
  { header: 'Horas', key: 'horas', width: 13, style: MOEDA },
  { header: '% do tempo do usuário', key: 'pct', width: 16, style: { numFmt: '0.0%' } },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Empresas atendidas', key: 'nemp', width: 16 },
  { header: 'Dias', key: 'dias', width: 9 },
], Object.values(porDeptUsu)
  .sort((a, b) => a.usuario.localeCompare(b.usuario) || b.aj - a.aj)
  .map(d => ({
    usuario: d.usuario, dep: d.dep, horas: h(d.aj),
    pct: totalUsu[d.usuario] > 0 ? d.aj / totalUsu[d.usuario] : 0,
    sess: d.sess, nemp: d.emps.size, dias: d.dias.size
  })), ROXO);

// ---------- 1. por empresa ----------
fazAba(wb, 'Resumo por empresa', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Horas atribuídas', key: 'horas', width: 14, style: MOEDA },
  { header: 'Horas brutas (soma das sessões)', key: 'brutas', width: 15, style: MOEDA },
  { header: 'Dias com acesso', key: 'dias', width: 13 },
  { header: 'Média h/dia', key: 'mdia', width: 12, style: MOEDA },
  { header: 'Usuários', key: 'nusu', width: 10 },
  ...MODS.map(m => ({ header: 'h ' + m, key: 'm_' + m, width: 13, style: MOEDA })),
  { header: 'Primeiro acesso', key: 'ini', width: 15 },
  { header: 'Último acesso', key: 'fim', width: 15 },
], Object.values(porEmp).sort((a, b) => b.aj - a.aj).map(x => {
  const e = E(x.cod);
  const o = {
    cod: e.cod, nome: e.nome, cnpj: e.cnpj, sess: x.sess,
    horas: h(x.aj), brutas: h(x.min), dias: x.dias.size,
    mdia: Math.round(x.aj / x.dias.size / 60 * 100) / 100,
    nusu: x.usus.size, ini: x.ini, fim: x.fim
  };
  MODS.forEach(m => o['m_' + m] = h(x.mods[m] || 0));
  return o;
}), ROXO);

// ---------- 2. por usuário ----------
// Aqui "Horas reais" é o tempo de relógio: a soma das fatias de um usuário
// é exatamente a união dos intervalos dele, sem contar minuto duas vezes.
fazAba(wb, 'Resumo por usuário', [
  { header: 'Usuário', key: 'usuario', width: 22 },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Horas reais (sem sobreposição)', key: 'horas', width: 16, style: MOEDA },
  { header: 'Horas brutas (soma das sessões)', key: 'brutas', width: 16, style: MOEDA },
  { header: 'Horas sobrepostas', key: 'sobre', width: 14, style: MOEDA },
  { header: '% sobreposto', key: 'pct', width: 12, style: { numFmt: '0.0%' } },
  { header: 'Empresas atendidas', key: 'emps', width: 16 },
  { header: 'Dias trabalhados', key: 'dias', width: 14 },
  { header: 'Média h/dia', key: 'mdia', width: 12, style: MOEDA },
  { header: 'Média min/sessão', key: 'msess', width: 15, style: { numFmt: '#,##0.0' } },
  ...MODS.map(m => ({ header: 'h ' + m, key: 'm_' + m, width: 13, style: MOEDA })),
], Object.values(porUsu).sort((a, b) => b.aj - a.aj).map(u => {
  const o = {
    usuario: u.usuario, sess: u.sess,
    horas: h(u.aj), brutas: h(u.min), sobre: h(u.min - u.aj),
    pct: u.min > 0 ? (u.min - u.aj) / u.min : 0,
    emps: u.emps.size, dias: u.dias.size,
    mdia: Math.round(u.aj / u.dias.size / 60 * 100) / 100,
    msess: Math.round(u.min / u.sess * 10) / 10
  };
  MODS.forEach(m => o['m_' + m] = h(u.mods[m] || 0));
  return o;
}), ROXO);

// ---------- 3. empresa x usuário ----------
fazAba(wb, 'Tempo por empresa e usuário', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'Usuário', key: 'usuario', width: 22 },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Horas atribuídas', key: 'horas', width: 14, style: MOEDA },
  { header: 'Horas brutas', key: 'brutas', width: 13, style: MOEDA },
  { header: 'Dias', key: 'dias', width: 9 },
  { header: 'Média min/sessão', key: 'msess', width: 15, style: { numFmt: '#,##0.0' } },
], Object.values(porEmpUsu).sort((a, b) => b.aj - a.aj).map(x => ({
  cod: E(x.cod).cod, nome: E(x.cod).nome, usuario: x.usuario, sess: x.sess,
  horas: h(x.aj), brutas: h(x.min), dias: x.dias.size,
  msess: Math.round(x.min / x.sess * 10) / 10
})), ROXO);

// ---------- 4. mensal: tempo x volume ----------
// "Lanç. por hora" compara tempo com CONTAGEM DE DOCUMENTO, não com esforço:
// empresa com importação de XML faz milhares de notas em minutos. Serve para
// achar o que destoa, não para medir gente.
const chaves = new Set([...Object.keys(porEmpMes), ...Object.keys(lanc)]);
fazAba(wb, 'Mensal tempo x lançamentos', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Horas no sistema', key: 'horas', width: 14, style: MOEDA },
  { header: 'Sessões', key: 'sess', width: 10 },
  { header: 'Dias com acesso', key: 'dias', width: 13 },
  { header: 'Usuários', key: 'nusu', width: 10 },
  { header: 'Notas de saída', key: 'fiscalSai', width: 13 },
  { header: 'Notas de entrada', key: 'fiscalEnt', width: 14 },
  { header: 'Lanç. contábeis', key: 'contabil', width: 14 },
  { header: 'Empregados na folha', key: 'folha', width: 16 },
  { header: 'Apurações fiscais', key: 'apur', width: 14 },
  { header: 'Total lançamentos', key: 'total', width: 15 },
  { header: 'Lanç. por hora', key: 'lancPorHora', width: 13, style: { numFmt: '#,##0.0' } },
], [...chaves].map(k => {
  const [cod, mes] = k.split('|');
  if (!mes) return null;
  const t = porEmpMes[k], l = lanc[k];
  const total = l ? l.fiscalSai + l.fiscalEnt + l.contabil + l.folha : 0;
  const horas = t ? h(t.aj) : 0;
  return {
    cod: E(cod).cod, nome: E(cod).nome, mes,
    horas, sess: t ? t.sess : 0, dias: t ? t.dias.size : 0, nusu: t ? t.usus.size : 0,
    fiscalSai: l ? l.fiscalSai : 0, fiscalEnt: l ? l.fiscalEnt : 0,
    contabil: l ? l.contabil : 0, folha: l ? l.folha : 0, apur: l ? l.apuracoes : 0,
    total, lancPorHora: horas > 0 ? Math.round(total / horas * 10) / 10 : ''
  };
}).filter(Boolean).sort((a, b) => a.cod - b.cod || a.mes.localeCompare(b.mes)), ROXO);

// ---------- 5. sessões ----------
// "Simultâneas" = bruto / atribuído. 1,0 = sessão sozinha; 2,0 = havia em
// média outra sessão aberta do mesmo usuário o tempo todo.
fazAba(wb, 'Sessões (detalhe)', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'Usuário', key: 'usuario', width: 20 },
  { header: 'Departamento', key: 'dep', width: 16 },
  { header: 'Módulo', key: 'mod', width: 16 },
  { header: 'Data', key: 'data', width: 12 },
  { header: 'Entrou', key: 'ini', width: 10 },
  { header: 'Saiu', key: 'fim', width: 10 },
  { header: 'Minutos (bruto)', key: 'min', width: 13, style: { numFmt: '#,##0.0' } },
  { header: 'Minutos atribuídos', key: 'aj', width: 15, style: { numFmt: '#,##0.0' } },
  { header: 'Sessões simultâneas', key: 'conc', width: 14, style: { numFmt: '#,##0.00' } },
], sess.map(s => ({
  cod: E(s.cod).cod, nome: E(s.cod).nome, usuario: s.usuario, dep: s.dep, mod: s.mod,
  data: s.data, ini: s.hIni, fim: s.hFim,
  min: Math.round(s.min * 10) / 10,
  aj: Math.round(s.minAjust * 10) / 10,
  conc: s.minAjust > 0 ? Math.round(s.min / s.minAjust * 100) / 100 : ''
})), ROXO);

// ---------- 6. lançamentos contábeis por operador ----------
fazAba(wb, 'Lanç. contábeis por operador', [
  { header: 'Código', key: 'cod', width: 9 },
  { header: 'Empresa', key: 'nome', width: 42 },
  { header: 'Competência', key: 'mes', width: 12 },
  { header: 'Operador', key: 'usu', width: 20 },
  { header: 'Lançamentos', key: 'n', width: 13 },
  { header: 'Valor', key: 'v', width: 16, style: MOEDA },
], readObj('11_lanc_contabeis').map(r => ({
  cod: E(r.codi_emp).cod, nome: E(r.codi_emp).nome, mes: mm(r.ano, r.mes),
  usu: (r.codi_usu || '').trim(), n: int(r.lancamentos), v: num(r.valor)
})).sort((a, b) => a.cod - b.cod || a.mes.localeCompare(b.mes)), ROXO);

wb.xlsx.writeFile(DEST).then(() => {
  const bruto = sess.reduce((a, s) => a + s.min, 0);
  const aj = sess.reduce((a, s) => a + s.minAjust, 0);
  const longas = sess.filter(s => s.min > 600).length;
  console.log('   sessões: %s | empresas: %d | usuários: %d',
    sess.length.toLocaleString('pt-BR'), Object.keys(porEmp).length, Object.keys(porUsu).length);
  console.log('   horas reais: %s | brutas: %s | sobreposição: %s (%s%%)%s',
    h(aj).toLocaleString('pt-BR'), h(bruto).toLocaleString('pt-BR'),
    h(bruto - aj).toLocaleString('pt-BR'), (100 * (bruto - aj) / bruto).toFixed(1),
    longas ? ` | ${longas} sessões acima de 10h (sistema aberto sem uso?)` : '');
});

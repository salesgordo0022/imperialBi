// 04-LIVRO-CAIXA.xlsx — movimento das contas de caixa
//
// A aba de detalhe passa de 880 mil linhas, então este gerador escreve em
// streaming (WorkbookWriter) e lê o .tsv linha a linha. Se carregasse tudo
// na memória, estouraria o heap do Node numa máquina com pouca RAM.
const fs = require('fs');
const readline = require('readline');
const ExcelJS = require('exceljs');
const L = require('./lib');
const { readObj, existe, num, int, MOEDA, destino, OUT } = L;

const { E } = L.carregaEmpresas();
const DEST = destino();
const AZUL = 'FF2F5496';
const LIMITE_EXCEL = 1040000;   // a planilha aguenta 1.048.576 linhas

const plano = {};
for (const p of readObj('20_plano_contas')) {
  plano[p.codi_emp + '|' + p.codi_cta] = { nome: (p.nome_cta || '').trim(), clas: (p.clas_cta || '').trim() };
}
const contaNome = (c, n) => (plano[c + '|' + n] || {}).nome || '';
const contasCaixa = new Set(readObj('21_contas_caixa').map(c => c.codi_emp + '|' + c.codi_cta));

// ---------- agregados ----------
const mensal = {};
for (const r of readObj('30_livro_caixa')) {
  const mes = r.ano + '-' + String(r.mes).padStart(2, '0');
  const k = r.codi_emp + '|' + mes;
  const o = mensal[k] || (mensal[k] = { cod: r.codi_emp, mes, ent: 0, sai: 0, n: 0 });
  o.ent += num(r.entradas); o.sai += num(r.saidas); o.n += int(r.lancamentos);
}
const listaMensal = Object.values(mensal).sort((a, b) => int(a.cod) - int(b.cod) || a.mes.localeCompare(b.mes));
let atual = '', saldo = 0;
for (const o of listaMensal) {
  if (o.cod !== atual) { atual = o.cod; saldo = 0; }
  saldo += o.ent - o.sai;
  o.saldoAcum = saldo;
}
const tot = {};
for (const o of listaMensal) {
  const t = tot[o.cod] || (tot[o.cod] = { cod: o.cod, ent: 0, sai: 0, n: 0, meses: 0, ini: o.mes, fim: o.mes });
  t.ent += o.ent; t.sai += o.sai; t.n += o.n; t.meses++;
  if (o.mes < t.ini) t.ini = o.mes;
  if (o.mes > t.fim) t.fim = o.mes;
}

// ---------- escrita em streaming ----------
const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: DEST, useStyles: true });

function novaAba(nome, cols) {
  const ws = wb.addWorksheet(nome, { views: [{ state: 'frozen', xSplit: 2, ySplit: 1 }] });
  ws.columns = cols;
  const r1 = ws.getRow(1);
  r1.values = cols.map(c => c.header);
  r1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  r1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AZUL } };
  r1.alignment = { vertical: 'middle', wrapText: true };
  r1.height = 32;
  r1.commit();
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  return ws;
}

(async () => {
  let ws = novaAba('Resumo por empresa', [
    { header: 'Código', key: 'cod', width: 9 },
    { header: 'Empresa', key: 'nome', width: 44 },
    { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
    { header: 'Cidade', key: 'cidade', width: 18 },
    { header: 'UF', key: 'uf', width: 6 },
    { header: 'Situação', key: 'status', width: 10 },
    { header: 'Entradas (recebimentos)', key: 'ent', width: 20, style: MOEDA },
    { header: 'Saídas (pagamentos)', key: 'sai', width: 19, style: MOEDA },
    { header: 'Saldo do período', key: 'saldo', width: 18, style: MOEDA },
    { header: 'Lançamentos', key: 'n', width: 12 },
    { header: 'Meses com movimento', key: 'meses', width: 15 },
    { header: 'Primeiro mês', key: 'ini', width: 12 },
    { header: 'Último mês', key: 'fim', width: 12 },
  ]);
  for (const t of Object.values(tot).sort((a, b) => (b.ent + b.sai) - (a.ent + a.sai))) {
    const e = E(t.cod);
    ws.addRow({
      cod: e.cod, nome: e.nome, cnpj: e.cnpj, cidade: e.cidade, uf: e.uf, status: e.status,
      ent: t.ent, sai: t.sai, saldo: t.ent - t.sai, n: t.n, meses: t.meses, ini: t.ini, fim: t.fim
    }).commit();
  }
  ws.commit();

  ws = novaAba('Movimento mensal', [
    { header: 'Código', key: 'cod', width: 9 },
    { header: 'Empresa', key: 'nome', width: 44 },
    { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
    { header: 'Competência', key: 'mes', width: 12 },
    { header: 'Entradas', key: 'ent', width: 17, style: MOEDA },
    { header: 'Saídas', key: 'sai', width: 17, style: MOEDA },
    { header: 'Saldo do mês', key: 'saldoMes', width: 17, style: MOEDA },
    { header: 'Saldo acumulado', key: 'saldoAcum', width: 17, style: MOEDA },
    { header: 'Lançamentos', key: 'n', width: 12 },
  ]);
  for (const o of listaMensal) {
    const e = E(o.cod);
    ws.addRow({
      cod: e.cod, nome: e.nome, cnpj: e.cnpj, mes: o.mes,
      ent: o.ent, sai: o.sai, saldoMes: o.ent - o.sai, saldoAcum: o.saldoAcum, n: o.n
    }).commit();
  }
  ws.commit();

  ws = novaAba('Contas de caixa', [
    { header: 'Código', key: 'cod', width: 9 },
    { header: 'Empresa', key: 'nome', width: 44 },
    { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
    { header: 'Conta', key: 'conta', width: 10 },
    { header: 'Classificação', key: 'clas', width: 16 },
    { header: 'Descrição', key: 'desc', width: 42 },
  ]);
  for (const c of readObj('21_contas_caixa').sort((a, b) => int(a.codi_emp) - int(b.codi_emp))) {
    const e = E(c.codi_emp), p = plano[c.codi_emp + '|' + c.codi_cta] || {};
    ws.addRow({ cod: e.cod, nome: e.nome, cnpj: e.cnpj, conta: int(c.codi_cta), clas: p.clas || '', desc: p.nome || '' }).commit();
  }
  ws.commit();

  // ---------- detalhe ----------
  let n = 0, cortou = false;
  if (existe('31_caixa_detalhe')) {
    ws = novaAba('Lançamentos (detalhe)', [
      { header: 'Código', key: 'cod', width: 9 },
      { header: 'Empresa', key: 'nome', width: 40 },
      { header: 'CNPJ/CPF', key: 'cnpj', width: 20 },
      { header: 'Data', key: 'data', width: 12 },
      { header: 'Nº lanç.', key: 'num', width: 10 },
      { header: 'Conta de caixa', key: 'caixa', width: 11 },
      { header: 'Descrição da conta de caixa', key: 'caixaNome', width: 32 },
      { header: 'Contrapartida', key: 'contra', width: 12 },
      { header: 'Descrição da contrapartida', key: 'contraNome', width: 36 },
      { header: 'Histórico', key: 'historico', width: 48 },
      { header: 'Documento', key: 'doc', width: 16 },
      { header: 'Entrada', key: 'entrada', width: 15, style: MOEDA },
      { header: 'Saída', key: 'saida', width: 15, style: MOEDA },
      { header: 'Operador', key: 'operador', width: 16 },
    ]);

    const rl = readline.createInterface({
      input: fs.createReadStream(OUT + '/31_caixa_detalhe.tsv', { encoding: 'latin1' }),
      crlfDelay: Infinity
    });
    for await (const line of rl) {
      if (!line) continue;
      if (n >= LIMITE_EXCEL) { cortou = true; break; }
      const [cod, data, numero, cdeb, ccre, valor, hist, doc, oper] = line.split('\\t');
      // conta de caixa no débito = entrada; no crédito = saída
      const ehDeb = contasCaixa.has(cod + '|' + cdeb);
      const e = E(cod), v = num(valor);
      ws.addRow({
        cod: e.cod, nome: e.nome, cnpj: e.cnpj, data: (data || '').trim(), num: int(numero),
        caixa: ehDeb ? int(cdeb) : int(ccre),
        caixaNome: contaNome(cod, ehDeb ? cdeb : ccre),
        contra: ehDeb ? int(ccre) : int(cdeb),
        contraNome: contaNome(cod, ehDeb ? ccre : cdeb),
        historico: (hist || '').trim(), doc: (doc || '').trim(),
        entrada: ehDeb ? v : 0, saida: ehDeb ? 0 : v,
        operador: (oper || '').trim()
      }).commit();
      n++;
    }
    rl.close();
    ws.commit();
  }

  await wb.commit();
  console.log('   empresas: %d | lançamentos no detalhe: %s %s',
    Object.keys(tot).length, n.toLocaleString('pt-BR'),
    cortou ? '(TRUNCADO no limite do Excel)' : (n ? '' : '(detalhe não extraído)'));
})();

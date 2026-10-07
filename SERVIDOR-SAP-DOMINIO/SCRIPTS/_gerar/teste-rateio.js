// Teste do rateio por concorrência do produtividade.js
//
// Invariante que tem que valer: para cada usuário, a soma das fatias
// (minAjust) é exatamente igual à UNIÃO dos intervalos dele — o tempo de
// relógio. A união é calculada aqui por outro caminho (merge de intervalos
// ordenados), de propósito: se os dois métodos concordam, o rateio está certo.
//
//   node teste-rateio.js x C:/dbtmp/out
const L = require('./lib');
const { readObj, hms } = L;

const diaAbs = d => {
  const [y, m, dd] = String(d).split('-').map(Number);
  return Date.UTC(y, m - 1, dd) / 60000;
};

// --- mesma construção de sessões do gerador ---
const sess = [];
for (const r of readObj('10_loguser')) {
  const a = hms(r.hora_ini), b = hms(r.hora_fim);
  if (a == null || b == null) continue;
  const ini = diaAbs(r.data_ini) + a;
  let fim = diaAbs(r.data_fim || r.data_ini) + b;
  if (fim < ini) fim = ini;
  sess.push({ usuario: (r.usuario || '').trim(), cod: r.codi_emp, ini, fim, min: fim - ini, minAjust: 0 });
}

// --- mesmo rateio do gerador ---
function rateiaConcorrencia(sessoes) {
  const porUsu = {};
  sessoes.forEach((s, i) => { (porUsu[s.usuario] = porUsu[s.usuario] || []).push(i); });
  for (const idxs of Object.values(porUsu)) {
    const pontos = [...new Set(idxs.flatMap(i => [sessoes[i].ini, sessoes[i].fim]))].sort((a, b) => a - b);
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

// --- união independente, por merge de intervalos ---
function uniao(ivs) {
  if (!ivs.length) return 0;
  const o = ivs.slice().sort((a, b) => a[0] - b[0]);
  let total = 0, ci = o[0][0], cf = o[0][1];
  for (let k = 1; k < o.length; k++) {
    const [i, f] = o[k];
    if (i <= cf) { if (f > cf) cf = f; }
    else { total += cf - ci; ci = i; cf = f; }
  }
  return total + (cf - ci);
}

// --- comparação ---
const porUsu = {};
for (const s of sess) (porUsu[s.usuario] = porUsu[s.usuario] || []).push(s);

let falhas = 0, maiorErro = 0, somaAj = 0, somaUni = 0, somaBruto = 0;
for (const [u, ss] of Object.entries(porUsu)) {
  const aj = ss.reduce((a, s) => a + s.minAjust, 0);
  const un = uniao(ss.map(s => [s.ini, s.fim]));
  const br = ss.reduce((a, s) => a + s.min, 0);
  somaAj += aj; somaUni += un; somaBruto += br;
  const erro = Math.abs(aj - un);
  if (erro > maiorErro) maiorErro = erro;
  if (erro > 1e-6) { falhas++; console.log('  FALHA', u, 'rateio', aj.toFixed(6), 'união', un.toFixed(6)); }
}

// --- nenhuma fatia pode passar da duração bruta nem ser negativa ---
const forado = sess.filter(s => s.minAjust < -1e-9 || s.minAjust > s.min + 1e-9).length;

const ok = (c, t) => console.log((c ? '  OK  ' : '  FALHA ') + t);
console.log('\nTeste do rateio por concorrência\n');
console.log('  sessões            ', sess.length.toLocaleString('pt-BR'));
console.log('  usuários           ', Object.keys(porUsu).length);
console.log('  horas brutas       ', (somaBruto / 60).toFixed(2));
console.log('  horas rateadas     ', (somaAj / 60).toFixed(2));
console.log('  horas de relógio   ', (somaUni / 60).toFixed(2));
console.log('  sobreposição       ', ((somaBruto - somaUni) / 60).toFixed(2),
            '(' + (100 * (somaBruto - somaUni) / somaBruto).toFixed(2) + '%)');
console.log('  maior erro (min)   ', maiorErro.toExponential(2));
console.log('');
ok(falhas === 0, `rateio bate com a união em todos os ${Object.keys(porUsu).length} usuários`);
ok(forado === 0, 'nenhuma fatia negativa ou maior que a sessão');
ok(Math.abs(somaAj - somaUni) < 1e-6, 'total rateado = total de relógio');
console.log('');
process.exit(falhas === 0 && forado === 0 ? 0 : 1);

/* Controle Financeiro — dashboard que lê a planilha Planejamento_Financeiro (Google Planilhas publicado como .xlsx) */
(function () {
  "use strict";
  Chart.register(ChartDataLabels);
  Chart.defaults.font.family = "Inter, system-ui, -apple-system, 'Segoe UI', Roboto, Arial, sans-serif";
  Chart.defaults.elements.line.tension = 0.3; Chart.defaults.elements.line.borderCapStyle = "round";
  const $ = (id) => document.getElementById(id);
  const MESES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  const MESESL = ["janeiro","fevereiro","março","abril","maio","junho","julho","agosto","setembro","outubro","novembro","dezembro"];
  const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const brl0 = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
  const pctf = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 });
  const R = (n) => (n == null || isNaN(n) ? "—" : brl.format(n));
  const R0 = (n) => (n == null || isNaN(n) ? "—" : brl0.format(n));
  const P = (n) => (n == null || !isFinite(n) ? "—" : pctf.format(n));
  const num = (v) => (typeof v === "number" && isFinite(v) ? v : 0);
  const txt = (v) => (v == null ? "" : String(v).trim());
  const esc = (s) => txt(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
  const CFG = window.DASH_CONFIG || {};

  // ---------- datas ----------
  const pad = (n) => String(n).padStart(2, "0");
  function toDate(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(v.getTime() + 12 * 3600e3);
    if (typeof v === "number" && v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return new Date(p.y, p.m - 1, p.d, 12); }
    return null;
  }
  const ym = (v) => { const d = toDate(v); return d ? d.getFullYear() + "-" + pad(d.getMonth() + 1) : null; };
  const ymd = (d) => d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  const lab = (k) => { if (!k) return "—"; const [y, m] = k.split("-"); return MESES[+m - 1] + "/" + y.slice(2); };
  const labL = (k) => { const [y, m] = k.split("-"); return MESESL[+m - 1] + " de " + y; };
  const addM = (k, n) => { let [y, m] = k.split("-").map(Number); m += n; y += Math.floor((m - 1) / 12); m = ((m - 1) % 12 + 12) % 12 + 1; return y + "-" + pad(m); };
  const diffM = (a, b) => { const [ya, ma] = a.split("-").map(Number), [yb, mb] = b.split("-").map(Number); return (yb - ya) * 12 + (mb - ma); };
  const fmtD = (d) => pad(d.getDate()) + "/" + pad(d.getMonth() + 1) + "/" + d.getFullYear();

  // 5º dia útil (sábado conta; domingo e feriado não). Se o 5º dia cair no sábado, o salário cai no dia útil anterior (sexta).
  function quintoDiaUtil(y, m0, fer) {
    let n = 0; const d = new Date(y, m0, 1, 12);
    while (true) { if (d.getDay() !== 0 && !fer.has(ymd(d))) { n++; if (n === 5) break; } d.setDate(d.getDate() + 1); }
    if (d.getDay() === 6) { do d.setDate(d.getDate() - 1); while (d.getDay() === 0 || d.getDay() === 6 || fer.has(ymd(d))); }
    return new Date(d);
  }
  function referencia(hoje, fer) {
    const y = hoje.getFullYear(), m = hoje.getMonth(); const q = quintoDiaUtil(y, m, fer);
    const hojeD = new Date(y, m, hoje.getDate(), 12);
    if (hojeD >= q) { const nx = quintoDiaUtil(m === 11 ? y + 1 : y, (m + 1) % 12, fer); return { mes: y + "-" + pad(m + 1), proximo: nx, virou: q }; }
    const pm = m === 0 ? 11 : m - 1, py = m === 0 ? y - 1 : y; return { mes: py + "-" + pad(pm + 1), proximo: q, virou: quintoDiaUtil(py, pm, fer) };
  }

  // ---------- leitura da planilha ----------
  function rows(wb, name) {
    const ws = wb.Sheets[name]; if (!ws) throw new Error('Aba "' + name + '" não encontrada');
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }
  const cell = (wb, sh, addr) => { const c = wb.Sheets[sh] && wb.Sheets[sh][addr]; return c ? c.v : null; };
  const findRow = (rs, pred, from = 0) => { for (let i = from; i < rs.length; i++) if (rs[i] && pred(txt(rs[i][0]))) return i; return -1; };

  function extract(wb) {
    const D = {};
    const C = "Configurações";
    D.ini = ym(cell(wb, C, "B3")); D.refPlanilha = ym(cell(wb, C, "B4"));
    D.versao = txt(cell(wb, C, "E3"));
    D.extras = [];
    if (wb.Sheets["Rendas Extras"]) rows(wb, "Rendas Extras").forEach((r, n) => {
      if (n < 4 || !r) return; const ini = ym(r[0]), v = num(r[2]); if (!ini || !v) return;
      const rec = txt(r[3]) === "Sim", fim = rec ? ym(r[4]) || "9999-12" : ini;
      D.extras.push({ ini, fim, desc: txt(r[1]), valor: v, vr: /VR|Flash/i.test(txt(r[6])) });
    });
    D.consumo = [];
    const HS = wb.Sheets["Histórico de contas"] ? "Histórico de contas" : wb.Sheets["Consumo"] ? "Consumo" : null;
    if (HS) { const cr = rows(wb, HS), ch = findRow(cr, (t) => t === "Mês da conta");
      if (ch >= 0) for (let i = ch + 1; i < cr.length; i++) { const r = cr[i]; if (!r) continue; const m = ym(r[0]); if (!m) continue;
        D.consumo.push({ mes: m, conta: txt(r[1]), ref: ym(r[2]), valor: num(r[3]), kwh: num(r[4]) }); } }
    D.tarifas = { hist: {}, band: {} };
    if (wb.Sheets["Tarifas Cemig"]) { const tr = rows(wb, "Tarifas Cemig"); let sec = 0;
      tr.forEach((r) => { if (!r) return; const t0 = txt(r[0]);
        if (/^Data da consulta/i.test(t0)) { sec = 1; return; } if (/^M[eê]s do consumo/i.test(t0)) { sec = 2; return; } if (/^BANDEIRA/i.test(t0)) { sec = 0; return; }
        if (sec === 1) { const d = toDate(r[0]), b = txt(r[1]), v = num(r[2]); if (d && b && v) { const k = ymd(d); (D.tarifas.hist[k] = D.tarifas.hist[k] || {})[b] = v; } }
        if (sec === 2) { const m = ym(r[0]), b = txt(r[1]); if (m && b) D.tarifas.band[m] = b; } }); }
    D.d13 = { "11": toDate(cell(wb, "13º e Férias", "C13")), "12": toDate(cell(wb, "13º e Férias", "C15")) };
    D.reservaInformada = num(cell(wb, C, "B5")); D.pct = num(cell(wb, C, "B6")); D.meta = D.pct > 1 ? D.pct : 0; D.planIni = ym(cell(wb, C, "B8")) || D.ini;
    D.cartoes = []; D.venc = {};
    for (let r = 11; r <= 15; r++) { const n = txt(cell(wb, C, "A" + r)); if (n && n !== "-") { D.cartoes.push(n); D.venc[n] = num(cell(wb, C, "B" + r)) || null; } }
    const PG = "Plano de Gastos";
    D.minSemana = num(cell(wb, PG, "B3")); D.fechamento = num(cell(wb, PG, "B14")) || null; D.minMes = num(cell(wb, PG, "B4")); D.saldo = num(cell(wb, PG, "B16"));
    D.feriados = [];
    D.ferNac = [];
    if (wb.Sheets["Feriados"]) rows(wb, "Feriados").slice(1).forEach((r) => { const d = toDate(r[0]); if (d) { D.feriados.push(ymd(d)); if (!txt(r[2]) || txt(r[2]) === "Nacional") D.ferNac.push(ymd(d)); } });

    // Projeção
    const pj = rows(wb, "Projeção");
    const hi = findRow(pj, (t) => t === "Item");
    D.meses = pj[hi].slice(1).map(ym).filter(Boolean);
    const n = D.meses.length;
    D.tipo = pj[hi - 1].slice(1, n + 1).map(txt);
    const line = (prefix) => { const i = findRow(pj, (t) => t.startsWith(prefix)); return i < 0 ? Array(n).fill(0) : pj[i].slice(1, n + 1).map(num); };
    D.proj = { renda: line("Renda ("), d13: line("(+) 13º"), ferias: line("(+) Férias"), ajuste: line("(−) Ajuste"), fixos: line("(−) Fixos"),
      parcelas: line("(−) Parcelas"), reemb: line("(+) Reembolsos"), variaveis: line("(−) Variáveis"), sobra: line("SOBRA DO MÊS"),
      saldo: line("Saldo acumulado"), vr: line("VR do mês"), comprom: line("TOTAL DE COMPROMISSOS") };
    D.proj.entradas = D.meses.map((_, i) => D.proj.renda[i] + D.proj.d13[i] + D.proj.ferias[i] + D.proj.ajuste[i] + D.proj.reemb[i]);
    D.faturas = {};
    let fi = findRow(pj, (t) => t.startsWith("FATURAS POR CARTÃO"));
    for (let i = fi + 1; i < pj.length; i++) { const t = txt(pj[i][0]); if (t.startsWith("TOTAL")) break; if (t && t !== "-") D.faturas[t] = pj[i].slice(1, n + 1).map(num); }

    // Plano de Gastos
    const pl = rows(wb, PG);
    const ph = findRow(pl, (t) => t === "Mês das compras");
    const K = ["mes","fatura","espaco","teto","semana","jaCartao","jaPix","ja","ainda","fora","vr","total","faturaInter","limite","reserva","reservaAc","meta","livre","livreResta","reservaRend","foraRes"];
    D.plano = {};
    for (let i = ph + 1; i < pl.length; i++) {
      const r = pl[i]; if (!r || txt(r[0]).startsWith("TOTAL")) break; const k = ym(r[0]); if (!k) continue;
      const o = {}; K.forEach((f, j) => (o[f] = j < 2 ? ym(r[j]) : num(r[j]))); D.plano[k] = o;
    }
    // Fixos
    const fx = rows(wb, "Fixos"); const fh = findRow(fx, (t) => t === "Descrição");
    D.fixos = [];
    for (let i = fh + 1; i < fx.length; i++) { const r = fx[i]; const d = txt(r && r[0]); if (d.startsWith("TOTAL")) break; if (!d) continue;
      D.fixos.push({ desc: d, cat: txt(r[1]), paga: txt(r[2]), valor: num(r[3]), r1: num(r[4]), q1: txt(r[5]), r2: num(r[6]), q2: txt(r[7]), dia: num(r[8]) || null, ini: ym(r[9]), fim: ym(r[10]), custo: num(r[11]) }); }
    // Parcelas
    const pa = rows(wb, "Parcelas"); const pah = findRow(pa, (t) => t.startsWith("Descrição"));
    D.dividas = [];
    for (let i = pah + 1; i < pa.length; i++) { const r = pa[i]; const d = txt(r && r[0]); if (d.startsWith("TOTAL")) break; if (!d) continue;
      D.dividas.push({ desc: d, cat: txt(r[1]), paga: txt(r[2]), valor: num(r[3]), ini: ym(r[4]), fim: ym(r[5]), r1: num(r[6]), q1: txt(r[7]), r2: num(r[8]), q2: txt(r[9]),
        custo: num(r[10]), totalParc: num(r[13]) || null, compra: toDate(r[19]) }); }
    // Lançamentos
    D.lanc = [];
    if (wb.Sheets["Lançamentos"]) { const lg = rows(wb, "Lançamentos"); const lh = findRow(lg, (t) => t === "Data da compra");
      for (let i = lh + 1; i < lg.length; i++) { const r = lg[i]; if (!r) continue; const dt = toDate(r[0]); if (!dt || !num(r[4])) continue;
        D.lanc.push({ data: ymd(dt), desc: txt(r[1]) || "(sem descrição)", cat: txt(r[2]), paga: txt(r[3]), total: num(r[4]), parcelado: txt(r[5]) === "Sim", n: num(r[6]) || 1,
          ini: ym(r[12]), fim: ym(r[13]), valor: num(r[14]), r1: num(r[15]), q1: txt(r[9]), r2: num(r[16]), q2: txt(r[11]), custo: num(r[17]), meuTotal: num(r[4]) - num(r[8]) - num(r[10]) }); } }
    // Categorias
    const rs = rows(wb, "Resumo mensal"); const rh = findRow(rs, (t) => t === "Mês");
    const rmeses = rs[rh].slice(1, 13).map(ym); const ci = findRow(rs, (t) => t.startsWith("MEU CUSTO REAL POR CATEGORIA"));
    D.cat = { meses: rmeses, itens: {} };
    for (let i = ci + 1; i < rs.length; i++) { const t = txt(rs[i] && rs[i][0]); if (!t) continue; if (t === "TOTAL") break; D.cat.itens[t] = { vals: rs[i].slice(1, 13), media: num(rs[i][13]) }; }
    // Metas e Cortes
    D.mc = { rend: 0, custo: 0, limites: [], metas: [], ass: [] };
    if (wb.Sheets["Metas e Cortes"]) { const mt = rows(wb, "Metas e Cortes");
      D.mc.rend = num(cell(wb, "Metas e Cortes", "B4")); D.mc.custo = num(cell(wb, "Metas e Cortes", "B5"));
      const sec = (pref) => findRow(mt, (t) => t.startsWith(pref));
      let i = sec("LIMITES POR CATEGORIA"); if (i >= 0) for (let r = i + 2; r < mt.length; r++) { const t = txt(mt[r] && mt[r][0]); if (!t || t === "TOTAL") break; D.mc.limites.push({ cat: t, lim: num(mt[r][1]) }); }
      i = sec("METAS DE RESERVA"); if (i >= 0) for (let r = i + 2; r < mt.length; r++) { const x = mt[r] || []; const t = txt(x[0]); if (t.startsWith("INDICADORES")) break; if (!t) continue;
        D.mc.metas.push({ nome: t, obj: num(x[1]), prazo: ym(x[2]), prio: num(x[3]) || 99 }); }
      i = sec("ASSINATURAS"); if (i >= 0) for (let r = i + 2; r < mt.length; r++) { const x = mt[r] || []; const t = txt(x[0]); if (t === "TOTAL") break; if (!t) continue;
        D.mc.ass.push({ nome: t, paga: txt(x[1]), valor: num(x[2]), ano: num(x[3]), ate: ym(x[4]) || txt(x[4]), manter: txt(x[5]) || "Avaliar" }); } }
    D.vrIni = 0;
    if (wb.Sheets["Metas e Cortes"]) { const mt2 = rows(wb, "Metas e Cortes"); const vi = findRow(mt2, (t) => t.startsWith("VR (FLASH)")); if (vi >= 0) D.vrIni = num(mt2[vi + 1] && mt2[vi + 1][1]); }
    // Recebimentos por item
    D.recItens = [];
    if (wb.Sheets["Recebimentos"]) { const rb = rows(wb, "Recebimentos"); const rh2 = findRow(rb, (t) => t.startsWith("Mês da conta"));
      for (let i = rh2 + 1; i < rb.length; i++) { const r = rb[i]; if (!r) continue; const m = ym(r[0]), p = txt(r[1]), it = txt(r[2]); if (!m || !p || !num(r[3])) continue;
        D.recItens.push({ mes: m, pessoa: p, item: it, valor: num(r[3]), data: toDate(r[4]) ? ymd(toDate(r[4])) : "" }); } }
    // Pagamentos feitos
    D.pagos = [];
    if (wb.Sheets["Pagamentos"]) { const pg = rows(wb, "Pagamentos"); const ph2 = findRow(pg, (t) => t.startsWith("Mês da conta"));
      for (let i = ph2 + 1; i < pg.length; i++) { const r = pg[i]; if (!r) continue; const m = ym(r[0]), n = txt(r[1]); if (!m || !n) continue;
        D.pagos.push({ mes: m, nome: n, valor: num(r[2]), data: toDate(r[3]) ? ymd(toDate(r[3])) : "", obs: txt(r[4]) }); } }
    // Faturas (previsto x real)
    D.fats = [];
    if (wb.Sheets["Faturas"]) rows(wb, "Faturas").slice(3).forEach((x) => { const m = ym(x && x[0]); if (m && txt(x[1])) D.fats.push({ mes: m, card: txt(x[1]), valor: num(x[2]), status: txt(x[3]), prev: num(x[4]) }); });
    // Reembolsos recebidos
    D.recebido = {};
    if (wb.Sheets["Reembolsos"]) { const re = rows(wb, "Reembolsos"); const ri = findRow(re, (t) => t.startsWith("RECEBIDO"));
      if (ri >= 0) { const ms = re[ri].slice(1).map(ym); for (let r = ri + 1; r < ri + 6; r++) { const p = txt(re[r] && re[r][0]); if (!p || p === "-") continue; D.recebido[p] = {}; ms.forEach((m, n) => { if (m && num(re[r][n + 1])) D.recebido[p][m] = num(re[r][n + 1]); }); } } }
    return D;
  }

  // ---------- carregamento ----------
  let DATA = null, REF = null;
  const setStatus = (t) => ($("status").textContent = t);
  function useWorkbook(buf, origem) {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    DATA = extract(wb);
    const stamp = new Date().toLocaleString("pt-BR");
    try { localStorage.setItem("dash-data", JSON.stringify({ DATA, stamp, origem })); } catch (e) {}
    setStatus("Atualizado de " + origem + " em " + stamp + (DATA.versao ? " · planilha versão " + DATA.versao : " · planilha sem número de versão (antiga)"));
    boot();
  }
  async function load() {
    const url = CFG.DATA_URL || "planilha.xlsx";
    setStatus("Buscando a planilha…");
    try {
      const res = await fetch(url + (url.includes("?") ? "&" : "?") + "v=" + Date.now(), { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      useWorkbook(await res.arrayBuffer(), url.startsWith("http") ? "Google Planilhas" : url);
    } catch (e) {
      let c = null; try { c = JSON.parse(localStorage.getItem("dash-data")); } catch (_) {}
      if (c && c.DATA) { DATA = c.DATA; setStatus("Sem conexão com a planilha. Mostrando a última versão salva (" + c.stamp + ")."); boot(); }
      else setStatus("Não consegui ler a planilha (" + url + "). Confira o link em config.js.");
      console.error(e);
    }
  }
  $("reload").addEventListener("click", load);
  if (CFG.EDIT_URL) { $("abrir").href = CFG.EDIT_URL; $("abrir").hidden = false; }

  // ---------- tema ----------
  function aplicaTema(t) { if (t === "light" || t === "dark") document.documentElement.setAttribute("data-theme", t); else document.documentElement.removeAttribute("data-theme");
    try { localStorage.setItem("dash-tema", t); } catch (_) {} if (DATA) renderAll(); }
  try { $("tema").value = localStorage.getItem("dash-tema") || "auto"; } catch (_) {}
  $("tema").addEventListener("change", (e) => aplicaTema(e.target.value));
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => DATA && renderAll());

  // ---------- abas ----------
  let TAB = "geral";
  document.querySelectorAll(".tabs button").forEach((b) => b.addEventListener("click", () => {
    TAB = b.dataset.tab; document.querySelectorAll(".tabs button").forEach((x) => x.classList.toggle("on", x === b));
    ["geral", "dividas", "rend", "metas", "sim"].forEach((t) => ($("tab-" + t).hidden = t !== TAB));
    try { localStorage.setItem("dash-aba", TAB); } catch (_) {} window.scrollTo({ top: 0 });
    $("filtros").style.display = TAB === "sim" ? "none" : ""; renderAll();
  }));

  // ---------- filtros ----------
  const F = { mes: null, de: null, ate: null, cartao: "", divida: "ativas", reemb: "", item: "", unid: "rs" };
  const opt = (sel, items, val) => { sel.innerHTML = items.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join(""); if (val != null) sel.value = val; };
  function itensLista() {
    const a = [["", "— nenhum —"]];
    DATA.fixos.forEach((f) => a.push(["F|" + f.desc, "Conta fixa: " + f.desc]));
    const seen = {};
    DATA.dividas.forEach((d, n) => { const key = d.desc + (d.compra ? " · compra " + fmtD(new Date(d.compra)) : ""); d.key = seen[key] ? key + " #" + n : key; seen[key] = 1; a.push(["P|" + d.key, "Parcelamento: " + d.key + " (" + d.paga + ")"]); });
    [...new Set(DATA.lanc.map((l) => l.desc))].forEach((d) => a.push(["L|" + d, "Lançamento: " + d]));
    return a;
  }
  function boot() {
    const fer = new Set(DATA.feriados);
    REF = referencia(new Date(), fer);
    const ms = DATA.meses.map((k, i) => [k, lab(k) + (DATA.tipo[i] === "Realizado" ? " (real)" : "")]);
    const clamp = (k) => (DATA.meses.includes(k) ? k : k < DATA.meses[0] ? DATA.meses[0] : DATA.meses[DATA.meses.length - 1]);
    const def = clamp(mesAtual());
    let saved = {}; try { saved = JSON.parse(localStorage.getItem("dash-filtros")) || {}; } catch (_) {}
    // o mês sempre abre no mês de referência (regra do 5º dia útil)
    F.mes = def;
    F.de = DATA.meses.includes(saved.de) ? saved.de : def;
    F.ate = DATA.meses.includes(saved.ate) ? saved.ate : clamp(addM(def, 14));
    ["cartao", "divida", "reemb", "unid"].forEach((k) => saved[k] != null && (F[k] = saved[k]));
    F.item = "";
    opt($("fMes"), ms, F.mes); opt($("fDe"), ms, F.de); opt($("fAte"), ms, F.ate);
    const formas = [["", "Todos"]].concat(DATA.cartoes.map((c) => [c, c]), [["Conta (Pix/Boleto/Débito)", "Conta (Pix/Boleto/Débito)"], ["VR", "VR"]]);
    if (!formas.some((f) => f[0] === F.cartao)) F.cartao = "";
    opt($("fCartao"), formas, F.cartao);
    $("fDivida").value = F.divida; $("fReemb").value = F.reemb; $("fUnid").value = F.unid;
    opt($("fItem"), itensLista(), F.item);
    document.querySelectorAll(".cardname").forEach((e) => (e.textContent = DATA.cartoes[0] || "Inter"));
    // simulação
    SIM_DEF = DATA.plano[def] && DATA.plano[def].teto > 0 ? def : clamp(addM(def, 1)); SIM = null;
    renderAll();
  }
  [["fMes","mes"],["fDe","de"],["fAte","ate"],["fCartao","cartao"],["fDivida","divida"],["fReemb","reemb"],["fItem","item"],["fUnid","unid"]].forEach(([id, k]) =>
    $(id).addEventListener("change", (e) => { F[k] = e.target.value;
      if (diffM(F.de, F.ate) < 0) { const t = F.de; F.de = F.ate; F.ate = t; $("fDe").value = F.de; $("fAte").value = F.ate; }
      try { const s = Object.assign({}, F); delete s.mes; delete s.item; localStorage.setItem("dash-filtros", JSON.stringify(s)); } catch (_) {}
      renderAll(); }));
  // filtros recolhíveis (fechados por padrão no celular)
  const isMobile = () => window.matchMedia("(max-width: 700px)").matches;
  let fAberto = null; try { fAberto = localStorage.getItem("dash-filtros-aberto"); } catch (_) {}
  function setFiltros(open) { $("filtros").classList.toggle("collapsed", !open);
    document.querySelectorAll("#filtros .f:not(.keep)").forEach((el) => (el.style.display = open ? "" : "none"));
    if (isMobile()) $("filtros").style.position = open ? "static" : ""; $("fToggle").setAttribute("aria-expanded", open); $("fToggle").textContent = open ? "Fechar filtros ▴" : "Filtros ▾"; }
  setFiltros(fAberto != null ? fAberto === "1" : !isMobile());
  $("fToggle").addEventListener("click", () => { const open = $("filtros").classList.contains("collapsed"); setFiltros(open); try { localStorage.setItem("dash-filtros-aberto", open ? "1" : "0"); } catch (_) {} });
  function resumoFiltros() {
    const t = [lab(F.de) + "–" + lab(F.ate)];
    if (F.cartao) t.push(F.cartao); if (F.divida !== "ativas") t.push($("fDivida").selectedOptions[0].text); if (F.reemb) t.push($("fReemb").selectedOptions[0].text);
    if (F.item) t.push(F.item.slice(2)); if (F.unid === "pct") t.push("em %");
    $("fAtivos").textContent = t.join(" · ");
  }
  $("extBtn").addEventListener("click", () => { EXT_OPEN = !EXT_OPEN; renderExtrato(F.mes); if (EXT_OPEN) $("extrato").scrollIntoView({ behavior: "smooth", block: "start" }); });
  $("recPessoa").addEventListener("change", (e) => { RECP = e.target.value; try { localStorage.setItem("rec-pessoa", RECP); } catch (_) {} renderDividas(); });
  $("cabeValor").addEventListener("input", () => renderCabe(F.mes));
  $("gastoManual").addEventListener("input", (e) => { const kk = addM(F.mes, -1); try { const v = e.target.value; v === "" ? localStorage.removeItem("gasto-" + kk) : localStorage.setItem("gasto-" + kk, v); } catch (_) {} renderGeral(true); });

  // ---------- cálculos ----------
  const idx = (k) => DATA.meses.indexOf(k);
  const ativo = (o, k) => (!o.ini || o.ini <= k) && (!o.fim || o.fim >= k);
  const range = () => DATA.meses.slice(idx(F.de), idx(F.ate) + 1);
  const passaForma = (paga) => !F.cartao || paga === F.cartao;
  function gastoInter(k) {
    let m = null; try { m = localStorage.getItem("gasto-" + k); } catch (_) {}
    if (m !== null && m !== "" && !isNaN(+m)) return { v: +m, fonte: "valor informado" };
    return { v: (DATA.plano[k] || {}).ja || 0, fonte: "aba Lançamentos" };
  }
  function mesAtual() { const h = new Date(); return h.getFullYear() + "-" + pad(h.getMonth() + 1); }
  function semanasRestantes(k) {
    if (!REF) return 0;
    const cal = mesAtual();
    if (k === cal) { const h = new Date(); const ult = new Date(h.getFullYear(), h.getMonth() + 1, 0).getDate(); return (ult - h.getDate() + 1) / 7; }
    if (k < cal) return 0;
    const [y, m] = k.split("-").map(Number); return new Date(y, m, 0).getDate() / 7;
  }
  function dividaInfo(d, k) {
    const at = ativo(d, k); let rest = null; const k1 = addM(k, 1);
    if (d.fim) { const i0 = d.ini && d.ini > k1 ? d.ini : k1; rest = Math.max(0, diffM(i0, d.fim) + 1); }
    const tot = d.totalParc; const pagas = tot && rest != null ? Math.max(0, tot - rest) : null;
    const atual = !tot || rest == null ? null : at ? pagas + "/" + tot : d.ini && d.ini > k ? "0/" + tot : "quitado";
    return { ativa: at, rest, pagas, atual, aPagar: rest == null ? null : rest * d.valor, totalCompra: tot ? tot * d.valor : null };
  }
  function diaVenc(k, dia) { if (!dia) return null; const [y, m] = k.split("-").map(Number); const ult = new Date(y, m, 0).getDate(); return new Date(y, m - 1, Math.min(dia, ult), 12); }
  function receber(k) {
    const out = [];
    const push = (pessoa, valor, item, origem, quando, base) => { if (pessoa && pessoa !== "-" && valor > 0.004) out.push({ pessoa, valor, item, origem, quando, base: base || item }); };
    DATA.fixos.filter((f) => ativo(f, k)).forEach((f) => { const q = diaVenc(k, f.dia); push(f.q1, f.r1, f.desc, "Conta fixa", q); push(f.q2, f.r2, f.desc, "Conta fixa", q); });
    DATA.dividas.filter((d) => ativo(d, k)).forEach((d) => { const q = diaVenc(k, DATA.venc[d.paga]); const inf = dividaInfo(d, k);
      const t = d.desc + (inf.atual ? " (" + inf.atual + ")" : ""); push(d.q1, d.r1, t, "Parcela " + d.paga, q, d.desc); push(d.q2, d.r2, t, "Parcela " + d.paga, q, d.desc); });
    DATA.lanc.filter((l) => l.ini && l.ini <= k && l.fim >= k).forEach((l) => { const q = diaVenc(k, DATA.venc[l.paga]); push(l.q1, l.r1, l.desc, "Lançamento", q); push(l.q2, l.r2, l.desc, "Lançamento", q); });
    return out;
  }
  function serieItem(key) {
    const [tp, desc] = [key.slice(0, 1), key.slice(2)];
    const rg = DATA.meses;
    if (tp === "F") { const f = DATA.fixos.find((x) => x.desc === desc); return { obj: f, tipo: "Conta fixa", vals: rg.map((k) => (ativo(f, k) ? f.valor : 0)), meu: rg.map((k) => (ativo(f, k) ? f.custo : 0)) }; }
    if (tp === "P") { const d = DATA.dividas.find((x) => x.key === desc); return { obj: d, tipo: "Parcelamento", vals: rg.map((k) => (ativo(d, k) ? d.valor : 0)), meu: rg.map((k) => (ativo(d, k) ? d.custo : 0)) }; }
    const ls = DATA.lanc.filter((x) => x.desc === desc); const on = (l, k) => l.ini && l.ini <= k && l.fim >= k;
    return { obj: ls[0], tipo: "Lançamento", vals: rg.map((k) => ls.reduce((s, l) => s + (on(l, k) ? l.valor : 0), 0)), meu: rg.map((k) => ls.reduce((s, l) => s + (on(l, k) ? l.custo : 0), 0)) };
  }

  // ---------- gráficos ----------
  const charts = {};
  const pal = () => [css("--s1"), css("--s2"), css("--s3"), css("--s4"), css("--s5"), css("--s6")];
  function baseOpts(stacked, fmtFn, extra = {}) {
    const ink = css("--ink2"), line = css("--grid") || css("--line");
    const o = Object.assign({
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 }, interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "top", labels: { color: ink, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: "rectRounded", padding: 14 } },
        tooltip: { backgroundColor: css("--tip") || "rgba(8,20,40,.92)", titleColor: "#fff", bodyColor: "#e6eefb", borderColor: css("--accent"), borderWidth: 1, padding: 10, cornerRadius: 10, boxPadding: 4, callbacks: { label: (c) => " " + c.dataset.label + ": " + fmtFn(c.raw) } },
        datalabels: { color: stacked ? "#fff" : ink, font: { size: 10, weight: "600" }, anchor: stacked ? "center" : "end", align: stacked ? "center" : "end", offset: 1, clamp: true,
          formatter: (v) => (v == null || Math.abs(v) < 0.005 ? "" : fmtFn(v)),
          display: (c) => { const v = c.dataset.data[c.dataIndex]; if (!v) return false; if (!stacked) return true;
            const max = Math.max(...c.chart.data.datasets.flatMap((d) => d.data.map((x) => Math.abs(x || 0)))); return Math.abs(v) > max * 0.06; } },
      },
      scales: { x: { stacked, grid: { display: false }, ticks: { color: ink } },
        y: { stacked, grid: { color: line }, border: { display: false }, ticks: { color: ink, callback: (v) => fmtFn(v) } } },
    }, extra);
    if (o.indexAxis === "y") { o.scales.y.ticks = { color: ink, autoSkip: false }; o.scales.x.ticks = { color: ink, callback: (v) => fmtFn(v) }; o.scales.x.grid = { color: line }; o.scales.y.grid = { display: false }; o.layout = { padding: { right: 60 } }; }
    return o;
  }
  function draw(id, type, labels, datasets, opts) {
    if (charts[id]) charts[id].destroy(); charts[id] = new Chart($(id), { type, data: { labels, datasets }, options: opts });
    charts[id]._raw = { type, labels, datasets, opts };
    const box = $(id).parentElement;
    if (!box.querySelector(".zoom")) { const b = document.createElement("button"); b.className = "zoom"; b.type = "button"; b.title = "Ampliar gráfico"; b.setAttribute("aria-label", "Ampliar gráfico"); b.textContent = "⤢"; b.addEventListener("click", (e) => { e.stopPropagation(); abrirZoom(id); }); box.appendChild(b); }
    $(id).onclick = () => abrirZoom(id);
    return charts[id];
  }
  let zoomChart = null;
  function fecharZoom() { if (zoomChart) { zoomChart.destroy(); zoomChart = null; } const o = document.querySelector(".zoomov"); if (o) o.remove(); document.body.style.overflow = ""; }
  function abrirZoom(id) {
    const src = charts[id]; if (!src || !src._raw) return; fecharZoom(); const raw = src._raw;
    const card = $(id).closest(".card"); const titulo = card && card.querySelector("h2") ? card.querySelector("h2").textContent : "Gráfico";
    const fonte = card && card.dataset.fonte ? card.dataset.fonte.split("|") : null;
    const ov = document.createElement("div"); ov.className = "zoomov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
    ov.innerHTML = `<div class="zh"><h3>${esc(titulo)}</h3><div class="zh-a">${fonte ? `<button class="btn ghost" type="button" id="zFonte">Ver detalhes em ${esc(NOME_ABA[fonte[0]] || "outra aba")} ›</button>` : ""}<button class="btn" type="button" id="zClose">Fechar ✕</button></div></div><div class="dica">Dica: gire o celular para ver o gráfico maior.</div><div class="zb"><div class="zc"><canvas id="zCanvas"></canvas></div></div>`;
    document.body.appendChild(ov); document.body.style.overflow = "hidden";
    $("zClose").addEventListener("click", fecharZoom);
    if (fonte) $("zFonte").addEventListener("click", () => { fecharZoom(); vaiPara(fonte[0], fonte[1]); });
    const horiz = raw.opts.indexAxis === "y", n = raw.labels.length, zc = ov.querySelector(".zc");
    if (horiz) zc.style.height = Math.max(ov.querySelector(".zb").clientHeight - 24, n * 30 + 80) + "px";
    else if (raw.type !== "doughnut") { const minW = n * 48; if (minW > zc.clientWidth) zc.style.width = minW + "px"; }
    const data = { labels: raw.labels.slice(), datasets: raw.datasets.map((d) => Object.assign({}, d, { data: d.data.slice() })) };
    const opts = Object.assign({}, raw.opts, { maintainAspectRatio: false, animation: false });
    opts.plugins = Object.assign({}, raw.opts.plugins, { legend: Object.assign({}, raw.opts.plugins.legend, { display: true }), datalabels: Object.assign({}, raw.opts.plugins.datalabels, { font: { size: 12, weight: "600" }, display: (c) => !!c.dataset.data[c.dataIndex] }) });
    zoomChart = new Chart($("zCanvas"), { type: raw.type, data, options: opts });
  }
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") fecharZoom(); });
  // barras com degradê (cor cheia no topo, mais clara na base)
  const grad = (color) => (ctx) => { const ch = ctx.chart, ar = ch.chartArea; if (!ar || typeof color !== "string" || !/^#[0-9a-f]{6}$/i.test(color)) return color;
    const h = ch.options.indexAxis === "y", g = h ? ch.ctx.createLinearGradient(ar.left, 0, ar.right, 0) : ch.ctx.createLinearGradient(0, ar.bottom, 0, ar.top);
    g.addColorStop(0, color + "70"); g.addColorStop(1, color); return g; };
  const ds = (label, data, color, extra = {}) => Object.assign({ label, data, backgroundColor: typeof color === "string" && !("fill" in extra) && extra.type !== "line" ? grad(color) : color, borderColor: color, borderWidth: 0, borderRadius: 7, borderSkipped: "start", maxBarThickness: 42,
    pointBackgroundColor: color, pointBorderColor: css("--surface-solid") || "#fff", pointBorderWidth: 2, pointRadius: 3.5, pointHoverRadius: 6 }, extra);

  // ---------- render ----------
  function renderAll() { if (!DATA) return; resumoFiltros(); renderTopo(); document.querySelectorAll(".mesSel").forEach((e) => (e.textContent = lab(F.mes)));
    if (!SIM) SIM = simLoad();
    if (TAB === "sim") renderSim(); else if (TAB === "geral") renderResumo(F.mes); else { renderGeral(); renderDividas(); renderMetas(); }
    renderSimBanner(); setTimeout(() => { exportaTudo(); rotulaTabelas(); }, 0); }

  function renderTopo() {
    const ms = DATA.meses.filter((k) => DATA.plano[k]); if (ms.length < 2) return;
    const v = ms.map((k) => DATA.plano[k].reservaAc || 0), W = 160, H = 40, mn = Math.min(0, ...v), mx = Math.max(...v) || 1;
    const pts = v.map((y, n) => [(n / (v.length - 1)) * W, H - 3 - ((y - mn) / (mx - mn || 1)) * (H - 6)]);
    const d = pts.map((p, n) => (n ? "L" : "M") + p[0].toFixed(1) + " " + p[1].toFixed(1)).join(" ");
    const iRef = ms.indexOf(REF && ms.includes(REF.mes) ? REF.mes : ms[0]), pr = pts[Math.max(0, iRef)];
    $("topSpark").innerHTML = `<div><div class="l">Reserva prevista</div><div class="v">${R0(v[v.length - 1])}</div><div class="l">em ${lab(ms[ms.length - 1])}</div></div>
      <svg viewBox="-3 -3 ${W + 6} ${H + 6}" width="${W}" height="${H}" aria-hidden="true"><path d="${d} L${W} ${H} L0 ${H} Z" fill="currentColor" opacity=".12"/><path d="${d}" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"/><circle cx="${pr[0]}" cy="${pr[1]}" r="3.2" fill="currentColor"/></svg>`;
    $("topSpark").hidden = false;
  }
  function renderHoje() {
    const hoje = new Date(), cm = DATA.meses.includes(mesAtual()) ? mesAtual() : F.mes, pc = DATA.plano[cm] || {}, fc = formas(cm);
    const dias = Math.max(0, Math.ceil((REF.proximo - new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 12)) / 864e5));
    const fechado = !pc.teto, card = DATA.cartoes[0] || "Inter", fch = DATA.fechamento;
    const antesCorte = fch && hoje.getDate() <= fch && !fechado;
    $("hoje").innerHTML = `
      <div class="it"><div class="l">Hoje</div><div class="v">${fmtD(hoje)}</div></div>
      <div class="it"><div class="l">Mês vigente</div><div class="v">${labL(cm)}</div><div class="muted small">salário em uso: ${labL(REF.mes)}</div></div>
      <div class="it"><div class="l">Saldo em conta</div><div class="v">${R(DATA.saldo)}</div></div>
      <div class="it"><div class="l">Próximo 5º dia útil (salário)</div><div class="v">${fmtD(REF.proximo)}</div><div class="muted small">${dias === 0 ? "hoje" : "em " + dias + " dia" + (dias > 1 ? "s" : "")}</div></div>
      ${(() => { const dmh = debitoMes(cm), d = dmh.livre, c = Math.max(0, (pc.teto || 0) - gastoInter(cm).v), exh = dmh.ex;
        return `<div class="it"><div class="l">Posso gastar em ${lab(cm)} (débito/Pix)</div><div class="v ${d <= 0 ? "neg" : ""}">${R(d)}</div><div class="muted small">${exh.valor ? (exh.caiu ? "inclui " + R(exh.valor) + " do 13º · " : "+ " + R(exh.valor) + " do 13º em " + (exh.data ? fmtD(exh.data) : "breve") + " · ") : ""}cartão só parcelado: ${R(c)} livres no dinheiro de ${lab(addM(cm, 1))}</div></div>`; })()}
      ${(() => { const pd = contasMes(cm).filter((c) => !c.pago); return `<div class="it"><div class="l">Contas a pagar em ${lab(cm)}</div><div class="v ${pd.length ? "neg" : ""}">${R(pd.reduce((s, c) => s + c.valor, 0))}</div><div class="muted small">${pd.length ? pd.map((c) => esc(c.nome.replace(/^Fatura /, ""))).join(", ") : "tudo pago"}</div></div>`; })()}
      <div class="it"><div class="l">${lab(cm)}: teto / livre</div><div class="v">${R0(pc.teto)} / ${R0(pc.livre)}</div><div class="muted small">o que não gastar vai para a reserva</div></div>`;
  }

  // 13º/férias livres do mês k: só entram no saldo a partir do dia em que caem
  function extraMes(k) {
    const v = (DATA.plano[k] || {}).fora || 0; if (!v) return { valor: 0, caiu: false, data: null };
    const d = (DATA.d13 || {})[k.slice(5)] || null, hoje = new Date(), h = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
    const caiu = d ? h >= new Date(d.getFullYear(), d.getMonth(), d.getDate()) : mesAtual() >= k;
    return { valor: v, caiu, data: d };
  }
  function debitoMes(k) { // dinheiro para débito/Pix no mês k = salário de k (linha k-1) + 13º já caído
    const pk = addM(k, -1), pd = DATA.plano[pk] || {}, ex = extraMes(k), teto = (pd.teto || 0) + (ex.caiu ? ex.valor : 0), usado = gastoInter(pk).v;
    const salData = dataSalario(k), h = new Date(); h.setHours(12, 0, 0, 0);
    return { teto, usado, livre: Math.max(0, teto - usado), ex, pd, salData, salCaiu: !salData || h >= salData };
  }
  function dataSalario(k) { if (!k) return null; const [y, m] = k.split("-").map(Number); return quintoDiaUtil(y, m - 1, new Set(DATA.feriados || [])); }
  function debitoDisp(k, rest) {
    const p = DATA.plano[k] || {}, prev = DATA.plano[addM(k, -1)] || {};
    const sobraSal = (prev.espaco || 0) - (p.jaPix || 0) - (k > DATA.planIni ? prev.jaCartao || 0 : 0);
    return { sobraSal, deb: Math.max(0, Math.min(rest, sobraSal)) };
  }
  function formas(k) {
    const p = DATA.plano[k] || {}, ja = gastoInter(k).v, fechado = !p.teto && k < DATA.planIni;
    return { fechado, ja, teto: p.teto || 0, ainda: fechado ? 0 : (p.teto || 0) - ja, fora: p.fora || 0, vr: p.vr || 0, jaCartao: p.jaCartao || 0, jaPix: p.jaPix || 0 };
  }
  // Quanto cabe numa compra no cartão: cada parcela ocupa o espaço de um mês de compras (k, k+1, ...)
  function renderCabe(k) {
    // linhas do plano = dinheiro de um salário: débito em k usa a linha k-1; parcela n no cartão usa a linha k+n-1
    const box = $("cabeCard"), card = esc(DATA.cartoes[0] || "Inter");
    const room = (m) => { const p = DATA.plano[m]; return p && p.teto ? Math.max(0, p.teto - gastoInter(m).v) : 0; };
    const sal = (m) => lab(addM(m, 1));
    const ms = []; for (let j = 0; j < 6; j++) { const m = addM(k, j); if (!DATA.plano[m] || !DATA.plano[m].teto) break; ms.push(m); }
    if (!ms.length) { box.hidden = true; return; } box.hidden = false;
    const val = parseFloat($("cabeValor").value) || 0, dmc = debitoMes(k);
    $("cabeInfo").innerHTML = `Prefira o débito/Pix. No cartão, use só parcelado: cada parcela cai numa fatura e usa o dinheiro daquele salário (já descontando o que você já comprou). A compra cabe se nenhum mês estoura.`;
    const roomBase = room; const room2 = (m) => (m === addM(k, -1) ? dmc.livre : roomBase(m));
    const opts = [{ nome: "À vista no débito/Pix", n: 1, ms: [addM(k, -1)], deb: true }].concat(ms.map((_, j) => ({ nome: j === 0 ? "À vista no cartão (evite)" : (j + 1) + "x no cartão", n: j + 1, ms: ms.slice(0, j + 1) })));
    let rows = `<tr><th>Forma</th><th>Compra máxima</th><th>Parcela máx.</th>${val ? "<th>Sua compra</th><th>Cabe?</th><th>Sobra no mês mais apertado</th>" : "<th>Mês que limita</th>"}</tr>`;
    opts.forEach((o) => {
      const lim = o.ms.reduce((a, m) => (room2(m) < room2(a) ? m : a), o.ms[0]), pmax = room2(lim), max = pmax * o.n;
      const cls = o.deb ? "pref" : o.n === 1 ? "evite" : "";
      if (!val) { rows += `<tr class="${cls}"><td>${o.nome}</td><td class="big3">${R(max)}</td><td>${R(pmax)}</td><td class="muted">dinheiro de ${sal(lim)}</td></tr>`; return; }
      const parc = val / o.n, sobra = Math.min(...o.ms.map((m) => room2(m) - parc)), ok = sobra >= -0.005;
      rows += `<tr class="${ok ? "ok" : "no"} ${cls}"><td>${o.nome}</td><td>${R(max)}</td><td>${R(pmax)}</td><td class="big3">${o.n === 1 ? R(val) : o.n + " × " + R(parc)}</td>` +
        `<td class="${ok ? "pos" : "neg"}">${ok ? "Cabe" : "Não cabe"}</td><td class="${sobra < 0 ? "neg" : ""}">${R(sobra)}</td></tr>`;
    });
    $("cabeTbl").innerHTML = rows;
    const all = [addM(k, -1)].concat(ms);
    $("cabeMeses").innerHTML = `<tr><th>Dinheiro do salário de</th>${all.map((m) => `<th>${sal(m)}</th>`).join("")}</tr>` +
      `<tr><td class="muted">Paga</td>${all.map((m, j) => `<td>${j === 0 ? "débito de " + lab(k) : "fatura de " + sal(m)}</td>`).join("")}</tr>` +
      `<tr><td>Ainda livre</td>${all.map((m) => `<td><b>${R(room2(m))}</b></td>`).join("")}</tr>`;
  }
  function renderFormas() { const el = $("formas"); if (el) { el.innerHTML = ""; el.hidden = true; } }

  function renderGeral(onlyHero) {
    const k = F.mes, i = idx(k), p = DATA.plano[k] || {}, pr = DATA.proj, pct = F.unid === "pct";
    const real = DATA.tipo[i] === "Realizado";
    $("hTag").innerHTML = real ? '<span class="tag">realizado</span>' : '<span class="tag prev">previsto</span>';
    $("hMes").textContent = lab(k);
    // dinheiro do mês = salário de k (linha k-1 do plano): paga a fatura de k e o débito/Pix de k
    const pk = addM(k, -1), pd = DATA.plano[pk] || {}, sem = semanasRestantes(k);
    const fechado = !(DATA.plano[k] || {}).teto && k < DATA.planIni;
    const gi = gastoInter(pk), dm = debitoMes(k), ex = dm.ex, teto = pd.teto || 0, deb = dm.livre;
    const card = DATA.cartoes[0] || "Inter", fch = DATA.fechamento, hj = new Date(), fatK = addM(k, 1), vc = DATA.venc[card];
    const credAinda = (p.teto || 0) - gastoInter(k).v;
    const antes = mesAtual() === k && fch && hj.getDate() <= fch;
    $("hRestante").textContent = R(deb); $("hRestante").classList.toggle("neg", deb <= 0);
    if (!dm.salCaiu && dm.salData) $("hTag").innerHTML += ` <span class="tag warn">entra em ${fmtD(dm.salData)}</span>`;
    { const sd = simDelta(pk); if (sd && Math.abs(sd.d) > 0.005) $("hTag").innerHTML += ` <span class="tag sim">🧪 com simulação: ${R(Math.max(0, deb + sd.d))}</span>`; }
    const uso = teto ? gi.v / teto : 1; $("hBar").style.width = Math.min(100, uso * 100) + "%"; $("hBar").classList.toggle("over", uso > 1);
    $("hBar").parentElement.hidden = false;
    $("hResumo").innerHTML = (teto
      ? `<b>Dinheiro de ${lab(k)}: ${R(teto)}</b> (o que o salário de ${lab(k)} pode pagar depois das contas, já separando a meta de reserva).<br>` +
        `Já usado ${R(gi.v)}: fatura de ${lab(k)} com compras do plano ${R(pd.jaCartao)} · débito/Pix em ${lab(k)} ${R(pd.jaPix)}.<br>` +
        (deb > 0 && sem >= 1 ? `Dá ${R(deb / sem)} por semana nas ${sem.toFixed(1).replace(".", ",")} semanas que faltam. ` : "")
      : `O salário de ${lab(k)} já foi todo para as contas e a fatura (sobra prevista ${R(pd.reserva || 0)}). Neste mês, não há dinheiro livre para o débito. `) +
      `O que não gastar vai para a reserva.`;
    {
      const el = $("hComo"); el.hidden = false; $("hBar").parentElement.hidden = true;
      const pkC = DATA.plano[k] || {}, tetoC = pkC.teto || 0, usoC = gastoInter(k).v;
      const barra = (cls, titulo, livre, tot, usado, nota) => `
        <div class="bx ${cls}"><div class="bx-h"><span>${titulo}</span><b>${R(Math.max(0, livre))}</b></div>
          <div class="bx-bar"><span style="width:${tot ? Math.min(100, (usado / tot) * 100) : 100}%"></span></div>
          <div class="bx-n">${nota}</div></div>`;
      el.innerHTML =
        barra("deb", `💵 Débito/Pix em ${lab(k)}`, deb, dm.teto, Math.min(gi.v, dm.teto || gi.v),
          (dm.teto ? `livre de ${R(dm.teto)} (salário de ${lab(k)}${ex.caiu ? " + 13º/férias" : ""}) · já usado ${R(gi.v)} (fatura ${R(pd.jaCartao)} + débito ${R(pd.jaPix)}). <b>Gaste primeiro aqui.</b>` : `o salário de ${lab(k)} já foi todo para as contas e a fatura.`)) +
        (ex.valor ? `<div class="bx-x">${ex.caiu ? `✓ Inclui <b>${R(ex.valor)}</b> do 13º/férias, que já caiu${ex.data ? " em " + fmtD(ex.data) : ""}.` : `⏳ <b>+ ${R(ex.valor)}</b> do 13º/férias entra aqui sozinho${ex.data ? " em " + fmtD(ex.data) : " quando cair"} (ainda não caiu).`}</div>` : "") +
        barra("cred", `💳 Cartão agora (cai na fatura de ${lab(fatK)})`, credAinda, tetoC, Math.min(usoC, tetoC || usoC),
          (antes ? `<span class="icon-warn">Até dia ${fch} a compra ainda cai na fatura de ${lab(k)} (usa a barra de cima).</span> ` : "") +
          `livre de ${R(tetoC)} do salário de ${lab(fatK)} · já comprometido ${R(usoC)}. Só parcelado ou assinatura: o que usar aqui sai do débito de ${lab(fatK)}.` +
          (ex.caiu && ex.valor ? ` Com o 13º já na conta, a fatura de ${lab(fatK)} também pode usar o que sobrar dele no débito (é o mesmo dinheiro: não some).` : ""));
    }
    let man = null; try { man = localStorage.getItem("gasto-" + addM(k, -1)); } catch (_) {}
    if (!onlyHero) $("gastoManual").value = man ?? "";
    const ent = pr.entradas[i] || 0;
    renderCabe(k); renderFormas(k); renderRend(k); renderSaude(k); renderRitmo(k); renderLimites(k); renderVR();
    const kp = [
      ["13º/férias livres", R(p.fora), "extra do mês, além do teto"],
      ["VR / Flash do mês", R(p.vr || pr.vr[i]), "usado nas compras do supermercado"],
      ["Limite do Inter a configurar", R(p.limite), "fatura de " + lab(p.fatura) + ": " + R(p.faturaInter)],
      ["Sobra estimada do mês", R(pr.sobra[i]), "antes das compras novas", pr.sobra[i] < 0 ? "neg" : ""],
      ["Renda comprometida", P(ent ? pr.comprom[i] / ent : 0), "fixos + parcelas: " + R(pr.comprom[i])],
      ["Vai para a reserva (se gastar o sugerido)", R(p.reserva), P(p.meta ? (p.reserva || 0) / p.meta : 0) + " da meta de " + R0(p.meta) + " (" + P(DATA.pct) + ")", (p.reserva || 0) < 0 ? "neg" : ""],
    ];
    $("kpis").innerHTML = kp.map(([l, v, s, c]) => `<div class="kpi"><div class="l">${l}</div><div class="v ${c || ""}">${v}</div><div class="s">${s}</div></div>`).join("");
    if (onlyHero) return;

    // item filtrado
    const rg = range(), L = rg.map(lab), ix = rg.map(idx), cor = pal();
    if (F.item) {
      const si = serieItem(F.item), o = si.obj || {}; $("itemCard").hidden = false;
      $("itemTitulo").textContent = si.tipo + ": " + F.item.slice(2);
      const extra = si.tipo === "Parcelamento" ? (() => { const inf = dividaInfo(o, k); return (inf.atual ? " · parcela " + inf.atual : "") + (o.fim ? " · última em " + lab(o.fim) : " · sem data de fim"); })() : o.dia ? " · vence dia " + o.dia : "";
      $("itemInfo").textContent = (o.paga || "") + (o.cat ? " · " + o.cat : "") + extra;
      const reb = si.vals[i] - si.meu[i];
      $("itemKpis").innerHTML = [["Valor em " + lab(k), R(si.vals[i])], ["Meu custo", R(si.meu[i])], ["Reembolso de terceiros", R(reb)], ["% das entradas do mês", P(ent ? si.vals[i] / ent : 0)], ["Total no período do gráfico", R(ix.reduce((s, j) => s + si.vals[j], 0))]]
        .map(([l, v]) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div></div>`).join("");
      draw("cItem", "bar", L, [ds("Valor cheio", ix.map((j) => si.vals[j]), cor[0]), ds("Meu custo", ix.map((j) => si.meu[j]), cor[2])], baseOpts(false, (v) => R0(v)));
    } else $("itemCard").hidden = true;

    // estimativa
    const pc = (v) => (ent ? P(v / ent) : "—");
    const lin = [["Salário e extras", pr.renda[i]], ["13º salário", pr.d13[i]], ["Férias", pr.ferias[i]], ["Ajuste pós-férias", pr.ajuste[i]], ["Reembolsos de terceiros", pr.reemb[i]]]
      .filter(([, v]) => v).map(([t, v]) => `<tr><td>${t}</td><td>${R(v)}</td><td>${pc(v)}</td></tr>`).join("");
    const sai = [["Contas fixas", pr.fixos[i]], ["Parcelas", pr.parcelas[i]], ["Outras cobranças e compras", pr.variaveis[i]]].map(([t, v]) => `<tr><td>${t}</td><td class="neg">${R(v)}</td><td>${pc(-v)}</td></tr>`).join("");
    const fat = Object.entries(DATA.faturas).filter(([c]) => passaForma(c)).map(([c, v]) => `<tr><td>Fatura ${esc(c)}</td><td>${R(v[i])}</td><td>${pc(v[i])}</td></tr>`).join("");
    $("tEstimativa").innerHTML = `<tr><th>${lab(k)} ${real ? '<span class="tag">realizado</span>' : '<span class="tag prev">previsto</span>'}</th><th>Valor</th><th class="pc">% das<br>entradas</th></tr>` + lin +
      `<tr class="total"><td>Entradas</td><td>${R(ent)}</td><td>${ent ? "100%" : "—"}</td></tr>` + sai +
      `<tr class="total"><td>Sobra do mês</td><td class="${pr.sobra[i] < 0 ? "neg" : "pos"}">${R(pr.sobra[i])}</td><td>${pc(pr.sobra[i])}</td></tr>` +
      (fat ? `<tr><td colspan="3" class="muted small" style="text-align:left">Faturas (já dentro das saídas)</td></tr>` + fat : "");

    // reserva
    const metaOk = rg.filter((m) => { const q = DATA.plano[m] || {}; return q.meta && q.reserva >= q.meta - 0.01; }).length;
    const ultimo = DATA.meses[idx(F.ate)], pu = DATA.plano[ultimo] || {};
    const aporte = p.reserva || 0, apPct = p.meta ? aporte / p.meta : 0;
    $("reserva").innerHTML = `<div class="kpis" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">
        <div class="kpi"><div class="l">Reserva hoje</div><div class="v">${R(DATA.reservaInformada)}</div></div>
        <div class="kpi"><div class="l">Prevista no fim de ${lab(k)}</div><div class="v">${R(p.reservaAc)}</div></div>
        <div class="kpi"><div class="l">Prevista em ${lab(ultimo)}</div><div class="v">${R(pu.reservaAc)}</div></div></div>
      <div class="meter"><div class="t"><span>Aporte de ${lab(k)}: ${R(aporte)}</span><span>${P(apPct)} da meta</span></div>
      <div class="bar"><span style="width:${Math.max(0, Math.min(100, apPct * 100))}%;background:${apPct >= 0.999 ? css("--good") : css("--s4")}"></span></div></div>
      <p class="${metaOk === rg.length ? "icon-ok" : "icon-warn"}">${metaOk} de ${rg.length} meses do período batem a meta de ${P(DATA.pct)} do salário líquido (${R0(p.meta)} em ${lab(k)}). O que não for gasto no mês também vai para a reserva.</p>`;

    // gráficos
    const den = (j) => (pct ? DATA.proj.entradas[j] || 1 : 1);
    const conv = (arr) => arr.map((v, n) => (v == null ? null : pct ? v / den(ix[n]) : v));
    const fmt = pct ? (v) => P(v) : (v) => R0(v);
    const sel = rg.indexOf(k), hl = (c) => rg.map((_, n) => (n === sel ? c : c + "B3"));
    const pv = (f) => rg.map((m) => (DATA.plano[m] || {})[f] || 0);
    draw("cGastar", "bar", L, [ds("Teto do mês (cartão + Pix/débito)", conv(pv("teto")), hl(cor[0])), ds("13º/férias livres", conv(pv("fora")), hl(cor[2])), ds("VR / Flash", conv(pv("vr")), hl(cor[3]))], baseOpts(true, fmt));
    draw("cReserva", "line", L, [ds("Reserva acumulada", pv("reservaAc"), cor[0], { fill: false, tension: 0.2, borderWidth: 2, pointRadius: 4, pointBackgroundColor: cor[0] })], baseOpts(false, (v) => R0(v)));
    draw("cAporte", "bar", L, [ds("Vai para a reserva", pv("reserva"), hl(cor[2])), { type: "line", label: "Meta (" + P(DATA.pct) + " do líquido)", data: pv("meta"), borderColor: css("--muted"), borderDash: [6, 4], borderWidth: 2, pointRadius: 0, datalabels: { display: false } }], baseOpts(false, (v) => R0(v)));
    draw("cLimite", "bar", L, [ds("Limite a configurar", pv("limite"), hl(cor[0])), ds("Fatura que vai chegar", pv("faturaInter"), hl(cor[1]))], baseOpts(false, (v) => R0(v)));
    const cards = Object.keys(DATA.faturas).filter(passaForma);
    draw("cFaturas", "bar", L, cards.map((c, n) => ds(c, conv(ix.map((j) => DATA.faturas[c][j])), cor[n % 6])), baseOpts(true, fmt));
    const pos = (a) => ix.map((j) => Math.abs(a[j]));
    draw("cRenda", "bar", L, [ds("Contas fixas", conv(pos(pr.fixos)), cor[0]), ds("Parcelas", conv(pos(pr.parcelas)), cor[1]), ds("Outras cobranças e compras", conv(pos(pr.variaveis)), cor[2]), ds("Sobra", conv(ix.map((j) => Math.max(0, pr.sobra[j]))), cor[3])], baseOpts(true, fmt));
    draw("cSobra", "bar", L, [ds("Sobra do mês", ix.map((j) => pr.sobra[j]), ix.map((j, n) => (pr.sobra[j] < 0 ? css("--bad") : n === sel ? cor[0] : cor[0] + "B3")))], baseOpts(false, (v) => R0(v)));
    // categorias: realizado do mês se houver; senão, meu custo previsto (fixos + parcelas) por categoria
    const ci = DATA.cat.meses.indexOf(k); let cats, nota;
    const realCats = ci >= 0 ? Object.entries(DATA.cat.itens).map(([n, o]) => [n, num(o.vals[ci])]).filter(([, v]) => v > 0.005) : [];
    if (realCats.length) { cats = realCats; nota = `Realizado em ${lab(k)}, já sem reembolsos.`; }
    else { const m = {}; DATA.fixos.filter((f) => ativo(f, k)).forEach((f) => (m[f.cat] = (m[f.cat] || 0) + f.custo)); DATA.dividas.filter((d) => ativo(d, k)).forEach((d) => (m[d.cat] = (m[d.cat] || 0) + d.custo));
      DATA.lanc.filter((l) => l.ini && l.ini <= k && l.fim >= k).forEach((l) => (m[l.cat || "Outros"] = (m[l.cat || "Outros"] || 0) + l.custo));
      cats = Object.entries(m).filter(([, v]) => v > 0.005); nota = `Previsto para ${lab(k)}: contas fixas, parcelas e lançamentos (meu custo, sem reembolsos).`; }
    cats.sort((a, b) => b[1] - a[1]); $("catNota").textContent = nota;
    const cc = draw("cCat", "bar", cats.map((c) => c[0]), [ds(pct ? "% das entradas" : "Meu custo", cats.map((c) => (pct ? c[1] / (ent || 1) : c[1])), cor[0])], baseOpts(false, fmt, { indexAxis: "y" }));
    cc.options.plugins.legend.display = false; cc.update();
  }

  function statusReceber(k, rc) {
    const n = (t) => txt(t).toLowerCase().replace(/\s+/g, " "), pool = DATA.recItens.filter((r) => r.mes === k).map((r) => Object.assign({ resto: r.valor }, r));
    rc.forEach((r) => { let pago = 0, dt = ""; pool.filter((p) => n(p.pessoa) === n(r.pessoa) && n(p.item) === n(r.base) && p.resto > 0.004).forEach((p) => { const u = Math.min(p.resto, r.valor - pago); if (u > 0) { p.resto -= u; pago += u; dt = dt || p.data; } });
      r.pago = pago; r.dataPago = dt; r.st = pago >= r.valor - 0.1 ? "pago" : pago > 0.004 ? "parcial" : "pendente"; });
    return pool.filter((p) => p.resto > 0.015);
  }
  function renderCompras(k, tid, rid) {
    let lc = DATA.lanc.filter((l) => l.data.slice(0, 7) === k && passaForma(l.paga));
    if (F.item && F.item.startsWith("L|")) lc = lc.filter((l) => l.desc === F.item.slice(2));
    lc.sort((a, b) => (a.data < b.data ? 1 : -1));
    const cart = lc.filter((l) => DATA.cartoes.includes(l.paga)), vr = lc.filter((l) => l.paga === "VR"), out = lc.filter((l) => !DATA.cartoes.includes(l.paga) && l.paga !== "VR");
    const sm = (a) => a.reduce((s, l) => s + l.total, 0);
    $(rid).textContent = lc.length ? `${lc.length} compra(s) em ${lab(k)} · cartão ${R(sm(cart))} · Flash ${R(sm(vr))}${out.length ? " · Pix/débito " + R(sm(out)) : ""}` : "Nenhuma compra lançada neste mês.";
    const cai = (l) => l.paga === "VR" ? "Flash (na hora)" : DATA.cartoes.includes(l.paga) ? "fatura de " + lab(l.ini) + (l.parcelado ? " até " + lab(l.fim) : "") : "na hora";
    $(tid).innerHTML = `<tr><th>Compra</th><th>Data</th><th>Forma</th><th>Categoria</th><th>Valor</th><th>Parcelas</th><th>Cai em</th></tr>` +
      lc.map((l) => `<tr><td>${esc(l.desc)}</td><td>${l.data.split("-").reverse().slice(0, 2).join("/")}</td><td>${esc(l.paga === "VR" ? "Flash (VR)" : l.paga)}</td><td>${esc(l.cat)}</td><td>${R(l.total)}</td><td>${l.parcelado ? l.n + "x de " + R(l.valor) : "à vista"}</td><td>${cai(l)}</td></tr>`).join("");
  }
  function renderDividas() {
    const k = F.mes, i = idx(k), ent = DATA.proj.entradas[i] || 0, cor = pal();
    // a receber
    renderRecPessoaSel();
    let rc = receber(k).filter((r) => passaPessoa(r.pessoa)); if (F.cartao) rc = rc.filter((r) => r.origem.includes(F.cartao) || (F.cartao.startsWith("Conta") && r.origem === "Conta fixa"));
    const porPessoa = {}; rc.forEach((r) => (porPessoa[r.pessoa] = (porPessoa[r.pessoa] || 0) + r.valor));
    $("rResumo").textContent = Object.entries(porPessoa).map(([p, v]) => p + ": " + R(v)).join(" · ") || "nada a receber";
    rc.sort((a, b) => a.pessoa.localeCompare(b.pessoa) || (a.quando || 0) - (b.quando || 0));
    const sobras = statusReceber(k, rc).filter((s) => passaPessoa(s.pessoa));
    rc.sort((a, b) => (a.st === "pago") - (b.st === "pago") || a.pessoa.localeCompare(b.pessoa) || (a.quando || 0) - (b.quando || 0));
    const tag = (r) => r.st === "pago" ? `<span class="tag good">Pago${r.dataPago ? " " + r.dataPago.split("-").reverse().slice(0, 2).join("/") : ""}</span>` : r.st === "parcial" ? `<span class="tag">Parcial: falta ${R(r.valor - r.pago)}</span>` : `<span class="tag bad">Pendente</span>`;
    const tot = {}; rc.forEach((r) => { const o = tot[r.pessoa] || (tot[r.pessoa] = { v: 0, p: 0 }); o.v += r.valor; o.p += r.pago; });
    $("rResumo").textContent = Object.entries(tot).map(([p, o]) => `${p}: recebido ${R(o.p)} · falta ${R(Math.max(0, o.v - o.p))}`).join("  |  ") || "nada a receber";
    $("tReceber").innerHTML = `<tr><th>Item</th><th>Pessoa</th><th>Origem</th><th>Valor</th><th>Receber até</th><th>Situação</th></tr>` +
      rc.map((r) => `<tr class="${r.st === "pago" ? "fora" : ""}"><td>${esc(r.item)}</td><td>${esc(r.pessoa)}</td><td>${esc(r.origem)}</td><td>${R(r.valor)}</td><td>${r.quando ? fmtD(r.quando) : "—"}</td><td>${tag(r)}</td></tr>`).join("") +
      Object.entries(tot).map(([p, o]) => `<tr class="total"><td>Total ${esc(p)}</td><td></td><td></td><td>${R(o.v)}</td><td></td><td>${o.p >= o.v - 0.015 ? '<span class="tag good">Tudo pago</span>' : '<span class="tag bad">Falta ' + R(o.v - o.p) + "</span>"}</td></tr>`).join("") +
      (sobras.length ? `<tr><td colspan="6" class="muted small" style="text-align:left">Recebido sem item correspondente: ${sobras.map((s) => esc(s.pessoa) + " " + R(s.resto) + " (" + esc(s.item) + ")").join(", ")}</td></tr>` : "");
    // fixos
    let fx = DATA.fixos.filter((f) => ativo(f, k) && passaForma(f.paga));
    if (F.reemb === "com") fx = fx.filter((f) => f.r1 + f.r2 > 0); else if (F.reemb === "sem") fx = fx.filter((f) => !(f.r1 + f.r2));
    const selItem = F.item && F.item.startsWith("F|") ? F.item.slice(2) : null;
    const tv = fx.reduce((s, f) => s + f.valor, 0), tm = fx.reduce((s, f) => s + f.custo, 0);
    $("fResumo").textContent = `${fx.length} contas · ${R(tv)} no total · minha parte ${R(tm)} (${P(ent ? tm / ent : 0)} das entradas)`;
    $("tFixos").innerHTML = `<tr><th>Conta</th><th>Categoria</th><th>Paga com</th><th>Valor</th><th>Divisão</th><th>Minha parte</th><th>Vence</th></tr>` +
      fx.map((f) => `<tr class="${f.desc === selItem ? "sel" : ""}"><td>${esc(f.desc)}</td><td>${esc(f.cat)}</td><td>${esc(f.paga)}</td><td>${R(f.valor)}</td><td>${f.r1 + f.r2 ? [f.r1 ? esc(f.q1) + " " + R(f.r1) : "", f.r2 ? esc(f.q2) + " " + R(f.r2) : ""].filter(Boolean).join(" + ") : "—"}</td><td>${R(f.custo)}</td><td>${f.dia ? "dia " + f.dia : "—"}</td></tr>`).join("") +
      `<tr class="total"><td>Total</td><td></td><td></td><td>${R(tv)}</td><td>${R(tv - tm)}</td><td>${R(tm)}</td><td></td></tr>`;
    // parcelamentos
    let dv = DATA.dividas.map((d) => Object.assign({}, d, dividaInfo(d, k))).filter((d) => passaForma(d.paga));
    if (F.reemb === "com") dv = dv.filter((d) => d.r1 + d.r2 > 0); else if (F.reemb === "sem") dv = dv.filter((d) => !(d.r1 + d.r2));
    if (F.divida === "ativas") dv = dv.filter((d) => d.ativa); else if (F.divida === "3m") dv = dv.filter((d) => d.ativa && d.fim && diffM(k, d.fim) <= 2);
    if (F.item && F.item.startsWith("P|")) dv = dv.filter((d) => d.key === F.item.slice(2));
    dv.sort((a, b) => ((a.fim || "9999") < (b.fim || "9999") ? -1 : (a.fim || "9999") > (b.fim || "9999") ? 1 : b.valor - a.valor));
    const act = dv.filter((d) => d.ativa); const tMes = act.reduce((s, d) => s + d.valor, 0), tCusto = act.reduce((s, d) => s + d.custo, 0), tPagar = dv.reduce((s, d) => s + (d.aPagar || 0), 0);
    $("dResumo").textContent = `${dv.length} parcelamento(s) · ${R(tMes)}/mês em ${lab(k)} (${P(ent ? tMes / ent : 0)} das entradas) · meu custo ${R(tCusto)}/mês · ${R(tPagar)} ainda a pagar`;
    const divis = (d) => { const a = []; if (d.r1) a.push(esc(d.q1) + ": " + R(d.r1) + "/parcela" + (d.rest ? " · " + R(d.r1 * d.rest) + " a receber" : "")); if (d.r2) a.push(esc(d.q2) + ": " + R(d.r2) + "/parcela" + (d.rest ? " · " + R(d.r2 * d.rest) + " a receber" : "")); return a.join("<br>") || "—"; };
    $("tDividas").innerHTML = `<tr><th>Descrição (como na fatura)</th><th>Compra</th><th>Cartão</th><th>Parcela</th><th>Total de parcelas</th><th>Pagas</th><th>Restantes</th><th>Paga no mês</th><th>Última</th><th>Valor total da compra</th><th>Ainda a pagar</th><th>Divisão / a receber</th><th>Meu custo/mês</th></tr>` +
      dv.map((d) => `<tr><td>${esc(d.desc)}</td><td>${d.compra ? fmtD(new Date(d.compra)) : "—"}</td><td>${esc(d.paga)}</td><td>${R(d.valor)}</td><td>${d.totalParc || "?"}</td><td>${d.pagas ?? "?"}</td><td>${d.rest ?? "—"}</td><td>${d.atual || "—"}</td><td>${d.fim ? lab(d.fim) : "sem data"}</td><td>${d.totalCompra ? R(d.totalCompra) : "—"}</td><td>${d.aPagar == null ? "—" : R(d.aPagar)}</td><td style="text-align:left">${divis(d)}</td><td>${R(d.custo)}</td></tr>`).join("") +
      `<tr class="total"><td>Total</td><td></td><td></td><td>${R(tMes)}</td><td></td><td></td><td></td><td></td><td></td><td></td><td>${R(tPagar)}</td><td></td><td>${R(tCusto)}</td></tr>`;
    const dvc = dv.filter((d) => d.aPagar).sort((a, b) => b.aPagar - a.aPagar);
    $("cDividas").parentElement.style.height = Math.max(260, dvc.length * 24 + 60) + "px";
    const ch = draw("cDividas", "bar", dvc.map((d) => { const t = d.desc.replace(" (a identificar)", "").replace(" — ", " · "); return (/fim /.test(t) ? t : t + " · até " + lab(d.fim)).slice(0, 50); }),
      [ds("Ainda a pagar", dvc.map((d) => d.aPagar), dvc.map((d) => cor[Math.max(0, DATA.cartoes.indexOf(d.paga)) % 6]))], baseOpts(false, (v) => R0(v), { indexAxis: "y", animation: false }));
    ch.options.plugins.legend.display = false; ch.update();
    renderCompras(k, "tLanc", "lResumo");
    renderContas(k); renderRecebidos(k); renderDevPessoa(k); renderExtrato(k); renderLibera(k); renderPrevReal();
  }

  // ---------- contas do mês: pago x pendente ----------
  const norm = (t) => txt(t).toLowerCase().replace(/\s+/g, " ");
  function contasMes(k) {
    const j = idx(k), out = [];
    DATA.fixos.filter((f) => ativo(f, k) && f.valor > 0 && !DATA.cartoes.includes(f.paga) && f.paga !== "VR").forEach((f) => out.push({ nome: f.desc, valor: f.valor, dia: f.dia, tipo: "Conta" }));
    DATA.dividas.filter((d) => ativo(d, k) && d.valor > 0 && !DATA.cartoes.includes(d.paga) && d.paga !== "VR").forEach((d) => out.push({ nome: d.desc, valor: d.valor, dia: null, tipo: "Parcela" }));
    DATA.cartoes.forEach((c) => { const v = j >= 0 ? (DATA.faturas[c] || [])[j] || 0 : 0; if (v > 0.005) out.push({ nome: "Fatura " + c, valor: v, dia: DATA.venc[c], tipo: "Fatura" }); });
    const pg = DATA.pagos.filter((p) => p.mes === k), usados = new Set();
    out.forEach((o) => { const p = pg.find((x, n) => !usados.has(n) && norm(x.nome) === norm(o.nome)); if (p) { usados.add(pg.indexOf(p)); o.pago = p; } });
    pg.forEach((p, n) => { if (!usados.has(n)) out.push({ nome: p.nome, valor: p.valor, dia: null, tipo: "Outro", pago: p }); });
    return out;
  }
  function renderContas(k, pre = "c") {
    const cs = contasMes(k), pagas = cs.filter((c) => c.pago), pend = cs.filter((c) => !c.pago);
    const tP = pagas.reduce((s, c) => s + (c.pago.valor || c.valor), 0), tF = pend.reduce((s, c) => s + c.valor, 0), tot = tP + tF;
    $(pre + "Resumo").textContent = `${lab(k)} · ${pagas.length} de ${cs.length} pagas`;
    const hoje = new Date(); hoje.setHours(0, 0, 0, 0);
    const venc = (c) => { const d = diaVenc(k, c.dia); if (!d) return "—"; const late = !c.pago && d < hoje; return `<span class="${late ? "neg" : ""}">${fmtD(d)}${late ? " (vencida)" : ""}</span>`; };
    $(pre + "Body").innerHTML = `<div class="kpis">
        <div class="kpi"><div class="l">Já pago</div><div class="v pos">${R(tP)}</div><div class="s">${pagas.length} conta(s)</div></div>
        <div class="kpi"><div class="l">Falta pagar</div><div class="v ${tF > 0.005 ? "neg" : ""}">${R(tF)}</div><div class="s">${pend.length} conta(s)</div></div></div>
      <div class="meter"><div class="t"><span>Pago</span><span>${P(tot ? tP / tot : 0)}</span></div><div class="bar"><span style="width:${tot ? (tP / tot) * 100 : 0}%;background:${css("--good")}"></span></div></div>
      <div class="tblwrap"><table class="tbl"><tr><th>Conta</th><th>Valor</th><th>Vence</th><th>Situação</th></tr>` +
      pend.concat(pagas).map((c) => `<tr class="clic" data-conta="${esc(c.nome)}" data-tipo="${c.tipo}" title="Ver evolução mês a mês"><td>${esc(c.nome)} <span class="muted small">📈</span></td><td>${R(c.pago && c.pago.valor ? c.pago.valor : c.valor)}</td><td>${venc(c)}</td><td>${c.pago ? `<span class="tag good">Pago${c.pago.data ? " " + c.pago.data.split("-").reverse().slice(0, 2).join("/") : ""}</span>` : `<span class="tag bad">Pendente</span>`}</td></tr>`).join("") +
      `</table></div><p class="muted small">Clique numa conta para ver a evolução mês a mês. Marque o que pagou na aba Pagamentos da planilha. Valores pendentes são os previstos.</p>`;
    $(pre + "Body").querySelectorAll("tr.clic").forEach((tr) => tr.addEventListener("click", () => abrirConta(tr.dataset.conta, tr.dataset.tipo, k)));
  }
  // ---------- evolução de uma conta ----------
  const baseConta = (n) => txt(n).replace(/\s*\((jan|fev|mar|abr|mai|jun|jul|ago|set|out|nov|dez)\/\d{2}\)$/i, "").trim();
  function serieConta(nome, tipo) {
    const b = baseConta(nome), nb = norm(b), energia = /luz|energia/i.test(b);
    const hist = (DATA.consumo || []).filter((x) => norm(baseConta(x.conta)) === nb);
    const meses = [...new Set(hist.map((x) => x.mes).concat(DATA.meses.filter((m) => m >= DATA.planIni)))].sort();
    return meses.map((m) => {
      let v = 0, real = false, est = false;
      if (m >= DATA.planIni) {
        if (tipo === "Fatura") { const card = nome.replace(/^Fatura /, ""); v = (DATA.faturas[card] || [])[idx(m)] || 0; const f = DATA.fats.find((x) => x.card === card && x.mes === m); real = !!(f && f.status === "Fechada"); est = !real; }
        else if (tipo === "Parcela") { const d = DATA.dividas.find((x) => x.desc === nome); if (d && ativo(d, m)) { v = d.valor; real = true; } }
        else { const fs = DATA.fixos.filter((f) => norm(baseConta(f.desc)) === nb && ativo(f, m));
          v = fs.reduce((s, f) => s + f.valor, 0); real = fs.some((f) => f.desc !== b); est = !real && /aluguel|condom|luz|energia/i.test(b); }
        const pg = DATA.pagos.find((p) => p.mes === m && norm(baseConta(p.nome)) === nb);
        if (pg && pg.valor) { v = pg.valor; real = true; est = false; }
      }
      const h = hist.find((x) => x.mes === m);
      if (h && h.valor) { v = h.valor; real = true; est = false; }
      return { m, v, real, est, kwh: energia && h && h.kwh ? h.kwh : null, ref: h ? h.ref : null };
    }).filter((x) => x.v > 0.004);
  }
  function abrirConta(nome, tipo, k) {
    fecharZoom();
    const s = serieConta(nome, tipo), energia = /luz|energia/i.test(nome), cor = pal();
    const ate = s.filter((x) => x.m <= addM(k, 3)); if (!ate.length) return;
    const pct = (a, b) => (a && b ? (b - a) / a : null);
    ate.forEach((x, n) => { const p = n ? ate[n - 1] : null; x.dv = p ? pct(p.v, x.v) : null; x.rk = x.kwh ? x.v / x.kwh : null;
      const pk = p && p.kwh ? p : null; x.dk = pk && x.kwh ? pct(pk.kwh, x.kwh) : null; x.drk = pk && x.rk ? pct(pk.v / pk.kwh, x.rk) : null; });
    const fP = (v) => (v == null ? "—" : (v > 0 ? "+" : "") + (v * 100).toFixed(1).replace(".", ",") + "%");
    const cl = (v) => (v == null ? "" : v > 0.0005 ? "neg" : v < -0.0005 ? "pos" : "");
    const reais = ate.filter((x) => x.real), ult = reais[reais.length - 1], pri = reais[0];
    const ov = document.createElement("div"); ov.className = "zoomov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
    const kpis = [["Último valor real", ult ? R(ult.v) : "—", ult ? lab(ult.m) : ""],
      ["Variação no último mês", fP(ult ? ult.dv : null), "em relação ao mês anterior"],
      ["Variação no período", pri && ult && pri !== ult ? fP(pct(pri.v, ult.v)) : "—", pri && ult ? lab(pri.m) + " → " + lab(ult.m) : ""],
      ["Média dos meses reais", reais.length ? R(reais.reduce((t, x) => t + x.v, 0) / reais.length) : "—", reais.length + " mês(es)"]];
    if (energia) { const ek = ate.filter((x) => x.kwh), ue = ek[ek.length - 1];
      kpis.push(["Consumo (último)", ue ? ue.kwh + " kWh" : "—", ue ? "variação " + fP(ue.dk) : "mande o kWh das contas"], ["Preço por kWh", ue ? "R$ " + ue.rk.toFixed(3).replace(".", ",") : "—", ue ? "variação " + fP(ue.drk) : ""]); }
    ov.innerHTML = `<div class="zh"><h3>${esc(baseConta(nome))} — evolução mês a mês</h3><button class="btn" type="button" id="zClose">Fechar ✕</button></div>
      <div class="zb"><h2 class="exp-tit">${esc(baseConta(nome))} — evolução mês a mês</h2><div class="kpis">${kpis.map(([l, v, sb]) => `<div class="kpi"><div class="l">${l}</div><div class="v ${/^[+-]/.test(v) ? cl(parseFloat(v.replace(",", ".").replace("%", ""))) : ""}">${v}</div><div class="s">${sb}</div></div>`).join("")}</div>
      ${energia ? `<h3 class="h3">Valor da conta (R$) x consumo (kWh)</h3>` : ""}<div class="chart"><canvas id="ctaC1"></canvas></div>${energia ? `<div id="ctaExp" class="cta-exp"></div><div id="ctaCemig" class="cemig"></div>` : ""}
      <div class="tblwrap"><table class="tbl"><tr><th>Mês</th><th>Valor</th><th>Variação</th><th>Tipo</th>${energia ? "<th>Referente a</th><th>kWh</th><th>Variação kWh</th><th>R$/kWh</th><th>Variação R$/kWh</th>" : ""}</tr>` +
      ate.map((x) => `<tr class="${x.m === k ? "sel" : ""}${energia && x.kwh ? " clic" : ""}" data-m="${x.m}"><td>${lab(x.m)}</td><td>${R(x.v)}${x.est ? "*" : ""}</td><td class="${cl(x.dv)}">${fP(x.dv)}</td><td>${x.real ? "real" : "previsto"}</td>${energia ? `<td>${x.ref ? lab(x.ref) : "—"}</td><td>${x.kwh || "—"}</td><td class="${cl(x.dk)}">${fP(x.dk)}</td><td>${x.rk ? "R$ " + x.rk.toFixed(3).replace(".", ",") : "—"}</td><td class="${cl(x.drk)}">${fP(x.drk)}</td>` : ""}</tr>`).join("") +
      `</table></div><p class="muted small">Variação = comparação com o mês anterior (vermelho = subiu, verde = caiu). * = valor previsto/estimado.${energia ? " Consumo e contas antigas ficam na aba Histórico de contas da planilha." : ""}</p></div>`;
    document.body.appendChild(ov); document.body.style.overflow = "hidden"; $("zClose").addEventListener("click", fecharZoom);
    const zb = ov.querySelector(".zb"); zb.classList.add("exp-alvo"); setTimeout(() => poeExport(zb), 0);
    const L = ate.map((x) => lab(x.m)), o1 = baseOpts(false, (v) => R0(v));
    o1.scales.y1 = { position: "right", grid: { drawOnChartArea: false }, ticks: { color: css("--ink2"), callback: (v) => v + "%" } };
    if (!energia) draw("ctaC1", "bar", L, [ds("Valor", ate.map((x) => x.v), ate.map((x) => (x.real ? cor[0] : cor[0] + "66"))),
      { type: "line", label: "Variação %", data: ate.map((x) => (x.dv == null ? null : Math.round(x.dv * 1000) / 10)), yAxisID: "y1", borderColor: cor[1], backgroundColor: cor[1], borderWidth: 2, pointRadius: 4, tension: 0.2, spanGaps: true,
        datalabels: { align: "top", color: cor[1], formatter: (v) => (v == null ? "" : (v > 0 ? "+" : "") + String(v).replace(".", ",") + "%") } }], o1);
    if (energia) { const ek = ate.filter((x) => x.kwh), o2 = baseOpts(false, (v) => R0(v));
      o2.scales.y.title = { display: true, text: "Valor (R$)", color: cor[0] };
      const mx = Math.max(...ek.map((x) => x.kwh), 1);
      o2.scales.y1 = { position: "right", min: 0, suggestedMax: mx * 1.12, grid: { drawOnChartArea: false }, title: { display: true, text: "Consumo (kWh)", color: cor[2] }, ticks: { color: css("--ink2"), callback: (v) => v + " kWh" } };
      o2.scales.y.suggestedMax = Math.max(...ate.map((x) => x.v)) * 1.7;
      o2.plugins.tooltip.callbacks.label = (c) => " " + c.dataset.label + ": " + (c.dataset.yAxisID === "y1" ? (c.raw == null ? "sem dado" : c.raw + " kWh") : R(c.raw));
      o2.plugins.tooltip.callbacks.afterBody = (it) => { const x = ate[it[0].dataIndex]; if (!x) return "";
        const l = []; if (x.rk) l.push("R$ por kWh: R$ " + x.rk.toFixed(3).replace(".", ",")); if (x.dv != null) l.push("Variação do valor: " + fP(x.dv)); if (x.dk != null) l.push("Variação do consumo: " + fP(x.dk));
        if (x.ref) l.push("Consumo referente a " + lab(x.ref)); if (!x.real) l.push("Valor previsto"); return l; };
      draw("ctaC1", "bar", L, [ds("Valor da conta (R$)", ate.map((x) => x.v), ate.map((x) => (x.real ? cor[0] : cor[0] + "66")), { datalabels: { anchor: "end", align: "end", formatter: (v) => R0(v) } }),
        { type: "line", label: "Consumo (kWh)", data: ate.map((x) => x.kwh || null), yAxisID: "y1", borderColor: cor[2], backgroundColor: cor[2], borderWidth: 3, pointRadius: 5, tension: 0.2, spanGaps: false,
          datalabels: { align: "top", offset: 6, color: "#fff", backgroundColor: cor[2], borderRadius: 4, padding: { top: 2, bottom: 2, left: 5, right: 5 }, font: { weight: "600", size: 11 }, formatter: (v) => (v == null ? "" : v + " kWh") } }], o2);
      const explica = (m) => {
        const i = ate.findIndex((x) => x.m === m), x = ate[i], box = $("ctaExp"); if (!box || !x) return;
        ov.querySelectorAll("tr[data-m]").forEach((tr) => tr.classList.toggle("sel", tr.dataset.m === m));
        const p = ate.slice(0, i).reverse().find((z) => z.kwh);
        if (!x.kwh) { box.innerHTML = `<b>${lab(m)}</b> — valor ${x.real ? "real" : "previsto"} de ${R(x.v)}. Sem kWh desse mês ainda, então não dá para separar consumo e preço. Mande a conta quando chegar.`; return; }
        if (!p) { box.innerHTML = `<b>${lab(m)}</b> — ${x.kwh} kWh a R$ ${x.rk.toFixed(3).replace(".", ",")} por kWh = ${R(x.v)}. É o primeiro mês com kWh, sem mês anterior para comparar.`; return; }
        const pk = p.v / p.kwh, dV = x.v - p.v, efC = (x.kwh - p.kwh) * pk, efP = (x.rk - pk) * x.kwh;
        const up = (v) => v > 0.0005, dn = (v) => v < -0.0005, dk = x.kwh - p.kwh, dr = x.rk - pk;
        const sg = (v, f) => (v > 0 ? "+" : v < 0 ? "−" : "") + f(Math.abs(v));
        let tit, ic;
        if (up(dk) && up(dr)) { tit = "Gastou mais energia e o kWh ficou mais caro"; ic = "🔺🔺"; }
        else if (dn(dk) && dn(dr)) { tit = "Gastou menos energia e o kWh ficou mais barato"; ic = "🔻🔻"; }
        else if (dn(dk) && up(dr)) { tit = dV > 0 ? "Pagou mais caro, mesmo reduzindo o consumo" : "Reduziu o consumo e isso compensou o kWh mais caro"; ic = "🔻🔺"; }
        else if (up(dk) && dn(dr)) { tit = dV > 0 ? "Pagou mais porque aumentou o consumo (o kWh até ficou mais barato)" : "Pagou mais barato, mesmo aumentando o consumo"; ic = "🔺🔻"; }
        else if (up(dk) || dn(dk)) { tit = up(dk) ? "Consumiu mais, com o mesmo preço do kWh" : "Consumiu menos, com o mesmo preço do kWh"; ic = up(dk) ? "🔺" : "🔻"; }
        else { tit = up(dr) ? "Mesmo consumo, mas o kWh ficou mais caro" : dn(dr) ? "Mesmo consumo, e o kWh ficou mais barato" : "Nada mudou em relação ao mês anterior"; ic = "="; }
        const tot = Math.abs(efC) + Math.abs(efP), shC = tot ? Math.abs(efC) / tot : 0;
        const motivo = !tot ? "" : shC >= 0.65 ? "O que mais pesou foi o <b>consumo</b>" : shC <= 0.35 ? "O que mais pesou foi o <b>preço do kWh</b>" : "<b>Consumo e preço</b> pesaram parecido";
        const c = (v) => (v > 0.005 ? "neg" : v < -0.005 ? "pos" : "");
        box.innerHTML = `<div class="cta-exp-t">${ic} <b>${lab(m)}: ${tit}</b></div>
          <div class="cta-exp-s">Comparado com ${lab(p.m)}: a conta foi de ${R(p.v)} para ${R(x.v)} (<span class="${c(dV)}">${sg(dV, R)}, ${fP(pct(p.v, x.v))}</span>).</div>
          <table class="tbl cta-exp-tb"><tr><th>Motivo</th><th>Mudança</th><th>Impacto na conta</th><th>Peso</th></tr>
          <tr><td>⚡ Consumo</td><td>${p.kwh} → ${x.kwh} kWh (${fP(pct(p.kwh, x.kwh))})</td><td class="${c(efC)}">${sg(efC, R)}</td><td>${tot ? Math.round(shC * 100) + "%" : "—"}</td></tr>
          <tr><td>💲 Preço do kWh</td><td>R$ ${pk.toFixed(3).replace(".", ",")} → R$ ${x.rk.toFixed(3).replace(".", ",")} (${fP(pct(pk, x.rk))})</td><td class="${c(efP)}">${sg(efP, R)}</td><td>${tot ? 100 - Math.round(shC * 100) + "%" : "—"}</td></tr>
          <tr><td><b>Total</b></td><td></td><td class="${c(dV)}"><b>${sg(dV, R)}</b></td><td></td></tr></table>
          <div class="cta-exp-s">${motivo}${motivo ? ". " : ""}Se o kWh tivesse ficado no preço de ${lab(p.m)}, essa conta seria ${R(x.kwh * pk)}.</div>
          <div class="muted small">Clique em outra coluna do gráfico ou linha da tabela para ver outro mês.</div>`;
      };
      const mostra = (m) => { explica(m); const cb = $("ctaCemig"); if (cb) cb.innerHTML = htmlCemig(m); setTimeout(rotulaTabelas, 0); };
      const ligaClique = (id, lista) => { const ch = charts[id]; if (!ch) return; $(id).onclick = null; $(id).style.cursor = "pointer";
        ch.options.onClick = (e, els) => { const el = els && els[0]; if (el) mostra(lista[el.index].m); }; ch.update("none"); };
      ligaClique("ctaC1", ate);
      ov.querySelectorAll("tr.clic[data-m]").forEach((tr) => tr.addEventListener("click", () => mostra(tr.dataset.m)));
      const ini = ek.find((x) => x.m === k) || ek[ek.length - 1]; if (ini) mostra(ini.m); }
  }

  // ================= VISÃO GERAL (consolidado) =================
  const NOME_ABA = { geral: "Visão geral", dividas: "Contas e dívidas", rend: "Rendimentos", metas: "Metas e cortes", sim: "Simulação" };
  const fPct = (v) => (v == null || !isFinite(v) ? "—" : (v > 0 ? "+" : "") + (v * 100).toFixed(1).replace(".", ",") + "%");
  const DIAS = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
  const hoje12 = () => { const h = new Date(); h.setHours(12, 0, 0, 0); return h; };
  const diasAte = (d) => Math.round((new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12) - hoje12()) / 864e5);
  const quando = (d) => { const n = diasAte(d); return n === 0 ? "hoje" : n === 1 ? "amanhã" : n > 0 ? "em " + n + " dias" : "há " + -n + " dia" + (n < -1 ? "s" : ""); };
  function vaiPara(tab, id) {
    const b = document.querySelector(`.tabs button[data-tab="${tab}"]`); if (b) b.click();
    setTimeout(() => { const el = $(id); if (!el) return; el.scrollIntoView({ behavior: "smooth", block: "start" }); el.classList.remove("piscar"); void el.offsetWidth; el.classList.add("piscar"); }, 120);
  }
  function donutOpts(fmtFn) {
    const ink = css("--ink2");
    return { responsive: true, maintainAspectRatio: false, cutout: "68%", animation: { duration: 300 },
      plugins: { legend: { position: "right", labels: { color: ink, boxWidth: 10, boxHeight: 10, usePointStyle: true, pointStyle: "rectRounded", padding: 10, font: { size: 12 } } },
        tooltip: { backgroundColor: css("--tip") || "rgba(8,20,40,.92)", padding: 10, cornerRadius: 10, callbacks: { label: (c) => " " + c.label + ": " + fmtFn(c.raw) } },
        datalabels: { color: "#fff", font: { size: 11, weight: "700" }, formatter: (v, c) => { const t = c.dataset.data.reduce((s, x) => s + (x || 0), 0); return t && v / t > 0.08 ? Math.round((v / t) * 100) + "%" : ""; } } } };
  }
  // dinheiro para débito/Pix considerando a data de hoje (salário, 13º e extras só depois que caem)
  function dinheiroAgora() {
    const cal = mesAtual(), dm = debitoMes(cal), prev = debitoMes(addM(cal, -1));
    const r = dm.salCaiu ? { agora: dm.livre, dm, futuro: null, cal } : { agora: prev.livre, dm, futuro: { valor: dm.livre, data: dm.salData, mes: cal }, cal };
    // o que posso gastar é só o dinheiro que está na conta (Plano de Gastos B16), sem o CDB (reserva de emergência),
    // menos as contas do mês que ainda faltam pagar (valor real ou previsto)
    if (DATA.saldo > 0.005) {
      const itens = contasMes(cal).filter((c) => !c.pago), pend = itens.reduce((s, c) => s + c.valor, 0);
      r.conta = { saldo: DATA.saldo, pend, itens, livre: Math.max(0, DATA.saldo - pend), plano: r.agora };
      r.agora = Math.min(r.agora, r.conta.livre);
    }
    return r;
  }
  function catsMes(k) {
    const ci = DATA.cat.meses.indexOf(k);
    const realCats = ci >= 0 ? Object.entries(DATA.cat.itens).map(([n, o]) => [n, num(o.vals[ci])]).filter(([, v]) => v > 0.005) : [];
    if (realCats.length) return { cats: realCats.sort((a, b) => b[1] - a[1]), real: true };
    const m = {}; DATA.fixos.filter((f) => ativo(f, k)).forEach((f) => (m[f.cat] = (m[f.cat] || 0) + f.custo)); DATA.dividas.filter((d) => ativo(d, k)).forEach((d) => (m[d.cat] = (m[d.cat] || 0) + d.custo));
    DATA.lanc.filter((l) => l.ini && l.ini <= k && l.fim >= k).forEach((l) => (m[l.cat || "Outros"] = (m[l.cat || "Outros"] || 0) + l.custo));
    return { cats: Object.entries(m).filter(([, v]) => v > 0.005).sort((a, b) => b[1] - a[1]), real: false };
  }
  function eventosMes(k) { // o que entra no mês k e quando
    const ev = [], r = rendimentos(k), sd = dataSalario(k), h = hoje12();
    if (r.sal) ev.push({ n: "Salário", v: r.sal, d: sd });
    r.itens.filter((i) => i.c !== "vr").forEach((i) => ev.push({ n: i.desc, v: i.valor, d: sd }));
    const d13 = (DATA.d13 || {})[k.slice(5)]; if (r.d13) ev.push({ n: "13º salário", v: r.d13, d: d13 || sd });
    if (r.fer) ev.push({ n: "Férias", v: r.fer, d: sd });
    const [y, m] = k.split("-").map(Number), vrd = ultimoDiaUtil(y, m - 1), vrv = vrMes(addM(k, 1)); if (vrv) ev.push({ n: "VR / Flash de " + lab(addM(k, 1)), v: vrv, d: vrd, vr: true });
    ev.forEach((e) => (e.caiu = e.d ? h >= new Date(e.d.getFullYear(), e.d.getMonth(), e.d.getDate(), 12) : false));
    return ev.sort((a, b) => (a.d || 0) - (b.d || 0));
  }
  function renderResumo(k) {
    const cor = pal(), h = hoje12(), r = rendimentos(k), da = dinheiroAgora(), cs = contasMes(k);
    const pagas = cs.filter((c) => c.pago), pend = cs.filter((c) => !c.pago);
    const tPago = pagas.reduce((s, c) => s + (c.pago.valor || c.valor), 0), tPend = pend.reduce((s, c) => s + c.valor, 0);
    const rc = receber(k); statusReceber(k, rc);
    const recTot = rc.reduce((s, x) => s + x.valor, 0), recPago = rc.reduce((s, x) => s + Math.min(x.valor, x.pago || 0), 0), recPend = Math.max(0, recTot - recPago);
    const outros = r.dinheiro - r.sal, entradas = r.dinheiro + recTot, saidas = tPago + tPend, sobra = entradas - saidas;
    const ev = eventosMes(k), vencidas = pend.filter((c) => { const d = diaVenc(k, c.dia); return d && d < h; });
    const ok = sobra >= -0.005, pc = DATA.plano[k] || {}, cartao = Math.max(0, (pc.teto || 0) - gastoInter(k).v);
    const sdl = simDelta(addM(da.futuro ? da.futuro.mes : da.cal, -1));
    const barra = (itens, tot) => `<div class="sbar">${itens.filter(([, v]) => v > 0.005).map(([n, v, c]) => `<span style="flex:${v};background:${c}" title="${esc(n)}: ${R(v)}"></span>`).join("")}</div>`;
    const maxT = Math.max(entradas, saidas, 1);
    $("sit").innerHTML = `
      <div class="sit-h"><div><p class="eyebrow">Minha situação financeira · ${labL(k)}</p><h2 class="sit-tit">${(() => { const fx = fluxoDiario(k), neg = fx.dias.filter((x) => x.saldo != null && x.saldo < -0.005);
          if (!ok) return "⚠️ As contas passam das entradas";
          if (neg.length) { const pior = neg.reduce((a, x) => (x.saldo < a.saldo ? x : a)); return `⚠️ Faltam ${R0(-pior.saldo)} para pagar tudo em dia (${pad(neg[0].d)}/${k.slice(5)} a ${pad(neg[neg.length - 1].d)}/${k.slice(5)})`; }
          return "✅ Dá para pagar todas as contas em dia"; })()}</h2>
        <p class="muted small">Hoje ${fmtD(h)} (${DIAS[h.getDay()]}) · salário de ${MESESL[+k.slice(5) - 1]} ${dataSalario(k) ? (h >= dataSalario(k) ? "caiu em " : "cai em ") + fmtD(dataSalario(k)) + " (" + DIAS[dataSalario(k).getDay()] + ")" : ""}</p>${checksMes(k, sobra)}</div>
        <div class="sit-res ${ok ? "pos" : "neg"}"><div class="l">${ok ? "Sobra no mês" : "Falta no mês"}</div><div class="v">${R(Math.abs(sobra))}</div><div class="s">entradas − contas do mês</div>${marcosSaldo(k)}</div></div>
      <div class="sit-g">
        <div class="blk blk-in"><div class="blk-t">💰 Entradas do mês</div><div class="blk-v">${R(entradas)}</div>
          <ul class="blk-l"><li><span>Salário líquido</span><b>${R(r.sal)}</b></li><li><span>Outros rendimentos</span><b>${R(outros)}</b></li><li><span>A receber de terceiros</span><b>${R(recTot)}</b></li></ul>
          ${barra([["Salário", r.sal, cor[0]], ["Outros", outros, cor[2]], ["Terceiros", recTot, cor[4]]])}
          <div class="blk-n">${r.vr ? "+ " + R(r.vr) + " no Flash (só alimentação)" : ""}</div></div>
        <div class="blk blk-out"><div class="blk-t">🧾 Contas do mês</div><div class="blk-v">${R(saidas)}</div>
          <ul class="blk-l"><li><span>Já pago</span><b class="pos">${R(tPago)}</b></li><li><span>Falta pagar</span><b class="${tPend > 0.005 ? "neg" : ""}">${R(tPend)}</b></li><li><span>Contas</span><b>${pagas.length} de ${cs.length} pagas</b></li></ul>
          <div class="pbar"><span style="width:${saidas ? (tPago / saidas) * 100 : 0}%"></span></div>
          <div class="blk-n ${vencidas.length ? "neg" : ""}">${vencidas.length ? "⚠️ " + vencidas.length + " vencida(s): " + vencidas.map((c) => esc(c.nome)).join(", ") : "nenhuma conta vencida"}</div></div>
        <div class="blk blk-rec"><div class="blk-t">🤝 Terceiros</div><div class="blk-v">${R(recPend)}</div>
          <ul class="blk-l"><li><span>Devem no mês</span><b>${R(recTot)}</b></li><li><span>Já recebi</span><b class="pos">${R(recPago)}</b></li><li><span>Falta receber</span><b class="${recPend > 0.005 ? "neg" : ""}">${R(recPend)}</b></li></ul>
          <div class="pbar alt"><span style="width:${recTot ? (recPago / recTot) * 100 : 0}%"></span></div><div class="blk-n">${Object.entries(devidoPessoa(k)).map(([p, v]) => esc(p) + " " + R0(v)).join(" · ") || "ninguém deve neste mês"}</div></div>
        <div class="blk blk-gasto"><div class="blk-t">💵 Posso gastar agora</div><div class="blk-v ${da.agora <= 0.005 ? "neg" : ""}">${R(da.agora)}</div>
          <ul class="blk-l">${da.conta ? `<li><span>🏦 Na conta</span><b>${R(da.conta.saldo)}</b></li><li><span>− Falta pagar${da.conta.itens.length ? " (" + da.conta.itens.map((c) => esc(c.nome)).join(", ") + ")" : ""}</span><b class="${da.conta.pend > 0.005 ? "neg" : ""}">${R(da.conta.pend)}</b></li>${da.conta.plano - da.conta.livre > 0.005 ? `<li><span>📊 Limite do plano</span><b>${R(da.conta.plano)}</b></li>` : ""}` : ""}${da.futuro ? `<li><span>⏳ ${fmtD(da.futuro.data)} (${DIAS[da.futuro.data.getDay()]})</span><b>+ ${R(da.futuro.valor)}</b></li>` : ""}
            ${da.dm.ex.valor && !da.dm.ex.caiu ? `<li><span>⏳ 13º livre em ${da.dm.ex.data ? fmtD(da.dm.ex.data) : "breve"}</span><b>+ ${R(da.dm.ex.valor)}</b></li>` : ""}
            <li><span>💳 Cartão (só parcelado)</span><b>${R(cartao)}</b></li></ul>
          ${da.conta && DATA.reservaInformada > 0.005 ? `<div class="blk-n">🛟 Reserva de emergência: ${R(DATA.reservaInformada)} (CDB), fora desta conta. Só para emergência.</div>` : ""}<div class="blk-n">no débito/Pix · conta só o que já caiu na conta${sdl && Math.abs(sdl.d) > 0.005 ? ` · <span class="sim-t">🧪 com simulação: ${R(Math.max(0, (da.futuro ? da.futuro.valor : da.agora) + sdl.d))}</span>` : ""}</div></div>
      </div>
      <div class="sit-cmp"><div class="cmp-r"><span>Entradas</span><div class="cmp-b"><span style="width:${(entradas / maxT) * 100}%"></span></div><b>${R0(entradas)}</b></div>
        <div class="cmp-r out"><span>Contas</span><div class="cmp-b"><span style="width:${(saidas / maxT) * 100}%"></span></div><b>${R0(saidas)}</b></div></div>
      <div class="tl">${ev.map((e) => `<div class="tl-i ${e.caiu ? "ok" : ""}"><i></i><div><b>${e.d ? pad(e.d.getDate()) + "/" + pad(e.d.getMonth() + 1) : "—"}</b> ${esc(e.n)}</div><span>${e.caiu ? "✓ caiu" : quando(e.d)} · ${R0(e.v)}</span></div>`).join("")}</div>`;

    renderFluxoDia(k);
    // ---- blocos de números (clicáveis: levam para a aba com os detalhes)
    const vr = (() => { const cal = mesAtual(); const m = DATA.meses.includes(cal) ? cal : k; return vrSaldoIni(m) + vrMes(m) - vrGasto(m); })();
    const Fm = addM(k, 1), j = idx(Fm), sal = j >= 0 ? DATA.proj.renda[j] : 0, parc = parcelasMeu(Fm), pp = sal ? parc / sal : 0;
    const as = DATA.mc.ass, assTot = as.filter((a) => !assOff(a.nome)).reduce((s, a) => s + a.valor, 0);
    const cem = checaCemig(), alE = cem ? cem.al.filter((x) => x.nv !== "info") : [];
    const tiles = [
      ["💵", "Posso gastar agora", R(da.agora), da.futuro ? "+ " + R0(da.futuro.valor) + " em " + fmtD(da.futuro.data) : da.conta ? "livre na conta, sem a reserva" : "débito/Pix", "metas|heroCard", da.agora <= 0.005 ? "neg" : "pos"],
      ["🧾", "Falta pagar em " + lab(k), R(tPend), pend.length + " conta(s) pendente(s)", "dividas|cardContas", tPend > 0.005 ? "neg" : "pos"],
      ["💰", "Rendimentos de " + lab(k), R(r.dinheiro), "salário + extras (líquido)", "rend|rendCard", ""],
      ["🤝", "Falta receber", R(recPend), "de " + R0(recTot) + " devidos no mês", "rend|cardReceber", recPend > 0.005 ? "warn" : "pos"],
      ["🍽️", "Flash (VR)", R(vr), "saldo para alimentação", "rend|vrCard", vr < 0 ? "neg" : ""],
      ["🛡️", "Reserva no fim de " + lab(k), R0(pc.reservaAc), "aporte do mês " + R0(pc.reserva), "metas|cardReserva", ""],
      ["📉", "Parcelas ÷ salário", P(pp), "ideal até 30%", "metas|saude", pp > 0.45 ? "neg" : pp > 0.3 ? "warn" : "pos"],
      ["🔁", "Assinaturas", R(assTot) + "/mês", as.length + " ativas" + (ASS_OFF.size ? " · " + ASS_OFF.size + " na prévia de corte" : ""), "metas|cardAss", ""],
    ];
    $("ovKpis").innerHTML = tiles.map(([ic, l, v, s, go, c]) => `<button type="button" class="tile ${c}" data-go="${go}"><span class="ic">${ic}</span><span class="l">${l}</span><span class="v">${v}</span><span class="s">${s}</span><span class="go">ver ›</span></button>`).join("");
    $("ovKpis").querySelectorAll("[data-go]").forEach((b) => b.addEventListener("click", () => { const [t, id] = b.dataset.go.split("|"); vaiPara(t, id); }));

    // ---- gráficos (clique = tela cheia com link para a fonte)
    const rg = range(), L = rg.map(lab), ix = rg.map(idx), pr = DATA.proj;
    const sai = ix.map((q) => Math.abs(pr.fixos[q] || 0) + Math.abs(pr.parcelas[q] || 0) + Math.abs(pr.variaveis[q] || 0));
    draw("ovFluxo", "bar", L, [ds("Entradas", ix.map((q) => pr.entradas[q] || 0), cor[0]), ds("Saídas", sai, cor[5]),
      ds("Sobra", ix.map((q) => pr.sobra[q] || 0), cor[1], { type: "line", fill: false, borderWidth: 2.5, tension: 0.35, datalabels: { display: false } })], baseOpts(false, (v) => R0(v)));
    charts.ovFluxo.options.plugins.datalabels.display = false; charts.ovFluxo.update("none");
    $("ovContasT").textContent = `${pagas.length} de ${cs.length} pagas`;
    draw("ovContas", "doughnut", ["Pago", "Falta pagar"], [{ data: [Math.round(tPago * 100) / 100, Math.round(tPend * 100) / 100], backgroundColor: [cor[1], cor[5]], borderWidth: 0, hoverOffset: 6 }], donutOpts(R));
    const rs = rg.map(rendimentos), cats = CAT_REND.filter(([c]) => c !== "vr" && rs.some((x) => x[c])), cc = [cor[0], cor[2], cor[3], cor[1], cor[4], css("--muted")];
    const or = draw("ovRend", "bar", L, cats.map(([c, n], z) => ds(n, rs.map((x) => Math.round(x[c] * 100) / 100), cc[z % cc.length])), baseOpts(true, (v) => R0(v)));
    or.options.plugins.datalabels.display = false; or.update("none");
    const ct = catsMes(k), top = ct.cats.slice(0, 6), resto = ct.cats.slice(6).reduce((s, c) => s + c[1], 0); if (resto > 0.005) top.push(["Outros", resto]);
    $("ovCatT").textContent = (ct.real ? "realizado" : "previsto") + " em " + lab(k);
    draw("ovCat", "doughnut", top.map((c) => c[0]), [{ data: top.map((c) => Math.round(c[1] * 100) / 100), backgroundColor: top.map((_, z) => [cor[0], cor[1], cor[2], cor[3], cor[4], cor[5], css("--muted")][z % 7]), borderWidth: 0, hoverOffset: 6 }], donutOpts(R));
    const ms = DATA.meses.filter((m) => DATA.plano[m] && m >= DATA.planIni).slice(0, 18);
    const orr = draw("ovRes", "line", ms.map(lab), [ds("Reserva acumulada", ms.map((m) => (DATA.plano[m] || {}).reservaAc || 0), cor[1], { fill: true, backgroundColor: (ctx) => { const ar = ctx.chart.chartArea; if (!ar) return cor[1] + "33"; const g = ctx.chart.ctx.createLinearGradient(0, ar.top, 0, ar.bottom); g.addColorStop(0, cor[1] + "66"); g.addColorStop(1, cor[1] + "00"); return g; }, borderWidth: 3, tension: 0.35, pointRadius: 0, datalabels: { display: (c) => c.dataIndex === ms.length - 1, align: "top" } })], baseOpts(false, (v) => R0(v)));
    orr.options.plugins.legend.display = false; orr.update("none");
    const pend2 = pendencias(k).mes, pes = Object.keys(pend2).filter((p) => pend2[p].d > 0.005);
    $("ovRecT").textContent = lab(k);
    const orc = draw("ovRec", "bar", pes, [ds("Recebido", pes.map((p) => Math.round(pend2[p].r * 100) / 100), cor[1]), ds("Falta", pes.map((p) => Math.round(pend2[p].pend * 100) / 100), cor[5])], baseOpts(true, (v) => R0(v), { indexAxis: "y" }));
    orc.update("none");
    const ml = DATA.meses.filter((m) => m >= addM(k, 1)).slice(0, 18), vl = ml.map((m) => Math.round(parcelasMeu(m) * 100) / 100);
    const ol = draw("ovLibera", "line", ml.map(lab), [ds("Minhas parcelas", vl, cor[3], { fill: true, backgroundColor: cor[3] + "26", stepped: true, borderWidth: 2.5, pointRadius: 0, datalabels: { display: (c) => c.dataIndex === 0 || vl[c.dataIndex - 1] - vl[c.dataIndex] > 150, align: "top" } })], baseOpts(false, (v) => R0(v)));
    ol.options.plugins.legend.display = false; ol.update("none");
    document.querySelectorAll("#tab-geral .ov-c").forEach((c) => { const cv = c.querySelector("canvas"); if (!cv) return; cv.onclick = () => abrirZoom(cv.id); });

    // ---- alertas
    const al = [];
    vencidas.forEach((c) => al.push(["bad", `Conta vencida: ${esc(c.nome)} (${R(c.valor)})`, "dividas|cardContas"]));
    alE.forEach((x) => al.push([x.nv, "⚡ " + x.t, "dividas|cardContas"]));
    const lim = DATA.mc.limites.filter((x) => x.lim > 0), gc = {}; lancMes(k).forEach((l) => (gc[l.cat || "Outros"] = (gc[l.cat || "Outros"] || 0) + l.meuTotal));
    const est = lim.filter((x) => (gc[x.cat] || 0) > x.lim + 0.005); if (est.length) al.push(["warn", `Limites estourados em ${lab(k)}: ${est.map((x) => esc(x.cat) + " (" + R0(gc[x.cat]) + " de " + R0(x.lim) + ")").join(", ")}`, "metas|cardLim"]);
    const aval = as.filter((a) => a.manter === "Avaliar"); if (aval.length) al.push(["info", `${aval.length} assinatura(s) marcadas "Avaliar" somam ${R(aval.reduce((s, a) => s + a.valor, 0))}/mês. Teste o corte na aba Metas e cortes.`, "metas|cardAss"]);
    if (recPend > 0.005) al.push(["info", `Falta receber ${R(recPend)} de terceiros em ${lab(k)}.`, "rend|cardReceber"]);
    $("ovAlertas").innerHTML = `<div class="row-between"><h2>Atenção</h2><span class="muted small">${al.length ? al.length + " ponto(s)" : "tudo em ordem"}</span></div>` +
      (al.length ? `<div class="alertas">${al.map(([nv, t, go]) => `<button type="button" class="al ${nv}" data-go="${go}">${t}<span>›</span></button>`).join("")}</div>` : `<p class="icon-ok">Nenhum alerta para ${lab(k)}.</p>`);
    $("ovAlertas").querySelectorAll("[data-go]").forEach((b) => b.addEventListener("click", () => { const [t, id] = b.dataset.go.split("|"); if (/⚡/.test(b.textContent)) { vaiPara(t, id); setTimeout(() => abrirConta("Luz / Energia", "Conta", k), 450); } else vaiPara(t, id); }));
  }

  // ================= FLUXO DIÁRIO DO MÊS =================
  function saldoInicio() { let v = null; try { v = localStorage.getItem("saldo-inicio"); } catch (_) {} return v != null && v !== "" && isFinite(+v) ? { v: +v, manual: true } : { v: 0, manual: false }; }
  // Fluxo do mês: tudo o que entra soma (acumula), cada conta sai no dia em que vence (paga em dia).
  // Começa do zero no dia 1 (ou do que sobrou do mês anterior, se informado). No fim do mês = sobra do mês.
  function fluxoDiario(k) {
    const [y, m] = k.split("-").map(Number), nd = new Date(y, m, 0).getDate(), h = hoje12(), cal = mesAtual();
    const atual = k === cal, passado = k < cal, dHoje = atual ? h.getDate() : passado ? nd + 1 : 0;
    const dia = (d) => Math.min(nd, Math.max(1, d.getDate())), ev = [];
    const naData = (d, fallback) => (d && d.getFullYear() === y && d.getMonth() + 1 === m ? dia(d) : fallback);
    const dt = (s) => (s ? new Date(s + "T12:00:00") : null);
    // entradas do mês (o Flash fica fora: é cartão de alimentação)
    eventosMes(k).filter((e) => !e.vr).forEach((e) => { const d = naData(e.d, 5); ev.push({ d, n: e.n, v: e.v, feito: atual ? d < dHoje : passado }); });
    // reembolsos de terceiros: o que já pagaram entra na data do pagamento; o que falta entra no vencimento
    const rc = receber(k); statusReceber(k, rc); const por = {};
    const soma = (d, n, v, feito) => { const key = n + "|" + d + "|" + feito; (por[key] = por[key] || { d, n, v: 0, feito }).v += v; };
    rc.forEach((r) => { const pago = Math.min(r.valor, r.pago || 0), falta = r.valor - pago, dv = naData(r.quando, 15);
      if (pago > 0.009) soma(naData(dt(r.dataPago), dv), "Reembolso " + r.pessoa, pago, true);
      if (falta > 0.1) soma(dv, "Reembolso " + r.pessoa, falta, false); });
    Object.values(por).forEach((o) => ev.push(o));
    // contas do mês: todas consideradas pagas no vencimento (as já pagas, na data do pagamento)
    contasMes(k).forEach((c) => { const dv = naData(diaVenc(k, c.dia), 10), v = c.pago && c.pago.valor ? c.pago.valor : c.valor;
      if (c.pago) ev.push({ d: naData(dt(c.pago.data), dv), n: c.nome, v: -v, feito: true });
      else ev.push({ d: dv, n: c.nome, v: -v, vencida: atual && dv < dHoje }); });
    const si = saldoInicio(), ini = si.v, dias = []; let s = ini;
    for (let d = 1; d <= nd; d++) {
      const doDia = ev.filter((e) => e.d === d), ent = doDia.filter((e) => e.v > 0).reduce((t, e) => t + e.v, 0), sai = doDia.filter((e) => e.v < 0).reduce((t, e) => t + e.v, 0);
      s += ent + sai; dias.push({ d, ent, sai, saldo: Math.round(s * 100) / 100, ev: doDia });
    }
    const d0 = Math.min(...ev.map((e) => e.d), nd), min = dias.filter((x) => x.d >= d0).reduce((a, x) => (a == null || x.saldo < a.saldo ? x : a), null);
    const hojeS = atual ? (dHoje > 1 ? dias[dHoje - 2].saldo : ini) : null;
    return { k, nd, dias, ev, ini, si, atual, dHoje, min, hojeS, fim: dias[nd - 1].saldo };
  }
  function checksMes(k, sobra) {
    const f = fluxoDiario(k), mm = k.slice(5), neg = f.dias.filter((x) => x.saldo < -0.005);
    const pior = neg.length ? neg.reduce((a, x) => (x.saldo < a.saldo ? x : a)) : null;
    // reserva: o plano separa a meta do salário deste mês (linha k-1); nunca mais do que sobra no fim do mês
    // reserva = sobra do mês − o que ainda está livre para gastar no débito (salário + 13º/férias livres do mês)
    const dm = debitoMes(k), livre = Math.max(0, (dm.pd.teto || 0) + (dm.ex.valor || 0) - dm.usado), metaRes = Math.max(0, f.fim - livre);
    const fim = (x) => x.fim === k && (x.custo || 0) > 0.005;
    const acabam = DATA.dividas.filter(fim).map((d) => ({ n: d.desc, v: d.custo })).concat(DATA.lanc.filter((l) => l.parcelado && fim(l)).map((l) => ({ n: l.desc, v: l.custo })));
    const libera = acabam.reduce((s, x) => s + x.v, 0);
    const it = [
      [pior ? "bad" : "ok", pior ? "⚠️" : "✅", pior ? `Faltam ${R0(-pior.saldo)} para pagar tudo em dia` : "Dá para pagar todas as contas em dia", pior ? `em ${pad(pior.d)}/${mm}` : `menor sobra ${R0(f.min ? f.min.saldo : 0)}${f.min ? " em " + pad(f.min.d) + "/" + mm : ""}`],
      [metaRes > 0.005 ? "ok" : "warn", "🛡️", metaRes > 0.005 ? `Dá para mandar ${R0(metaRes)} para a reserva` : "Sem sobra para a reserva neste mês", `sobra ${R0(f.fim)} − ${R0(livre)} livres para gastar · sem gastar nada, até ${R0(Math.max(0, f.fim))}`],
      [acabam.length ? "ok" : "info", "🏁", acabam.length ? `${acabam.length} parcelamento${acabam.length > 1 ? "s acabam" : " acaba"} este mês` : "Nenhum parcelamento acaba este mês", acabam.length ? `libera ${R0(libera)}/mês a partir de ${lab(addM(k, 1))}: ${acabam.slice(0, 3).map((x) => esc(x.n)).join(", ")}${acabam.length > 3 ? " +" + (acabam.length - 3) : ""}` : (() => { const prox = DATA.meses.find((m) => m > k && DATA.dividas.some((d) => d.fim === m && d.custo > 0)); return prox ? "próximo término em " + lab(prox) : ""; })()],
    ];
    return `<div class="checks">${it.map(([c, ic, t, s]) => `<div class="ck ${c}"><span class="ic">${ic}</span><div><b>${t}</b><small>${s}</small></div></div>`).join("")}</div>`;
  }
  function marcosSaldo(k) { // sobra acumulada no fim dos dias 5, 10, 15, 20 e fim do mês, com as contas pagas em dia
    const f = fluxoDiario(k), mm = k.slice(5), cortes = [5, 10, 15, 20, f.nd]; let de = 1;
    const cards = cortes.map((d) => { const ini = de; de = d + 1;
      const ev = f.ev.filter((e) => e.d >= ini && e.d <= d), ent = ev.filter((e) => e.v > 0).reduce((s, e) => s + e.v, 0), pag = ev.filter((e) => e.v < 0).reduce((s, e) => s - e.v, 0);
      const v = f.dias[d - 1].saldo, contas = ev.filter((e) => e.v < 0).map((e) => pad(e.d) + "/" + mm + " " + e.n + " " + R(-e.v));
      const tip = `Do dia ${ini} ao ${d}: entrou ${R(ent)} e paguei ${R(pag)}` + (contas.length ? "\n" + contas.join("\n") : "") + `\nSobra acumulada no fim do dia ${d}: ${R(v)}`;
      return `<div class="mc ${v < -0.005 ? "neg" : "pos"} ${f.atual && d < f.dHoje ? "ja" : ""}" title="${esc(tip)}"><span>dia ${d}</span><b>${R0(v)}</b><em>+${R0(ent)}<br>−${R0(pag)}</em>${v < -0.005 ? `<i>faltam ${R0(-v)}</i>` : ""}</div>`; });
    return `<div class="marcos"><div class="mc-t">Sobra acumulada no fim do dia (tudo o que entrou − contas pagas em dia)</div><div class="mc-g">${cards.join("")}</div><div class="mc-l">+ entrou no período · − contas que venceram no período</div></div>`;
  }
  function renderFluxoDia(k) {
    const box = $("fluxoDia"); if (!box) return; const f = fluxoDiario(k), cor = pal(), mm = k.slice(5);
    const L = f.dias.map((x) => pad(x.d) + "/" + mm);
    const prox = f.ev.filter((e) => !f.atual || e.d >= f.dHoje).sort((a, b) => a.d - b.d || b.v - a.v);
    let run = f.atual ? f.hojeS : f.ini; const lista = prox.map((e) => { run += e.v; return Object.assign({ depois: run }, e); });
    const neg = f.min && f.min.saldo < -0.005;
    $("fdInfo").innerHTML = `<div class="fd-k"><span>Início do mês</span><b>${R(f.ini)}</b><em>${f.si.manual ? "sobra do mês anterior" : "começa do zero"}</em></div>
      ${f.atual ? `<div class="fd-k"><span>Hoje (${pad(f.dHoje)}/${mm})</span><b class="${f.hojeS < -0.005 ? "neg" : ""}">${R(f.hojeS)}</b><em>entrou − pago até ontem</em></div>` : ""}
      <div class="fd-k"><span>Menor sobra</span><b class="${neg ? "neg" : ""}">${f.min ? R(f.min.saldo) : "—"}</b><em>${f.min ? "em " + pad(f.min.d) + "/" + mm : ""}</em></div>
      <div class="fd-k"><span>Sobra no fim do mês</span><b class="${f.fim < 0 ? "neg" : "pos"}">${R(f.fim)}</b><em>antes do débito/Pix do dia a dia</em></div>
      <label class="fd-in">Sobrou do mês anterior (opcional) <input id="fdSaldo" type="number" step="0.01" inputmode="decimal" value="${f.si.manual ? f.si.v : ""}" placeholder="R$ 0,00"></label>`;
    const ds1 = ds("Entradas", f.dias.map((x) => (x.ent ? Math.round(x.ent * 100) / 100 : null)), cor[1], { stack: "m", datalabels: { display: false } });
    const ds2 = ds("Saídas", f.dias.map((x) => (x.sai ? Math.round(x.sai * 100) / 100 : null)), cor[5], { stack: "m", datalabels: { display: false } });
    const ds3 = { type: "line", label: "Sobra acumulada (contas pagas em dia)", data: f.dias.map((x) => x.saldo), borderColor: cor[0], backgroundColor: cor[0], borderWidth: 3, stepped: true, pointRadius: f.dias.map((x) => (x.ev.length ? 4 : 0)), pointBackgroundColor: f.dias.map((x) => (x.saldo != null && x.saldo < 0 ? css("--bad") : cor[0])), pointBorderColor: css("--surface-solid"), pointBorderWidth: 2, spanGaps: false, order: 0,
      datalabels: { align: "top", offset: 4, color: (c) => (c.dataset.data[c.dataIndex] < 0 ? css("--bad") : css("--ink2")), font: { size: 10, weight: "700" }, display: (c) => { const x = f.dias[c.dataIndex]; return !isMobile() && x.ev.length && Math.abs(x.ent + x.sai) > 150; }, formatter: (v) => R0(v) } };
    const o = baseOpts(false, (v) => R0(v));
    o.scales.x.stacked = true; o.scales.y.stacked = false; o.scales.x.ticks = { color: css("--ink2"), maxRotation: 0, autoSkip: true, autoSkipPadding: 8, callback: (v, i) => L[i].slice(0, 2) };
    o.plugins.tooltip.callbacks = { title: (it) => "Dia " + L[it[0].dataIndex], label: () => null,
      afterBody: (it) => { const x = f.dias[it[0].dataIndex], out = x.ev.map((e) => (e.v > 0 ? "▲ " : "▼ ") + e.n + ": " + R(Math.abs(e.v)) + (e.feito ? " ✓" : e.vencida ? " (vencida)" : ""));
        out.push("", "Sobra acumulada no fim do dia: " + R(x.saldo)); return out.length ? out : ["Sem movimento"]; } };
    o.plugins.tooltip.filter = (t) => t.datasetIndex === 0; o.interaction = { mode: "index", intersect: false };
    if (f.atual) o.plugins.annotation = undefined;
    const ch = draw("cFluxoDia", "bar", L, [ds3, ds1, ds2], o);
    ch.$hoje = f.atual ? f.dHoje - 1 : null;
    $("fdLista").innerHTML = lista.length ? lista.slice(0, 10).map((e) => `<div class="fd-ev ${e.v > 0 ? "in" : "out"} ${e.vencida ? "venc" : ""}"><span class="d">${pad(e.d)}/${mm}</span><span class="n">${esc(e.n)}${e.vencida ? " · vencida" : ""}</span><span class="v">${e.v > 0 ? "+" : "−"} ${R(Math.abs(e.v))}</span><span class="s ${e.depois < 0 ? "neg" : ""}">saldo ${R(e.depois)}</span></div>`).join("") + (lista.length > 10 ? `<p class="muted small">+ ${lista.length - 10} movimento(s) no gráfico</p>` : "")
      : `<p class="muted small">Nenhum movimento previsto para o resto do mês.</p>`;
    $("fdAviso").innerHTML = neg ? `<span class="icon-bad">Em ${pad(f.min.d)}/${mm} o que entrou até ali não cobre as contas que vencem até esse dia: faltam ${R(-f.min.saldo)}. Só fica negativo quando não tem dinheiro para pagar em dia.</span>` : `<span class="icon-ok">Todas as contas do mês cabem no que entra até o vencimento de cada uma.</span>`;
    const inp = $("fdSaldo"); inp.onchange = () => { try { inp.value === "" ? localStorage.removeItem("saldo-inicio") : localStorage.setItem("saldo-inicio", inp.value); } catch (_) {} renderResumo(k); };
  }
  // linha vertical "hoje" no fluxo diário
  Chart.register({ id: "linhaHoje", afterDatasetsDraw(ch) { if (ch.$hoje == null || !ch.scales.x) return; const x = ch.scales.x.getPixelForValue(ch.$hoje), a = ch.chartArea, c = ch.ctx;
    c.save(); c.strokeStyle = css("--warn"); c.setLineDash([4, 4]); c.lineWidth = 1.5; c.beginPath(); c.moveTo(x, a.top); c.lineTo(x, a.bottom); c.stroke();
    c.setLineDash([]); c.fillStyle = css("--warn"); c.font = "600 11px Inter, sans-serif"; c.fillText("hoje", x + 4, a.top + 12); c.restore(); } });

  // ================= TARIFAS DA CEMIG =================
  const ORD_B = ["verde", "amarela", "vermelha 1", "vermelha 2"];
  const bandK = (b) => { const t = txt(b).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, ""); return /verm.*(2|ii)/.test(t) ? "vermelha 2" : /verm/.test(t) ? "vermelha 1" : /amarel/.test(t) ? "amarela" : /verde/.test(t) ? "verde" : null; };
  const f5 = (v) => (v == null ? "—" : "R$ " + v.toFixed(5).replace(".", ","));
  const f3 = (v) => (v == null ? "—" : "R$ " + v.toFixed(3).replace(".", ","));
  function tarifaCemig() {
    const T = DATA.tarifas; if (!T || !T.hist) return null; const ds_ = Object.keys(T.hist).sort(); if (!ds_.length) return null;
    const pega = (d) => { const o = {}; Object.entries(T.hist[d]).forEach(([b, v]) => { const k = bandK(b); if (k) o[k] = v; }); return o; };
    return { data: ds_[ds_.length - 1], t: pega(ds_[ds_.length - 1]), ant: ds_.length > 1 ? { data: ds_[ds_.length - 2], t: pega(ds_[ds_.length - 2]) } : null, band: T.band || {} };
  }
  function checaCemig() {
    const tc = tarifaCemig(); if (!tc) return null;
    const al = [], dt = (s) => s.split("-").reverse().join("/");
    for (let i = 1; i < ORD_B.length; i++) { const x = tc.t[ORD_B[i - 1]], y = tc.t[ORD_B[i]]; if (x && y && y <= x) al.push({ nv: "bad", t: `A tabela da Cemig não faz sentido: ${ORD_B[i]} (${f5(y)}) deveria ser mais cara que ${ORD_B[i - 1]} (${f5(x)}). Confira a aba Tarifas Cemig.` }); }
    if (!tc.t.verde) al.push({ nv: "bad", t: "Falta a tarifa da bandeira verde na aba Tarifas Cemig." });
    if (tc.t.verde && (tc.t.verde < 0.4 || tc.t.verde > 2)) al.push({ nv: "bad", t: `Tarifa verde de ${f5(tc.t.verde)} por kWh parece errada (esperado entre R$ 0,40 e R$ 2,00). Confira se não foi copiada por 1.000 kWh.` });
    if (tc.ant && tc.ant.t.verde && tc.t.verde) { const v = tc.t.verde / tc.ant.t.verde - 1;
      if (Math.abs(v) > 0.15) al.push({ nv: "bad", t: `A tarifa verde mudou ${fPct(v)} entre ${dt(tc.ant.data)} e ${dt(tc.data)}. Reajuste da Cemig costuma ser bem menor: confira se o valor está certo.` });
      else if (Math.abs(v) > 0.00001) al.push({ nv: "info", t: `Tarifa nova desde ${dt(tc.data)}: verde ${fPct(v)} em relação à tabela de ${dt(tc.ant.data)}.` }); }
    const serie = serieConta("Luz / Energia", "Conta").filter((x) => x.kwh && x.real);
    const L = serie.map((x) => { const ref = x.ref || addM(x.m, -1), bd = tc.band[ref] ? bandK(tc.band[ref]) : null, base = bd ? tc.t[bd] || null : null, rk = x.v / x.kwh;
      return { m: x.m, ref, bd, base, rk, v: x.v, kwh: x.kwh, fator: base ? rk / base : null }; });
    const med = (arr) => { if (!arr.length) return null; const s = arr.slice().sort((a, b) => a - b), n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; };
    L.forEach((l, i) => { l.fatorRef = med(L.filter((o, j) => j !== i && o.fator).map((o) => o.fator));
      if (!l.bd) l.st = "semband"; else if (l.base && l.fatorRef) { l.esp = l.base * l.fatorRef * l.kwh; l.dif = l.v - l.esp; l.difP = l.dif / l.esp; l.st = l.difP > 0.04 ? "acima" : l.difP < -0.04 ? "abaixo" : "ok"; } else l.st = "ok"; });
    L.forEach((l, i) => { const p = L[i - 1], msgs = [];
      if (l.st === "semband") al.push({ nv: "warn", m: l.m, t: `Conta de ${lab(l.m)}: falta a bandeira de ${lab(l.ref)} na aba Tarifas Cemig, então não dá para conferir.` });
      if (l.st === "acima") msgs.push(`veio ${R(l.dif)} (${fPct(l.difP)}) acima do esperado pela tarifa com bandeira ${l.bd} em ${lab(l.ref)} (esperado ≈ ${R(l.esp)})`);
      if (p && p.bd && l.bd && ORD_B.indexOf(l.bd) < ORD_B.indexOf(p.bd) && l.rk > p.rk * 1.01) msgs.push(`a bandeira ficou mais barata (${p.bd} → ${l.bd}), mas o kWh pago subiu ${fPct(l.rk / p.rk - 1)}; pela tabela deveria ter caído ${fPct(l.base / p.base - 1)}`);
      if (p && p.bd && l.bd && ORD_B.indexOf(l.bd) > ORD_B.indexOf(p.bd) && l.rk < p.rk * 0.99) msgs.push(`a bandeira ficou mais cara (${p.bd} → ${l.bd}), mas o kWh pago caiu: confira se a bandeira do mês está certa`);
      if (l.fator && (l.fator < 1.15 || l.fator > 1.5)) msgs.push(`impostos e encargos equivalem a ${fPct(l.fator - 1)} sobre a tarifa (o normal fica perto de 25% a 35%)`);
      if (msgs.length) { l.alerta = msgs; al.push({ nv: "bad", m: l.m, t: `Conta de energia de ${lab(l.m)}: ${msgs.join("; ")}. Confira na fatura ICMS, PIS/COFINS, iluminação pública, multa/juros ou outros itens.` }); }
      else if (l.st === "abaixo") al.push({ nv: "info", m: l.m, t: `Conta de energia de ${lab(l.m)} veio ${R(-l.dif)} abaixo do esperado pela tarifa.` });
    });
    return { tc, L, al };
  }
  function htmlCemig(sel) {
    const c = checaCemig(); if (!c) return `<p class="muted small">Cadastre a aba Tarifas Cemig na planilha para conferir as contas com a tarifa.</p>`;
    const dt = c.tc.data.split("-").reverse().join("/"), cl = (l) => (l.alerta ? "neg" : l.st === "abaixo" ? "pos" : "");
    const tag = (l) => l.alerta ? `<span class="tag bad">Conferir</span>` : l.st === "semband" ? `<span class="tag">Sem bandeira</span>` : l.st === "abaixo" ? `<span class="tag good">Abaixo</span>` : `<span class="tag good">Dentro</span>`;
    return `<h3 class="h3">⚡ Conferência com a tarifa da Cemig</h3>
      <div class="tarifas">${ORD_B.filter((b) => c.tc.t[b]).map((b) => `<span class="tf tf-${b.replace(" ", "")}"><i></i>${b[0].toUpperCase() + b.slice(1)} <b>${f5(c.tc.t[b])}</b></span>`).join("")}<span class="muted small">por kWh, sem impostos · tabela de ${dt}</span></div>
      ${c.al.length ? `<div class="alertas">${c.al.map((x) => `<div class="al ${x.nv}">${x.nv === "bad" ? "⚠️ " : x.nv === "warn" ? "❕ " : "ℹ️ "}${x.t}</div>`).join("")}</div>` : `<p class="icon-ok">Todas as contas batem com a tarifa da Cemig.</p>`}
      <div class="tblwrap"><table class="tbl"><tr><th>Conta</th><th>Consumo de</th><th>Bandeira</th><th>Tarifa (sem impostos)</th><th>R$/kWh pago</th><th>Impostos e encargos</th><th>Esperado</th><th>Pago</th><th>Diferença</th><th>Situação</th></tr>` +
      c.L.map((l) => `<tr class="${l.m === sel ? "sel" : ""}"><td>${lab(l.m)}</td><td>${lab(l.ref)}</td><td>${l.bd ? `<span class="tf tf-${l.bd.replace(" ", "")}"><i></i>${l.bd}</span>` : "—"}</td><td>${f5(l.base)}</td><td>${f3(l.rk)}</td><td>${l.fator ? fPct(l.fator - 1) : "—"}</td><td>${l.esp ? R(l.esp) : "—"}</td><td>${R(l.v)}</td><td class="${cl(l)}">${l.dif != null ? (l.dif > 0 ? "+" : "") + R(l.dif) : "—"}</td><td>${tag(l)}</td></tr>`).join("") + `</table></div>
      <p class="muted small">Esperado = kWh × tarifa da bandeira do mês do consumo × impostos e encargos típicos (mediana das outras contas). Contas de consumo antes de jun/26 podem ter outra tarifa (a Cemig reajusta todo ano no fim de maio). Quando a Cemig mudar a tarifa, mande o PDF que eu atualizo a aba Tarifas Cemig.</p>`;
  }

  // ================= SIMULAÇÃO EM QUALQUER QUADRO =================
  function simDelta(row) { // diferença no dinheiro livre da linha do plano com a simulação ativa
    if (!SIM || !SIM.itens.length || !DATA.plano[row]) return null;
    const ln = simCalc().linhas, l = ln.find((x) => x.k === row), u = ln[ln.length - 1];
    return l ? { d: l.tetoD - l.tetoA, accA: u.accA, accD: u.accD, u } : null;
  }
  function renderSimBanner() {
    const el = $("simBanner"); if (!el) return; const n = SIM && SIM.itens.length;
    if (!n || TAB === "sim") { el.hidden = true; return; }
    const tot = SIM.itens.reduce((s, it) => s + (it.t === "div" ? (+it.valor || 0) * (+it.n || 1) : +it.valor || 0), 0), ln = simCalc().linhas, u = ln[ln.length - 1];
    el.hidden = false;
    el.innerHTML = `<span>🧪 <b>Prévia de simulação ativa</b>: ${n} item(ns), ${R(tot)} no total · reserva em ${lab(u.k)}: ${R0(u.accA)} → <b>${R0(u.accD)}</b>. Nada mudou na planilha.</span><span class="sb-a"><button type="button" class="btn sm" id="sbVer">Ver detalhes</button><button type="button" class="btn ghost sm" id="sbLimpar">Limpar</button></span>`;
    $("sbVer").onclick = () => vaiPara("sim", "sRes"); $("sbLimpar").onclick = () => { SIM.itens = []; simSave(); renderAll(); };
  }
  function abrirSimRapida(alvo) {
    fecharSimRapida(); const h = alvo.querySelector("h2, h3"), tit = h ? h.textContent.trim() : "este quadro";
    const ov = document.createElement("div"); ov.className = "qs-ov"; ov.innerHTML = `<form class="qs" id="qsF"><div class="row-between"><h3>🧪 Simular a partir de “${esc(tit.slice(0, 40))}”</h3><button type="button" class="btn ghost sm" id="qsX">✕</button></div>
      <p class="muted small">Teste antes de fazer. É uma prévia: a planilha não muda.</p>
      <div class="qs-g"><div class="f"><label>O que</label><select id="qsT"><option value="pix">Compra no débito/Pix</option><option value="cartao">Compra no cartão</option><option value="div">Nova dívida (parcela mensal)</option></select></div>
      <div class="f"><label id="qsVl">Valor total (R$)</label><input id="qsV" type="number" inputmode="decimal" step="0.01" min="0" required placeholder="ex.: 600"></div>
      <div class="f"><label>Nº de vezes</label><input id="qsN" type="number" inputmode="numeric" min="1" max="60" value="1"></div>
      <div class="f"><label>Mês</label><select id="qsM">${DATA.meses.filter((k) => k >= DATA.planIni && DATA.plano[k]).map((k) => `<option value="${k}" ${k === SIM_DEF ? "selected" : ""}>${lab(k)}</option>`).join("")}</select></div>
      <div class="f wide"><label>Descrição (opcional)</label><input id="qsD" type="text" placeholder="ex.: tênis"></div></div>
      <div id="qsOut" class="qs-out"></div>
      <div class="qs-a"><button class="btn" type="submit">Ver prévia</button><button class="btn ghost" type="button" id="qsDesf" hidden>Desfazer</button><button class="btn ghost" type="button" id="qsOk" hidden>Manter na simulação</button></div></form>`;
    document.body.appendChild(ov); setTimeout(() => $("qsV").focus(), 50);
    let ult = null;
    $("qsT").onchange = () => { $("qsVl").textContent = $("qsT").value === "div" ? "Valor da parcela (R$)" : "Valor total (R$)"; if ($("qsT").value === "div" && +$("qsN").value < 2) $("qsN").value = 12; };
    $("qsX").onclick = fecharSimRapida; ov.addEventListener("click", (e) => { if (e.target === ov) fecharSimRapida(); });
    $("qsDesf").onclick = () => { if (ult) { SIM.itens = SIM.itens.filter((x) => x.id !== ult); simSave(); ult = null; renderAll(); } $("qsOut").innerHTML = ""; $("qsDesf").hidden = $("qsOk").hidden = true; };
    $("qsOk").onclick = fecharSimRapida;
    $("qsF").addEventListener("submit", (e) => { e.preventDefault(); const t = $("qsT").value, v = +$("qsV").value || 0; if (!v) return;
      const antes = simCalc().linhas;
      if (ult) SIM.itens = SIM.itens.filter((x) => x.id !== ult);
      const it = novoItem(t === "div" ? "div" : "compra"); it.valor = v; it.n = Math.max(1, +$("qsN").value || 1); it.mes = $("qsM").value; it.desc = $("qsD").value || "Simulação rápida";
      if (t === "pix") it.cartao = PIX; SIM.itens.push(it); ult = it.id; simSave();
      const depois = simCalc().linhas, cb = cobrancas(it).lista.map((c) => c.row), ms = [...new Set(cb)].filter((m) => depois.some((x) => x.k === m));
      const u0 = antes[antes.length - 1], u1 = depois[depois.length - 1];
      const piores = ms.map((m) => { const a0 = antes.find((x) => x.k === m) || {}, a1 = depois.find((x) => x.k === m) || {}; return { m, a: a0.tetoD ?? a0.tetoA, d: a1.tetoD }; });
      const estoura = piores.some((x) => x.d < -0.005);
      $("qsOut").innerHTML = `<div class="qs-res ${estoura ? "no" : "ok"}"><b>${estoura ? "⚠️ Não cabe: algum mês fica no vermelho" : "✅ Cabe no planejamento"}</b>
        <table class="tbl"><tr><th>Dinheiro de</th><th>Livre antes</th><th>Livre depois</th></tr>${piores.slice(0, 6).map((x) => `<tr><td>${lab(addM(x.m, 1))}</td><td>${R(x.a)}</td><td class="${x.d < 0 ? "neg" : "pos"}">${R(x.d)}</td></tr>`).join("")}</table>
        <p class="small">Reserva em ${lab(u1.k)}: ${R0(u0.accD)} → <b>${R0(u1.accD)}</b></p></div>`;
      $("qsDesf").hidden = $("qsOk").hidden = false; renderAll(); });
  }
  function fecharSimRapida() { const o = document.querySelector(".qs-ov"); if (o) o.remove(); }

  // ================= ASSINATURAS: prévia de cancelamento =================
  let ASS_OFF = new Set(); try { ASS_OFF = new Set(JSON.parse(localStorage.getItem("ass-off") || "[]")); } catch (_) {}
  const assOff = (n) => ASS_OFF.has(n);
  const assSave = () => { try { localStorage.setItem("ass-off", JSON.stringify([...ASS_OFF])); } catch (_) {} };
  function assToggle(n) { ASS_OFF.has(n) ? ASS_OFF.delete(n) : ASS_OFF.add(n); assSave(); }

  // ================= TABELAS SEM ROLAGEM LATERAL =================
  function rotulaTabelas() {
    document.querySelectorAll("table.tbl").forEach((t) => {
      const rows = [...t.rows]; const hr = rows.find((r) => r.querySelector("th")); if (!hr) return;
      const heads = []; [...hr.cells].forEach((c) => { for (let i = 0; i < (c.colSpan || 1); i++) heads.push(c.textContent.trim()); });
      t.classList.toggle("tstack", heads.length >= 4); t.classList.toggle("wide", heads.length >= 8); t.classList.toggle("xwide", heads.length >= 11); hr.classList.add("hrow");
      rows.forEach((r) => { if (r === hr || r.querySelector("th")) return; let i = 0; [...r.cells].forEach((c) => { if (!c.dataset.label) c.dataset.label = heads[i] || ""; if ((c.colSpan || 1) > 1) c.classList.add("span"); i += c.colSpan || 1; }); });
    });
  }

  // ---------- VR (Flash) ----------
  const isVR = (l) => l.paga === "VR";
  function ultimoDiaUtil(y, m0) { const fer = new Set(DATA.ferNac || DATA.feriados || []); const d = new Date(y, m0 + 1, 0, 12);
    while (d.getDay() === 0 || d.getDay() === 6 || fer.has(ymd(d))) d.setDate(d.getDate() - 1); return d; }
  function proxRecargaVR() { const h = new Date(); h.setHours(12, 0, 0, 0); let d = ultimoDiaUtil(h.getFullYear(), h.getMonth());
    if (d < h) d = ultimoDiaUtil(h.getFullYear(), h.getMonth() + 1);
    const uso = addM(d.getFullYear() + "-" + pad(d.getMonth() + 1), 1), dias = Math.round((d - h) / 864e5);
    return { data: d, uso, valor: vrMes(uso), dias }; }
  const CAT_REND = [["sal", "Salário líquido"], ["d13", "13º salário"], ["fer", "Férias"], ["abono", "Abono"], ["retro", "Retroativos"], ["outras", "Outras rendas"], ["vr", "VR / Flash"]];
  function rendimentos(m) {
    const j = idx(m), pr = DATA.proj, o = { sal: 0, d13: 0, fer: 0, abono: 0, retro: 0, outras: 0, vr: 0, itens: [] };
    if (j < 0) return o;
    const ex = (DATA.extras || []).filter((e) => e.ini <= m && e.fim >= m);
    let exConta = 0;
    ex.forEach((e) => { if (e.vr) return; exConta += e.valor; const c = /abono/i.test(e.desc) ? "abono" : /retroativ/i.test(e.desc) ? "retro" : "outras"; o[c] += e.valor; o.itens.push({ c, desc: e.desc, valor: e.valor }); });
    o.sal = (pr.renda[j] || 0) - exConta + (pr.ajuste[j] || 0); o.d13 = pr.d13[j] || 0; o.fer = pr.ferias[j] || 0; o.vr = pr.vr[j] || 0;
    o.dinheiro = o.sal + o.d13 + o.fer + o.abono + o.retro + o.outras; o.total = o.dinheiro + o.vr;
    ex.filter((e) => e.vr).forEach((e) => o.itens.push({ c: "vr", desc: e.desc, valor: e.valor }));
    return o;
  }
  function renderRend(k) {
    const r = rendimentos(k), cor = pal(), cores = [cor[0], cor[3], cor[5], cor[1], cor[4], css("--muted"), cor[2]];
    $("rendTit").textContent = lab(k) + (DATA.tipo[idx(k)] === "Realizado" ? " (realizado)" : " (previsto)");
    $("rendKpis").innerHTML = `<div class="kpi"><div class="l">Total em dinheiro</div><div class="v">${R(r.dinheiro)}</div><div class="s">tudo o que cai na conta no mês (líquido)</div></div>` +
      `<div class="kpi"><div class="l">+ VR / Flash</div><div class="v">${R(r.vr)}</div><div class="s">só para alimentação · total ${R(r.total)}</div></div>` +
      CAT_REND.filter(([c]) => c !== "vr" && r[c]).map(([c, n]) => `<div class="kpi"><div class="l">${n}</div><div class="v">${R(r[c])}</div><div class="s">${c === "sal" ? "salário fixo do mês" : (r.itens.filter((i) => i.c === c).map((i) => esc(i.desc)).join(" · ") || "")}</div></div>`).join("");
    const rg = range(), L = rg.map(lab), rs = rg.map(rendimentos);
    draw("cRend", "bar", L, CAT_REND.filter(([c]) => rs.some((x) => x[c])).map(([c, n]) => ds(n, rs.map((x) => Math.round(x[c] * 100) / 100), cores[CAT_REND.findIndex((z) => z[0] === c)])), baseOpts(true, (v) => R0(v)));
    $("tRend").innerHTML = `<tr><th>Mês</th>${CAT_REND.map(([, n]) => `<th>${n}</th>`).join("")}<th>Total em dinheiro</th></tr>` +
      rg.map((m, n) => `<tr class="${m === k ? "sel" : ""}"><td>${lab(m)}</td>${CAT_REND.map(([c]) => `<td>${rs[n][c] ? R(rs[n][c]) : "—"}</td>`).join("")}<td><b>${R(rs[n].dinheiro)}</b></td></tr>`).join("");
  }
  function vrMes(m) { const j = idx(m); return j >= 0 ? DATA.proj.vr[j] || 0 : 0; }
  const vrGasto = (m) => DATA.lanc.filter((l) => isVR(l) && l.ini === m).reduce((s, l) => s + l.total, 0);
  function vrSaldoIni(m) { let s = DATA.vrIni || 0; for (const k of DATA.meses) { if (k < DATA.planIni) continue; if (k >= m) break; s += vrMes(k) - vrGasto(k); } return s; }
  function diasRestantes(m) { const [y, mo] = m.split("-").map(Number), fim = new Date(y, mo, 0, 12), hoje = new Date(); hoje.setHours(12, 0, 0, 0);
    const ini = hoje.getFullYear() === y && hoje.getMonth() + 1 === mo ? hoje : new Date(y, mo - 1, 1, 12); const fer = new Set(DATA.ferNac || DATA.feriados);
    let cor = 0, ut = 0; for (let d = new Date(ini); d <= fim; d.setDate(d.getDate() + 1)) { cor++; const w = d.getDay(); if (w > 0 && w < 6 && !fer.has(ymd(d))) ut++; } return { cor, ut, total: fim.getDate() }; }
  function renderVR() {
    const hoje = new Date(), cal = hoje.getFullYear() + "-" + pad(hoje.getMonth() + 1);
    let m = F.mes > cal ? F.mes : cal; if (m < DATA.planIni) m = DATA.planIni; if (!DATA.meses.includes(m)) { $("vrCard").hidden = true; return; } $("vrCard").hidden = false;
    const futuro = m > cal, ini = futuro ? 0 : vrSaldoIni(m), rec = vrMes(m), g = vrGasto(m), saldo = ini + rec - g, dr = diasRestantes(m);
    const usado = ini + rec ? g / (ini + rec) : 0;
    $("vrTit").textContent = lab(m) + (futuro ? " (previsto)" : "");
    const pv = proxRecargaVR();
    $("vrBody").innerHTML = `<div class="vr-prox">📅 <b>Próxima recarga: ${fmtD(pv.data)}</b> (${["dom", "seg", "ter", "qua", "qui", "sex", "sáb"][pv.data.getDay()]}, ${pv.dias === 0 ? "hoje" : "em " + pv.dias + " dia" + (pv.dias > 1 ? "s" : "")}) · <b>${R(pv.valor)}</b> para usar em ${lab(pv.uso)} <span class="muted small">(último dia útil do mês)</span></div><div class="kpis">
        <div class="kpi"><div class="l">${futuro ? "VR previsto do mês" : "Saldo no Flash agora"}</div><div class="v ${saldo < 0 ? "neg" : ""}">${R(saldo)}</div><div class="s">${futuro ? "mais o que sobrar dos meses anteriores" : R(ini) + " que sobrou + " + R(rec) + " do mês" + (g ? " − " + R(g) + " gastos" : "")}</div></div>
        <div class="kpi"><div class="l">Por semana</div><div class="v">${R(saldo / Math.max(1, dr.cor / 7))}</div><div class="s">${(dr.cor / 7).toFixed(1).replace(".", ",")} semanas ${futuro ? "no mês" : "até o fim do mês"}</div></div>
        <div class="kpi"><div class="l">Por dia útil</div><div class="v">${R(dr.ut ? saldo / dr.ut : 0)}</div><div class="s">${dr.ut} dias úteis (seg a sex, sem feriados nacionais) ${futuro ? "no mês" : "restantes, com hoje"}</div></div>
        <div class="kpi"><div class="l">Por dia corrido</div><div class="v">${R(saldo / Math.max(1, dr.cor))}</div><div class="s">${dr.cor} dias ${futuro ? "no mês" : "restantes (com hoje)"}</div></div></div>
      <div class="meter"><div class="t"><span>Usado em ${lab(m)}: ${R(g)}</span><span>${P(usado)} do disponível</span></div><div class="bar"><span style="width:${Math.min(100, usado * 100)}%;background:${usado > 1 ? css("--bad") : css("--s3")}"></span></div></div>
      <p class="muted small">Só o VR do mês daria ${R(rec / (dr.total / 7))}/semana. O que sobra no Flash passa para o mês seguinte (não vira reserva em dinheiro).${DATA.vrIni ? "" : " Saldo que sobrou antes do plano: R$ 0 (ajuste em Metas e Cortes se tinha algo)."}</p>` +
      (DATA.lanc.some((l) => isVR(l) && l.ini === m) ? `<div class="tblwrap"><table class="tbl"><tr><th>Compra no Flash</th><th>Data</th><th>Valor</th></tr>` + DATA.lanc.filter((l) => isVR(l) && l.ini === m).map((l) => `<tr><td>${esc(l.desc)}</td><td>${l.data.split("-").reverse().join("/")}</td><td>${R(l.total)}</td></tr>`).join("") + `</table></div>` : "");
  }

  // ---------- saúde financeira, ritmo, limites ----------
  const semVR = (l) => l.paga !== "VR";
  const lancMes = (k) => DATA.lanc.filter((l) => l.data.slice(0, 7) === k && semVR(l));
  const custoEss = () => DATA.mc.custo || (DATA.fixos.filter((f) => ativo(f, F.mes)).reduce((s, f) => s + f.custo, 0) + (DATA.minMes || 0));
  function parcelasMeu(m) { return DATA.dividas.filter((d) => ativo(d, m)).reduce((s, d) => s + d.custo, 0) + DATA.lanc.filter((l) => l.parcelado && l.ini && l.ini <= m && l.fim >= m).reduce((s, l) => s + l.custo, 0); }
  function semaforo(v, bom, medio, maior) { const ok = maior ? v >= bom : v <= bom, mid = maior ? v >= medio : v <= medio; return ok ? "ok" : mid ? "mid" : "bad"; }
  function renderSaude(k) {
    const Fm = addM(k, 1), j = idx(Fm), sal = j >= 0 ? DATA.proj.renda[j] : 0, p = DATA.plano[k] || {};
    const parc = parcelasMeu(Fm), pp = sal ? parc / sal : 0, meses = (p.reservaAc || 0) / (custoEss() || 1), poup = sal ? (p.reserva || 0) / sal : 0;
    const it = [
      ["Parcelas ÷ salário", P(pp), "minha parte " + R0(parc) + " de " + R0(sal) + " · ideal até 30%", semaforo(pp, 0.3, 0.45, false)],
      ["Meses cobertos pela reserva", meses.toFixed(1).replace(".", ","), "reserva " + R0(p.reservaAc) + " ÷ custo essencial " + R0(custoEss()) + " · ideal 6", semaforo(meses, 6, 3, true)],
      ["Salário que vai para a reserva", P(poup), R0(p.reserva) + " no mês · meta " + P(DATA.pct), semaforo(poup, DATA.pct - 0.001, 0.05, true)]];
    $("saude").innerHTML = `<div class="row-between"><h2>Saúde financeira</h2><p class="muted small">compras de ${lab(k)}, pagas com o salário de ${lab(Fm)}</p></div><div class="kpis">` +
      it.map(([l, v, s, c]) => `<div class="kpi sem ${c}"><div class="l"><i></i>${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("") + "</div>";
  }
  function renderRitmo(k) {
    const [y, m] = k.split("-").map(Number), dias = new Date(y, m, 0).getDate(), hoje = new Date();
    const atual = hoje.getFullYear() === y && hoje.getMonth() + 1 === m, passado = new Date(y, m, 0, 23) < hoje;
    const ate = atual ? hoje.getDate() : passado ? dias : 0;
    const p = DATA.plano[k] || {}, ref = p.teto || DATA.minMes || 0, refNome = p.teto ? "sugerido" : "mínimo";
    const ls = lancMes(k), porDia = Array(dias).fill(0); ls.forEach((l) => (porDia[+l.data.slice(8, 10) - 1] += l.meuTotal));
    let acc = 0; const cum = porDia.map((v, n) => (acc += v, n < ate ? Math.round(acc * 100) / 100 : null));
    const gasto = ate ? cum[ate - 1] : 0, ideal = porDia.map((_, n) => Math.round((ref * (n + 1)) / dias * 100) / 100);
    const cor = pal(), labs = porDia.map((_, n) => String(n + 1));
    draw("cRitmo", "line", labs, [
      ds("Gasto acumulado", cum, cor[0], { fill: false, borderWidth: 2.5, pointRadius: 0, tension: 0.15, datalabels: { display: (c) => c.dataIndex === ate - 1, align: "top" } }),
      ds("Ritmo ideal (" + refNome + " " + R0(ref) + ")", ideal, css("--muted"), { fill: false, borderWidth: 2, borderDash: [6, 4], pointRadius: 0, datalabels: { display: false } })],
      baseOpts(false, (v) => R0(v)));
    let t;
    if (!ate) t = `Mês ainda não começou. O ritmo ideal é ${R0(ref / (dias / 7))} por semana.`;
    else { const proj = atual ? (gasto / ate) * dias : gasto, idealHoje = ideal[ate - 1];
      t = `Até ${atual ? "hoje (dia " + ate + ")" : "o fim do mês"}: ${R(gasto)} gastos (sem VR), contra ${R(idealHoje)} do ritmo ideal. ` +
        (atual ? `Nesse ritmo você fecha ${lab(k)} em <b>${R(proj)}</b> (${refNome} ${R0(ref)}). ` : "") +
        (proj > ref + 0.5 ? `<span class="icon-bad">${atual ? "Vai passar" : "Passou"} ${R(proj - ref)} do ${refNome}.</span>` : `<span class="icon-ok">Dentro do ${refNome}.</span>`); }
    $("ritmoTxt").innerHTML = t + (ls.length ? "" : ` <span class="muted">Sem lançamentos neste mês na aba Lançamentos.</span>`);
  }
  function renderLimites(k) {
    const ls = lancMes(k), g = {}; ls.forEach((l) => (g[l.cat || "Outros"] = (g[l.cat || "Outros"] || 0) + l.meuTotal));
    const lims = DATA.mc.limites.slice(); Object.keys(g).forEach((c) => { if (!lims.some((x) => x.cat === c)) lims.push({ cat: c, lim: 0, extra: 1 }); });
    const totL = lims.reduce((s, x) => s + x.lim, 0), totG = Object.values(g).reduce((s, v) => s + v, 0);
    $("limResumo").textContent = `${lab(k)} · ${R0(totG)} de ${R0(totL)} (${P(totL ? totG / totL : 0)})`;
    $("limites").innerHTML = lims.filter((x) => x.lim > 0 || g[x.cat]).map((x) => { const v = g[x.cat] || 0, u = x.lim ? v / x.lim : v > 0 ? 2 : 0;
      const c = u > 1 ? css("--bad") : u >= 0.8 ? css("--s4") : css("--good");
      return `<div class="meter"><div class="t"><span>${esc(x.cat)}${x.extra ? ' <span class="muted small">(sem limite)</span>' : ""}</span><span>${R0(v)} / ${x.lim ? R0(x.lim) : "—"}${u > 1 && x.lim ? ' <b style="color:' + css("--bad") + '">+' + R0(v - x.lim) + "</b>" : ""}</span></div>
        <div class="bar"><span style="width:${Math.min(100, u * 100)}%;background:${c}"></span></div></div>`; }).join("") ||
      `<p class="muted">Defina os limites na aba Metas e Cortes da planilha.</p>`;
  }
  // ---------- reembolsos recebidos, parcelas que acabam, previsto x real ----------
  function devidoPessoa(m) { const o = {}; receber(m).forEach((r) => (o[r.pessoa] = (o[r.pessoa] || 0) + r.valor)); return o; }
  function pendencias(k) {
    const pessoas = new Set(Object.keys(DATA.recebido)); DATA.meses.forEach((m) => Object.keys(devidoPessoa(m)).forEach((p) => pessoas.add(p)));
    const mes = {}, atras = {};
    pessoas.forEach((p) => { const d = devidoPessoa(k)[p] || 0, r = (DATA.recebido[p] || {})[k] || 0; mes[p] = { d, r, pend: Math.max(0, d - r) };
      atras[p] = DATA.meses.filter((m) => m < REF.mes).reduce((s, m) => s + Math.max(0, (devidoPessoa(m)[p] || 0) - ((DATA.recebido[p] || {})[m] || 0)), 0); });
    return { mes, atras };
  }
  let RECP = ""; try { RECP = localStorage.getItem("rec-pessoa") || ""; } catch (_) {}
  const passaPessoa = (p) => !RECP || p === RECP;
  function pessoasTodas() {
    const s = new Set(); const add = (p) => { if (p && p !== "-") s.add(p); };
    DATA.fixos.forEach((f) => { add(f.q1); add(f.q2); }); DATA.dividas.forEach((d) => { add(d.q1); add(d.q2); }); DATA.lanc.forEach((l) => { add(l.q1); add(l.q2); }); DATA.recItens.forEach((r) => add(r.pessoa));
    return [...s].sort();
  }
  function renderRecPessoaSel() {
    const el = $("recPessoa"), ps = pessoasTodas(); if (RECP && !ps.includes(RECP)) RECP = "";
    el.innerHTML = `<option value="">Todas as pessoas</option>` + ps.map((p) => `<option ${p === RECP ? "selected" : ""}>${esc(p)}</option>`).join("");
  }
  // dívidas das pessoas comigo (itens com fim): parcelas pagas x pendentes
  function renderDevPessoa(k) {
    const n = (t) => txt(t).toLowerCase().replace(/\s+/g, " "), itens = [];
    const add = (p, v, desc, ini, fim, tot) => { if (!p || p === "-" || !(v > 0.004) || !fim) return; if (!passaPessoa(p)) return;
      const N = tot || diffM(ini, fim) + 1, i0 = tot ? addM(fim, -(N - 1)) : ini || addM(fim, -(N - 1));
      itens.push({ p, v, desc, ini: i0, fim, N }); };
    DATA.dividas.forEach((d) => { add(d.q1, d.r1, d.desc, d.ini, d.fim, d.totalParc); add(d.q2, d.r2, d.desc, d.ini, d.fim, d.totalParc); });
    DATA.lanc.filter((l) => l.parcelado).forEach((l) => { add(l.q1, l.r1, l.desc, l.ini, l.fim, l.n); add(l.q2, l.r2, l.desc, l.ini, l.fim, l.n); });
    const rows = itens.map((it) => {
      let antes = 0; for (let m = it.ini; m <= it.fim; m = addM(m, 1)) if (m < DATA.planIni) antes++;
      const rec = DATA.recItens.filter((r) => n(r.pessoa) === n(it.p) && n(r.item) === n(it.desc)).reduce((s, r) => s + r.valor, 0);
      const pagas = Math.min(it.N, antes + Math.floor((rec + 0.05) / it.v)), pend = it.N - pagas;
      let devidas = 0; for (let m = it.ini; m <= it.fim && m <= k; m = addM(m, 1)) devidas++;
      const st = pend === 0 ? `<span class="tag good">Quitado</span>` : pagas > devidas ? `<span class="tag good">Adiantado ${pagas - devidas}</span>` : pagas < devidas ? `<span class="tag bad">Atrasado ${devidas - pagas}</span>` : `<span class="tag">Em dia</span>`;
      const falta = pend ? (diffM(k, it.fim) >= 0 ? `${lab(it.fim)} (${diffM(k, it.fim) ? "faltam " + (diffM(k, it.fim) + 1) + " meses" : "último mês"})` : lab(it.fim)) : "—";
      return { it, pagas, pend, pago: antes * it.v + rec, pendV: Math.max(0, it.N * it.v - antes * it.v - rec), st, falta };
    }).sort((a, b) => a.it.p.localeCompare(b.it.p) || (a.pend === 0) - (b.pend === 0) || a.it.fim.localeCompare(b.it.fim));
    $("tDevPessoa").innerHTML = `<tr><th>Pessoa</th><th>Dívida</th><th>Parcela</th><th>Pagas</th><th>Pendentes</th><th>Total pago</th><th>Total pendente</th><th>Última parcela</th><th>Situação</th></tr>` +
      (rows.map((r) => `<tr class="${r.pend === 0 ? "fora" : ""}"><td>${esc(r.it.p)}</td><td style="text-align:left">${esc(r.it.desc)}</td><td>${R(r.it.v)}</td><td>${r.pagas} de ${r.it.N}</td><td>${r.pend}</td><td class="pos">${R(r.pago)}</td><td class="${r.pendV > 0.005 ? "neg" : ""}">${R(r.pendV)}</td><td>${r.falta}</td><td>${r.st}</td></tr>`).join("") || `<tr><td colspan="9" class="muted">Nenhuma dívida parcelada${RECP ? " de " + esc(RECP) : ""}.</td></tr>`) +
      (rows.length > 1 ? `<tr class="total"><td>Total</td><td></td><td></td><td></td><td></td><td>${R(rows.reduce((s, r) => s + r.pago, 0))}</td><td>${R(rows.reduce((s, r) => s + r.pendV, 0))}</td><td></td><td></td></tr>` : "");
  }
  // extrato para enviar à pessoa (mês a mês)
  let EXT_OPEN = false;
  function carregaLib(nome, cdn) { return new Promise((ok, err) => { if (window.html2canvas) return ok();
    const tries = ["vendor/" + nome, nome, cdn]; (function t(i) { if (i >= tries.length) return err(new Error("lib")); const s = document.createElement("script"); s.src = tries[i]; s.onload = () => ok(); s.onerror = () => t(i + 1); document.head.appendChild(s); })(0); }); }
  // exportar qualquer quadro: copiar imagem / baixar imagem / PDF
  async function capturaImagem(alvo, modo, nome, st) {
    const msg = (t) => { if (st) st.textContent = t; else toast(t); };
    try { msg("Gerando imagem…"); await carregaLib("html2canvas.min.js", "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
      document.body.classList.add("capturando"); const mk = marcaRelatorio(alvo);
      const cv = await window.html2canvas(alvo, { backgroundColor: getComputedStyle(document.body).backgroundColor, scale: 2, width: alvo.scrollWidth, windowWidth: Math.max(alvo.scrollWidth + 40, window.innerWidth) });
      document.body.classList.remove("capturando"); mk.remove();
      const blob = await new Promise((r) => cv.toBlob(r, "image/png"));
      if (modo === "copiar" && navigator.clipboard && window.ClipboardItem) { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); msg("Imagem copiada! É só colar."); }
      else { const u = URL.createObjectURL(blob), l = document.createElement("a"); l.href = u; l.download = (nome || "dashboard") + ".png"; l.click(); setTimeout(() => URL.revokeObjectURL(u), 2000); msg(modo === "copiar" ? "Este navegador não deixa copiar: baixei a imagem." : "Imagem baixada."); }
    } catch (e) { document.body.classList.remove("capturando"); document.querySelectorAll(".cap-brand").forEach((x) => x.remove()); msg("Não consegui gerar a imagem. Use o print ou o PDF."); }
  }
  function marcaRelatorio(alvo) { const d = document.createElement("div"); d.className = "cap-brand";
    d.innerHTML = `<img src="img/marca.png" alt=""><div><b>Controle Financeiro</b><span>Nathan · ${fmtD(new Date())}</span></div>`; alvo.insertBefore(d, alvo.firstChild); return d; }
  function imprimeAlvo(alvo) { const mk = marcaRelatorio(alvo); setTimeout(() => mk.remove(), 900); alvo.classList.add("print-alvo"); document.body.classList.add("print-one"); window.print(); setTimeout(() => { document.body.classList.remove("print-one"); alvo.classList.remove("print-alvo"); }, 600); }
  function toast(t) { let el = document.querySelector(".toast"); if (!el) { el = document.createElement("div"); el.className = "toast"; document.body.appendChild(el); } el.textContent = t; el.classList.add("on"); clearTimeout(el._t); el._t = setTimeout(() => el.classList.remove("on"), 2500); }
  const slug = (t) => txt(t).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40);
  function poeExport(alvo) {
    if (!alvo || alvo.hidden || alvo.querySelector(":scope > .exp-bar")) return;
    const bar = document.createElement("div"); bar.className = "exp-bar";
    const podeSim = !alvo.closest("#tab-sim") && !alvo.closest(".zoomov");
    bar.innerHTML = (podeSim ? `<button type="button" title="Simular uma compra ou dívida a partir daqui" data-x="sim">🧪</button>` : "") + `<button type="button" title="Copiar como imagem" data-x="copiar">📋</button><button type="button" title="Baixar imagem (PNG)" data-x="baixar">🖼️</button><button type="button" title="Imprimir / salvar PDF" data-x="pdf">PDF</button>`;
    bar.addEventListener("click", (e) => { const b = e.target.closest("button"); if (!b) return; e.stopPropagation();
      const h = alvo.querySelector("h2, h3"), nome = slug((h ? h.textContent : "quadro") + " " + lab(F.mes));
      if (b.dataset.x === "sim") abrirSimRapida(alvo); else if (b.dataset.x === "pdf") imprimeAlvo(alvo); else capturaImagem(alvo, b.dataset.x, nome); });
    alvo.appendChild(bar);
  }
  function exportaTudo() { document.querySelectorAll("section.card, .grid2 > .card, .ov-grid > .card").forEach(poeExport); }
  async function extratoImagem(modo) {
    const el = $("extrato"), st = $("extStatus");
    try { st.textContent = "Gerando imagem…"; await carregaLib("html2canvas.min.js", "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js");
      el.classList.add("ext-cap");
      const alvo = el.querySelector(".ext-docs"); const cv = await window.html2canvas(alvo, { width: alvo.scrollWidth, windowWidth: alvo.scrollWidth + 40, backgroundColor: getComputedStyle(document.body).backgroundColor, scale: 2 });
      el.classList.remove("ext-cap");
      const blob = await new Promise((r) => cv.toBlob(r, "image/png"));
      if (modo === "copiar" && navigator.clipboard && window.ClipboardItem) { await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]); st.textContent = "Imagem copiada! É só colar no WhatsApp."; }
      else { const u = URL.createObjectURL(blob), l = document.createElement("a"); l.href = u; l.download = "extrato-" + (RECP || "todos").toLowerCase() + ".png"; l.click(); setTimeout(() => URL.revokeObjectURL(u), 2000); st.textContent = modo === "copiar" ? "Este navegador não deixa copiar: baixei a imagem." : "Imagem baixada."; }
    } catch (e) { el.classList.remove("ext-cap"); st.textContent = "Não consegui gerar a imagem (" + e.message + "). Use o print ou o PDF."; }
  }
  function renderExtrato(k) {
    const box = $("extrato"); $("extBtn").textContent = EXT_OPEN ? "Fechar extrato" : "📄 Extrato para enviar";
    if (!EXT_OPEN) { box.hidden = true; box.innerHTML = ""; return; } box.hidden = false;
    const ini = addM(k, -1), fimJ = addM(k, 6), ms = DATA.meses.filter((m) => m >= ini);
    const porMes = ms.map((m) => { const rc = receber(m); statusReceber(m, rc); return { m, rc }; });
    const pessoas = RECP ? [RECP] : [...new Set(porMes.filter((x) => x.m <= fimJ).flatMap((x) => x.rc.map((r) => r.pessoa)))].sort();
    const hoje = new Date(), cor = pal();
    const estim = (b) => DATA.fixos.some((f) => f.desc === b && !f.fim && /aluguel|condom|luz|energia/i.test(b));
    const fixo = (b) => DATA.fixos.some((f) => f.desc === b && !f.fim);
    const curto = (b) => b.replace(/\s*—.*$/, "").replace(/\s*\((anual|ingressos)[^)]*\)/gi, "").replace(/\s*\(\w{3}\/\d{2}\)$/, "").trim();
    const charts = [];
    box.innerHTML = `<div class="ext-docs">` + pessoas.map((p, pi) => {
      const linhas = porMes.map((x) => ({ m: x.m, it: x.rc.filter((r) => r.pessoa === p) })).filter((x) => x.it.length);
      const itens = {}; linhas.forEach((x) => x.it.forEach((r) => { const o = itens[r.base] || (itens[r.base] = { base: r.base, fim: x.m, valores: [], n: 0, pend: 0 });
        o.fim = x.m; if (r.valor - (r.pago || 0) > 0.015) o.n++; o.valores.push(r.valor); o.pend += r.valor - (r.pago || 0); }));
      const finitos = Object.values(itens).filter((o) => !fixo(o.base)), recorr = Object.values(itens).filter((o) => fixo(o.base));
      const totFin = finitos.reduce((s, o) => s + o.pend, 0);
      const jan = linhas.filter((x) => x.m <= fimJ);
      const cols = [...new Set(jan.flatMap((x) => x.it.map((r) => r.base)))].sort((a2, b2) => fixo(a2) - fixo(b2));
      const nomeCol = {}; cols.forEach((c) => (nomeCol[c] = curto(c))); const uniq = [...new Set(cols.map((c) => nomeCol[c]))];
      const resumoM = jan.map((x) => { const tm = x.it.reduce((s, r) => s + r.valor, 0), pg = x.it.reduce((s, r) => s + (r.st === "pago" ? r.valor : Math.min(r.valor, r.pago || 0)), 0); return { m: x.m, tm, pg, pend: Math.max(0, tm - pg) }; });
      const totJan = resumoM.reduce((s, x) => s + x.pend, 0), pagoJan = resumoM.reduce((s, x) => s + x.pg, 0);
      const temEst = jan.some((x) => x.it.some((r) => estim(r.base) && r.st !== "pago"));
      charts.push({ id: "extC" + pi, resumoM });
      return `<div class="ext-doc">
        <div class="ext-h"><div><h3>Valores a receber — ${esc(p)}</h3><div class="muted small">${lab(jan.length ? jan[0].m : ini)} a ${lab(jan.length ? jan[jan.length - 1].m : fimJ)} · gerado em ${fmtD(hoje)}</div></div>
          <div class="ext-tots"><div class="ext-tot"><div class="l">Já pago no período</div><div class="v pos">${R(pagoJan)}</div></div>
          <div class="ext-tot"><div class="l">Falta pagar no período${temEst ? "*" : ""}</div><div class="v neg">${R(totJan)}</div></div></div></div>
        <div class="ext-chart"><canvas id="extC${pi}"></canvas></div>
        <div class="tblwrap"><table class="tbl ext-t"><tr><th>Mês</th>${uniq.map((u) => `<th>${esc(u)}</th>`).join("")}<th>Total do mês</th><th>Situação</th></tr>` +
        jan.map((x, n) => { const rm = resumoM[n];
          const cel = (u) => { const rs = x.it.filter((r) => nomeCol[r.base] === u); if (!rs.length) return "—"; const v = rs.reduce((s, r) => s + r.valor, 0), pago = rs.every((r) => r.st === "pago"),
            par = rs.map((r) => (r.item.match(/\((\d+\/\d+)\)/) || [])[1]).filter(Boolean)[0], est = rs.some((r) => estim(r.base)) && !pago;
            return `<span class="${pago ? "ok" : "pend"}">${R(v)}${est ? "*" : ""}</span>` + (par ? `<div class="muted small">${par}</div>` : ""); };
          const st = rm.pend < 0.015 ? '<span class="tag good">Pago</span>' : rm.pg > 0.004 ? `<span class="tag bad">Falta ${R(rm.pend)}</span>` : x.m < mesAtual() ? '<span class="tag bad">Atrasado</span>' : '<span class="tag bad">Pendente</span>';
          return `<tr class="${x.m === k ? "sel" : ""}"><td><b>${lab(x.m)}</b></td>${uniq.map((u) => `<td>${cel(u)}</td>`).join("")}<td><b>${R(rm.tm)}</b></td><td>${st}</td></tr>`; }).join("") +
        `<tr class="total"><td>Total</td>${uniq.map(() => "<td></td>").join("")}<td>${R(resumoM.reduce((s, x) => s + x.tm, 0))}</td><td><span class="pend">falta ${R(totJan)}</span></td></tr></table></div>
        <p class="ext-leg"><span class="ok">■ verde = pago</span> · <span class="pend">■ vermelho = pendente</span>${temEst ? " · * valor estimado (aluguel, condomínio e luz mudam todo mês; vale o valor da conta quando chegar)" : ""}</p>
        <h4>Resumo por item</h4>
        <table class="tbl ext-t"><tr><th>Item</th><th>Valor por mês</th><th>Até</th><th>Meses restantes</th><th>Total restante</th></tr>` +
        finitos.filter((o) => o.pend > 0.5).map((o) => `<tr><td style="text-align:left">${esc(curto(o.base))}</td><td>${R(o.valores[o.valores.length - 1])}</td><td><b>${lab(o.fim)}</b> (última)</td><td>${o.n}</td><td class="pend"><b>${R(o.pend)}</b></td></tr>`).join("") +
        recorr.map((o) => `<tr><td style="text-align:left">${esc(curto(o.base))}</td><td>${R(o.valores[o.valores.length - 1])}${estim(o.base) ? "*" : ""}</td><td>todo mês (sem fim)</td><td>—</td><td>${R(o.valores[o.valores.length - 1])}/mês</td></tr>`).join("") +
        `</table>${finitos.some((o) => o.pend > 0.5) ? `<p class="muted small">Parcelas com fim: falta <b>${R(totFin)}</b> no total até a última parcela.</p>` : ""}
      </div>`;
    }).join("") + `</div><div class="ext-act"><span class="muted small" id="extStatus"></span><button type="button" class="btn sm" id="extCopy">📋 Copiar imagem</button><button type="button" class="btn ghost sm" id="extPng">Baixar imagem</button><button type="button" class="btn ghost sm" id="extPrint">Imprimir / salvar PDF</button></div>`;
    const ok = css("--good"), bad = css("--bad");
    charts.forEach((c) => { const ch = draw(c.id, "bar", c.resumoM.map((x) => lab(x.m)), [ds("Pago", c.resumoM.map((x) => Math.round(x.pg * 100) / 100), ok), ds("Pendente", c.resumoM.map((x) => Math.round(x.pend * 100) / 100), bad)], baseOpts(true, (v) => R0(v)));
      ch.options.animation = false; ch.update(); });
    $("extPrint").addEventListener("click", () => { document.body.classList.add("print-ext"); window.print(); setTimeout(() => document.body.classList.remove("print-ext"), 500); });
    $("extCopy").addEventListener("click", () => extratoImagem("copiar"));
    $("extPng").addEventListener("click", () => extratoImagem("baixar"));
  }
  function renderRecebidos(k) {
    const { mes, atras } = pendencias(k);
    $("tRecebido").innerHTML = `<tr><th>Pessoa</th><th>Devido em ${lab(k)}</th><th>Recebido</th><th>Pendente</th><th>Atrasado (meses anteriores)</th></tr>` +
      Object.entries(mes).filter(([p, o]) => passaPessoa(p) && (o.d || o.r || atras[p])).map(([p, o]) => `<tr><td>${esc(p)}</td><td>${R(o.d)}</td><td class="${o.r >= o.d - 0.005 && o.d ? "pos" : ""}">${R(o.r)}</td><td class="${o.pend > 0.005 ? "neg" : ""}">${R(o.pend)}</td><td class="${atras[p] > 0.005 ? "neg" : ""}">${R(atras[p])}</td></tr>`).join("") || `<tr><td colspan="5" class="muted">Nada a receber neste mês.</td></tr>`;
    const tA = Object.entries(atras).filter(([p]) => passaPessoa(p)).reduce((s, [, v]) => s + v, 0);
    $("recAviso").innerHTML = tA > 0.005 ? `<span class="icon-bad">Tem ${R(tA)} de reembolso atrasado. Cobre e marque no bloco RECEBIDO da aba Reembolsos.</span>` : `<span class="muted small">Quando alguém te pagar, preencha o bloco RECEBIDO da aba Reembolsos da planilha.</span>`;
  }
  function renderLibera(k) {
    const ms = DATA.meses.filter((m) => m >= addM(k, 1)), v = ms.map((m) => Math.round(parcelasMeu(m) * 100) / 100), cor = pal();
    const cl = draw("cLibera", "line", ms.map(lab), [ds("Minhas parcelas por mês", v, cor[1], { fill: true, backgroundColor: cor[1] + "33", stepped: true, borderWidth: 2, pointRadius: 0,
      datalabels: { display: (c) => c.dataIndex === 0 || v[c.dataIndex] !== v[c.dataIndex - 1], align: "top" } })], baseOpts(false, (x) => R0(x)));
    cl.options.plugins.legend.display = false; cl.update();
    const quedas = [];
    for (let n = 1; n < ms.length; n++) { const dif = v[n - 1] - v[n]; if (dif > 5) {
      const fim = ms[n - 1], itens = DATA.dividas.filter((d) => d.fim === fim && d.custo > 0).map((d) => d.desc).concat(DATA.lanc.filter((l) => l.parcelado && l.fim === fim).map((l) => l.desc));
      quedas.push(`<tr><td>${lab(ms[n])}</td><td class="pos">+${R(dif)}/mês</td><td>${R(v[n])}</td><td style="text-align:left">${esc(itens.slice(0, 4).join(", "))}${itens.length > 4 ? " +" + (itens.length - 4) : ""}</td></tr>`); } }
    $("tLibera").innerHTML = `<tr><th>A partir de</th><th>Libera</th><th>Parcelas no mês</th><th>O que termina</th></tr>` + (quedas.join("") || `<tr><td colspan="4" class="muted">Nenhuma parcela termina no período.</td></tr>`);
  }
  function renderPrevReal() {
    const fs = DATA.fats.filter((f) => f.prev > 0);
    const linhas = fs.map((f) => { const j = idx(f.mes), atual = f.status === "Fechada" ? f.valor : (DATA.faturas[f.card] || [])[j] ?? f.valor, dif = atual - f.prev;
      return `<tr><td>${lab(f.mes)}</td><td>${esc(f.card)}</td><td>${R(f.prev)}</td><td>${R(atual)}</td><td class="${dif > 0.005 ? "neg" : dif < -0.005 ? "pos" : ""}">${dif > 0 ? "+" : ""}${R(dif)} (${P(f.prev ? dif / f.prev : 0)})</td><td>${f.status === "Fechada" ? "real (fechada)" : "previsão atual"}</td></tr>`; });
    const ms = DATA.meses.filter((m) => m <= REF.mes && (DATA.plano[m] || {}).mes);
    const gl = ms.map((m) => { const p = DATA.plano[m], g = lancMes(m).reduce((s, l) => s + l.meuTotal, 0), ref = p.teto || DATA.minMes || 0;
      return `<tr><td>${lab(m)}</td><td>${R(ref)}${p.teto ? "" : " (mínimo)"}</td><td>${R(g)}</td><td class="${g > ref + 0.005 ? "neg" : "pos"}">${g > ref ? "+" : ""}${R(g - ref)}</td></tr>`; });
    $("tPrevFat").innerHTML = `<tr><th>Fatura</th><th>Cartão</th><th>Previsto antes</th><th>Agora</th><th>Diferença</th><th>Situação</th></tr>` + (linhas.join("") || `<tr><td colspan="6" class="muted">Preencha "Previsão anterior" na aba Faturas para comparar.</td></tr>`);
    $("tPrevGasto").innerHTML = `<tr><th>Mês das compras</th><th>Planejado</th><th>Gasto lançado (valor das compras, sem VR)</th><th>Diferença</th></tr>` + gl.join("");
  }
  // ---------- metas e cortes ----------
  function renderMetas() {
    const k = F.mes, ms = DATA.meses.filter((m) => DATA.plano[m]), cor = pal();
    const metas = DATA.mc.metas.slice().sort((a, b) => a.prio - b.prio); let cum = 0;
    const acM = (m) => (DATA.plano[m] || {}).reservaAc || 0, rdM = (m) => (DATA.plano[m] || {}).reservaRend || 0;
    const quando = (alvo, f) => ms.find((m) => f(m) >= alvo - 0.005);
    $("metas").innerHTML = metas.map((mt) => { const ini = cum; cum += mt.obj;
      const hoje = Math.max(0, Math.min(mt.obj, DATA.reservaInformada - ini)), noMes = Math.max(0, Math.min(mt.obj, acM(k) - ini));
      const q = quando(cum, acM), qr = quando(cum, rdM), noPrazo = mt.prazo ? acM(mt.prazo) : null, ok = noPrazo != null && noPrazo >= cum - 0.005;
      return `<div class="meta-it"><div class="row-between"><b>${esc(mt.nome)}</b><span class="muted small">prioridade ${mt.prio} · objetivo ${R0(mt.obj)}${mt.prazo ? " até " + lab(mt.prazo) : ""}</span></div>
        <div class="bar"><span style="width:${mt.obj ? (noMes / mt.obj) * 100 : 0}%;background:${noMes >= mt.obj - 0.01 ? css("--good") : cor[0]}"></span></div>
        <div class="meta-l"><span>Hoje: <b>${R0(hoje)}</b> (${P(mt.obj ? hoje / mt.obj : 0)})</span><span>Fim de ${lab(k)}: <b>${R0(noMes)}</b> (${P(mt.obj ? noMes / mt.obj : 0)})</span>
        <span>Atinge em <b>${q ? lab(q) : "depois de " + lab(ms[ms.length - 1])}</b>${qr && qr !== q ? ` <span class="muted">(com rendimento estimado: ${lab(qr)})</span>` : ""}</span>
        ${mt.prazo ? `<span class="${ok ? "icon-ok" : "icon-warn"}">${ok ? "No prazo" : "Faltariam " + R0(cum - (noPrazo || 0)) + " no prazo"}</span>` : ""}</div></div>`; }).join("") || `<p class="muted">Cadastre metas na aba Metas e Cortes da planilha.</p>`;
    const alvo = []; cum = 0; metas.forEach((mt) => { cum += mt.obj; alvo.push([mt.nome, cum]); });
    draw("cMetas", "line", ms.map(lab), [
      ds("Reserva prevista (plano)", ms.map(acM), cor[0], { fill: false, borderWidth: 2.5, pointRadius: 2, tension: 0.2, datalabels: { display: (c) => c.dataIndex % 4 === 3 || c.dataIndex === ms.length - 1, align: "top" } }),
      ds("Com rendimento estimado de " + P(DATA.mc.rend) + " a.a. (projeção)", ms.map(rdM), cor[2], { fill: false, borderWidth: 2, borderDash: [6, 4], pointRadius: 0, tension: 0.2, datalabels: { display: false } })]
      .concat(alvo.map(([n, v], z) => ds("Meta: " + n.replace(/\s*\(.*\)/, ""), ms.map(() => v), css("--muted"), { fill: false, borderWidth: 1.5, borderDash: [2, 3], pointRadius: 0, datalabels: { display: false } }))),
      baseOpts(false, (v) => R0(v), { layout: { padding: { right: 44, top: 8 } } }));
    $("rendNota").textContent = `A linha tracejada verde é só uma projeção com ${P(DATA.mc.rend)} ao ano (ajuste na aba Metas e Cortes). O rendimento real depende da economia e não entra no plano.`;
    // assinaturas
    const as = DATA.mc.ass, tot = as.reduce((s, a) => s + a.valor, 0), ano = as.reduce((s, a) => s + a.ano, 0);
    const eco = as.filter((a) => a.manter === "Não").reduce((s, a) => s + a.ano, 0), aval = as.filter((a) => a.manter === "Avaliar").reduce((s, a) => s + a.ano, 0);
    const off = as.filter((a) => assOff(a.nome)), ecoM = off.reduce((s, a) => s + a.valor, 0), ecoA = off.reduce((s, a) => s + a.ano, 0);
    $("assResumo").textContent = `${as.length} assinaturas · ${R(tot)}/mês · ${R0(ano)} nos próximos 12 meses`;
    $("assEco").innerHTML = off.length ? `<div class="eco"><div><div class="l">Se cancelar ${off.length === 1 ? "esta" : "estas " + off.length}</div><div class="v pos">− ${R(ecoM)}/mês</div><div class="s">${R(ecoA)} nos próximos 12 meses · assinaturas passam de ${R(tot)} para <b>${R(tot - ecoM)}</b>/mês</div></div><button type="button" class="btn ghost sm" id="assReset">Manter todas de novo</button></div>` : "";
    $("tAss").innerHTML = `<tr><th>Assinatura</th><th>Paga com</th><th>Por mês</th><th>Próximos 12 meses</th><th>Até</th><th>Na planilha</th><th>Manter?</th></tr>` +
      as.slice().sort((a, b) => b.ano - a.ano).map((a) => { const o = assOff(a.nome); return `<tr class="${o ? "fora riscado" : ""}"><td>${esc(a.nome)}</td><td>${esc(a.paga)}</td><td>${R(a.valor)}</td><td>${R(a.ano)}</td><td>${/^\d{4}-\d\d$/.test(a.ate) ? lab(a.ate) : esc(a.ate)}</td><td><span class="tag ${a.manter === "Não" ? "bad" : a.manter === "Sim" ? "good" : ""}">${esc(a.manter)}</span></td><td><button type="button" class="sw ${o ? "" : "on"}" data-ass="${esc(a.nome)}" aria-pressed="${!o}"><i></i>${o ? "Cancelar" : "Manter"}</button></td></tr>`; }).join("") +
      `<tr class="total"><td>Total</td><td></td><td>${off.length ? `<span class="muted">${R(tot)}</span> → ${R(tot - ecoM)}` : R(tot)}</td><td>${off.length ? `<span class="muted">${R(ano)}</span> → ${R(ano - ecoA)}` : R(ano)}</td><td></td><td></td><td></td></tr>`;
    $("tAss").querySelectorAll("[data-ass]").forEach((b) => b.addEventListener("click", () => { assToggle(b.dataset.ass); renderMetas(); setTimeout(rotulaTabelas, 0); }));
    if ($("assReset")) $("assReset").addEventListener("click", () => { ASS_OFF = new Set(); assSave(); renderMetas(); setTimeout(rotulaTabelas, 0); });
    $("assAviso").innerHTML = (eco ? `<span class="icon-ok">Cancelando as marcadas "Não" você economiza ${R0(eco)} em 12 meses.</span><br>` : "") +
      (aval ? `<span class="icon-warn">${R0(aval)} em 12 meses estão marcadas "Avaliar". Marque Sim ou Não na aba Metas e Cortes.</span>` : "");
    // recorrentes
    const fim = REF.mes, ini = addM(fim, -2), g = {};
    DATA.lanc.filter((l) => semVR(l) && l.data.slice(0, 7) >= ini && l.data.slice(0, 7) <= fim).forEach((l) => { const n = l.desc.toUpperCase().replace(/\s+/g, " ").trim();
      const o = g[n] || (g[n] = { n: l.desc, cat: l.cat, qt: 0, tot: 0 }); o.qt++; o.tot += l.meuTotal; });
    const lst = Object.values(g).sort((a, b) => b.tot - a.tot).slice(0, 12);
    $("recResumo").textContent = `compras de ${lab(ini)} a ${lab(fim)} (aba Lançamentos, sem VR)`;
    $("tRec").innerHTML = `<tr><th>Onde</th><th>Categoria</th><th>Vezes</th><th>Total</th><th>Média por compra</th><th>Por mês</th></tr>` +
      (lst.map((o) => `<tr><td>${esc(o.n)}</td><td>${esc(o.cat)}</td><td>${o.qt}</td><td>${R(o.tot)}</td><td>${R(o.tot / o.qt)}</td><td>${R(o.tot / Math.max(1, new Set(DATA.lanc.filter((l) => l.desc.toUpperCase().replace(/\s+/g, " ").trim() === o.n.toUpperCase().replace(/\s+/g, " ").trim()).map((l) => l.data.slice(0, 7))).size))}</td></tr>`).join("") || `<tr><td colspan="6" class="muted">Sem lançamentos no período.</td></tr>`);
    const freq = lst.filter((o) => o.qt >= 3);
    $("recAviso2").innerHTML = freq.length ? `<span class="icon-warn">Compras frequentes: ${freq.map((o) => esc(o.n) + " (" + o.qt + "x, " + R0(o.tot) + ")").join(" · ")}. Pequenos gastos repetidos somam rápido.</span>` : "";
  }

  // ---------- simulação: novas compras e dívidas ----------
  const TIPOS = { compra: "Compra (cartão, Pix ou débito — à vista ou parcelada)", div: "Nova dívida (boleto, empréstimo, financiamento)" };
  const PIX = "Pix/débito";
  let SIM = null, SIM_DEF = null;
  const cent = (x) => Math.round(x * 100) / 100;
  function simLoad() {
    let s = null; try { s = JSON.parse(localStorage.getItem("sim-v2")); } catch (_) {}
    if (!s || !Array.isArray(s.itens)) s = { itens: [], modo: "orc" };
    s.itens.forEach((it) => { if (it.t === "parc" || it.t === "vista") { if (it.t === "vista") it.n = 1; it.t = "compra"; }
      if (!DATA.meses.includes(it.mes)) it.mes = SIM_DEF; if (it.cartao && it.cartao !== PIX && !DATA.cartoes.includes(it.cartao)) it.cartao = DATA.cartoes[0]; });
    return s;
  }
  const simSave = () => { try { localStorage.setItem("sim-v2", JSON.stringify(SIM)); } catch (_) {} };
  function novoItem(t) {
    const base = { id: Date.now() + "" + Math.random().toString(36).slice(2, 6), t, desc: "", valor: 0, pct: 0 };
    if (t === "compra") return Object.assign(base, { n: 1, cartao: DATA.cartoes[0], mes: SIM_DEF });
    return Object.assign(base, { n: 12, mes: DATA.meses.includes(addM(SIM_DEF, 1)) ? addM(SIM_DEF, 1) : SIM_DEF });
  }
  // cobranças do item: row = mês do plano (teto) que a parcela ocupa; pag = salário que paga
  function cobrancas(it) {
    const n = Math.max(1, Math.round(+it.n || 1)), out = [], pix = it.cartao === PIX;
    const parcela = it.t === "div" ? +it.valor || 0 : (+it.valor || 0) / n;
    const meu = parcela * (1 - Math.min(100, Math.max(0, +it.pct || 0)) / 100);
    for (let i = 0; i < n; i++) {
      if (it.t === "div") out.push({ row: addM(it.mes, i - 1), pag: addM(it.mes, i), valor: parcela, meu, div: true });
      else out.push({ row: pix ? addM(it.mes, i - 1) : addM(it.mes, i), pag: pix ? addM(it.mes, i) : addM(it.mes, i + 1), valor: parcela, meu, div: false });
    }
    return { n, parcela, meu, lista: out };
  }
  function simCalc() {
    const minMes = DATA.minMes || 0, pool = {}, divs = {};
    SIM.itens.forEach((it) => cobrancas(it).lista.forEach((c) => { const o = c.div ? divs : pool; o[c.row] = (o[c.row] || 0) + c.meu; }));
    const meses = DATA.meses.filter((k) => DATA.plano[k] && k >= DATA.planIni);
    const cal = mesAtual();
    let acc = 0;
    const linhas = meses.map((k) => {
      const p = DATA.plano[k], F = addM(k, 1), j = idx(F), pr = DATA.proj;
      const nPool = cent(pool[k] || 0), nDiv = cent(divs[k] || 0), nova = cent(nPool + nDiv);
      const ja = k === cal ? gastoInter(k).v : p.ja || 0;
      const shares = (p.reserva || 0) - ((p.espaco || 0) - (p.teto || 0));
      const aindaA = (p.teto || 0) - ja;
      let aindaD, resD, tetoD;
      if (SIM.modo === "res") { tetoD = p.teto || 0; aindaD = aindaA; resD = (p.reserva || 0) - nova; }
      else {
        const espD = (p.espaco || 0) - nDiv;
        tetoD = p.teto > 0 ? Math.max(0, Math.min(espD, p.teto, Math.max(minMes, p.teto - nDiv))) : 0;
        aindaD = tetoD - ja - nPool;
        const estouro = Math.max(0, nPool - Math.max(0, tetoD - ja));
        resD = espD - tetoD + shares - estouro;
      }
      acc += resD - (p.reserva || 0);
      const ent = j >= 0 ? pr.entradas[j] : 0;
      const obrig = j >= 0 ? Math.abs(pr.fixos[j]) + Math.abs(pr.parcelas[j]) + Math.abs(pr.variaveis[j]) - pr.reemb[j] : 0;
      return { k, F, nova, nPool, nDiv, ent, obrig, compA: ent ? obrig / ent : null, compD: ent ? (obrig + nova) / ent : null,
        sobraA: j >= 0 ? pr.sobra[j] : null, sobraD: j >= 0 ? pr.sobra[j] - nova : null,
        tetoA: cent(aindaA), tetoD: cent(aindaD), resA: p.reserva || 0, resD: cent(resD), accA: p.reservaAc || 0, accD: cent((p.reservaAc || 0) + acc), meta: p.meta || 0, p };
    });
    return { linhas };
  }
  function renderSimEditor() {
    const mesOpts = (sel) => DATA.meses.filter((k) => k >= DATA.planIni || k === sel).map((k) => `<option value="${k}" ${k === sel ? "selected" : ""}>${lab(k)}</option>`).join("");
    const carOpts = (sel) => DATA.cartoes.concat([PIX]).map((c) => `<option ${c === sel ? "selected" : ""}>${esc(c)}</option>`).join("");
    $("sItens").innerHTML = SIM.itens.length ? SIM.itens.map((it, n) => `
      <div class="simit" data-i="${n}">
        <div class="simit-h"><b>${n + 1}. ${TIPOS[it.t]}</b><button type="button" class="btn ghost sm" data-rm="${n}" aria-label="Remover">Remover</button></div>
        <div class="simit-g">
          <div class="f wide"><label>Descrição</label><input data-k="desc" type="text" value="${esc(it.desc)}" placeholder="${it.t === "div" ? "ex.: empréstimo, financiamento" : "ex.: tênis, passagem"}"></div>
          <div class="f"><label>${it.t === "div" ? "Valor da parcela (R$)" : "Valor total da compra (R$)"}</label><input data-k="valor" type="number" inputmode="decimal" step="0.01" min="0" value="${it.valor || ""}"></div>
          <div class="f"><label>${it.t === "div" ? "Nº de parcelas" : "Nº de vezes (1 = à vista)"}</label><input data-k="n" type="number" inputmode="numeric" step="1" min="1" max="60" value="${it.n}"></div>
          ${it.t === "div" ? "" : `<div class="f"><label>Forma de pagamento</label><select data-k="cartao">${carOpts(it.cartao)}</select></div>`}
          <div class="f"><label>${it.t === "div" ? "1ª parcela paga em" : "Mês da compra"}</label><select data-k="mes">${mesOpts(it.mes)}</select></div>
          <div class="f"><label>Outra pessoa me paga (%)</label><input data-k="pct" type="number" inputmode="numeric" step="1" min="0" max="100" value="${it.pct || ""}" placeholder="0"></div>
        </div>
        <p class="muted small" id="sInfo${n}"></p>
      </div>`).join("") : `<p class="muted">Nenhuma simulação ainda. Escolha abaixo o que você quer testar.</p>`;
    $("sModo").value = SIM.modo;
    $("sItens").querySelectorAll(".simit").forEach((box) => {
      const it = SIM.itens[+box.dataset.i];
      box.querySelectorAll("[data-k]").forEach((el) => el.addEventListener(el.tagName === "SELECT" ? "change" : "input", () => {
        const k = el.dataset.k; it[k] = k === "desc" || k === "cartao" || k === "mes" ? el.value : +el.value || 0; simSave(); renderSimResultado(); }));
      box.querySelector("[data-rm]").addEventListener("click", () => { SIM.itens.splice(+box.dataset.i, 1); simSave(); renderSimEditor(); renderSimResultado(); });
    });
  }
  const ab = (a, b, f = R0) => (Math.abs(a - b) < 0.005 ? f(b) : `<span class="muted">${f(a)}</span> → <b>${f(b)}</b>`);
  function renderSimResultado() {
    const cor = pal(), minMes = DATA.minMes || 0, card0 = DATA.cartoes[0] || "Inter";
    // informações por item + regras
    const av = [];
    SIM.itens.forEach((it, n) => {
      const c = cobrancas(it), nome = it.desc ? "“" + esc(it.desc) + "”" : "Item " + (n + 1), el = $("sInfo" + n);
      if (!it.valor) { if (el) el.textContent = "Preencha o valor."; return; }
      const ini = c.lista[0].pag, fim = c.lista[c.lista.length - 1].pag;
      const r0 = c.lista[0].row, r1 = c.lista[c.lista.length - 1].row;
      if (el) el.innerHTML = (it.t === "div" ? `${c.n}x de ${R(c.parcela)} (total ${R(c.parcela * c.n)}) · pago com o salário de ${lab(ini)}${c.n > 1 ? " até " + lab(fim) : ""}`
        : (c.n > 1 ? `${c.n}x de ${R(c.parcela)}` : `${R(c.parcela)} à vista`) + ` no ${esc(it.cartao)} · usa o dinheiro do salário de ${lab(addM(r0, 1))}${c.n > 1 ? " até " + lab(addM(r1, 1)) : ""}`) +
        (it.pct ? ` · meu custo ${R(c.meu)}/mês (${esc(it.pct)}% volta pra mim)` : "");
      const p = DATA.plano[it.mes] || {};
      if (it.t !== "div" && it.cartao !== card0 && it.cartao !== PIX) av.push(`<span class="icon-warn">${nome}: a combinação era parar de usar ${esc(it.cartao)} e concentrar no ${esc(card0)}.</span>`);
      if (fim > DATA.meses[DATA.meses.length - 1]) av.push(`<span class="icon-warn">${nome}: as últimas parcelas passam de ${lab(DATA.meses[DATA.meses.length - 1])}, fim da projeção da planilha.</span>`);
    });
    const temValor = SIM.itens.some((it) => +it.valor > 0);
    $("sRes").hidden = !temValor;
    if (!temValor) return;
    const { linhas } = simCalc();
    const afet = linhas.filter((l) => l.nova > 0), primeiro = afet.length ? afet[0].k : linhas[0].k;
    const vis = linhas.filter((l) => l.k >= primeiro);
    const ult = linhas[linhas.length - 1];
    const totMeu = SIM.itens.reduce((s, it) => { const c = cobrancas(it); return s + c.meu * c.n; }, 0);
    const totCheio = SIM.itens.reduce((s, it) => { const c = cobrancas(it); return s + c.parcela * c.n; }, 0);
    const maxN = afet.reduce((a, l) => (l.nova > a.nova ? l : a), { nova: 0 });
    const maxC = afet.reduce((a, l) => (l.compD != null && l.compD > (a.compD || 0) ? l : a), { compD: 0 });
    const cortOrc = vis.reduce((s, l) => s + (l.tetoA - l.tetoD), 0), difRes = ult.accD - ult.accA;
    $("sKpis").innerHTML = [
      ["Custo total para mim", R(totMeu), totCheio - totMeu > 0.005 ? "valor cheio " + R(totCheio) : afet.length + " mês(es) com cobrança"],
      ["Maior cobrança no mês", R(maxN.nova), maxN.k ? "compras de " + lab(maxN.k) + " (salário de " + lab(maxN.F) + ")" : ""],
      ["Teto dos meses", "−" + R(cortOrc), SIM.modo === "res" ? "teto mantido" : "a menos para gastar no período"],
      ["Reserva em " + lab(ult.k), R(ult.accD), (difRes < 0 ? "−" : "+") + R(Math.abs(difRes)) + " vs. sem a simulação", ult.accD < 0 ? "neg" : difRes < -0.005 ? "" : "pos"],
      ["Entrada comprometida", maxC.k ? P(maxC.compD) : "—", maxC.k ? "no pior mês (" + lab(maxC.F) + "), antes " + P(maxC.compA) : ""],
    ].map(([l, v, s, c]) => `<div class="kpi"><div class="l">${l}</div><div class="v ${c || ""}">${v}</div><div class="s">${s}</div></div>`).join("");
    // avisos do planejamento
    const neg = vis.filter((l) => l.resD < -0.005), noMin = vis.filter((l) => l.nPool > 0 && SIM.modo === "orc" && l.tetoD < -0.005);
    const abaixoMeta = vis.filter((l) => l.nova > 0 && l.resD >= -0.005 && l.resD < l.meta - 0.01 && l.resA >= l.meta - 0.01);
    const resNeg = vis.find((l) => l.accD < -0.005);
    if (neg.length) av.unshift(`<span class="icon-bad">Em ${neg.length} mês(es) a conta não fecha (${neg.map((l) => lab(l.k) + " " + R0(l.resD)).join(", ")}): a diferença teria que sair da reserva.</span>`);
    if (resNeg) av.unshift(`<span class="icon-bad">A reserva ficaria negativa a partir de ${lab(resNeg.k)} (${R(resNeg.accD)}).</span>`);
    if (noMin.length) av.push(`<span class="icon-bad">Passa do teto em ${noMin.map((l) => lab(l.k) + " (falta " + R0(-l.tetoD) + ")").join(", ")}: essa diferença sairia da reserva.</span>`);
    if (abaixoMeta.length) av.push(`<span class="icon-warn">A reserva do mês fica abaixo da meta de ${P(DATA.pct)} em ${abaixoMeta.map((l) => lab(l.k)).join(", ")}.</span>`);
    if (!av.length) av.push(`<span class="icon-ok">Cabe no planejamento: ${SIM.modo === "res" ? "o teto continua o mesmo e a reserva absorve a compra" : "a compra cabe no teto dos meses e a meta da reserva continua batida"}.</span>`);
    av.push(`<span class="muted small">Se decidir fazer, lance na aba Lançamentos da planilha para entrar no planejamento de verdade.</span>`);
    $("sAviso").innerHTML = av.join("<br>");
    // gráficos
    const L = vis.map((l) => lab(l.k));
    const ultAf = afet.length ? afet[afet.length - 1].k : primeiro, vl = vis.filter((l) => l.k <= addM(ultAf, 2)), bad = css("--bad");
    const cl = draw("cSimLivre", "bar", vl.map((l) => lab(l.k)), [ds("Nova parcela", vl.map((l) => l.nova), cor[1]), ds("Ainda posso gastar", vl.map((l) => Math.max(0, l.tetoD)), cor[0]),
      ds("Vai para a reserva", vl.map((l) => Math.max(0, l.resD)), cor[2])].concat(vl.some((l) => l.resD < -0.005) ? [ds("Falta (sai da reserva)", vl.map((l) => Math.min(0, l.resD)), bad)] : []), baseOpts(true, (v) => R0(v)));
    if (vl.length > 8) { cl.options.plugins.datalabels.display = (c) => c.datasetIndex === 0 && !!c.raw && (c.dataIndex === 0 || vl[c.dataIndex].nova !== vl[c.dataIndex - 1].nova); cl.update(); }
    draw("cSimRes", "line", L, [
      ds("Sem a simulação", vis.map((l) => l.accA), css("--muted"), { fill: false, tension: 0.2, borderWidth: 2, borderDash: [6, 4], pointRadius: 0, datalabels: { display: false } }),
      ds("Com a simulação", vis.map((l) => l.accD), cor[0], { fill: false, tension: 0.2, borderWidth: 2, pointRadius: 3, pointBackgroundColor: cor[0] })],
      baseOpts(false, (v) => R0(v), {}));
    const cr = charts.cSimRes; cr.options.plugins.datalabels.display = (c) => c.datasetIndex === 1 && (c.dataIndex === vis.length - 1 || c.dataIndex % 3 === 0); cr.update();
    // tabela
    $("tSim").innerHTML = `<tr><th>Dinheiro do salário de</th><th>Nova parcela</th><th>Entrada comprometida</th><th>Ainda posso gastar</th><th>Vai para a reserva</th><th>Reserva acumulada</th><th>Sobra do salário seguinte</th></tr>` +
      vis.map((l) => `<tr class="${l.nova > 0 ? "" : "fora"}"><td>${lab(l.F)}</td><td>${l.nova ? R(l.nova) : "—"}</td><td>${l.compA == null ? "—" : ab(l.compA, l.compD, P)}</td>
        <td class="${l.tetoD < 0 ? "neg" : ""}">${ab(l.tetoA, l.tetoD)}</td><td class="${l.resD < 0 ? "neg" : ""}">${ab(l.resA, l.resD)}</td><td class="${l.accD < 0 ? "neg" : ""}">${ab(l.accA, l.accD)}</td><td class="${l.sobraD != null && l.sobraD < 0 ? "neg" : ""}">${l.sobraA == null ? "—" : ab(l.sobraA, l.sobraD)}</td></tr>`).join("");
  }
  function renderSim() { if (!SIM) SIM = simLoad(); renderSimEditor(); renderSimResultado(); }
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") fecharSimRapida(); });
  document.querySelectorAll("[data-add]").forEach((b) => b.addEventListener("click", () => { SIM.itens.push(novoItem(b.dataset.add)); simSave(); renderSimEditor(); renderSimResultado(); }));
  $("sModo").addEventListener("change", (e) => { SIM.modo = e.target.value; simSave(); renderSimResultado(); });
  $("sLimpar").addEventListener("click", () => { SIM.itens = []; simSave(); renderSimEditor(); renderSimResultado(); });

  // ---------- senha de acesso (proteção simples, só na tela) ----------
  async function hashSenha(t) { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("meu-financeiro:" + t)); return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join(""); }
  function portao() {
    const H = (CFG.PASS_HASH || "").trim().toLowerCase(); let lembrado = null; try { lembrado = localStorage.getItem("dash-acesso"); } catch (_) {}
    if (H && lembrado === H) { document.body.classList.remove("travado"); load(); return; }
    const ov = document.createElement("div"); ov.className = "lock"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
    ov.innerHTML = H ? `<form class="lock-box" id="lkF"><img class="logo-lk" src="img/logo.png" alt=""><h2>Controle Financeiro</h2><p class="muted small">Digite a senha para ver o dashboard.</p>
        <input type="password" id="lkS" autocomplete="current-password" placeholder="Senha" required><label class="lk-rem"><input type="checkbox" id="lkR" checked> Lembrar neste aparelho</label>
        <button class="btn" type="submit">Entrar</button><p class="lk-err" id="lkE" role="alert"></p></form>`
      : `<form class="lock-box" id="lkF"><div class="logo-lk" aria-hidden="true">🔑</div><h2>Criar senha</h2><p class="muted small">Ainda não há senha. Digite a senha que você quer usar e copie a linha gerada para o arquivo config.js no GitHub.</p>
        <input type="password" id="lkS" autocomplete="new-password" placeholder="Nova senha" required minlength="4"><button class="btn" type="submit">Gerar</button>
        <textarea id="lkO" readonly rows="3" hidden></textarea><p class="muted small" id="lkI" hidden>Cole essa linha dentro de config.js (no lugar de PASS_HASH: "") e salve. A senha em si não fica salva em lugar nenhum.</p></form>`;
    document.body.appendChild(ov); document.body.classList.add("travado"); setTimeout(() => $("lkS").focus(), 50);
    $("lkF").addEventListener("submit", async (e) => { e.preventDefault(); const v = $("lkS").value; if (!v) return;
      if (!window.crypto || !crypto.subtle) { (H ? $("lkE") : $("lkI")).textContent = "Abra pelo endereço https do GitHub Pages."; return; }
      const h = await hashSenha(v);
      if (!H) { $("lkO").hidden = false; $("lkI").hidden = false; $("lkO").value = 'PASS_HASH: "' + h + '",'; $("lkO").select(); return; }
      if (h === H) { try { $("lkR").checked ? localStorage.setItem("dash-acesso", h) : localStorage.removeItem("dash-acesso"); } catch (_) {} ov.remove(); document.body.classList.remove("travado"); load(); }
      else { $("lkE").textContent = "Senha incorreta."; $("lkS").value = ""; $("lkS").focus(); } });
  }
  $("sair").addEventListener("click", () => { try { localStorage.removeItem("dash-acesso"); localStorage.removeItem("dash-data"); } catch (_) {} location.reload(); });
  if (!(CFG.PASS_HASH || "").trim()) $("sair").hidden = true;
  portao();
})();

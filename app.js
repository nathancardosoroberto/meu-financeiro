/* Meu Financeiro — dashboard que lê a planilha Planejamento_Financeiro (Google Planilhas publicado como .xlsx) */
(function () {
  "use strict";
  Chart.register(ChartDataLabels);
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

  // 5º dia útil (sábado conta; domingo e feriado não)
  function quintoDiaUtil(y, m0, fer) {
    let n = 0; const d = new Date(y, m0, 1, 12);
    while (true) { if (d.getDay() !== 0 && !fer.has(ymd(d))) { n++; if (n === 5) return new Date(d); } d.setDate(d.getDate() + 1); }
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
    D.reservaInformada = num(cell(wb, C, "B5")); D.pct = num(cell(wb, C, "B6")); D.meta = D.pct > 1 ? D.pct : 0; D.planIni = ym(cell(wb, C, "B8")) || D.ini;
    D.cartoes = []; D.venc = {};
    for (let r = 11; r <= 15; r++) { const n = txt(cell(wb, C, "A" + r)); if (n && n !== "-") { D.cartoes.push(n); D.venc[n] = num(cell(wb, C, "B" + r)) || null; } }
    const PG = "Plano de Gastos";
    D.minSemana = num(cell(wb, PG, "B3")); D.fechamento = num(cell(wb, PG, "B14")) || null; D.minMes = num(cell(wb, PG, "B4")); D.saldo = num(cell(wb, PG, "B16"));
    D.feriados = [];
    if (wb.Sheets["Feriados"]) rows(wb, "Feriados").slice(1).forEach((r) => { const d = toDate(r[0]); if (d) D.feriados.push(ymd(d)); });

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
    const K = ["mes","fatura","espaco","teto","semana","vista","parc","vezes","parcMax","simult","fora","vr","total","faturaInter","limite","reserva","reservaAc","ja","ainda","meta","livre","livreResta"];
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
          ini: ym(r[12]), fim: ym(r[13]), valor: num(r[14]), r1: num(r[15]), q1: txt(r[9]), r2: num(r[16]), q2: txt(r[11]), custo: num(r[17]) }); } }
    // Categorias
    const rs = rows(wb, "Resumo mensal"); const rh = findRow(rs, (t) => t === "Mês");
    const rmeses = rs[rh].slice(1, 13).map(ym); const ci = findRow(rs, (t) => t.startsWith("MEU CUSTO REAL POR CATEGORIA"));
    D.cat = { meses: rmeses, itens: {} };
    for (let i = ci + 1; i < rs.length; i++) { const t = txt(rs[i] && rs[i][0]); if (!t) continue; if (t === "TOTAL") break; D.cat.itens[t] = { vals: rs[i].slice(1, 13), media: num(rs[i][13]) }; }
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
    setStatus("Atualizado de " + origem + " em " + stamp);
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
    ["geral", "dividas", "sim"].forEach((t) => ($("tab-" + t).hidden = t !== TAB));
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
    const def = clamp(REF.mes);
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
    const simDef = DATA.plano[def] && DATA.plano[def].teto > 0 ? def : clamp(addM(def, 1));
    opt($("sMes"), ms, simDef);
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
  $("gastoManual").addEventListener("input", (e) => { try { const v = e.target.value; v === "" ? localStorage.removeItem("gasto-" + F.mes) : localStorage.setItem("gasto-" + F.mes, v); } catch (_) {} renderGeral(true); });

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
  function semanasRestantes(k) {
    if (!REF) return 0;
    if (k === REF.mes) { const hoje = new Date(); const d = Math.max(0, Math.ceil((REF.proximo - new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 12)) / 864e5)); return d / 7; }
    if (k < REF.mes) return 0;
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
    const push = (pessoa, valor, item, origem, quando) => { if (pessoa && pessoa !== "-" && valor > 0.004) out.push({ pessoa, valor, item, origem, quando }); };
    DATA.fixos.filter((f) => ativo(f, k)).forEach((f) => { const q = diaVenc(k, f.dia); push(f.q1, f.r1, f.desc, "Conta fixa", q); push(f.q2, f.r2, f.desc, "Conta fixa", q); });
    DATA.dividas.filter((d) => ativo(d, k)).forEach((d) => { const q = diaVenc(k, DATA.venc[d.paga]); const inf = dividaInfo(d, k);
      const t = d.desc + (inf.atual ? " (" + inf.atual + ")" : ""); push(d.q1, d.r1, t, "Parcela " + d.paga, q); push(d.q2, d.r2, t, "Parcela " + d.paga, q); });
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
    const ink = css("--ink2"), line = css("--line");
    const o = Object.assign({
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 }, interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "top", labels: { color: ink, boxWidth: 12, boxHeight: 12 } },
        tooltip: { callbacks: { label: (c) => " " + c.dataset.label + ": " + fmtFn(c.raw) } },
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
    const ov = document.createElement("div"); ov.className = "zoomov"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
    ov.innerHTML = `<div class="zh"><h3>${esc(titulo)}</h3><button class="btn" type="button" id="zClose">Fechar ✕</button></div><div class="dica">Dica: gire o celular para ver o gráfico maior.</div><div class="zb"><div class="zc"><canvas id="zCanvas"></canvas></div></div>`;
    document.body.appendChild(ov); document.body.style.overflow = "hidden";
    $("zClose").addEventListener("click", fecharZoom);
    const horiz = raw.opts.indexAxis === "y", n = raw.labels.length, zc = ov.querySelector(".zc");
    if (horiz) zc.style.height = Math.max(ov.querySelector(".zb").clientHeight - 24, n * 30 + 80) + "px";
    else { const minW = n * 64; if (minW > zc.clientWidth) zc.style.width = minW + "px"; }
    const data = { labels: raw.labels.slice(), datasets: raw.datasets.map((d) => Object.assign({}, d, { data: d.data.slice() })) };
    const opts = Object.assign({}, raw.opts, { maintainAspectRatio: false, animation: false });
    opts.plugins = Object.assign({}, raw.opts.plugins, { legend: Object.assign({}, raw.opts.plugins.legend, { display: true }), datalabels: Object.assign({}, raw.opts.plugins.datalabels, { font: { size: 12, weight: "600" }, display: (c) => !!c.dataset.data[c.dataIndex] }) });
    zoomChart = new Chart($("zCanvas"), { type: raw.type, data, options: opts });
  }
  document.addEventListener("keydown", (e) => { if (e.key === "Escape") fecharZoom(); });
  const ds = (label, data, color, extra = {}) => Object.assign({ label, data, backgroundColor: color, borderColor: color, borderWidth: 0, borderRadius: 4, borderSkipped: "start", maxBarThickness: 46 }, extra);

  // ---------- render ----------
  function renderAll() { if (!DATA) return; resumoFiltros(); document.querySelectorAll(".mesSel").forEach((e) => (e.textContent = lab(F.mes)));
    if (TAB === "geral") renderGeral(); else if (TAB === "dividas") renderDividas(); else renderSim(); }

  function renderHoje() {
    const hoje = new Date(); const prox = addM(REF.mes, 1); const pp = DATA.plano[prox] || {};
    const dias = Math.max(0, Math.ceil((REF.proximo - new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 12)) / 864e5));
    const pr = DATA.plano[REF.mes] || {};
    const fechado = !pr.teto;
    $("hoje").innerHTML = `
      <div class="it"><div class="l">Hoje</div><div class="v">${fmtD(hoje)}</div></div>
      <div class="it"><div class="l">Mês de referência</div><div class="v">${labL(REF.mes)}</div></div>
      <div class="it"><div class="l">Saldo em conta</div><div class="v">${R(DATA.saldo)}</div></div>
      <div class="it"><div class="l">Próximo 5º dia útil (salário)</div><div class="v">${fmtD(REF.proximo)}</div><div class="muted small">${dias === 0 ? "hoje" : "em " + dias + " dia" + (dias > 1 ? "s" : "")}</div></div>
      <div class="it"><div class="l">Agora no crédito do ${esc(DATA.cartoes[0] || "Inter")}</div><div class="v">${R(formas(REF.mes).credSug)}</div><div class="muted small">${fechado ? "compras agora caem na próxima fatura, que já está no limite" : "sugerido"}</div></div>
      <div class="it"><div class="l">Agora em dinheiro (Pix/débito)</div><div class="v">${R(fechado ? DATA.saldo : formas(REF.mes).cashSug)}</div><div class="muted small">${fechado ? "saldo em conta" : "sugerido"}</div></div>
      <div class="it"><div class="l">${labL(prox).split(" ")[0]}: sugerido / livre</div><div class="v">${R0(pp.teto)} / ${R0(pp.livre)}</div><div class="muted small">o que não gastar vai para a reserva</div></div>`;
  }

  function formas(k) {
    const p = DATA.plano[k] || {}, prev = DATA.plano[addM(k, -1)] || {};
    const ja = gastoInter(k).v, fechado = !p.teto && k < DATA.planIni;
    const cashJa = DATA.lanc.filter((l) => l.data.slice(0, 7) === k && !DATA.cartoes.includes(l.paga) && l.paga !== "VR").reduce((s, l) => s + l.total, 0);
    const sobraMes = prev.reserva != null ? prev.reserva : 0; // o que sobra na conta neste mês depois de pagar contas e faturas
    return {
      fechado, ja, cashJa, sobraMes,
      credSug: fechado ? 0 : (p.teto || 0) - ja, credMax: fechado ? 0 : (p.espaco || 0) - ja,
      cashSug: Math.max(0, (p.fora || 0) + Math.max(0, sobraMes - (prev.meta || 0)) - cashJa),
      cashMax: Math.max(0, (p.fora || 0) + Math.max(0, sobraMes) - cashJa),
      fora: p.fora || 0, vr: p.vr || 0, fatura: p.fatura, meta: prev.meta || 0,
    };
  }
  function renderFormas(k) {
    const f = formas(k), p = DATA.plano[k] || {}, card = DATA.cartoes[0] || "Inter";
    const salPaga = lab(addM(k, 1)), mesL = lab(k);
    const cred = `<div class="card forma">
      <h2>No crédito do ${esc(card)}</h2>
      <p class="quando">Compras de ${mesL} caem na fatura de ${salPaga} e são pagas com o salário de ${salPaga} (5º dia útil). ${DATA.fechamento ? "Compre até o dia " + DATA.fechamento + " para cair nessa fatura." : "Atenção ao dia de fechamento: depois dele a compra vai para a fatura seguinte."}</p>
      <div class="muted small">Sugerido (já separa os ${P(DATA.pct)} da reserva)</div>
      <div class="big2 ${f.credSug < 0 ? "neg" : ""}">${R(f.credSug)}</div>
      <div class="linha"><span>Máximo sem estourar o mês de ${salPaga}</span><b>${R(f.credMax)}</b></div>
      <div class="linha"><span>À vista / parcelas novas</span><b>${R(p.vista)} / ${R(p.parc)} por mês (até ${p.vezes || "—"}x)</b></div>
      <div class="linha"><span>Já lançado no ${esc(card)} neste mês</span><b>${R(f.ja)}</b></div></div>`;
    const cash = `<div class="card forma cash">
      <h2>Em dinheiro (Pix / débito)</h2>
      <p class="quando">Sai da conta na hora, então usa o que sobrou do salário de ${mesL} depois das contas e faturas de ${mesL}${f.fora ? ", mais o 13º livre/férias que entram neste mês" : ""}.</p>
      <div class="muted small">Sugerido (sem mexer nos ${P(DATA.pct)} da reserva)</div>
      <div class="big2">${R(f.cashSug)}</div>
      <div class="linha"><span>Máximo (usando o que iria para a reserva)</span><b>${R(f.cashMax)}</b></div>
      <div class="linha"><span>Sobra na conta depois de pagar tudo</span><b class="${f.sobraMes < 0 ? "neg" : ""}">${R(f.sobraMes)}</b></div>
      ${f.fora ? `<div class="linha"><span>13º livre / férias do mês</span><b>${R(f.fora)}</b></div>` : ""}
      <div class="linha"><span>VR / Flash (supermercado)</span><b>${R(f.vr)}</b></div>
      ${f.sobraMes < 0 ? `<p class="icon-bad">A conta fecha negativa em ${R(f.sobraMes)} neste mês: não use Pix/débito e cubra a diferença com a reserva.</p>` : ""}</div>`;
    $("formas").innerHTML = cred + cash;
  }

  function renderGeral(onlyHero) {
    renderHoje();
    const k = F.mes, i = idx(k), p = DATA.plano[k] || {}, pr = DATA.proj, pct = F.unid === "pct";
    const real = DATA.tipo[i] === "Realizado";
    $("hTag").innerHTML = real ? '<span class="tag">realizado</span>' : '<span class="tag prev">previsto</span>';
    $("hMes").textContent = lab(k);
    const fechado = !p.teto && k < DATA.planIni;
    const gi = gastoInter(k), teto = p.teto || 0, rest = fechado ? DATA.saldo : teto - gi.v, sem = semanasRestantes(k);
    $("hRestante").textContent = R(rest); $("hRestante").classList.toggle("neg", rest < 0);
    const uso = teto ? gi.v / teto : 0; $("hBar").style.width = (fechado ? 0 : Math.min(100, uso * 100)) + "%"; $("hBar").classList.toggle("over", uso > 1);
    $("hResumo").innerHTML = fechado
      ? `Mês fechado: tudo já foi pago e a renda está zerada (saldo em conta ${R(DATA.saldo)}). O teto volta no 5º dia útil (${fmtD(REF.proximo)}): ${lab(addM(k, 1))} começa com ${R((DATA.plano[addM(k, 1)] || {}).teto)}.`
      : `<b>Livre no mês: ${R(p.livre)}</b> (inclui a meta de reserva de ${R(p.meta)}, ${P(DATA.pct)} do salário líquido). Ainda livre: ${R((p.livre || 0) - gi.v)} — o que você não gastar vai para a reserva.<br>` +
        `Sugerido ${R(teto)} (livre − meta) · já gasto ${R(gi.v)} (${P(uso)} do sugerido, ${gi.fonte})<br>` + (sem > 0 ? `Dá ${R(rest / sem)} por semana nas ${sem.toFixed(1).replace(".", ",")} semanas que faltam · ` : "") +
        `à vista até ${R(p.vista)} · parcelas novas até ${R(p.parc)}/mês em até ${p.vezes || "—"}x (máx. ${p.simult || "—"} ao mesmo tempo)`;
    let man = null; try { man = localStorage.getItem("gasto-" + k); } catch (_) {}
    if (!onlyHero) $("gastoManual").value = man ?? "";
    const ent = pr.entradas[i] || 0;
    renderFormas(k);
    const kp = [
      ["Fora do Inter disponível", R(p.fora), "13º livre + férias (Pix/débito)"],
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
    $("tEstimativa").innerHTML = `<tr><th>${lab(k)} ${real ? '<span class="tag">realizado</span>' : '<span class="tag prev">previsto</span>'}</th><th>Valor</th><th>% das entradas</th></tr>` + lin +
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
    draw("cGastar", "bar", L, [ds("Inter à vista", conv(pv("vista")), hl(cor[0])), ds("Inter parcelado (limite)", conv(pv("parc")), hl(cor[1])), ds("Fora do Inter", conv(pv("fora")), hl(cor[2])), ds("VR / Flash", conv(pv("vr")), hl(cor[3]))], baseOpts(true, fmt));
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

  function renderDividas() {
    const k = F.mes, i = idx(k), ent = DATA.proj.entradas[i] || 0, cor = pal();
    // a receber
    let rc = receber(k); if (F.cartao) rc = rc.filter((r) => r.origem.includes(F.cartao) || (F.cartao.startsWith("Conta") && r.origem === "Conta fixa"));
    const porPessoa = {}; rc.forEach((r) => (porPessoa[r.pessoa] = (porPessoa[r.pessoa] || 0) + r.valor));
    $("rResumo").textContent = Object.entries(porPessoa).map(([p, v]) => p + ": " + R(v)).join(" · ") || "nada a receber";
    rc.sort((a, b) => a.pessoa.localeCompare(b.pessoa) || (a.quando || 0) - (b.quando || 0));
    $("tReceber").innerHTML = `<tr><th>Item</th><th>Pessoa</th><th>Origem</th><th>Valor</th><th>Receber até</th></tr>` +
      rc.map((r) => `<tr><td>${esc(r.item)}</td><td>${esc(r.pessoa)}</td><td>${esc(r.origem)}</td><td>${R(r.valor)}</td><td>${r.quando ? fmtD(r.quando) : "vencimento (dia não informado)"}</td></tr>`).join("") +
      Object.entries(porPessoa).map(([p, v]) => `<tr class="total"><td>Total ${esc(p)}</td><td></td><td></td><td>${R(v)}</td><td></td></tr>`).join("");
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
    // lançamentos
    let lc = DATA.lanc.filter((l) => l.ini && l.ini <= k && l.fim >= k && passaForma(l.paga));
    if (F.item && F.item.startsWith("L|")) lc = lc.filter((l) => l.desc === F.item.slice(2));
    const tl = lc.reduce((s, l) => s + l.valor, 0);
    $("lResumo").textContent = lc.length ? `${lc.length} lançamento(s) cobrados em ${lab(k)} · ${R(tl)}` : "Nenhum lançamento cobrado neste mês. Adicione na aba Lançamentos da planilha.";
    $("tLanc").innerHTML = `<tr><th>Descrição</th><th>Data</th><th>Forma</th><th>Valor total</th><th>Parcelas</th><th>Cobrado no mês</th><th>Meu custo</th></tr>` +
      lc.map((l) => `<tr><td>${esc(l.desc)}</td><td>${l.data.split("-").reverse().join("/")}</td><td>${esc(l.paga)}</td><td>${R(l.total)}</td><td>${l.parcelado ? (diffM(l.ini, k) + 1) + "/" + l.n : "à vista"}</td><td>${R(l.valor)}</td><td>${R(l.custo)}</td></tr>`).join("");
  }

  // ---------- simulação ----------
  const VARS = [["Lazer", 0.35], ["Delivery/Restaurante", 0.25], ["Compras/Vestuário", 0.2], ["Outros", 0.2], ["Mercado (além do VR)", 0]];
  let SIM = null;
  function simBase(k) {
    const i = idx(k), pr = DATA.proj, p = DATA.plano[k] || {};
    const F = DATA.meses.includes(addM(k, 1)) ? addM(k, 1) : k, j = idx(F);
    // reembolsos de contas que eu não pago mais (ex.: Prime anual já pago) entram como entrada
    const creditos = DATA.fixos.filter((f) => ativo(f, F) && f.custo < 0).reduce((s, f) => s - f.custo, 0);
    const renda = pr.renda[j] + pr.ajuste[j] + pr.d13[i] + pr.ferias[i] + creditos;
    const ob = [];
    DATA.fixos.filter((f) => ativo(f, F) && f.custo > 0).forEach((f) => ob.push([f.desc, f.custo]));
    let outros = 0;
    DATA.dividas.filter((d) => ativo(d, F) && d.custo > 0).forEach((d) => (/a identificar/.test(d.desc) ? (outros += d.custo) : ob.push([d.desc, d.custo])));
    if (outros) ob.push(["Outras parcelas nos cartões", outros]);
    const lanc = DATA.lanc.filter((l) => l.ini && l.ini <= F && l.fim >= F).reduce((s, l) => s + l.custo, 0);
    const ja = p.ja || 0;
    if (lanc - ja > 0.5) ob.push(["Compras já lançadas (outros meses)", lanc - ja]);
    const x = Math.abs(pr.variaveis[j]) - lanc; if (x > 0.5) ob.push(["Outras cobranças previstas nas faturas", x]);
    // meta = 15% do líquido + parte do 13º/férias que o plano manda para a reserva
    const cent = (x) => Math.round(x * 100) / 100;
    const obTot = ob.reduce((s, o) => s + o[1], 0);
    const shares = Math.max(0, (p.reserva || 0) - ((p.espaco || 0) - (p.teto || 0)));
    const metaT = cent((p.meta || 0) + shares);
    const sug = p.teto ? cent(Math.max(DATA.minMes || 0, renda - obTot - metaT)) : 0;
    // padrão: gastar o sugerido; o que não gastar vai para a reserva
    const livre = Math.max(0, sug - ja);
    const vars = VARS.map(([n, w]) => ({ n, v: cent(livre * w) }));
    const dif = cent(livre - vars.reduce((s, o) => s + o.v, 0)); vars[3].v = cent(vars[3].v + dif);
    if (ja) vars.unshift({ n: "Já gasto no Inter neste mês", v: Math.round(ja * 100) / 100 });
    return { F, renda: Math.round(renda * 100) / 100, ob: ob.map(([n, v]) => ({ n, min: Math.round(v * 100) / 100, v: Math.round(v * 100) / 100 })), vars, meta: metaT, meta15: p.meta || 0, shares: cent(shares), sugerido: sug, vr: p.vr || pr.vr[i] || 0 };
  }
  function simLoad(k) { let s = null; try { s = JSON.parse(localStorage.getItem("sim-" + k)); } catch (_) {} const b = simBase(k);
    if (s && s.ob && s.ob.length === b.ob.length && s.vars && s.vars.length === b.vars.length && s.meta != null) { s.ob.forEach((o, n) => (o.min = b.ob[n].min, o.v = Math.max(o.v, b.ob[n].min))); s.vr = b.vr; s.F = b.F; s.meta = b.meta; s.meta15 = b.meta15; s.shares = b.shares; s.sugerido = b.sugerido; return s; } return b; }
  const simSave = (k) => { try { localStorage.setItem("sim-" + k, JSON.stringify(SIM)); } catch (_) {} };
  function slider(id, label, val, min, max, hint) {
    return `<div class="sl"><label for="${id}n">${esc(label)}</label><input type="range" id="${id}r" min="${min}" max="${max}" step="1" value="${val}"><input type="number" id="${id}n" min="${min}" step="0.01" value="${val}">${hint ? `<div class="hint">${hint}</div>` : ""}</div>`;
  }
  function renderSim(soResultado) {
    const k = $("sMes").value;
    if (!SIM || SIM.k !== k) { SIM = simLoad(k); SIM.k = k; soResultado = false; }
    const minMes = DATA.minMes || 0, meta = SIM.meta || 0;
    if (!soResultado) {
      $("sRenda").value = SIM.renda;
      $("sOb").innerHTML = SIM.ob.map((o, n) => slider("ob" + n, o.n + " (mínimo " + R(o.min) + ")", o.v, Math.floor(o.min), Math.ceil(Math.max(o.min * 2, o.min + 300)))).join("") || '<p class="muted">Nenhum obrigatório neste mês.</p>';
      $("sVar").innerHTML = SIM.vars.map((o, n) => slider("va" + n, o.n, o.v, 0, Math.ceil(Math.max(2000, SIM.renda)), o.n.startsWith("Mercado") ? "O Flash (VR) do mês (" + R(SIM.vr) + ") paga o supermercado; aqui entra só o que passar dele." : "")).join("") ;
      const bind = (pref, arr, isOb) => arr.forEach((o, n) => { const r = $(pref + n + "r"), x = $(pref + n + "n");
        const set = (v) => { v = +v || 0; if (isOb) v = Math.max(o.min, v); o.v = v; r.value = v; x.value = v; simSave(k); renderSim(true); };
        r.addEventListener("input", () => set(r.value)); x.addEventListener("change", () => set(x.value)); });
      bind("ob", SIM.ob, true); bind("va", SIM.vars, false);
    }
    const ob = SIM.ob.reduce((s, o) => s + o.v, 0), va = SIM.vars.reduce((s, o) => s + o.v, 0), ent = SIM.renda;
    const livre = ent - ob, res = livre - va, sobra = res;
    $("sObTot").textContent = R(ob); $("sVarTot").textContent = R(va) + " em gastos · livre " + R(livre);
    $("sVarNota").textContent = `Começa no sugerido (${R(SIM.sugerido)} = livre − meta de ${R(meta)}${SIM.shares ? ": " + P(DATA.pct) + " do líquido " + R(SIM.meta15) + " + parte do 13º/férias " + R(SIM.shares) : ", " + P(DATA.pct) + " do salário líquido"}). Não é obrigação: o que você não gastar vai para a reserva. Mínimo prioritário: ${R(minMes)}/mês.`;
    $("sNota").textContent = `Compras de ${lab(k)} são pagas com o salário de ${lab(SIM.F)} (5º dia útil). Por isso a entrada e os obrigatórios são os de ${lab(SIM.F)}: salário, contas fixas e faturas daquele mês, já descontados os reembolsos.`;
    $("sKpis").innerHTML = [["Entrada do mês", R(ent), ""], ["Obrigatórios", R(ob), P(ent ? ob / ent : 0) + " comprometido"], ["Livre no mês", R(livre), "sugerido gastar " + R(SIM.sugerido)],
      ["Gastos do mês", R(va), P(livre > 0 ? va / livre : 0) + " do livre"], [res >= 0 ? "Vai para a reserva" : "Déficit", R(res), res < 0 ? "falta dinheiro" : P(meta ? res / meta : 0) + " da meta de " + R0(meta), res < 0 ? "neg" : res >= meta - 0.01 ? "pos" : ""]]
      .map(([l, v, s, c]) => `<div class="kpi"><div class="l">${l}</div><div class="v ${c || ""}">${v}</div><div class="s">${s}</div></div>`).join("");
    const cor = pal(), tot = Math.max(ent, ob + va) || 1;
    const parts = [["Obrigatórios", ob, cor[0]], ["Gastos do mês", va, cor[1]], ["Vai para a reserva", Math.max(0, res), cor[2]]];
    $("sStack").innerHTML = parts.map(([n, v, c]) => (v > 0 ? `<span title="${n}: ${R(v)}" style="width:${(v / tot) * 100}%;background:${c}"></span>` : "")).join("");
    $("sLegend").innerHTML = parts.map(([n, v, c]) => `<span><i style="background:${c}"></i>${n}: ${R(v)} (${P(ent ? v / ent : 0)})</span>`).join("");
    const av = [];
    if (res < 0) av.push(`<span class="icon-bad">Faltam ${R(-res)}: os gastos passam do livre do mês.</span>`);
    else if (res < meta - 0.01) av.push(`<span class="icon-warn">A reserva fica em ${R(res)}, abaixo da meta de ${R(meta)} (${P(DATA.pct)} do salário líquido).</span>`);
    if (va < minMes) av.push(`<span class="icon-warn">Gastos abaixo do mínimo prioritário de ${R(minMes)}.</span>`);
    if (!av.length) av.push(`<span class="icon-ok">Cabe tudo: ${R(res)} vão para a reserva (meta de ${R(meta)} batida).</span>`);
    $("sAviso").innerHTML = av.join("<br>");
  }
  $("sMes").addEventListener("change", () => { SIM = null; renderSim(); });
  $("sRenda").addEventListener("change", (e) => { SIM.renda = +e.target.value || 0; simSave(SIM.k); renderSim(true); });
  $("sReset").addEventListener("click", () => { try { localStorage.removeItem("sim-" + $("sMes").value); } catch (_) {} SIM = null; renderSim(); });

  load();
})();

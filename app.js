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
    const K = ["mes","fatura","espaco","teto","semana","vista","parc","vezes","parcMax","simult","fora","vr","total","faturaInter","limite","reserva","reservaAc","ja","ainda","meta","livre","livreResta","reservaRend"];
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
    ["geral", "dividas", "metas", "sim"].forEach((t) => ($("tab-" + t).hidden = t !== TAB));
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
  function renderAll() { if (!DATA) return; resumoFiltros(); renderTopo(); document.querySelectorAll(".mesSel").forEach((e) => (e.textContent = lab(F.mes)));
    if (TAB === "geral") renderGeral(); else if (TAB === "dividas") renderDividas(); else if (TAB === "metas") renderMetas(); else renderSim(); }

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
      <div class="it"><div class="l">Crédito do ${esc(card)} em ${lab(cm)}</div><div class="v">${R(fechado ? 0 : fc.credSug)}</div><div class="muted small">${fechado ? "mês fechado" : antesCorte ? "sugerido · até dia " + fch + " a compra ainda cai na fatura de " + lab(cm) + ", que já está no limite: espere o dia " + (fch + 1) : "sugerido (cai na fatura de " + lab(addM(cm, 1)) + ")"}</div></div>
      <div class="it"><div class="l">Em dinheiro (Pix/débito)</div><div class="v">${R(fechado ? DATA.saldo : fc.cashSug)}</div><div class="muted small">${fechado ? "saldo em conta" : fc.sobraMes < 0 ? "a conta de " + lab(cm) + " fecha negativa: evite" : "sugerido"}</div></div>
      ${(() => { const pd = contasMes(cm).filter((c) => !c.pago); return `<div class="it"><div class="l">Contas a pagar em ${lab(cm)}</div><div class="v ${pd.length ? "neg" : ""}">${R(pd.reduce((s, c) => s + c.valor, 0))}</div><div class="muted small">${pd.length ? pd.map((c) => esc(c.nome.replace(/^Fatura /, ""))).join(", ") : "tudo pago"}</div></div>`; })()}
      <div class="it"><div class="l">${lab(cm)}: sugerido / livre</div><div class="v">${R0(pc.teto)} / ${R0(pc.livre)}</div><div class="muted small">o que não gastar vai para a reserva</div></div>`;
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
    renderContas(k, "g"); renderFormas(k); renderSaude(k); renderRitmo(k); renderLimites(k); renderVR();
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

  function statusReceber(k, rc) {
    const n = (t) => txt(t).toLowerCase().replace(/\s+/g, " "), pool = DATA.recItens.filter((r) => r.mes === k).map((r) => Object.assign({ resto: r.valor }, r));
    rc.forEach((r) => { let pago = 0, dt = ""; pool.filter((p) => n(p.pessoa) === n(r.pessoa) && n(p.item) === n(r.base) && p.resto > 0.004).forEach((p) => { const u = Math.min(p.resto, r.valor - pago); if (u > 0) { p.resto -= u; pago += u; dt = dt || p.data; } });
      r.pago = pago; r.dataPago = dt; r.st = pago >= r.valor - 0.015 ? "pago" : pago > 0.004 ? "parcial" : "pendente"; });
    return pool.filter((p) => p.resto > 0.015);
  }
  function renderDividas() {
    const k = F.mes, i = idx(k), ent = DATA.proj.entradas[i] || 0, cor = pal();
    // a receber
    let rc = receber(k); if (F.cartao) rc = rc.filter((r) => r.origem.includes(F.cartao) || (F.cartao.startsWith("Conta") && r.origem === "Conta fixa"));
    const porPessoa = {}; rc.forEach((r) => (porPessoa[r.pessoa] = (porPessoa[r.pessoa] || 0) + r.valor));
    $("rResumo").textContent = Object.entries(porPessoa).map(([p, v]) => p + ": " + R(v)).join(" · ") || "nada a receber";
    rc.sort((a, b) => a.pessoa.localeCompare(b.pessoa) || (a.quando || 0) - (b.quando || 0));
    const sobras = statusReceber(k, rc);
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
    // lançamentos
    let lc = DATA.lanc.filter((l) => l.ini && l.ini <= k && l.fim >= k && passaForma(l.paga));
    if (F.item && F.item.startsWith("L|")) lc = lc.filter((l) => l.desc === F.item.slice(2));
    const tl = lc.reduce((s, l) => s + l.valor, 0);
    $("lResumo").textContent = lc.length ? `${lc.length} lançamento(s) cobrados em ${lab(k)} · ${R(tl)}` : "Nenhum lançamento cobrado neste mês. Adicione na aba Lançamentos da planilha.";
    $("tLanc").innerHTML = `<tr><th>Descrição</th><th>Data</th><th>Forma</th><th>Valor total</th><th>Parcelas</th><th>Cobrado no mês</th><th>Meu custo</th></tr>` +
      lc.map((l) => `<tr><td>${esc(l.desc)}</td><td>${l.data.split("-").reverse().join("/")}</td><td>${esc(l.paga)}</td><td>${R(l.total)}</td><td>${l.parcelado ? (diffM(l.ini, k) + 1) + "/" + l.n : "à vista"}</td><td>${R(l.valor)}</td><td>${R(l.custo)}</td></tr>`).join("");
    renderContas(k); renderRecebidos(k); renderLibera(k); renderPrevReal();
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
      pend.concat(pagas).map((c) => `<tr><td>${esc(c.nome)}</td><td>${R(c.pago && c.pago.valor ? c.pago.valor : c.valor)}</td><td>${venc(c)}</td><td>${c.pago ? `<span class="tag good">Pago${c.pago.data ? " " + c.pago.data.split("-").reverse().slice(0, 2).join("/") : ""}</span>` : `<span class="tag bad">Pendente</span>`}</td></tr>`).join("") +
      `</table></div><p class="muted small">Marque o que pagou na aba Pagamentos da planilha (mês, nome da conta e valor). Valores pendentes são os previstos: condomínio e energia ainda estimados.</p>`;
  }

  // ---------- VR (Flash) ----------
  const isVR = (l) => l.paga === "VR";
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
    $("vrBody").innerHTML = `<div class="kpis">
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
  function renderRecebidos(k) {
    const { mes, atras } = pendencias(k);
    $("tRecebido").innerHTML = `<tr><th>Pessoa</th><th>Devido em ${lab(k)}</th><th>Recebido</th><th>Pendente</th><th>Atrasado (meses anteriores)</th></tr>` +
      Object.entries(mes).filter(([p, o]) => o.d || o.r || atras[p]).map(([p, o]) => `<tr><td>${esc(p)}</td><td>${R(o.d)}</td><td class="${o.r >= o.d - 0.005 && o.d ? "pos" : ""}">${R(o.r)}</td><td class="${o.pend > 0.005 ? "neg" : ""}">${R(o.pend)}</td><td class="${atras[p] > 0.005 ? "neg" : ""}">${R(atras[p])}</td></tr>`).join("") || `<tr><td colspan="5" class="muted">Nada a receber neste mês.</td></tr>`;
    const tA = Object.values(atras).reduce((s, v) => s + v, 0);
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
    $("assResumo").textContent = `${as.length} assinaturas · ${R(tot)}/mês · ${R0(ano)} nos próximos 12 meses`;
    $("tAss").innerHTML = `<tr><th>Assinatura</th><th>Paga com</th><th>Por mês</th><th>Próximos 12 meses</th><th>Até</th><th>Manter?</th></tr>` +
      as.slice().sort((a, b) => b.ano - a.ano).map((a) => `<tr><td>${esc(a.nome)}</td><td>${esc(a.paga)}</td><td>${R(a.valor)}</td><td>${R(a.ano)}</td><td>${/^\d{4}-\d\d$/.test(a.ate) ? lab(a.ate) : esc(a.ate)}</td><td><span class="tag ${a.manter === "Não" ? "bad" : a.manter === "Sim" ? "good" : ""}">${esc(a.manter)}</span></td></tr>`).join("") +
      `<tr class="total"><td>Total</td><td></td><td>${R(tot)}</td><td>${R(ano)}</td><td></td><td></td></tr>`;
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
  const TIPOS = { parc: "Compra parcelada no cartão", vista: "Compra à vista no cartão", div: "Nova dívida (boleto, empréstimo, financiamento)" };
  let SIM = null, SIM_DEF = null;
  const cent = (x) => Math.round(x * 100) / 100;
  function simLoad() {
    let s = null; try { s = JSON.parse(localStorage.getItem("sim-v2")); } catch (_) {}
    if (!s || !Array.isArray(s.itens)) s = { itens: [], modo: "orc" };
    s.itens.forEach((it) => { if (!DATA.meses.includes(it.mes)) it.mes = SIM_DEF; if (it.cartao && !DATA.cartoes.includes(it.cartao)) it.cartao = DATA.cartoes[0]; });
    return s;
  }
  const simSave = () => { try { localStorage.setItem("sim-v2", JSON.stringify(SIM)); } catch (_) {} };
  function novoItem(t) {
    const base = { id: Date.now() + "" + Math.random().toString(36).slice(2, 6), t, desc: "", valor: 0, pct: 0 };
    if (t === "parc") return Object.assign(base, { n: 3, cartao: DATA.cartoes[0], mes: SIM_DEF });
    if (t === "vista") return Object.assign(base, { n: 1, cartao: DATA.cartoes[0], mes: SIM_DEF });
    return Object.assign(base, { n: 12, mes: DATA.meses.includes(addM(SIM_DEF, 1)) ? addM(SIM_DEF, 1) : SIM_DEF });
  }
  // cobranças do item por mês de pagamento (salário que paga)
  function cobrancas(it) {
    const n = it.t === "vista" ? 1 : Math.max(1, Math.round(+it.n || 1)), out = [];
    const parcela = it.t === "div" ? +it.valor || 0 : (+it.valor || 0) / n;
    const meu = parcela * (1 - Math.min(100, Math.max(0, +it.pct || 0)) / 100);
    for (let i = 0; i < n; i++) out.push({ pag: it.t === "div" ? addM(it.mes, i) : addM(it.mes, i + 1), valor: parcela, meu });
    return { n, parcela, meu, lista: out };
  }
  function simCalc() {
    const minMes = DATA.minMes || 0, porPag = {};
    SIM.itens.forEach((it) => cobrancas(it).lista.forEach((c) => (porPag[c.pag] = (porPag[c.pag] || 0) + c.meu)));
    const meses = DATA.meses.filter((k) => DATA.plano[k] && k >= DATA.planIni);
    let acc = 0;
    const linhas = meses.map((k) => {
      const p = DATA.plano[k], F = addM(k, 1), j = idx(F), pr = DATA.proj;
      const nova = cent(porPag[F] || 0);
      const shares = (p.reserva || 0) - ((p.espaco || 0) - (p.teto || 0));
      let tetoD, resD;
      if (SIM.modo === "res") { tetoD = p.teto || 0; resD = (p.reserva || 0) - nova; }
      else { const espD = (p.espaco || 0) - nova; tetoD = p.teto > 0 ? Math.max(minMes, espD - (p.meta || 0)) : 0; resD = espD - tetoD + shares; }
      acc += resD - (p.reserva || 0);
      const ent = j >= 0 ? pr.entradas[j] : 0;
      const obrig = j >= 0 ? Math.abs(pr.fixos[j]) + Math.abs(pr.parcelas[j]) + Math.abs(pr.variaveis[j]) - pr.reemb[j] : 0;
      return { k, F, nova, ent, obrig, compA: ent ? obrig / ent : null, compD: ent ? (obrig + nova) / ent : null,
        sobraA: j >= 0 ? pr.sobra[j] : null, sobraD: j >= 0 ? pr.sobra[j] - nova : null,
        tetoA: p.teto || 0, tetoD: cent(tetoD), resA: p.reserva || 0, resD: cent(resD), accA: p.reservaAc || 0, accD: cent((p.reservaAc || 0) + acc), meta: p.meta || 0, p };
    });
    return { linhas, porPag };
  }
  function renderSimEditor() {
    const mesOpts = (sel) => DATA.meses.filter((k) => k >= DATA.planIni || k === sel).map((k) => `<option value="${k}" ${k === sel ? "selected" : ""}>${lab(k)}</option>`).join("");
    const carOpts = (sel) => DATA.cartoes.map((c) => `<option ${c === sel ? "selected" : ""}>${esc(c)}</option>`).join("");
    $("sItens").innerHTML = SIM.itens.length ? SIM.itens.map((it, n) => `
      <div class="simit" data-i="${n}">
        <div class="simit-h"><b>${n + 1}. ${TIPOS[it.t]}</b><button type="button" class="btn ghost sm" data-rm="${n}" aria-label="Remover">Remover</button></div>
        <div class="simit-g">
          <div class="f wide"><label>Descrição</label><input data-k="desc" type="text" value="${esc(it.desc)}" placeholder="${it.t === "div" ? "ex.: empréstimo, financiamento" : "ex.: tênis, passagem"}"></div>
          <div class="f"><label>${it.t === "div" ? "Valor da parcela (R$)" : "Valor total da compra (R$)"}</label><input data-k="valor" type="number" inputmode="decimal" step="0.01" min="0" value="${it.valor || ""}"></div>
          ${it.t === "vista" ? "" : `<div class="f"><label>Nº de parcelas</label><input data-k="n" type="number" inputmode="numeric" step="1" min="1" max="60" value="${it.n}"></div>`}
          ${it.t === "div" ? "" : `<div class="f"><label>Cartão</label><select data-k="cartao">${carOpts(it.cartao)}</select></div>`}
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
      if (el) el.innerHTML = (it.t === "div" ? `${c.n}x de ${R(c.parcela)} (total ${R(c.parcela * c.n)})` : c.n > 1 ? `${c.n}x de ${R(c.parcela)} no ${esc(it.cartao)}` : `${R(c.parcela)} à vista no ${esc(it.cartao)}`) +
        ` · pago com o salário de ${lab(ini)}${c.n > 1 ? " até " + lab(fim) : ""}` + (it.pct ? ` · meu custo ${R(c.meu)}/mês (${esc(it.pct)}% volta pra mim)` : "");
      const p = DATA.plano[it.mes] || {};
      if (it.t !== "div" && it.cartao !== card0) av.push(`<span class="icon-warn">${nome}: a combinação era parar de usar ${esc(it.cartao)} e concentrar no ${esc(card0)}.</span>`);
      if (it.t === "parc" && p.parc != null) {
        if (p.parc && c.parcela > p.parc + 0.005) av.push(`<span class="icon-warn">${nome}: parcela de ${R(c.parcela)} passa do limite de parcelas novas de ${lab(it.mes)} (${R(p.parc)}/mês).</span>`);
        if (p.vezes && c.n > p.vezes) av.push(`<span class="icon-warn">${nome}: ${c.n}x passa da regra de até ${p.vezes}x em ${lab(it.mes)}.</span>`);
      }
      if (it.t === "vista" && p.vista != null && c.parcela > (p.teto || 0) + 0.005) av.push(`<span class="icon-warn">${nome}: ${R(c.parcela)} à vista é mais do que o orçamento sugerido de ${lab(it.mes)} (${R(p.teto)}).</span>`);
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
      ["Orçamento de gastos", "−" + R(cortOrc), SIM.modo === "res" ? "orçamento mantido" : "a menos para gastar no período"],
      ["Reserva em " + lab(ult.k), R(ult.accD), (difRes < 0 ? "−" : "+") + R(Math.abs(difRes)) + " vs. sem a simulação", ult.accD < 0 ? "neg" : difRes < -0.005 ? "" : "pos"],
      ["Entrada comprometida", maxC.k ? P(maxC.compD) : "—", maxC.k ? "no pior mês (" + lab(maxC.F) + "), antes " + P(maxC.compA) : ""],
    ].map(([l, v, s, c]) => `<div class="kpi"><div class="l">${l}</div><div class="v ${c || ""}">${v}</div><div class="s">${s}</div></div>`).join("");
    // avisos do planejamento
    const neg = vis.filter((l) => l.resD < -0.005), noMin = vis.filter((l) => l.nova > 0 && SIM.modo === "orc" && l.tetoD <= minMes + 0.005 && l.resD < l.resA - 0.005);
    const abaixoMeta = vis.filter((l) => l.nova > 0 && l.resD >= -0.005 && l.resD < l.meta - 0.01 && l.resA >= l.meta - 0.01);
    const resNeg = vis.find((l) => l.accD < -0.005);
    if (neg.length) av.unshift(`<span class="icon-bad">Em ${neg.length} mês(es) a conta não fecha (${neg.map((l) => lab(l.k) + " " + R0(l.resD)).join(", ")}): a diferença teria que sair da reserva.</span>`);
    if (resNeg) av.unshift(`<span class="icon-bad">A reserva ficaria negativa a partir de ${lab(resNeg.k)} (${R(resNeg.accD)}).</span>`);
    if (noMin.length) av.push(`<span class="icon-warn">Em ${noMin.map((l) => lab(l.k)).join(", ")} o orçamento já fica no mínimo de ${R0(minMes)}/mês, então a parcela passa a sair da reserva.</span>`);
    if (abaixoMeta.length) av.push(`<span class="icon-warn">A reserva do mês fica abaixo da meta de ${P(DATA.pct)} em ${abaixoMeta.map((l) => lab(l.k)).join(", ")}.</span>`);
    if (!av.length) av.push(`<span class="icon-ok">Cabe no planejamento: ${SIM.modo === "res" ? "o orçamento continua o mesmo e a reserva absorve a parcela" : "a parcela sai do orçamento de gastos e a meta da reserva continua batida"}.</span>`);
    av.push(`<span class="muted small">Se decidir fazer, lance na aba Lançamentos da planilha para entrar no planejamento de verdade.</span>`);
    $("sAviso").innerHTML = av.join("<br>");
    // gráficos
    const L = vis.map((l) => lab(l.k));
    const ultAf = afet.length ? afet[afet.length - 1].k : primeiro, vl = vis.filter((l) => l.k <= addM(ultAf, 2)), bad = css("--bad");
    const cl = draw("cSimLivre", "bar", vl.map((l) => lab(l.k)), [ds("Nova parcela", vl.map((l) => l.nova), cor[1]), ds("Para gastar", vl.map((l) => l.tetoD), cor[0]),
      ds("Vai para a reserva", vl.map((l) => Math.max(0, l.resD)), cor[2])].concat(vl.some((l) => l.resD < -0.005) ? [ds("Falta (sai da reserva)", vl.map((l) => Math.min(0, l.resD)), bad)] : []), baseOpts(true, (v) => R0(v)));
    if (vl.length > 8) { cl.options.plugins.datalabels.display = (c) => c.datasetIndex === 0 && !!c.raw && (c.dataIndex === 0 || vl[c.dataIndex].nova !== vl[c.dataIndex - 1].nova); cl.update(); }
    draw("cSimRes", "line", L, [
      ds("Sem a simulação", vis.map((l) => l.accA), css("--muted"), { fill: false, tension: 0.2, borderWidth: 2, borderDash: [6, 4], pointRadius: 0, datalabels: { display: false } }),
      ds("Com a simulação", vis.map((l) => l.accD), cor[0], { fill: false, tension: 0.2, borderWidth: 2, pointRadius: 3, pointBackgroundColor: cor[0] })],
      baseOpts(false, (v) => R0(v), {}));
    const cr = charts.cSimRes; cr.options.plugins.datalabels.display = (c) => c.datasetIndex === 1 && (c.dataIndex === vis.length - 1 || c.dataIndex % 3 === 0); cr.update();
    // tabela
    $("tSim").innerHTML = `<tr><th>Mês das compras</th><th>Nova parcela</th><th>Entrada comprometida</th><th>Orçamento para gastar</th><th>Vai para a reserva</th><th>Reserva acumulada</th><th>Sobra do salário seguinte</th></tr>` +
      vis.map((l) => `<tr class="${l.nova > 0 ? "" : "fora"}"><td>${lab(l.k)}<div class="muted small">salário de ${lab(l.F)}</div></td><td>${l.nova ? R(l.nova) : "—"}</td><td>${l.compA == null ? "—" : ab(l.compA, l.compD, P)}</td>
        <td>${ab(l.tetoA, l.tetoD)}</td><td class="${l.resD < 0 ? "neg" : ""}">${ab(l.resA, l.resD)}</td><td class="${l.accD < 0 ? "neg" : ""}">${ab(l.accA, l.accD)}</td><td class="${l.sobraD != null && l.sobraD < 0 ? "neg" : ""}">${l.sobraA == null ? "—" : ab(l.sobraA, l.sobraD)}</td></tr>`).join("");
  }
  function renderSim() { if (!SIM) SIM = simLoad(); renderSimEditor(); renderSimResultado(); }
  document.querySelectorAll("[data-add]").forEach((b) => b.addEventListener("click", () => { SIM.itens.push(novoItem(b.dataset.add)); simSave(); renderSimEditor(); renderSimResultado(); }));
  $("sModo").addEventListener("change", (e) => { SIM.modo = e.target.value; simSave(); renderSimResultado(); });
  $("sLimpar").addEventListener("click", () => { SIM.itens = []; simSave(); renderSimEditor(); renderSimResultado(); });

  // ---------- senha de acesso (proteção simples, só na tela) ----------
  async function hashSenha(t) { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode("meu-financeiro:" + t)); return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join(""); }
  function portao() {
    const H = (CFG.PASS_HASH || "").trim().toLowerCase(); let lembrado = null; try { lembrado = localStorage.getItem("dash-acesso"); } catch (_) {}
    if (H && lembrado === H) { document.body.classList.remove("travado"); load(); return; }
    const ov = document.createElement("div"); ov.className = "lock"; ov.setAttribute("role", "dialog"); ov.setAttribute("aria-modal", "true");
    ov.innerHTML = H ? `<form class="lock-box" id="lkF"><div class="logo-lk" aria-hidden="true">🔒</div><h2>Meu Financeiro</h2><p class="muted small">Digite a senha para ver o dashboard.</p>
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

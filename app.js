/* Meu Financeiro — dashboard que lê a planilha Planejamento_Financeiro.xlsx */
(function () {
  "use strict";
  Chart.register(ChartDataLabels);
  const $ = (id) => document.getElementById(id);
  const MESES = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"];
  const brl = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  const brl0 = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL", maximumFractionDigits: 0 });
  const pctf = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 });
  const R = (n) => (n == null || isNaN(n) ? "—" : brl.format(n));
  const R0 = (n) => (n == null || isNaN(n) ? "—" : brl0.format(n));
  const P = (n) => (n == null || !isFinite(n) ? "—" : pctf.format(n));
  const num = (v) => (typeof v === "number" && isFinite(v) ? v : 0);
  const css = (v) => getComputedStyle(document.documentElement).getPropertyValue(v).trim();

  // ---------- datas ----------
  function ym(v) {
    if (v instanceof Date && !isNaN(v)) { const d = new Date(v.getTime() + 12 * 3600e3); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); }
    if (typeof v === "number" && v > 20000 && v < 80000) { const p = XLSX.SSF.parse_date_code(v); return p.y + "-" + String(p.m).padStart(2, "0"); }
    return null;
  }
  const lab = (k) => { if (!k) return "—"; const [y, m] = k.split("-"); return MESES[+m - 1] + "/" + y.slice(2); };
  const addM = (k, n) => { let [y, m] = k.split("-").map(Number); m += n; y += Math.floor((m - 1) / 12); m = ((m - 1) % 12 + 12) % 12 + 1; return y + "-" + String(m).padStart(2, "0"); };
  const diffM = (a, b) => { const [ya, ma] = a.split("-").map(Number), [yb, mb] = b.split("-").map(Number); return (yb - ya) * 12 + (mb - ma); };
  const todayKey = () => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0"); };

  // ---------- leitura da planilha ----------
  function rows(wb, name) {
    const ws = wb.Sheets[name]; if (!ws) throw new Error('Aba "' + name + '" não encontrada');
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }
  const cell = (wb, sh, addr) => { const c = wb.Sheets[sh] && wb.Sheets[sh][addr]; return c ? c.v : null; };
  const txt = (v) => (v == null ? "" : String(v).trim());
  const findRow = (rs, pred, from = 0) => { for (let i = from; i < rs.length; i++) if (rs[i] && pred(txt(rs[i][0]))) return i; return -1; };

  function extract(wb) {
    const D = {};
    // Configurações
    D.atual = ym(cell(wb, "Configurações", "B4"));
    D.reservaInformada = num(cell(wb, "Configurações", "B5"));
    D.meta = num(cell(wb, "Configurações", "B6"));
    D.cartoes = [];
    for (let r = 11; r <= 15; r++) { const n = txt(cell(wb, "Configurações", "A" + r)); if (n && n !== "-") D.cartoes.push(n); }
    D.minSemana = num(cell(wb, "Plano de Gastos", "B3"));

    // Projeção
    const pj = rows(wb, "Projeção");
    const hi = findRow(pj, (t) => t === "Item");
    D.meses = pj[hi].slice(1).map(ym).filter(Boolean);
    const n = D.meses.length;
    D.tipo = pj[hi - 1].slice(1, n + 1).map(txt);
    const line = (prefix) => { const i = findRow(pj, (t) => t.startsWith(prefix)); return i < 0 ? Array(n).fill(0) : pj[i].slice(1, n + 1).map(num); };
    D.proj = {
      renda: line("Renda ("), d13: line("(+) 13º"), ferias: line("(+) Férias"), ajuste: line("(−) Ajuste"),
      fixos: line("(−) Fixos"), parcelas: line("(−) Parcelas"), reemb: line("(+) Reembolsos"), variaveis: line("(−) Variáveis"),
      sobra: line("SOBRA DO MÊS"), saldo: line("Saldo acumulado"), vr: line("VR do mês"),
      comprom: line("TOTAL DE COMPROMISSOS"), pctComprom: line("% da renda do mês"),
    };
    D.proj.entradas = D.meses.map((_, i) => D.proj.renda[i] + D.proj.d13[i] + D.proj.ferias[i] + D.proj.ajuste[i] + D.proj.reemb[i]);
    D.faturas = {};
    let fi = findRow(pj, (t) => t.startsWith("FATURAS POR CARTÃO"));
    for (let i = fi + 1; i < pj.length; i++) { const t = txt(pj[i][0]); if (t.startsWith("TOTAL")) break; if (t && t !== "-") D.faturas[t] = pj[i].slice(1, n + 1).map(num); }
    D.reembPessoa = {};
    fi = findRow(pj, (t) => t.startsWith("REEMBOLSOS POR PESSOA"));
    for (let i = fi + 1; i < fi + 6 && i < pj.length; i++) { const t = txt(pj[i] && pj[i][0]); if (t && t !== "-") D.reembPessoa[t] = pj[i].slice(1, n + 1).map(num); }

    // Plano de Gastos
    const pl = rows(wb, "Plano de Gastos");
    const ph = findRow(pl, (t) => t === "Mês das compras");
    const K = ["mes","fatura","espaco","teto","semana","vista","parc","vezes","parcMax","simult","fora","vr","total","faturaInter","limite","reserva","reservaAc"];
    D.plano = {};
    for (let i = ph + 1; i < pl.length; i++) {
      const r = pl[i]; if (!r || txt(r[0]).startsWith("TOTAL")) break;
      const k = ym(r[0]); if (!k) continue;
      const o = {}; K.forEach((f, j) => (o[f] = j < 2 ? ym(r[j]) : num(r[j]))); D.plano[k] = o;
    }

    // Parcelas
    const pa = rows(wb, "Parcelas");
    const pah = findRow(pa, (t) => t === "Descrição");
    D.dividas = [];
    for (let i = pah + 1; i < pa.length; i++) {
      const r = pa[i]; const d = txt(r && r[0]); if (!d) continue; if (d.startsWith("TOTAL")) break;
      D.dividas.push({ desc: d, cat: txt(r[1]), cartao: txt(r[2]), valor: num(r[3]), ini: ym(r[4]), fim: ym(r[5]),
        reemb: num(r[6]) + num(r[8]), quem: [txt(r[7]), txt(r[9])].filter((x) => x && x !== "-").join(" e "), custo: num(r[10]) });
    }

    // Gastos avulsos
    const gs = rows(wb, "Gastos");
    const gh = findRow(gs, (t) => t === "Data");
    D.gastos = [];
    for (let i = gh + 1; i < gs.length; i++) { const r = gs[i]; if (!r) continue; const k = ym(r[0]); if (!k || !num(r[4])) continue; D.gastos.push({ mes: k, cat: txt(r[2]), paga: txt(r[3]), valor: num(r[4]) }); }

    // Categorias (Resumo mensal)
    const rs = rows(wb, "Resumo mensal");
    const rh = findRow(rs, (t) => t === "Mês");
    const rmeses = rs[rh].slice(1, 13).map(ym);
    const ci = findRow(rs, (t) => t.startsWith("MEU CUSTO REAL POR CATEGORIA"));
    D.cat = { meses: rmeses, itens: {} };
    for (let i = ci + 1; i < rs.length; i++) { const t = txt(rs[i] && rs[i][0]); if (!t) continue; if (t === "TOTAL") break; D.cat.itens[t] = { vals: rs[i].slice(1, 13), media: num(rs[i][13]) }; }
    return D;
  }

  // ---------- carregamento ----------
  let DATA = null;
  function setStatus(t) { $("status").textContent = t; }
  function useWorkbook(buf, origem) {
    const wb = XLSX.read(buf, { type: "array", cellDates: true });
    DATA = extract(wb);
    const stamp = new Date().toLocaleString("pt-BR");
    try { localStorage.setItem("dash-data", JSON.stringify({ DATA, stamp, origem })); } catch (e) {}
    setStatus("Planilha carregada de " + origem + " em " + stamp + " · mês atual da planilha: " + lab(DATA.atual));
    setupFilters(); render();
  }
  async function load() {
    const url = (window.DASH_CONFIG && window.DASH_CONFIG.DATA_URL) || "planilha.xlsx";
    try {
      const res = await fetch(url + (url.includes("?") ? "&" : "?") + "v=" + Date.now(), { cache: "no-store" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      useWorkbook(await res.arrayBuffer(), url.startsWith("http") ? "link configurado" : url);
    } catch (e) {
      let c = null; try { c = JSON.parse(localStorage.getItem("dash-data")); } catch (_) {}
      if (c && c.DATA) { DATA = c.DATA; setStatus("Usando a última planilha salva neste navegador (" + c.stamp + "). Não consegui ler " + url + "."); setupFilters(); render(); }
      else setStatus('Não encontrei a planilha (' + url + '). Envie planilha.xlsx ao repositório ou use "Carregar planilha".');
    }
  }
  $("file").addEventListener("change", (e) => { const f = e.target.files[0]; if (!f) return; f.arrayBuffer().then((b) => useWorkbook(b, "arquivo " + f.name)); });
  $("reload").addEventListener("click", load);

  // ---------- filtros ----------
  const F = { mes: null, de: null, ate: null, cartao: "", divida: "ativas", reemb: "", unid: "rs" };
  function opt(sel, items, val) { sel.innerHTML = items.map(([v, t]) => `<option value="${v}">${t}</option>`).join(""); if (val != null) sel.value = val; }
  function setupFilters() {
    const ms = DATA.meses.map((k) => [k, lab(k) + (DATA.tipo[DATA.meses.indexOf(k)] === "Realizado" ? " (real)" : "")]);
    const tk = todayKey();
    const def = DATA.meses.includes(tk) ? tk : (DATA.meses.includes(DATA.atual) ? DATA.atual : DATA.meses[0]);
    let saved = {}; try { saved = JSON.parse(localStorage.getItem("dash-filtros")) || {}; } catch (e) {}
    F.mes = DATA.meses.includes(saved.mes) ? saved.mes : def;
    F.de = DATA.meses.includes(saved.de) ? saved.de : DATA.meses[0];
    F.ate = DATA.meses.includes(saved.ate) ? saved.ate : DATA.meses[DATA.meses.length - 1];
    ["cartao", "divida", "reemb", "unid"].forEach((k) => saved[k] != null && (F[k] = saved[k]));
    opt($("fMes"), ms, F.mes); opt($("fDe"), ms, F.de); opt($("fAte"), ms, F.ate);
    opt($("fCartao"), [["", "Todos"]].concat(DATA.cartoes.map((c) => [c, c])), F.cartao);
    $("fDivida").value = F.divida; $("fReemb").value = F.reemb; $("fUnid").value = F.unid;
    document.querySelectorAll(".cardname").forEach((e) => (e.textContent = DATA.cartoes[0] || "Inter"));
  }
  [["fMes", "mes"], ["fDe", "de"], ["fAte", "ate"], ["fCartao", "cartao"], ["fDivida", "divida"], ["fReemb", "reemb"], ["fUnid", "unid"]].forEach(([id, k]) =>
    $(id).addEventListener("change", (e) => { F[k] = e.target.value; if (diffM(F.de, F.ate) < 0) { const t = F.de; F.de = F.ate; F.ate = t; $("fDe").value = F.de; $("fAte").value = F.ate; }
      try { localStorage.setItem("dash-filtros", JSON.stringify(F)); } catch (_) {} render(); }));
  $("gastoManual").addEventListener("input", (e) => { try { const v = e.target.value; v === "" ? localStorage.removeItem("gasto-" + F.mes) : localStorage.setItem("gasto-" + F.mes, v); } catch (_) {} render(true); });

  // ---------- cálculos ----------
  const idx = (k) => DATA.meses.indexOf(k);
  function gastoNoMes(k, pred) { return DATA.gastos.filter((g) => g.mes === k && pred(g)).reduce((s, g) => s + g.valor, 0); }
  function gastoInter(k) {
    let m = null; try { m = localStorage.getItem("gasto-" + k); } catch (_) {}
    if (m !== null && m !== "" && !isNaN(+m)) return { v: +m, fonte: "valor informado" };
    return { v: gastoNoMes(k, (g) => g.paga === DATA.cartoes[0]), fonte: "aba Gastos" };
  }
  function semanasRestantes(k) {
    const [y, m] = k.split("-").map(Number); const fim = new Date(y, m, 0); const tk = todayKey();
    if (k < tk) return 0;
    const ini = k === tk ? new Date() : new Date(y, m - 1, 1);
    const dias = Math.floor((fim - new Date(ini.getFullYear(), ini.getMonth(), ini.getDate())) / 864e5) + 1;
    return Math.max(dias / 7, 0);
  }
  function dividaInfo(d, k) {
    const ativa = (!d.ini || d.ini <= k) && (!d.fim || d.fim >= k);
    let rest = null;
    if (d.fim) { const ini = d.ini && d.ini > k ? d.ini : k; rest = Math.max(0, diffM(ini, d.fim) + 1); }
    return { ativa, rest, aPagar: rest == null ? null : rest * d.valor };
  }
  function range() { const a = idx(F.de), b = idx(F.ate); return DATA.meses.slice(a, b + 1); }

  // ---------- charts ----------
  const charts = {};
  function palette() { return [css("--s1"), css("--s2"), css("--s3"), css("--s4"), css("--s5")]; }
  function baseOpts(stacked, fmtFn, extra = {}) {
    const ink = css("--ink2"), line = css("--line");
    const o = Object.assign({
      responsive: true, maintainAspectRatio: false, animation: { duration: 250 },
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { position: "top", labels: { color: ink, boxWidth: 12, boxHeight: 12, usePointStyle: false } },
        tooltip: { callbacks: { label: (c) => " " + c.dataset.label + ": " + fmtFn(c.raw) } },
        datalabels: {
          color: stacked ? "#fff" : ink, font: { size: 10, weight: "600" },
          anchor: stacked ? "center" : "end", align: stacked ? "center" : "end", offset: 1, clamp: true,
          formatter: (v, c) => (v == null || Math.abs(v) < 0.005 ? "" : fmtFn(v)),
          display: (c) => { const v = c.dataset.data[c.dataIndex]; if (!v) return false; if (!stacked) return true;
            const max = Math.max(...c.chart.data.datasets.flatMap((d) => d.data.map((x) => Math.abs(x || 0)))); return Math.abs(v) > max * 0.06; },
        },
      },
      scales: {
        x: { stacked, grid: { display: false }, ticks: { color: ink } },
        y: { stacked, grid: { color: line }, border: { display: false }, ticks: { color: ink, callback: (v) => fmtFn(v, true) } },
      },
    }, extra);
    if (o.indexAxis === "y") { o.scales.y.ticks = { color: ink, autoSkip: false }; o.scales.x.ticks = { color: ink, callback: (v) => fmtFn(v, true) }; o.scales.x.grid = { color: line }; o.scales.y.grid = { display: false }; o.layout = { padding: { right: 56 } }; }
    return o;
  }
  function draw(id, type, labels, datasets, opts) {
    if (charts[id]) charts[id].destroy();
    charts[id] = new Chart($(id), { type, data: { labels, datasets }, options: opts });
  }
  function ds(label, data, color, extra = {}) { return Object.assign({ label, data, backgroundColor: color, borderColor: color, borderWidth: 0, borderRadius: 4, borderSkipped: "start", maxBarThickness: 46 }, extra); }

  // ---------- render ----------
  function render(onlyHero) {
    if (!DATA) return;
    const k = F.mes, i = idx(k), p = DATA.plano[k] || {}, pr = DATA.proj;
    const pct = F.unid === "pct";
    const tipo = DATA.tipo[i] === "Realizado" ? '<span class="tag">realizado</span>' : '<span class="tag prev">previsto</span>';

    // Hero
    const gi = gastoInter(k), teto = p.teto || 0, rest = teto - gi.v;
    $("hMes").textContent = lab(k);
    $("hRestante").textContent = R(rest); $("hRestante").classList.toggle("neg", rest < 0);
    const uso = teto ? gi.v / teto : 0; const bar = $("hBar"); bar.style.width = Math.min(100, uso * 100) + "%"; bar.classList.toggle("over", uso > 1);
    const sem = semanasRestantes(k);
    $("hResumo").innerHTML = `Teto do mês ${R(teto)} · já gasto ${R(gi.v)} (${P(uso)} do teto, ${gi.fonte})<br>` +
      (sem > 0 ? `Dá ${R(rest / sem)} por semana nas ${sem.toFixed(1).replace(".", ",")} semanas que faltam · ` : "") +
      `à vista até ${R(p.vista)} · parcelas novas até ${R(p.parc)}/mês em até ${p.vezes || "—"}x (máx. ${p.simult || "—"} ao mesmo tempo)`;
    let man = null; try { man = localStorage.getItem("gasto-" + k); } catch (_) {}
    if (!onlyHero) $("gastoManual").value = man ?? "";

    // KPIs
    const fora = (p.fora || 0) - gastoNoMes(k, (g) => g.paga.startsWith("Conta") || g.paga === "Dinheiro");
    const vr = (p.vr || 0) - gastoNoMes(k, (g) => g.paga === "VR");
    const kp = [
      ["Fora do Inter disponível", R(fora), "13º livre + férias (Pix/débito)"],
      ["VR disponível", R(vr), "só alimentação"],
      ["Limite do Inter a configurar", R(p.limite), "fatura de " + lab(p.fatura) + ": " + R(p.faturaInter)],
      ["Sobra estimada do mês", R(pr.sobra[i]), "antes das compras novas " + (DATA.tipo[i] === "Realizado" ? "(real)" : "(previsto)")],
      ["Renda comprometida", P(pr.entradas[i] ? pr.comprom[i] / pr.entradas[i] : 0), "fixos + parcelas: " + R(pr.comprom[i]) + " das entradas"],
      ["Vai para a reserva no mês", R(p.reserva), P(DATA.meta ? p.reserva / DATA.meta : 0) + " da meta de " + R0(DATA.meta)],
    ];
    $("kpis").innerHTML = kp.map(([l, v, s]) => `<div class="kpi"><div class="l">${l}</div><div class="v">${v}</div><div class="s">${s}</div></div>`).join("");
    if (onlyHero) return;

    // Estimativa
    const ent = pr.entradas[i] || 0; const pc = (v) => (ent ? P(v / ent) : "—");
    const lin = [["Salário e extras", pr.renda[i]], ["13º salário", pr.d13[i]], ["Férias", pr.ferias[i]], ["Ajuste pós-férias", pr.ajuste[i]], ["Reembolsos de terceiros", pr.reemb[i]]]
      .filter(([, v]) => v).map(([t, v]) => `<tr><td>${t}</td><td>${R(v)}</td><td>${pc(v)}</td></tr>`).join("");
    const sai = [["Fixos (boletos, assinaturas)", pr.fixos[i]], ["Parcelas", pr.parcelas[i]], ["Demais parcelas e compras", pr.variaveis[i]]]
      .map(([t, v]) => `<tr><td>${t}</td><td class="neg">${R(v)}</td><td>${pc(-v)}</td></tr>`).join("");
    const fat = Object.entries(DATA.faturas).filter(([c]) => !F.cartao || c === F.cartao).map(([c, v]) => `<tr><td>Fatura ${c}</td><td>${R(v[i])}</td><td>${pc(v[i])}</td></tr>`).join("");
    $("tEstimativa").innerHTML = `<tr><th>${lab(k)} ${tipo}</th><th>Valor</th><th>% das entradas</th></tr>` + lin +
      `<tr class="total"><td>Entradas</td><td>${R(ent)}</td><td>100%</td></tr>` + sai +
      `<tr class="total"><td>Sobra do mês</td><td class="${pr.sobra[i] < 0 ? "neg" : "pos"}">${R(pr.sobra[i])}</td><td>${pc(pr.sobra[i])}</td></tr>` +
      `<tr><td colspan="3" class="muted small" style="text-align:left">Faturas incluídas nas saídas</td></tr>` + fat;

    // Reserva
    const rg = range(); const metaOk = rg.filter((m) => (DATA.plano[m] || {}).reserva >= DATA.meta - 0.01).length;
    const last = DATA.plano[DATA.meses[DATA.meses.length - 1]] || {};
    const aporte = p.reserva || 0, apPct = DATA.meta ? aporte / DATA.meta : 0;
    $("reserva").innerHTML =
      `<div class="kpis" style="grid-template-columns:repeat(auto-fit,minmax(150px,1fr))">
        <div class="kpi"><div class="l">Reserva hoje (Configurações)</div><div class="v">${R(DATA.reservaInformada)}</div></div>
        <div class="kpi"><div class="l">Prevista no fim de ${lab(k)}</div><div class="v">${R(p.reservaAc)}</div></div>
        <div class="kpi"><div class="l">Prevista em ${lab(DATA.meses[DATA.meses.length - 1])}</div><div class="v">${R(last.reservaAc)}</div></div>
      </div>
      <div class="meter"><div class="t"><span>Aporte de ${lab(k)}: ${R(aporte)}</span><span>${P(apPct)} da meta</span></div>
      <div class="bar"><span style="width:${Math.min(100, apPct * 100)}%;background:${apPct >= 0.999 ? css("--good") : css("--s4")}"></span></div></div>
      <p class="${metaOk === rg.length ? "icon-ok" : "icon-warn"}">${metaOk} de ${rg.length} meses do período batem a meta de ${R0(DATA.meta)}.</p>`;

    // Gráficos
    const L = rg.map(lab), ix = rg.map(idx), pal = palette();
    const den = (j) => (pct ? DATA.proj.entradas[j] || 1 : 1);
    const conv = (arr, byIdx) => arr.map((v, n) => (v == null ? null : pct ? v / den(byIdx[n]) : v));
    const fmt = pct ? (v) => P(v) : (v, axis) => (axis ? R0(v) : R0(v));
    const sel = rg.indexOf(k);
    const hl = (color) => rg.map((_, n) => (n === sel ? color : color + "B3"));
    const pv = (f) => rg.map((m) => (DATA.plano[m] || {})[f] || 0);

    draw("cGastar", "bar", L, [
      ds("Inter à vista", conv(pv("vista"), ix), hl(pal[0])), ds("Inter parcelado (limite)", conv(pv("parc"), ix), hl(pal[1])),
      ds("Fora do Inter", conv(pv("fora"), ix), hl(pal[2])), ds("VR", conv(pv("vr"), ix), hl(pal[3])),
    ], baseOpts(true, fmt));
    draw("cReserva", "line", L, [ds("Reserva acumulada", pv("reservaAc"), pal[0], { fill: false, tension: 0.2, borderWidth: 2, pointRadius: 4, pointBackgroundColor: pal[0] })],
      baseOpts(false, (v) => R0(v), {}));
    draw("cAporte", "bar", L, [
      ds("Vai para a reserva", pv("reserva"), hl(pal[2])),
      { type: "line", label: "Meta", data: rg.map(() => DATA.meta), borderColor: css("--muted"), borderDash: [6, 4], borderWidth: 2, pointRadius: 0, datalabels: { display: false } },
    ], baseOpts(false, (v) => R0(v)));
    draw("cLimite", "bar", L, [ds("Limite a configurar", pv("limite"), hl(pal[0])), ds("Fatura que vai chegar", pv("faturaInter"), hl(pal[1]))], baseOpts(false, (v) => R0(v)));
    const cards = Object.keys(DATA.faturas).filter((c) => !F.cartao || c === F.cartao);
    draw("cFaturas", "bar", L, cards.map((c, n) => ds(c, conv(ix.map((j) => DATA.faturas[c][j]), ix), pal[n % 5])), baseOpts(true, fmt));
    const pos = (a) => ix.map((j) => Math.abs(a[j]));
    draw("cRenda", "bar", L, [
      ds("Fixos", conv(pos(pr.fixos), ix), pal[0]), ds("Parcelas", conv(pos(pr.parcelas), ix), pal[1]),
      ds("Demais parcelas e compras", conv(pos(pr.variaveis), ix), pal[2]), ds("Sobra", conv(ix.map((j) => Math.max(0, pr.sobra[j])), ix), pal[3]),
    ], baseOpts(true, fmt));
    draw("cSobra", "bar", L, [ds("Sobra do mês", ix.map((j) => pr.sobra[j]), ix.map((j, n) => (pr.sobra[j] < 0 ? css("--bad") : n === sel ? pal[0] : pal[0] + "B3")))],
      baseOpts(false, (v) => R0(v)));

    // Dívidas
    let dv = DATA.dividas.map((d) => Object.assign({}, d, dividaInfo(d, k)));
    if (F.cartao) dv = dv.filter((d) => d.cartao === F.cartao);
    if (F.reemb === "com") dv = dv.filter((d) => d.reemb > 0); else if (F.reemb === "sem") dv = dv.filter((d) => !d.reemb);
    if (F.divida === "ativas") dv = dv.filter((d) => d.ativa);
    else if (F.divida === "3m") dv = dv.filter((d) => d.ativa && d.fim && diffM(k, d.fim) <= 2);
    dv.sort((a, b) => (a.fim || "9999") < (b.fim || "9999") ? -1 : (a.fim || "9999") > (b.fim || "9999") ? 1 : b.valor - a.valor);
    const totMes = dv.filter((d) => d.ativa).reduce((s, d) => s + d.valor, 0), totCusto = dv.filter((d) => d.ativa).reduce((s, d) => s + d.custo, 0);
    const totPagar = dv.reduce((s, d) => s + (d.aPagar || 0), 0);
    $("dResumo").textContent = `${dv.length} parcelamento(s) · ${R(totMes)}/mês em ${lab(k)} (${P(ent ? totMes / ent : 0)} das entradas) · meu custo ${R(totCusto)}/mês · ${R(totPagar)} ainda a pagar`;
    $("tDividas").innerHTML = `<tr><th>Descrição</th><th>Cartão</th><th>Parcela</th><th>Última</th><th>Restantes</th><th>Ainda a pagar</th><th>Reembolso/mês</th><th>Meu custo/mês</th><th>Status em ${lab(k)}</th></tr>` +
      dv.map((d) => `<tr><td>${d.desc}</td><td>${d.cartao}</td><td>${R(d.valor)}</td><td>${d.fim ? lab(d.fim) : "sem data"}</td><td>${d.rest == null ? "—" : d.rest}</td>` +
        `<td>${d.aPagar == null ? "—" : R(d.aPagar)}</td><td>${d.reemb ? R(d.reemb) + (d.quem ? " (" + d.quem + ")" : "") : "—"}</td><td>${R(d.custo)}</td><td>${d.ativa ? "ativa" : d.fim && d.fim < k ? "quitada" : "futura"}</td></tr>`).join("") +
      `<tr class="total"><td>Total</td><td></td><td>${R(totMes)}</td><td></td><td></td><td>${R(totPagar)}</td><td></td><td>${R(totCusto)}</td><td></td></tr>`;
    const dvc = dv.filter((d) => d.aPagar);
    draw("cDividas", "bar", dvc.map((d) => { const t = d.desc.replace(" (a identificar)", "").replace(" — ", " · "); return (/fim /.test(t) ? t : t + " · até " + lab(d.fim)).slice(0, 48); }),
      [ds("Ainda a pagar", dvc.map((d) => d.aPagar), dvc.map((d) => pal[Math.max(0, DATA.cartoes.indexOf(d.cartao)) % 5]))],
      baseOpts(false, (v) => R0(v), { indexAxis: "y" }));
    charts.cDividas.options.plugins.legend.display = false; charts.cDividas.options.scales.x.grid.display = true; charts.cDividas.update();

    // Categorias
    const ci = DATA.cat.meses.indexOf(k);
    const cats = Object.entries(DATA.cat.itens).map(([n, o]) => [n, ci >= 0 && typeof o.vals[ci] === "number" ? o.vals[ci] : o.media]).filter(([, v]) => v > 0.005).sort((a, b) => b[1] - a[1]);
    $("catNota").textContent = ci >= 0 ? `Meses realizados: valores de ${lab(k)}, já sem reembolsos.` : `${lab(k)} ainda não foi lançado: mostrando a média dos meses realizados.`;
    draw("cCat", "bar", cats.map((c) => c[0]), [ds(pct ? "% das entradas" : "Meu custo", cats.map((c) => (pct ? c[1] / (ent || 1) : c[1])), pal[0])],
      baseOpts(false, fmt, { indexAxis: "y" }));
    charts.cCat.options.plugins.legend.display = false; charts.cCat.update();
  }
  window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", () => render());
  load();
})();

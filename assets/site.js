/* FFL Magnificos site -- shared helpers. Plain JavaScript, no libraries. */
(function () {
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

  const NAV = [["index.html", "Home"], ["standings.html", "Standings"], ["owners.html", "Owners"], ["alltime.html", "All-Time"],
    ["records.html", "Records"], ["power.html", "Power"], ["h2h.html", "Head-to-Head"], ["draft.html", "Draft"],
    ["payouts.html", "Payouts"], ["money.html", "Money"], ["rules.html", "Rules"], ["photos.html", "Photos 🔒"]];

  async function getJSON(path) {
    const r = await fetch(path, { cache: "no-cache" });
    if (!r.ok) throw new Error(path + " (" + r.status + ")");
    return r.json();
  }

  const rec = (w, l, t) => (t ? w + "-" + l + "-" + t : w + "-" + l);
  const pct = (p) => (p == null ? "–" : (p * 100).toFixed(1) + "%");
  const num = (n, d) => (n == null ? "–" : Number(n).toLocaleString(undefined,
    { minimumFractionDigits: d == null ? 1 : d, maximumFractionDigits: d == null ? 1 : d }));
  const ownerLink = (slug, nick, scope) => slug
    ? '<a href="owner.html?o=' + encodeURIComponent(slug) + (scope ? "&scope=" + encodeURIComponent(scope) : "") + '">' + esc(nick) + "</a>" : esc(nick);
  const ordinal = (n) => { const s = ["th", "st", "nd", "rd"], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); };

  const RESULT_BADGE = (r) => {
    if (!r) return "";
    if (r === "Champion") return '<span class="badge gold">🏆 Champion</span>';
    if (r === "Runner-Up") return '<span class="badge blue">Runner-up</span>';
    if (r === "Missed Playoffs") return '<span class="badge out">Missed playoffs</span>';
    return '<span class="badge">' + esc(r) + "</span>";
  };

  function chrome(active, title, sub) {
    document.body.insertAdjacentHTML("afterbegin",
      '<header class="topbar"><div class="wrap"><a class="brand" href="index.html">🏈 <span id="brandName">FFL Magnificos</span></a>' +
      '<nav class="nav">' + NAV.map(([h, t]) =>
        '<a href="' + h + '"' + (h === active ? ' class="on"' : "") + ">" + t + "</a>").join("") +
      "</nav></div></header>");
    if (title) {
      $("main").insertAdjacentHTML("beforebegin",
        '<div class="page-head"><div class="wrap"><h1 id="pageTitle">' + esc(title) + "</h1>" +
        (sub ? '<p id="pageSub">' + esc(sub) + "</p>" : '<p id="pageSub"></p>') + "</div></div>");
    }
    document.body.insertAdjacentHTML("beforeend",
      '<footer><div class="wrap" id="foot">FFL Magnificos</div></footer>');
  }

  function fail(err) {
    if (window.console) console.error(err && err.stack ? err.stack : err);
    const m = $("main .wrap") || $("main");
    const local = location.protocol === "file:";
    m.innerHTML = '<div class="error"><b>Could not load the league data.</b><br>' +
      (local ? "This page has to be opened through a web server (GitHub Pages, or run <code>python -m http.server</code> in the site folder and visit <code>http://localhost:8000</code>)."
             : "Try refreshing in a minute.") +
      '<br><span class="note">' + esc(err && err.message) + "</span></div>";
  }

  /* boot(pageFile, title, subtitleFn, async (meta) => {...}) */
  async function boot(active, title, fn) {
    chrome(active, title);
    try {
      const meta = await getJSON("data/meta.json");
      $("#brandName").textContent = meta.leagueName.replace(/ League$/i, "");
      document.title = (title ? title + " · " : "") + meta.leagueName;
      $("#foot").textContent = meta.leagueName + " · data through " +
        meta.currentSeason + " week " + meta.currentWeek + " · updated " + meta.updated;
      await fn(meta);
    } catch (e) { fail(e); }
  }

  /* click-to-sort for tables: <th data-sort> and cells with data-v */
  function sortable(table) {
    const ths = $$("thead th", table);
    ths.forEach((th, i) => {
      if (th.dataset.sort === undefined) return;
      th.classList.add("sortable");
      th.addEventListener("click", () => {
        const asc = th.classList.contains("sorted") && !th.classList.contains("asc");
        ths.forEach((x) => x.classList.remove("sorted", "asc"));
        th.classList.add("sorted"); if (asc) th.classList.add("asc");
        const tb = $("tbody", table);
        const rows = $$("tr", tb);
        const val = (r) => { const c = r.children[i]; const v = c.dataset.v !== undefined ? c.dataset.v : c.textContent; const n = parseFloat(v); return isNaN(n) || v === "" ? String(v).toLowerCase() : n; };
        rows.sort((a, b) => { const x = val(a), y = val(b); const r = x < y ? -1 : x > y ? 1 : 0; return asc ? r : -r; });
        rows.forEach((r) => tb.appendChild(r));
      });
    });
  }

  /* "Which games count" buttons shared by the pages that offer them.
     scopes: [{key,label}], current: key, onPick(key) is called after the buttons redraw. */
  const SCOPE_NOTES = {
    reg: "Showing regular-season games only.",
    regpo: "Showing regular-season and playoff games (quarterfinals, semifinals and championship). The 3rd-place game and consolation games are not counted.",
    po: "Showing playoff games only (quarterfinals, semifinals and championship; the 3rd-place game is not counted).",
    all: "Showing every game, including the 3rd-place game and consolation-bracket games."
  };
  function scopePills(el, scopes, current, onPick) {
    el.innerHTML = scopes.map((s) => '<button class="pill' + (s.key === current ? " on" : "") + '" data-k="' + s.key + '">' + esc(s.label) + "</button>").join("");
    $$(".pill", el).forEach((b) => b.addEventListener("click", () => onPick(b.dataset.k)));
  }
  const pickScope = (scopes, def) => {
    const q = new URLSearchParams(location.search).get("scope");
    return (scopes.find((s) => s.key === q) || scopes.find((s) => s.key === def) || scopes[0]).key;
  };

  window.L = { SCOPE_NOTES, scopePills, pickScope, $, $$, esc, getJSON, rec, pct, num, ownerLink, ordinal, RESULT_BADGE, boot, sortable, chrome, fail };
})();

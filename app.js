(function (root) {
  'use strict';

  var GEST = 283, SECADO = 60, DX_MIN = 35;
  var pad = function (n) { return String(n).padStart(2, '0'); };

  /* ---------- fechas ---------- */
  function ymd(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function dateToISO(d) { return ymd(new Date(d.getTime() + 6 * 3600 * 1000)); }
  function parseDMY(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return isNaN(v) ? null : dateToISO(v);
    var m = String(v).trim().match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
    if (!m) return null;
    var y = +m[3]; if (y < 100) y += 2000;
    return y + '-' + pad(+m[2]) + '-' + pad(+m[1]);
  }
  function P(s) { var a = s.split('-').map(Number); return new Date(a[0], a[1] - 1, a[2], 12); }
  function diffDays(a, b) { return Math.round((P(a) - P(b)) / 864e5); }
  function addDays(s, n) { var d = P(s); d.setDate(d.getDate() + n); return ymd(d); }
  function fmt(s) { return s ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : ''; }
  function rel(n) {
    if (n === 0) return 'hoy';
    if (n === 1) return 'mañana';
    if (n === -1) return 'ayer';
    return n > 0 ? 'en ' + n + ' días' : 'hace ' + (-n) + ' días';
  }

  /* ---------- lectura del Excel ---------- */
  function txt(v) { if (v == null) return null; var s = String(v).trim(); return s === '' ? null : s; }
  function num(v) {
    if (v == null || v === '') return null;
    var n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
    return isFinite(n) ? n : null;
  }
  function numOrTxt(v) { var n = num(v); return n != null ? n : txt(v); }

  function parseLibreta(XLSX, buf) {
    var wb = XLSX.read(buf, { type: 'array', cellDates: true });
    for (var si = 0; si < wb.SheetNames.length; si++) {
      var ws = wb.Sheets[wb.SheetNames[si]];
      var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      var h = -1;
      for (var r = 0; r < Math.min(rows.length, 40); r++) {
        var cells = rows[r].map(function (c) { return c == null ? '' : String(c).trim(); });
        if (cells.indexOf('RP') >= 0 && cells.some(function (c) { return /^nomb/i.test(c); })) { h = r; break; }
      }
      if (h < 0) continue;
      var head = rows[h].map(function (c) { return c == null ? '' : String(c).trim(); });
      var idx = function (name, nth) {
        var c = 0; nth = nth || 0;
        for (var i = 0; i < head.length; i++) { if (head[i] === name) { if (c === nth) return i; c++; } }
        return -1;
      };
      var idxRe = function (re) { for (var i = 0; i < head.length; i++) { if (re.test(head[i])) return i; } return -1; };
      var col = {
        nombre: idxRe(/^nomb/i), edad: idx('Edad'), pFecha: idx('Fecha', 0), pN: idxRe(/^N[°º]$/),
        r1: idx('R1'), r2: idx('R2'), sFecha: idx('Fecha', 1), sN: idx('N Serv'), toro: idxRe(/^toro/i),
        eg: idx('E.G'), ipp: idx('Ipp'), tFecha: idx('Fecha', 2), result: idx('Result'), rech: idx('Rech'),
        rp: idx('RP'), idElec: idx('Id Elect'), rod: idx('Rod'), dLac: idx('D Lac'),
        prodLts: idx('Prod Lts'), mej: idxRe(/^mej/i)
      };
      var prCols = [];
      for (var k = 1; k <= 10; k++) { var pc = idx('Pr' + k); if (pc >= 0) prCols.push(pc); }
      var get = function (row, i) { return i >= 0 ? row[i] : null; };
      var animales = [];
      for (var j = h + 1; j < rows.length; j++) {
        var row = rows[j];
        var rpRaw = get(row, col.rp);
        if (rpRaw == null || String(rpRaw).trim() === '') continue;
        var rp = typeof rpRaw === 'number' ? String(Math.round(rpRaw)) : String(rpRaw).trim();
        var tfRaw = get(row, col.tFecha), resRaw = get(row, col.result);
        var tFecha = null, partoProb = null, result = txt(resRaw);
        if (typeof tfRaw === 'string' && /^pare/i.test(tfRaw.trim())) { partoProb = parseDMY(resRaw); result = null; }
        else tFecha = parseDMY(tfRaw);
        var estado = 'sin';
        if (partoProb) estado = 'prenada';
        else if (result && /^pre/i.test(result)) estado = 'prenada';
        else if (result && /^vac/i.test(result)) estado = 'vacia';
        else if (result && /^cic/i.test(result)) estado = 'ciclando';
        else if (result && /^apta/i.test(result)) estado = 'apta';
        else if (result) estado = 'otro';
        var pr = prCols.map(function (c) { return numOrTxt(row[c]); });
        while (pr.length && pr[pr.length - 1] == null) pr.pop();
        var idE = get(row, col.idElec);
        animales.push({
          rp: rp,
          nombre: txt(get(row, col.nombre)),
          edad: num(get(row, col.edad)),
          rod: txt(get(row, col.rod)),
          idElec: idE == null ? null : ((typeof idE === 'number' ? String(Math.round(idE)) : String(idE).trim()).replace(/^0+/, '') || null),
          pFecha: parseDMY(get(row, col.pFecha)),
          pN: num(get(row, col.pN)),
          r1: txt(get(row, col.r1)),
          r2: txt(get(row, col.r2)),
          sFecha: parseDMY(get(row, col.sFecha)),
          sN: num(get(row, col.sN)),
          toro: txt(get(row, col.toro)),
          eg: numOrTxt(get(row, col.eg)),
          ipp: numOrTxt(get(row, col.ipp)),
          tFecha: tFecha,
          result: result,
          partoProb: partoProb,
          estado: estado,
          rech: txt(get(row, col.rech)),
          pr: pr,
          dLac: num(get(row, col.dLac)),
          prodLts: num(get(row, col.prodLts)),
          mej: num(get(row, col.mej))
        });
      }
      animales.forEach(function (a) {
        if (a.estado !== 'prenada' && a.sFecha && a.pFecha && !/^celo/i.test(a.toro || '') && a.sFecha > a.pFecha && (!a.tFecha || a.sFecha > a.tFecha)) a.estado = 'servida';
      });
      if (!animales.length) continue;
      return { animales: animales, hoja: wb.SheetNames[si] };
    }
    throw new Error('No encontré la tabla de la libreta. Debe tener las columnas "Nomb./Orig." y "RP".');
  }

  /* ---------- cálculos ---------- */
  function esServicio(a) { return !!a.sFecha && !/^celo/i.test(a.toro || ''); }
  function derive(a, today) {
    var o = {
      partoProb: a.partoProb, partoEstimado: false, secado: null,
      diasParto: a.pFecha ? diffDays(today, a.pFecha) : null,
      diasServ: a.sFecha ? diffDays(today, a.sFecha) : null,
      pendienteTacto: false, servDespuesTacto: false
    };
    if (a.estado === 'prenada') {
      if (!o.partoProb && esServicio(a)) { o.partoProb = addDays(a.sFecha, GEST); o.partoEstimado = true; }
      if (o.partoProb) o.secado = addDays(o.partoProb, -SECADO);
    } else if (esServicio(a)) {
      o.servDespuesTacto = !!(a.tFecha && a.sFecha > a.tFecha);
      if ((!a.tFecha || a.tFecha < a.sFecha) && o.diasServ >= DX_MIN) o.pendienteTacto = true;
    }
    return o;
  }

  function norm(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim(); }
  function search(animales, q) {
    q = norm(q);
    if (!q) return [];
    var tokens = q.split(/\s+/), out = [];
    animales.forEach(function (a) {
      var n = a._n || (a._n = norm(a.nombre)), score = 0;
      if (a.rp === q) score = 100;
      else if (a.rp.indexOf(q) === 0) score = 80;
      else if (q.length >= 5 && a.idElec && a.idElec.slice(-q.length) === q) score = 70;
      else if (n.indexOf(q) >= 0) score = 60;
      else if (tokens.every(function (t) { return n.indexOf(t) >= 0 || a.rp.indexOf(t) === 0; })) score = 50;
      else if (q.length >= 3 && a.rp.indexOf(q) >= 0) score = 40;
      if (score) out.push([score, a]);
    });
    out.sort(function (x, y) {
      return (y[0] - x[0]) || (x[1].rp.length - y[1].rp.length) || x[1].rp.localeCompare(y[1].rp, 'es', { numeric: true });
    });
    return out.map(function (x) { return x[1]; });
  }
  function conteo(animales) {
    var c = { prenada: 0, servida: 0, vacia: 0, ciclando: 0, apta: 0, sin: 0, otro: 0 };
    animales.forEach(function (a) { c[a.estado] = (c[a.estado] || 0) + 1; });
    return c;
  }

  var api = { parseLibreta: parseLibreta, derive: derive, search: search, conteo: conteo, fmt: fmt, rel: rel, diffDays: diffDays, addDays: addDays, ymd: ymd, GEST: GEST, SECADO: SECADO, DX_MIN: DX_MIN };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof document !== 'undefined') initUI(api);

  /* ================= interfaz ================= */
  function initUI(A) {
    var KEY = 'libreta:v1', KEYR = 'libreta:recientes';
    var DB = { animales: [], cargada: null, archivo: '' };
    var state = { q: '', rp: null, error: '', persisted: true };
    var $ = function (s) { return document.querySelector(s); };
    var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
    var today = function () { return A.ymd(new Date()); };
    var nf = function (n) { return n == null ? '' : Number(n).toLocaleString('es-UY', { maximumFractionDigits: 0 }); };

    var LABEL = { prenada: 'Preñada', servida: 'Servida', vacia: 'Vacía', ciclando: 'Ciclando', apta: 'Apta para servicio', sin: 'Sin tacto', otro: 'Otros' };

    function loadDB() {
      try { var r = localStorage.getItem(KEY); if (r) { var o = JSON.parse(r); if (o && Array.isArray(o.animales)) DB = o; } } catch (e) { state.persisted = false; }

    }
    function saveDB() {
      try { localStorage.setItem(KEY, JSON.stringify(DB)); state.persisted = true; } catch (e) { state.persisted = false; }
    }
    function recientes() { try { return JSON.parse(localStorage.getItem(KEYR)) || []; } catch (e) { return []; } }
    function pushReciente(rp) {
      try { var l = recientes().filter(function (x) { return x !== rp; }); l.unshift(rp); localStorage.setItem(KEYR, JSON.stringify(l.slice(0, 6))); } catch (e) {}
    }
    function byRp(rp) { for (var i = 0; i < DB.animales.length; i++) if (DB.animales[i].rp === rp) return DB.animales[i]; return null; }

    function toast(msg) {
      var t = $('#toast'); t.textContent = msg; t.classList.add('on');
      clearTimeout(toast._t); toast._t = setTimeout(function () { t.classList.remove('on'); }, 3200);
    }

    /* ----- archivo ----- */
    function handleFile(f) {
      if (!f) return;
      var fr = new FileReader();
      fr.onload = function () {
        try {
          var res = A.parseLibreta(root.XLSX, fr.result);
          DB = { animales: res.animales, cargada: today(), archivo: f.name };
          saveDB(); state.q = ''; state.rp = null; state.error = '';
          toast(res.animales.length + ' animales cargados');
        } catch (e) { state.error = e.message || 'No pude leer el archivo.'; }
        render();
      };
      fr.onerror = function () { state.error = 'No pude abrir el archivo.'; render(); };
      fr.readAsArrayBuffer(f);
    }

    /* ----- navegación ----- */
    function openFicha(rp) {
      state.rp = rp; pushReciente(rp);
      try { location.hash = 'ficha-' + rp; } catch (e) {}
      render(); window.scrollTo(0, 0);
    }
    function back() {
      state.rp = null;
      try { location.hash = ''; } catch (e) {}
      render();
    }
    window.addEventListener('hashchange', function () {
      var m = /^#ficha-(\d+)$/.exec(location.hash);
      state.rp = m && byRp(m[1]) ? m[1] : null; render();
    });

    /* ----- vistas ----- */
    function labelOf(a) { return a.estado === 'otro' ? (a.result || 'Otro') : LABEL[a.estado]; }
    function chip(a) { return '<span class="chip ' + a.estado + '">' + esc(labelOf(a)) + '</span>'; }

    function rowHTML(a) {
      var d = A.derive(a, today()), sub = '';
      if (a.estado === 'prenada' && d.partoProb) sub = (d.partoEstimado ? 'Parto estimado ' : 'Parto probable ') + A.fmt(d.partoProb);
      else if (a.tFecha) sub = 'Tacto ' + A.fmt(a.tFecha);
      else if (a.sFecha) sub = (/^celo/i.test(a.toro || '') ? 'Celo ' : 'Servicio ') + A.fmt(a.sFecha);
      return '<button class="row" data-rp="' + esc(a.rp) + '"><span class="rp">' + esc(a.rp) + '</span>' +
        '<span class="who">' + esc(a.nombre || 'Sin nombre') + '</span>' + chip(a) +
        '<span class="sub">' + esc(sub) + (a.rech ? ' <span class="dsc">Descarte programado</span>' : '') + '</span></button>';
    }

    function emptyHTML() {
      return '<section class="empty"><h2>Cargá tu libreta</h2>' +
        '<p>Elegí el archivo Excel (.xls o .xlsx) exportado de SysCaw. Se lee en este teléfono y no se sube a ningún servidor.</p>' +
        (state.error ? '<p class="error" role="alert">' + esc(state.error) + '</p>' : '') +
        '<button class="btn primary big" data-a="load">Elegir archivo</button></section>';
    }

    function searchHTML() {
      var c = A.conteo(DB.animales);
      var stats = ['prenada', 'servida', 'vacia', 'ciclando', 'apta', 'sin', 'otro'].filter(function (k) { return c[k]; }).map(function (k) {
        return '<div class="stat"><b>' + c[k] + '</b><span>' + LABEL[k] + '</span></div>';
      }).join('');
      var nDesc = DB.animales.filter(function (a) { return a.rech; }).length;
      if (nDesc) stats += '<div class="stat dscstat"><b>' + nDesc + '</b><span>Descarte programado</span></div>';
      return '<div class="searchbox"><label class="sr" for="q">Caravana o nombre</label>' +
        '<input id="q" type="search" enterkeyhint="search" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="Caravana (RP) o nombre" value="' + esc(state.q) + '">' +
        '</div><div id="results"></div><div class="stats" id="stats">' + stats + '</div>';
    }

    function updateResults() {
      var box = $('#results'); if (!box) return;
      var q = state.q.trim(), html = '';
      if (!q) {
        var rec = recientes().map(byRp).filter(Boolean);
        $('#stats').hidden = false;
        html = rec.length ? '<h2 class="h">Consultadas hace poco</h2>' + rec.map(rowHTML).join('') : '<p class="hint">Escribí la caravana de la vaca para ver su ficha.</p>';
      } else {
        $('#stats').hidden = true;
        var res = A.search(DB.animales, q);
        html = res.length
          ? '<h2 class="h">' + res.length + (res.length === 1 ? ' resultado' : ' resultados') + (res.length > 25 ? ', mostrando 25' : '') + '</h2>' + res.slice(0, 25).map(rowHTML).join('')
          : '<p class="hint">No hay animales con "' + esc(q) + '". Probá con otra caravana o con parte del nombre.</p>';
      }
      box.innerHTML = html;
    }

    function kv(rows) {
      var items = rows.filter(function (r) { return r[1] != null && r[1] !== ''; });
      if (!items.length) return '';
      return '<dl class="kv">' + items.map(function (r) { return '<div><dt>' + r[0] + '</dt><dd>' + r[1] + '</dd></div>'; }).join('') + '</dl>';
    }
    function withRel(iso, today_) { return A.fmt(iso) + ' <small>' + A.rel(A.diffDays(iso, today_)) + '</small>'; }

    function bars(pr) {
      var vals = pr.slice().reverse(), n = vals.length;
      var nums = vals.filter(function (v) { return typeof v === 'number'; });
      if (!nums.length) return '';
      var max = Math.max.apply(null, nums), W = 300, H = 118, padx = 6, gap = 6;
      var bw = (W - padx * 2 - gap * (n - 1)) / n, s = '';
      vals.forEach(function (v, i) {
        var x = padx + i * (bw + gap), label = 'Pr' + (n - i);
        if (typeof v === 'number') {
          var h = Math.max(3, (v / max) * (H - 40)), y = H - 18 - h;
          s += '<rect class="bar' + (i === n - 1 ? ' last' : '') + '" x="' + x.toFixed(1) + '" y="' + y.toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + h.toFixed(1) + '" rx="3"/>' +
            '<text class="val" x="' + (x + bw / 2).toFixed(1) + '" y="' + (y - 4).toFixed(1) + '" text-anchor="middle">' + v + '</text>';
        } else if (v) {
          s += '<text class="val" x="' + (x + bw / 2).toFixed(1) + '" y="' + (H - 24) + '" text-anchor="middle">' + esc(v) + '</text>';
        }
        s += '<text class="lbl" x="' + (x + bw / 2).toFixed(1) + '" y="' + (H - 4) + '" text-anchor="middle">' + label + '</text>';
      });
      return '<svg class="bars" viewBox="0 0 ' + W + ' ' + H + '" role="img" aria-label="Litros por control lechero">' + s + '</svg>' +
        '<p class="cap">Litros por control, del más viejo al más reciente (Pr1).</p>';
    }

    function fichaHTML(a) {
      var t = today(), d = A.derive(a, t), banner = '';
      var lines = [];
      if (a.estado === 'prenada') {
        if (d.partoProb) {
          lines.push((d.partoEstimado ? 'Parto estimado (servicio + ' + A.GEST + ' días): ' : 'Parto probable: ') + '<b>' + A.fmt(d.partoProb) + '</b> <small>' + A.rel(A.diffDays(d.partoProb, t)) + '</small>');
          var ds = A.diffDays(d.secado, t), dp = A.diffDays(d.partoProb, t);
          lines.push('Secado (' + A.SECADO + ' días antes): <b>' + A.fmt(d.secado) + '</b> <small>' + (ds < 0 && dp >= 0 ? 'ya pasó, hace ' + (-ds) + ' días' : A.rel(ds)) + '</small>');
          if (dp < 0) lines.push('<span class="warn">La fecha de parto probable ya pasó.</span>');
        }
        if (a.tFecha) lines.push('Tacto ' + A.fmt(a.tFecha));
      } else {
        if (a.tFecha) lines.push((a.estado === 'servida' ? 'Último tacto: ' : 'Tacto: ') + '<b>' + A.fmt(a.tFecha) + '</b> <small>' + A.rel(A.diffDays(a.tFecha, t)) + '</small>');
        if (a.estado === 'servida') lines.unshift('Servicio: <b>' + A.fmt(a.sFecha) + '</b> <small>' + A.rel(A.diffDays(a.sFecha, t)) + '</small>');
        if (d.pendienteTacto) lines.push('<span class="warn">Pasaron ' + d.diasServ + ' días del servicio: corresponde tacto.</span>');
        if (a.estado === 'sin' && !a.sFecha) lines.push('Sin tacto ni servicio registrados.');
      }
      banner = '<section class="banner ' + a.estado + '"><div class="state">' + esc(labelOf(a)) + '</div>' +
        lines.map(function (l) { return '<p>' + l + '</p>'; }).join('') + '</section>';

      var chips = [a.rod ? 'Rodeo ' + esc(a.rod) : '', a.edad != null ? a.edad + ' años' : '', a.pN != null ? 'Parto ' + a.pN : ''].filter(Boolean)
        .map(function (c) { return '<span class="tag">' + c + '</span>'; }).join('');

      var toro = a.toro ? (a.toro === '???..' ? 'Sin dato' : esc(a.toro)) : null;
      var egTxt = a.eg == null ? null : (typeof a.eg === 'number' ? a.eg + ' días' : esc(a.eg));
      var servLabel = /^celo/i.test(a.toro || '') ? 'Último celo' : 'Último servicio';
      var repro = kv([
        ['Último parto', a.pFecha ? withRel(a.pFecha, t) : null],
        ['Resultado del parto (R1, R2)', a.r1 ? esc(a.r1) + (a.r2 ? ' y ' + esc(a.r2) : '') : null],
        [servLabel, a.sFecha ? withRel(a.sFecha, t) : null],
        ['N° de servicio', a.sN != null ? a.sN : null],
        ['Toro o embrión', toro],
        ['E.G', egTxt],
        ['Tacto', a.tFecha ? A.fmt(a.tFecha) : null],
        ['Resultado del tacto', a.result ? esc(a.result) : (a.partoProb ? 'Preñada' : null)],
        ['Ipp', a.ipp != null ? esc(a.ipp) : null]
      ]);
      var prodRows = kv([
        ['Días en lactancia', a.dLac != null ? nf(a.dLac) : null],
        ['Producción total', a.prodLts != null ? nf(a.prodLts) + ' lts' : null],
        ['Mej 305', a.mej != null ? nf(a.mej) + ' lts' : null]
      ]);
      var ident = kv([
        ['RP', esc(a.rp)], ['Nombre / origen', esc(a.nombre)], ['Id electrónica', esc(a.idElec)]
      ]);
      return '<button class="backbtn" data-a="back"><svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true"><path d="M15 5l-7 7 7 7" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>Volver</button>' +
        '<header class="fh"><div class="big">' + esc(a.rp) + '</div><div class="nm">' + esc(a.nombre || 'Sin nombre') + '</div><div class="tags">' + chips + '</div></header>' +
        banner +
        (a.rech ? '<section class="descarte"><b>Descarte programado</b><span>Código en Rech: ' + esc(a.rech) + '</span></section>' : '') +
        '<section class="card"><h2>Reproducción</h2>' + repro + '</section>' +
        '<section class="card"><h2>Producción</h2>' + prodRows + bars(a.pr) + '</section>' +
        '<section class="card"><h2>Identificación</h2>' + ident + '</section>';
    }

    function render() {
      var main = $('#main'), meta = $('#meta'), btn = $('#btnLoad');
      if (!DB.animales.length) {
        meta.textContent = 'Sin libreta cargada'; btn.hidden = true;
        main.innerHTML = emptyHTML(); return;
      }
      btn.hidden = false;
      meta.textContent = DB.animales.length + ' animales, cargada el ' + A.fmt(DB.cargada) + (state.persisted ? '' : ' (no se guarda en este navegador)');
      if (state.rp) {
        var a = byRp(state.rp);
        if (a) { main.innerHTML = fichaHTML(a); return; }
        state.rp = null;
      }
      main.innerHTML = (state.error ? '<p class="error" role="alert">' + esc(state.error) + '</p>' : '') + searchHTML();
      updateResults();
    }

    /* ----- eventos ----- */
    document.addEventListener('click', function (e) {
      var row = e.target.closest('[data-rp]');
      if (row) { openFicha(row.getAttribute('data-rp')); return; }
      var b = e.target.closest('[data-a]'); if (!b) return;
      if (b.dataset.a === 'load') $('#file').click();
      if (b.dataset.a === 'back') back();
    });
    document.addEventListener('input', function (e) {
      if (e.target.id === 'q') { state.q = e.target.value; updateResults(); }
    });
    document.addEventListener('keydown', function (e) {
      if (e.target.id === 'q' && e.key === 'Enter') {
        var res = A.search(DB.animales, state.q);
        if (res.length) { e.target.blur(); openFicha(res[0].rp); }
      }
    });
    $('#btnLoad').addEventListener('click', function () { $('#file').click(); });
    $('#file').addEventListener('change', function (e) { handleFile(e.target.files[0]); e.target.value = ''; });

    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) navigator.serviceWorker.register('sw.js').catch(function () {});
    loadDB();
    (function () { var m = /^#ficha-(\d+)$/.exec(location.hash); if (m && byRp(m[1])) state.rp = m[1]; })();
    render();
  }
})(typeof window !== 'undefined' ? window : this);

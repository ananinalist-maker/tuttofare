/* ==========================================================================
   TUTTOFARE 2026 — интерактивы лендингов
   1) scroll-expand hero  2) интерактивный таймлайн  3) калькулятор preventivo
   Чистый JS, без зависимостей и сборки.
   ========================================================================== */
(function () {
  'use strict';

  var reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  /* ---------- аналитика: GA4 + запас под Meta Pixel ---------- */
  function track(name, params) {
    try { if (typeof gtag === 'function') gtag('event', name, params || {}); } catch (e) {}
    try { if (typeof fbq === 'function') fbq('trackCustom', name, params || {}); } catch (e) {}
  }
  window.tfTrack = track;

  var clamp = function (v, a, b) { return Math.min(Math.max(v, a), b); };
  // Итальянский формат с точками. toLocaleString('it-IT') не группирует
  // четырёхзначные числа (8800 вместо 8.800), из-за чего диапазон
  // «8800 – 12.400» выглядит рассогласованно. Группируем вручную.
  var euro = function (n) {
    var v = String(Math.round(n / 100) * 100);
    return v.replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  };

  /* ======================================================================
     1) SCROLL-EXPAND HERO
     Прогресс берём из позиции скролла внутри .xwrap (sticky), а не
     перехватом колеса: не ломает нативную прокрутку, тачпад и доступность.
     ====================================================================== */
  function initHero() {
    var wrap = document.querySelector('.xwrap');
    if (!wrap) return;
    var media = wrap.querySelector('.xhero__media');
    var bg = wrap.querySelector('.xhero__bg');
    var w1 = wrap.querySelector('.xhero__w1');
    var w2 = wrap.querySelector('.xhero__w2');
    var sub = wrap.querySelector('.xhero__sub');
    var hint = wrap.querySelector('.xhero__hint');
    if (!media) return;

    var mob = window.innerWidth < 768;
    var ticking = false;

    function render() {
      ticking = false;
      var rect = wrap.getBoundingClientRect();
      var total = wrap.offsetHeight - window.innerHeight;
      var p = total > 0 ? clamp(-rect.top / total, 0, 1) : 0;
      var e = p * p * (3 - 2 * p); // мягкое ускорение

      media.style.width = (mob ? 240 + e * 620 : 380 + e * 1180) + 'px';
      media.style.height = (mob ? 340 + e * 260 : 440 + e * 380) + 'px';
      if (bg) bg.style.opacity = String(1 - e * 0.85);

      var shift = e * (mob ? 42 : 46);
      if (w1) w1.style.transform = 'translateX(-' + shift + 'vw)';
      if (w2) w2.style.transform = 'translateX(' + shift + 'vw)';
      if (sub) sub.style.opacity = String(clamp(1 - e * 2.2, 0, 1));
      if (hint) hint.style.opacity = String(clamp(1 - e * 3, 0, 1));

      if (p > 0.98 && !wrap.dataset.done) {
        wrap.dataset.done = '1';
        track('hero_expanded');
      }
    }

    function onScroll() {
      if (!ticking) { ticking = true; requestAnimationFrame(render); }
    }

    if (reduce) {
      media.style.width = 'min(1180px,92vw)';
      media.style.height = 'min(70vh,640px)';
      if (bg) bg.style.opacity = '.2';
      return;
    }

    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('resize', function () { mob = window.innerWidth < 768; render(); }, { passive: true });
    render();
  }

  /* ======================================================================
     2) ИНТЕРАКТИВНЫЙ ТАЙМЛАЙН
     Данные — в <script type="application/json" id="tl-data"> на странице.
     ====================================================================== */
  function initTimeline() {
    var root = document.querySelector('[data-timeline]');
    var raw = document.getElementById('tl-data');
    if (!root || !raw) return;

    var phases;
    try { phases = JSON.parse(raw.textContent); } catch (e) { return; }
    if (!phases || !phases.length) return;

    var mode = 'do';
    var idx = 0;

    var bar = root.querySelector('.tl__bar');
    var body = root.querySelector('.tl__body');
    var prog = root.querySelector('.tl__prog i');
    var btns = root.querySelectorAll('.tl__switch button');

    phases.forEach(function (ph, i) {
      var b = document.createElement('button');
      b.type = 'button';
      b.innerHTML = '<span class="tn">' + ph.days + '</span><span class="td">' + ph.name + '</span>';
      b.addEventListener('click', function () { idx = i; paint(); track('timeline_phase', { phase: ph.name }); });
      bar.appendChild(b);
    });

    btns.forEach(function (b) {
      b.addEventListener('click', function () {
        mode = b.dataset.mode;
        btns.forEach(function (x) { x.classList.toggle('on', x === b); });
        paint();
        track('timeline_mode', { mode: mode });
      });
    });

    function paint() {
      var ph = phases[idx];
      Array.prototype.forEach.call(bar.children, function (c, i) { c.classList.toggle('on', i === idx); });
      var txt = mode === 'do' ? ph.do : ph.get;
      body.innerHTML =
        '<h4>' + ph.name + '</h4>' +
        '<p class="tl__days">' + ph.days + ' · ' + ph.range + '</p>' +
        '<p>' + txt + '</p>' +
        (ph.pause && mode === 'do' ? '<div class="tl__pause"><b>Tempo tecnico:</b> ' + ph.pause + '</div>' : '');
      if (prog) prog.style.width = ((idx + 1) / phases.length * 100) + '%';
      var active = bar.children[idx];
      if (active && active.scrollIntoView) active.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: reduce ? 'auto' : 'smooth' });
    }
    paint();
  }

  /* ======================================================================
     3) КАЛЬКУЛЯТОР PREVENTIVO
     Цены — стартовая линия 2026 (нижняя граница рынка Лацио).
     Всегда отдаёт ВИЛКУ и всегда с оговоркой: точная цена только
     после видео-сметы и техаудита.
     ====================================================================== */
  var MODEL = {
    'bagno-standard': { base: [7900, 11000], perM2: [450, 700], baseM2: 4, days: [24, 28], label: 'Bagno chiavi in mano' },
    'bagno-grande':   { base: [12900, 16000], perM2: [700, 950], baseM2: 4, days: [28, 35], label: 'Bagno · grande formato' },
    'completa':       { base: [0, 0], perM2: [700, 950], baseM2: 0, days: [60, 90], label: 'Ristrutturazione completa' },
    'affitto':        { base: [0, 0], perM2: [480, 650], baseM2: 0, days: [35, 50], label: 'Immobile pronto da affittare' }
  };
  var ZONE = { roma: 1, viterbo: 0.72 };
  var STATO = { completo: 1, superficie: 0.72 };

  function initCalc() {
    var root = document.querySelector('[data-calc]');
    if (!root) return;

    var out = root.querySelector('.calc__out');
    var resEl = out.querySelector('.res');
    var subEl = out.querySelector('.resub');
    var rowsEl = out.querySelector('.calc__rows');
    var range = root.querySelector('input[type=range]');
    var rangeOut = root.querySelector('.rng output');
    var waBtn = out.querySelector('[data-wa]');

    function val(n) {
      var el = root.querySelector('input[name="' + n + '"]:checked');
      return el ? el.value : null;
    }

    function calc() {
      var tipo = val('tipo') || 'bagno-standard';
      var zona = val('zona') || 'roma';
      var stato = val('stato') || 'completo';
      var m2 = parseInt(range.value, 10);
      var m = MODEL[tipo];
      var z = ZONE[zona];
      var s = tipo.indexOf('bagno') === 0 ? STATO[stato] : 1;

      var extra = Math.max(0, m2 - m.baseM2);
      var lo = (m.base[0] + extra * m.perM2[0]) * z * s;
      var hi = (m.base[1] + extra * m.perM2[1]) * z * s;

      var d0 = m.days[0], d1 = m.days[1];
      if (tipo.indexOf('bagno') === 0 && stato === 'superficie') { d0 = Math.round(d0 * 0.6); d1 = Math.round(d1 * 0.6); }

      rangeOut.textContent = m2 + ' m²';
      resEl.textContent = euro(lo) + ' – ' + euro(hi) + ' €';
      subEl.textContent = 'Stima orientativa · prezzi di lancio 2026 · IVA 10% esclusa';

      rowsEl.innerHTML = '';
      [
        ['Intervento', m.label],
        ['Superficie', m2 + ' m²'],
        ['Zona', zona === 'roma' ? 'Roma e provincia' : 'Viterbo e provincia'],
        ['Durata stimata', d0 + '–' + d1 + ' giorni lavorativi'],
        ['Nel prezzo', 'Contratto a corpo, video report, documenti, Buon Vicinato']
      ].forEach(function (r) {
        var d = document.createElement('div');
        d.innerHTML = '<span>' + r[0] + '</span><b>' + r[1] + '</b>';
        rowsEl.appendChild(d);
      });

      if (waBtn) {
        var msg = 'Buongiorno, ho usato il calcolatore sul sito.%0A' +
          'Intervento: ' + encodeURIComponent(m.label) + '%0A' +
          'Superficie: ' + m2 + ' m²%0A' +
          'Zona: ' + (zona === 'roma' ? 'Roma' : 'Viterbo') + '%0A' +
          'Stima vista: ' + encodeURIComponent(euro(lo) + '–' + euro(hi) + ' €') + '%0A' +
          'Vorrei il preventivo video.';
        waBtn.setAttribute('href', 'https://wa.me/393894481853?text=' + msg);
      }
      return { tipo: tipo, m2: m2, lo: lo, hi: hi };
    }

    root.addEventListener('change', function () { calc(); });
    if (range) range.addEventListener('input', calc);
    if (waBtn) waBtn.addEventListener('click', function () {
      var r = calc();
      track('calc_to_whatsapp', { tipo: r.tipo, m2: r.m2, stima_min: Math.round(r.lo) });
    });

    var once = false;
    root.addEventListener('input', function () {
      if (!once) { once = true; track('calc_started'); }
    });

    calc();
  }

  /* ---------- клики в WhatsApp по всей странице ---------- */
  function initWa() {
    document.addEventListener('click', function (e) {
      var a = e.target.closest && e.target.closest('a[href*="wa.me"]');
      if (!a) return;
      track('whatsapp_click', { position: a.dataset.pos || 'page', page: document.body.dataset.page || '' });
    });
  }

  /* ---------- стрелка «наверх» + плавающая кнопка калькулятора ---------- */
  function initFloats() {
    var top = document.querySelector('.toTop');
    var fab = document.querySelector('.calcFab');
    var calc = document.getElementById('calcolatore');
    if (!top && !fab) return;

    if (top) {
      top.addEventListener('click', function (e) {
        e.preventDefault();
        window.scrollTo({ top: 0, behavior: reduce ? 'auto' : 'smooth' });
        track('scroll_to_top');
      });
    }
    if (fab) {
      fab.addEventListener('click', function () { track('calc_fab_click'); });
    }

    var ticking = false;
    function upd() {
      ticking = false;
      var y = window.scrollY || window.pageYOffset;
      if (top) top.classList.toggle('on', y > 700);
      if (fab) {
        // прячем кнопку, когда калькулятор и так на экране
        var near = false;
        if (calc) {
          var r = calc.getBoundingClientRect();
          near = r.top < window.innerHeight * 0.85 && r.bottom > 0;
        }
        fab.classList.toggle('on', y > 500 && !near);
      }
    }
    window.addEventListener('scroll', function () {
      if (!ticking) { ticking = true; requestAnimationFrame(upd); }
    }, { passive: true });
    upd();
  }

  /* ======================================================================
     4) АНИМАЦИИ
     Классы навешиваются здесь, а не в разметке: если скрипт не выполнится,
     контент виден сразу и ничего не пропадает.
     ====================================================================== */
  // первый экран внутренних страниц: каскадный вход вместо появления по скроллу
  function markPageHero() {
    if (reduce) return;
    if (document.querySelector('.hfeat')) return;
    var s = document.querySelector('body > section');
    if (!s) return;
    var w = s.querySelector('.wrap');
    if (!w) return;
    s.classList.add('pgSec');
    w.classList.add('pgHero');
  }

  function initReveal() {
    if (reduce || !('IntersectionObserver' in window)) return;

    var marked = [];

    // одиночные элементы
    var SOLO = [
      'section:not(.hfeat):not(.pgSec) .kicker',
      'section:not(.hfeat):not(.pgSec) h2',
      'section:not(.hfeat):not(.pgSec) > .wrap > .lede',
      '.tl', '.calc', '.launch', '.tw', '.smallprint'
    ].join(',');
    Array.prototype.forEach.call(document.querySelectorAll(SOLO), function (el) {
      el.classList.add('rv'); marked.push(el);
    });

    // группы — дети появляются каскадом
    var GROUPS = '.grid, .paths, .stats, .hmos__grid, .steps, .shield, .docs, .media-grid';
    Array.prototype.forEach.call(document.querySelectorAll(GROUPS), function (box) {
      Array.prototype.forEach.call(box.children, function (ch, i) {
        ch.classList.add('rv', 'rv-' + Math.min(i + 1, 6));
        marked.push(ch);
      });
    });

    // FAQ — каскадом среди соседей
    var faqs = document.querySelectorAll('details');
    Array.prototype.forEach.call(faqs, function (d, i) {
      d.classList.add('rv', 'rv-' + Math.min((i % 6) + 1, 6));
      marked.push(d);
    });

    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (e) {
        if (e.isIntersecting) { e.target.classList.add('in'); io.unobserve(e.target); }
      });
    }, { rootMargin: '0px 0px -8% 0px', threshold: 0.05 });

    marked.forEach(function (el) { io.observe(el); });

    // страховка: что не поймал наблюдатель — показать через 2,5 с
    setTimeout(function () {
      marked.forEach(function (el) { el.classList.add('in'); });
    }, 2500);
  }

  /* --- параллакс на первом экране + состояние шапки --- */
  function initHeroMotion() {
    var nav = document.querySelector('.nav');
    var frame = document.querySelector('.hfeat__frame');
    var img = frame ? frame.querySelector('img') : null;
    var inner = frame ? frame.querySelector('.hfeat__in') : null;
    if (!nav && !frame) return;

    var ticking = false;
    function upd() {
      ticking = false;
      var y = window.scrollY || window.pageYOffset;
      if (nav) nav.classList.toggle('scrolled', y > 40);
      if (reduce || !frame) return;
      var h = frame.offsetHeight || 1;
      var p = clamp(y / h, 0, 1);
      // translate отдельным свойством, чтобы не спорить с CSS-анимацией kenburns
      if (img) img.style.translate = '0 ' + (y * 0.16) + 'px';
      if (inner) {
        inner.style.translate = '0 ' + (y * 0.06) + 'px';
        inner.style.opacity = String(clamp(1 - p * 1.25, 0, 1));
      }
    }
    window.addEventListener('scroll', function () {
      if (!ticking) { ticking = true; requestAnimationFrame(upd); }
    }, { passive: true });
    upd();
  }

  function boot() {
    markPageHero();
    initHero(); initTimeline(); initCalc(); initWa(); initFloats();
    initReveal(); initHeroMotion();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();

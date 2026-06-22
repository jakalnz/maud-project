/*
 * Shared theme switcher for all MAud pages. Persists choice in localStorage
 * under 'maud_theme' so it carries across pages (dashboard, forms, home).
 *
 * Call applyStoredMaudTheme() as early as possible (before paint) and
 * injectMaudThemePicker() once the body exists.
 */
(function () {
  var THEMES = [
    { id: 'default',     label: 'Default',          swatch: '#185FA5' },
    { id: 'dark',        label: 'Dark',              swatch: '#2a2a2a' },
    { id: 'low-vision',  label: 'Low Vision',        swatch: '#000000' },
    { id: 'colorblind',  label: 'Colour-blind Safe', swatch: '#0077BB' },
    { id: 'uoa',         label: 'UoA Colours',       swatch: '#00457D' }
  ];

  function applyStoredMaudTheme() {
    var stored = localStorage.getItem('maud_theme') || 'default';
    document.documentElement.setAttribute('data-theme', stored);
  }

  function injectMaudThemePicker() {
    if (document.querySelector('.maud-theme-picker')) return;
    var current = localStorage.getItem('maud_theme') || 'default';

    var wrap = document.createElement('div');
    wrap.className = 'maud-theme-picker';

    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'maud-theme-picker-btn';
    btn.setAttribute('aria-label', 'Change colour theme');
    btn.setAttribute('title', 'Change colour theme');
    btn.textContent = '◑'; // half-shaded circle, theme-neutral icon

    var menu = document.createElement('div');
    menu.className = 'maud-theme-picker-menu';

    THEMES.forEach(function (t) {
      var opt = document.createElement('div');
      opt.className = 'maud-theme-picker-opt' + (t.id === current ? ' active' : '');
      opt.setAttribute('role', 'button');
      opt.dataset.themeId = t.id;
      var sw = document.createElement('span');
      sw.className = 'maud-theme-picker-swatch';
      sw.style.background = t.swatch;
      opt.appendChild(sw);
      opt.appendChild(document.createTextNode(t.label));
      opt.addEventListener('click', function () {
        localStorage.setItem('maud_theme', t.id);
        document.documentElement.setAttribute('data-theme', t.id);
        menu.querySelectorAll('.maud-theme-picker-opt').forEach(function (o) {
          o.classList.toggle('active', o.dataset.themeId === t.id);
        });
        menu.classList.remove('open');
      });
      menu.appendChild(opt);
    });

    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      menu.classList.toggle('open');
    });
    document.addEventListener('click', function () { menu.classList.remove('open'); });

    wrap.appendChild(menu);
    wrap.appendChild(btn);
    document.body.appendChild(wrap);
  }

  applyStoredMaudTheme();
  window.injectMaudThemePicker = injectMaudThemePicker;
  if (document.body) injectMaudThemePicker();
  else document.addEventListener('DOMContentLoaded', injectMaudThemePicker);
})();

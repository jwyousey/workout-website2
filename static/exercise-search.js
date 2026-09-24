// Type-to-search exercise picker for the new/edit workout pages.
// Works on any .exercise-block with an .ex-search input and an .ex-name <select>:
// typing lists matching exercises under the box; picking one selects it in the <select>.
(function () {
  const MAX_RESULTS = 8;
  // Spelled-out words people type -> the abbreviations used in exercise names
  const ABBREVIATIONS = [
    [/\bdumb ?bells?\b/g, 'db'],
    [/\bbar ?bells?\b/g, 'bb'],
    [/\bkettle ?bells?\b/g, 'kb'],
    [/\bclose grip\b/g, 'cg'],
    [/\bsingle leg\b/g, 'sl'],
    [/\bbehind the back\b/g, 'btb'],
    [/\bromanian deadlifts?\b/g, 'rdl'],
  ];

  function normalize(text) {
    let s = text.toLowerCase();
    for (const [re, abbr] of ABBREVIATIONS) s = s.replace(re, abbr);
    return s.replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function findMatches(block, query) {
    const sel = block.querySelector('.ex-name');
    const group = block.querySelector('.ex-muscle-group')?.value || '';
    const q = normalize(query);
    const tokens = q.split(' ').filter(Boolean);
    const results = [];
    for (const opt of sel.options) {
      if (!opt.value || opt.value === '__other__') continue;
      const optGroup = opt.dataset.group || 'Other';
      if (group && optGroup !== group) continue;
      const name = normalize(opt.value);
      const squashed = name.replace(/ /g, '');  // so "pullup" finds "Pull-Up"
      if (!tokens.every(t => squashed.includes(t))) continue;
      const qSquashed = q.replace(/ /g, '');
      const exact = squashed === qSquashed;
      const rank = exact ? -1 : squashed.startsWith(qSquashed) ? 0 : name.split(' ').some(w => w.startsWith(tokens[0])) ? 1 : 2;
      results.push({ name: opt.value, group: optGroup, rank, exact });
    }
    results.sort((a, b) => a.rank - b.rank || a.name.localeCompare(b.name));
    return results;
  }

  function getPanel(input) {
    let panel = input.nextElementSibling;
    if (!panel || !panel.classList.contains('ex-search-results')) {
      panel = document.createElement('div');
      panel.className = 'ex-search-results';
      panel.setAttribute('role', 'listbox');
      input.after(panel);
    }
    return panel;
  }

  function close(input) {
    const panel = input.nextElementSibling;
    if (panel && panel.classList.contains('ex-search-results')) panel.remove();
  }

  function render(input) {
    const query = input.value.trim();
    if (!query) return close(input);
    const block = input.closest('.exercise-block');
    const matches = findMatches(block, query);
    const panel = getPanel(input);
    panel.innerHTML = '';

    const items = matches.slice(0, MAX_RESULTS).map(m => ({ label: m.name, meta: m.group, name: m.name, isNew: false }));
    if (!matches.some(m => m.exact)) {
      items.push({ label: `+ Add "${query}"`, meta: 'new exercise', name: query, isNew: true });
    }
    items.forEach((item, i) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'ex-search-item' + (i === 0 ? ' active' : '') + (item.isNew ? ' is-new' : '');
      btn.setAttribute('role', 'option');
      const label = document.createElement('span');
      label.textContent = item.label;
      const meta = document.createElement('span');
      meta.className = 'ex-search-meta';
      meta.textContent = item.meta;
      btn.append(label, meta);
      // Keep focus in the input so the tap isn't lost to a blur
      btn.addEventListener('pointerdown', e => e.preventDefault());
      btn.addEventListener('click', () => choose(input, item.name, item.isNew));
      panel.appendChild(btn);
    });
    if (matches.length > MAX_RESULTS) {
      const more = document.createElement('div');
      more.className = 'ex-search-more';
      more.textContent = `${matches.length - MAX_RESULTS} more, keep typing to narrow down`;
      panel.appendChild(more);
    }
  }

  function choose(input, name, isNew) {
    const block = input.closest('.exercise-block');
    const sel = block.querySelector('.ex-name');
    let opt = [...sel.options].find(o => o.value === name);
    if (!opt) {
      const group = block.querySelector('.ex-muscle-group')?.value || 'Other';
      opt = new Option(name, name);
      opt.dataset.group = group;
      sel.insertBefore(opt, sel.lastElementChild);
      if (isNew) {
        fetch('/exercises/add', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          credentials: 'same-origin',
          body: JSON.stringify({ name, muscle_group: group }),
        });
      }
    }
    opt.hidden = false;
    sel.value = name;
    sel.dispatchEvent(new Event('change'));
    input.value = '';
    close(input);
    input.blur();
  }

  function move(input, step) {
    const panel = input.nextElementSibling;
    if (!panel || !panel.classList.contains('ex-search-results')) return;
    const items = [...panel.querySelectorAll('.ex-search-item')];
    const cur = items.findIndex(b => b.classList.contains('active'));
    const next = (cur + step + items.length) % items.length;
    items.forEach((b, i) => b.classList.toggle('active', i === next));
    items[next]?.scrollIntoView({ block: 'nearest' });
  }

  const isSearch = el => el && el.classList && el.classList.contains('ex-search');

  document.addEventListener('input', e => { if (isSearch(e.target)) render(e.target); });
  document.addEventListener('focusin', e => { if (isSearch(e.target)) render(e.target); });
  document.addEventListener('focusout', e => {
    if (isSearch(e.target)) setTimeout(() => { if (document.activeElement !== e.target) close(e.target); }, 150);
  });
  document.addEventListener('keydown', e => {
    if (!isSearch(e.target)) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); move(e.target, 1); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); move(e.target, -1); }
    else if (e.key === 'Escape') { close(e.target); }
    else if (e.key === 'Enter') {
      const active = e.target.nextElementSibling?.querySelector?.('.ex-search-item.active');
      if (active) { e.preventDefault(); active.click(); }
    }
  });
  // Changing the muscle group narrows an open search
  document.addEventListener('change', e => {
    if (e.target.classList?.contains('ex-muscle-group')) {
      const input = e.target.closest('.exercise-block')?.querySelector('.ex-search');
      if (input && input.value.trim()) render(input);
    }
  });

  const style = document.createElement('style');
  style.textContent = `
    .ex-search-results { display:flex; flex-direction:column; max-height:320px; overflow-y:auto; margin:-2px 0 8px;
      border:1px solid var(--border); border-radius:3px; background:var(--surface2); }
    .ex-search-item { display:flex; justify-content:space-between; align-items:center; gap:10px; width:100%;
      padding:10px 12px; border:none; border-bottom:1px solid var(--border); background:transparent; color:var(--text);
      font:inherit; font-size:14px; text-align:left; cursor:pointer; }
    .ex-search-item:last-child { border-bottom:none; }
    .ex-search-item:hover, .ex-search-item.active { background:rgba(255,85,0,0.12); }
    .ex-search-item.is-new { color:var(--accent); font-weight:600; }
    .ex-search-meta { flex-shrink:0; font-size:11px; color:var(--text-3); text-transform:uppercase; letter-spacing:.05em; }
    .ex-search-more { padding:8px 12px; font-size:12px; color:var(--text-3); }
  `;
  document.head.appendChild(style);
})();

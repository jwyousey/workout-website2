// Move exercise cards up/down and remove them on the new/edit workout pages,
// keeping the "Superset" link buttons between cards in sync.
// Relies on the page's insertSupersetBtn(block); saving reads cards in on-screen order.
(function () {
  const container = () => document.getElementById('exercises');
  const blocksOf = () => [...container().querySelectorAll('.exercise-block')];

  function updateMoveButtons() {
    const blocks = blocksOf();
    blocks.forEach((b, i) => {
      const up = b.querySelector('.ex-move-up');
      const down = b.querySelector('.ex-move-down');
      if (up) up.disabled = i === 0;
      if (down) down.disabled = i === blocks.length - 1;
    });
  }

  // Rebuild the link buttons between cards. A card stays in a superset only
  // while a neighbouring card shares its group.
  function refreshSupersetLinks() {
    const blocks = blocksOf();
    blocks.forEach((b, i) => {
      const g = b.dataset.supersetGroup;
      if (g && blocks[i - 1]?.dataset.supersetGroup !== g && blocks[i + 1]?.dataset.supersetGroup !== g) {
        delete b.dataset.supersetGroup;
      }
    });
    blocks.forEach(b => b.classList.toggle('ss-grouped', !!b.dataset.supersetGroup));
    container().querySelectorAll('.ss-link-btn').forEach(btn => btn.remove());
    blocks.slice(1).forEach((b, i) => {
      insertSupersetBtn(b);
      const above = blocks[i];
      if (b.dataset.supersetGroup && above.dataset.supersetGroup === b.dataset.supersetGroup) {
        const btn = b.previousElementSibling;
        btn.classList.add('active');
        btn.innerHTML = btn.innerHTML.replace('Superset', 'Linked');
      }
    });
    updateMoveButtons();
  }

  window.refreshSupersetLinks = refreshSupersetLinks;

  window.moveExercise = function (btn, dir) {
    const block = btn.closest('.exercise-block');
    const blocks = blocksOf();
    let idx = blocks.indexOf(block) + dir;
    // Hop over a whole superset rather than landing inside it and breaking the link
    const group = blocks[idx]?.dataset.supersetGroup;
    if (group && group !== block.dataset.supersetGroup) {
      while (blocks[idx + dir]?.dataset.supersetGroup === group) idx += dir;
    }
    const target = blocks[idx];
    if (!target) return;
    if (dir < 0) target.before(block); else target.after(block);
    refreshSupersetLinks();
    block.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    block.classList.remove('ex-moved');
    void block.offsetWidth;  // restart the highlight animation
    block.classList.add('ex-moved');
  };

  window.removeExercise = function (btn) {
    btn.closest('.exercise-block').remove();
    refreshSupersetLinks();
  };

  document.addEventListener('DOMContentLoaded', () => {
    const el = container();
    if (!el) return;
    new MutationObserver(updateMoveButtons).observe(el, { childList: true });
    updateMoveButtons();
  });

  const style = document.createElement('style');
  style.textContent = `
    .ex-move-btn { border:1px solid var(--border); background:transparent; color:var(--text-2); border-radius:3px;
      width:30px; height:30px; display:flex; align-items:center; justify-content:center; cursor:pointer; flex-shrink:0; padding:0; }
    .ex-move-btn:hover:not(:disabled) { border-color:var(--accent); color:var(--accent); }
    .ex-move-btn:disabled { opacity:.25; cursor:default; }
    @keyframes ex-moved-flash { from { background-color:rgba(255,85,0,0.18); } to { background-color:var(--card); } }
    .exercise-block.ex-moved { animation: ex-moved-flash .9s ease-out; }
  `;
  document.head.appendChild(style);
})();

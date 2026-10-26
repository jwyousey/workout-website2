// Auto-saves the in-progress lift on the Log Lift page to this browser's storage, so it
// survives the phone unloading the page in the background. Restored on the next visit,
// cleared once the workout is saved (or the user taps "Start over").
(function () {
  const KEY = 'progresslog:draft:new-workout';
  const VERSION = 1;
  let ready = false;  // don't save until any restore has finished
  let timer = null;

  function read() {
    try {
      const d = JSON.parse(localStorage.getItem(KEY));
      return d && d.v === VERSION ? d : null;
    } catch { return null; }
  }
  function write(d) { try { localStorage.setItem(KEY, JSON.stringify(d)); } catch { /* storage full or blocked */ } }
  window.clearWorkoutDraft = function () { try { localStorage.removeItem(KEY); } catch { /* ignore */ } };
  // Called once the workout is stored on the server: stop auto-saving and drop the draft
  window.workoutSaved = function () { ready = false; clearTimeout(timer); window.clearWorkoutDraft(); };

  const val = id => document.getElementById(id)?.value || '';

  function snapshot() {
    const exercises = [...document.querySelectorAll('#exercises .exercise-block')].map(b => {
      const name = b.querySelector('.ex-name').value;
      return {
        name: name === '__other__' ? '' : name,
        group: b.querySelector('.ex-muscle-group')?.value || '',
        note: b.querySelector('.ex-note')?.value || '',
        superset: b.dataset.supersetGroup || '',
        sets: [...b.querySelectorAll('.set-row')].map(r => ({
          reps: r.querySelector('.set-reps').value,
          weight: r.querySelector('.set-weight').value,
          unit: r.querySelector('.set-unit').value,
          rpe: r.querySelector('.set-rpe').value,
          warmup: r.querySelector('.warmup-btn')?.classList.contains('active') || false,
        })),
      };
    });
    return {
      v: VERSION, savedAt: Date.now(),
      date: val('workoutDate'), name: val('workoutName'), customName: val('workoutNameCustom'),
      notes: val('workoutNotes'), duration: val('durationMin'), rating: val('workoutRating'),
      exercises,
    };
  }

  function hasContent(d) {
    return d.name || d.customName.trim() || d.notes.trim() || d.duration || d.rating ||
      d.exercises.some(e => e.name || e.note.trim() || e.sets.length);
  }

  function save() {
    if (!ready) return;
    clearTimeout(timer);
    const d = snapshot();
    if (hasContent(d)) write(d); else window.clearWorkoutDraft();
  }
  function saveSoon() { clearTimeout(timer); timer = setTimeout(save, 300); }

  function restore(d) {
    const nameSel = document.getElementById('workoutName');
    if (d.date) document.getElementById('workoutDate').value = d.date;
    if (d.name) {
      nameSel.value = d.name;
      if (nameSel.value !== d.name) nameSel.value = '__other__';
    }
    document.getElementById('workoutNameCustom').value = d.customName || '';
    handleWorkoutNameSelect(nameSel);
    document.getElementById('workoutNotes').value = d.notes || '';
    document.getElementById('durationMin').value = d.duration || '';
    if (d.rating) setRating(parseInt(d.rating));

    document.getElementById('exercises').innerHTML = '';
    let maxGroup = 0;
    for (const ex of d.exercises) {
      const block = addExercise(ex.name || '');
      const exSel = block.querySelector('.ex-name');
      const mgSel = block.querySelector('.ex-muscle-group');
      if (ex.group && mgSel) {
        mgSel.value = ex.group;
        filterExercisesByGroup(mgSel);
      }
      if (ex.name) {  // keep the exercise selected even if a filter hid it
        const opt = [...exSel.options].find(o => o.value === ex.name);
        if (opt) opt.hidden = false;
        exSel.value = ex.name;
        updateLastTime(block);
      }
      const note = block.querySelector('.ex-note');
      if (note) note.value = ex.note || '';
      if (ex.superset) {
        block.dataset.supersetGroup = ex.superset;
        maxGroup = Math.max(maxGroup, parseInt(ex.superset) || 0);
      }
      const addBtn = block.querySelector('[onclick="addSet(this)"]');
      const setsContainer = block.querySelector('.sets-container');
      for (const s of ex.sets) {
        addSet(addBtn);
        const row = setsContainer.lastElementChild;
        row.querySelector('.set-reps').value = s.reps;
        row.querySelector('.set-weight').value = s.weight;
        row.querySelector('.set-unit').value = s.unit || 'lbs';
        row.querySelector('.set-rpe').value = s.rpe;
        if (s.warmup) toggleWarmup(row.querySelector('.warmup-btn'));
      }
    }
    if (!d.exercises.length) addExercise();
    nextSSGroup = maxGroup + 1;
    if (window.refreshSupersetLinks) window.refreshSupersetLinks();
    showBanner(d.savedAt);
  }

  function showBanner(savedAt) {
    const when = new Date(savedAt);
    const sameDay = when.toDateString() === new Date().toDateString();
    const label = when.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) +
      (sameDay ? '' : ' on ' + when.toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' }));
    const banner = document.createElement('div');
    banner.className = 'card draft-banner';
    banner.innerHTML = `<div style="flex:1;min-width:0;"><div style="font-size:14px;font-weight:700;">Unsaved lift restored</div>
      <div style="font-size:12px;color:var(--text-3);margin-top:2px;">Last change at ${label}. Pick up where you left off.</div></div>
      <button type="button" class="btn btn-ghost" style="flex-shrink:0;">Start over</button>`;
    banner.querySelector('button').onclick = () => {
      if (!confirm('Discard this unsaved lift and start over?')) return;
      window.workoutSaved();  // stops auto-save so leaving the page can't write it back
      location.reload();
    };
    const anchor = document.querySelector('.page-title')?.parentElement;
    if (anchor) anchor.after(banner); else document.querySelector('main')?.prepend(banner);
  }

  document.addEventListener('DOMContentLoaded', () => {
    // Run after the page's own setup (e.g. the date defaulting to today)
    setTimeout(() => {
      const d = read();
      if (d && hasContent(d)) {
        try { restore(d); } catch (e) { console.error('Could not restore draft', e); }
      }
      ready = true;
    }, 0);
  });

  ['input', 'change'].forEach(t => document.addEventListener(t, saveSoon, true));
  document.addEventListener('click', saveSoon, true);  // add/remove set, warm-up, superset, move, rating
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden') save(); });
  window.addEventListener('pagehide', save);

  const style = document.createElement('style');
  style.textContent = `.draft-banner { display:flex; align-items:center; gap:12px; margin-bottom:16px; padding:14px 16px;
    border-left-color:#10b981 !important; }`;
  document.head.appendChild(style);
})();

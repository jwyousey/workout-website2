// "Last time" hints for the Log Lift and Edit Workout pages. Each page defines
// LAST_SETS and LAST_BY_REPS (see get_previous_sets in app.py).

function formatLastTime(name) {
  const sets = LAST_SETS[name];
  if (!sets || !sets.length) return '';
  const parts = sets.map(s => s.weight ? `${s.weight}${s.unit}×${s.reps ?? '?'}` : `${s.reps ?? '?'} reps`);
  let text = `Last (${sets[0].date}): ${parts.join(', ')}`;
  const top = sets.reduce((best, s) => (s.weight || 0) > (best.weight || 0) ? s : best, sets[0]);
  if (top.weight) {
    if (top.rpe && top.rpe >= 9) {
      text += `  →  Stay at ${top.weight}${top.unit}`;
    } else {
      const step = top.unit === 'kg' ? 2.5 : 5;
      const next = Math.ceil((top.weight + 0.1) / step) * step;
      text += `  →  Try ${next}${top.unit}`;
    }
  }
  return text;
}

function updateLastTime(block) {
  const name = block.querySelector('.ex-name').value;
  const hint = block.querySelector('.last-time');
  const text = formatLastTime(name);
  if (text) {
    hint.textContent = text;
    hint.style.display = '';
  } else {
    hint.style.display = 'none';
  }
}

function updateSetPlaceholder(row, name) {
  const weightInput = row.querySelector('.set-weight');
  const reps = parseInt(row.querySelector('.set-reps').value);
  const match = LAST_BY_REPS[name] && LAST_BY_REPS[name][reps];
  weightInput.placeholder = match ? match.weight : '0';
}

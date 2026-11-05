/** Nearest-rank percentiles; no interpolation or discarded failed samples. */
export function distribution(values) {
  if (!values.length) return { n: 0, p50: null, p95: null, max: null };
  const ordered = values.toSorted((a, b) => a - b);
  if (ordered.some(value => !Number.isFinite(value) || value < 0)) throw new Error('Invalid timing sample');
  const at = p => ordered[Math.ceil(ordered.length * p) - 1];
  return { n: ordered.length, p50: at(0.5), p95: at(0.95), max: ordered.at(-1) };
}

export function summarize(runs) {
  const valid = runs.filter(run => run.valid);
  const keys = valid.flatMap(run => run.recording.keys);
  const phases = [...new Set(keys.map(key => key.phase))];
  return {
    runs: runs.length, validRuns: valid.length,
    // This includes driver connection/preparation delay. It is an observed
    // successful edit from navigation, never an estimate of the earliest ready time.
    navigationToFirstEditRenderOpportunityMs: distribution(valid
      .filter(run => run.recording.boundary === 'plugin-navigation-start')
      .map(run => run.recording.firstEdit.renderOpportunityAt)),
    phases: Object.fromEntries(phases.map(phase => {
      const samples = keys.filter(key => key.phase === phase && key.validated);
      return [phase, {
        keyToVisibleDomMs: distribution(samples.map(key => key.visibleDomAt - key.downAt)),
        keyToRenderOpportunityMs: distribution(samples.map(key => key.renderOpportunityAt - key.downAt)),
        coalesced: samples.filter(key => key.coalesced).length,
        savingAtKeydown: samples.filter(key => ['exporting', 'uploading', 'committing'].includes(key.saveState)).length,
      }];
    })),
    longTasks: { count: valid.reduce((sum, run) => sum + run.recording.longTasks.length, 0),
      durationMs: valid.reduce((sum, run) => sum + run.recording.longTasks.reduce((n, task) => n + task.duration, 0), 0) },
  };
}

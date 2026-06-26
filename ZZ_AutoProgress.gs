function getGenerationStatus() {
  const props = PropertiesService.getUserProperties();
  const raw = props.getProperty(STATE_KEY);
  if (!raw) return { status: 'idle', currentIndex: 0, totalItems: 0, results: [], folders: [] };
  let state = JSON.parse(raw);
  const kicking = props.getProperty('generationStatusKickInProgress') === '1';
  if (state.status === 'running' && !kicking) {
    props.setProperty('generationStatusKickInProgress', '1');
    try {
      continueGeneration();
      const nextRaw = props.getProperty(STATE_KEY);
      if (nextRaw) state = JSON.parse(nextRaw);
    } finally {
      props.deleteProperty('generationStatusKickInProgress');
    }
  }
  state.progress = state.totalItems ? Math.round((state.currentIndex / state.totalItems) * 100) : 0;
  return state;
}

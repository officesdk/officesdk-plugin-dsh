// A remounted editor shares this queue. A lost acknowledgement retries the same
// revision, so an earlier response can never replace a newer list of cards.
export function createPendingReviews(initial, send, report) {
  let saved = initial.changes, current = saved, revision = initial.revision;
  let pending, running, failed = false;

  async function drain() {
    while (pending || saved !== current) {
      if (!pending) {
        const target = current;
        const previous = new Map(saved.map(card => [card.id, card]));
        const put = target.filter(card => previous.get(card.id) !== card);
        const remaining = new Set(target.map(card => card.id));
        const remove = saved.filter(card => !remaining.has(card.id)).map(card => card.id);
        if (!put.length && !remove.length) { saved = target; continue; }
        pending = { target, patch: { base: revision, revision: crypto.randomUUID(), put, remove } };
      }
      const reply = await send(pending.patch);
      if (reply.revision !== pending.patch.revision) throw new Error('Missing change history acknowledgement');
      revision = reply.revision;
      saved = pending.target;
      pending = undefined;
    }
  }

  function flush() {
    if (running) return running;
    if (saved === current && !pending) return Promise.resolve();
    running = (async () => {
      try {
        do { await drain(); } while (pending || saved !== current);
        if (failed) report(false);
        failed = false;
      } catch (error) {
        failed = true;
        report(true);
        throw error;
      } finally {
        running = undefined;
      }
    })();
    return running;
  }

  return {
    get current() { return current; },
    stage(cards) {
      if (current === cards) return;
      current = cards;
      void flush().catch(() => {}); // The sidebar keeps a retry available.
    },
    flush,
  };
}

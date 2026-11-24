// DSH records plugin commands in the conversation without submitting them to the main model.
export const PARALLEL_EDIT_COMMAND = 'officesdk-edit';

export function parallelMessage(node) {
  if (node?.kind !== 'command' || node.name !== PARALLEL_EDIT_COMMAND || typeof node.args !== 'string') return;
  return { text:node.args, error:node.outcome?.kind === 'error' ? node.outcome.text : undefined };
}

// Project the user's instruction at its send event, outside the foldable agent
// process. A later result updates that node without moving it in the chat.
export const parallelInputDefinition = {
  kind: 'officesdk-edit-input',
  target: 'chat',
  match(event) {
    const { commandId, name } = event.data ?? {};
    if (event.type === 'command/run' && name === PARALLEL_EDIT_COMMAND) return { id: commandId, role: 'start' };
    if (event.type === 'command/done' && commandId?.startsWith('cmd-officesdk-')) return { id: commandId, role: 'update' };
    return null;
  },
  start(_context, { event }) {
    return { kind: 'command', name: event.data.name, args: event.data.args, seq: event.seq, time:event.time, outcome: null };
  },
  update({ state }, { event }) {
    return state && { ...state, outcome: { kind: event.data.kind, text: event.data.text } };
  },
  buildViewNode({ state, key, id, start }) {
    if (!state) return null;
    return { key, id, kind: 'officesdk-edit-input', target: 'chat', groupTree: 'input',
      anchorSeq: state.seq, location: start?.location ?? { kind: 'unresolved' }, visibility: 'visible', data: state };
  },
};

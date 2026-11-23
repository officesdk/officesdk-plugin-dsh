import { validatePptxOps } from '../agent/pptx.mjs';
import { LABELS, createOverlay, inspectDeck, rollback, runSteps } from './presentation-agent.mjs';

const MAX_UNDO = 20;
const EDIT_ID = /^[A-Za-z0-9_-]{1,64}$/;
const agentError = (code, message) => Object.assign(new Error(message), { code });

/**
 * AI editing on a Kit page (spec §12.5): the same plan, confirmation and checks as §10, through the editor
 * package's Presentation API. With motion (v1.1, spec §7) every target is selected before it is written and text grows
 * in at most four real writes; without it each step writes its final value once. Every change the editor makes reaches Kit through the page's
 * change queue; the agent never stages anything itself. A hand edit hands the undo history to the editor.
 */
export function createKitPresentationAgent({ editor, document, animate = () => true, locale = 'zh', wait }) {
  const presentation = () => editor.api.presentation;
  const api = () => presentation().getJsApi();
  const internals = () => presentation().editorInternals();
  const store = () => internals().getDocument().store;
  // The editor's undo — not the store's — reaches subscribeLocalChanges, so Kit follows the rollback.
  const rollBack = baseline => rollback(store(), baseline, 2000, () => internals().undo());
  const labels = LABELS[locale] ?? LABELS.zh;
  const overlay = createOverlay(document, { animate, pause: wait });
  const undoStack = [];
  let applying = false;
  return {
    /** True while an AI edit or undo is writing: its changes are not the user's. */
    get applying() { return applying; },
    /** The user edited by hand: AI undo would also take that back, so it ends here. */
    handOff() { undoStack.length = 0; },
    get undoable() { return undoStack.at(-1)?.editId ?? null; },
    inspect: () => inspectDeck(api()),
    async apply({ ops, editId }, signal) {
      if (applying) throw agentError('busy', 'Another edit is still running.');
      applying = true;
      overlay.busy(true, labels.editing);
      try {
        const jsApi = api();
        const baseline = store().encodeBinary();
        const refs = {};
        const deck = await inspectDeck(jsApi);
        const { steps } = validatePptxOps(ops, deck);
        const geometry = new Map(deck.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape])));
        // runSteps says whether the slides were written at the same time.
        const report = {};
        let after;
        try {
          // v1.1 (spec §7): select before write, text growing in a few real writes, page by page with agent tracks.
          const moving = animate();
          await runSteps(jsApi, steps, { overlay, document, signal, labels, wait, animate: moving, typing: moving ? 'chunks' : 'overlay', locale, refs, report, geometryOf: (slideId, shapeId) => geometry.get(`${slideId}/${shapeId}`) });
          after = await inspectDeck(jsApi);
          const texts = new Map(after.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape.text])));
          for (const step of steps) if (step.op === 'setText' && texts.get(`${step.slideId}/${step.shapeId}`) !== step.text) throw agentError('editor-failed', 'The presentation did not take the planned text.');
        } catch (error) {
          // The editor's own history takes the deck back; those undo steps reach Kit as changes too.
          if (!rollBack(baseline)) throw agentError('editor-failed', 'The edit failed and the editor could not return to the state before it; reload the document.');
          throw signal?.aborted ? agentError('cancelled', 'The edit was cancelled; nothing was changed.') : error;
        }
        if (EDIT_ID.test(editId ?? '')) { undoStack.push({ editId, baseline }); if (undoStack.length > MAX_UNDO) undoStack.shift(); }
        const touched = [...new Set(steps.map(step => step.createdId && step.op === 'addSlide' ? step.createdId : step.slideId).filter(Boolean))];
        let focus = touched.find(id => steps.every(step => !(step.op === 'deleteSlide' && step.slideId === id)));
        // Several slides written at once: the user may have picked one of them meanwhile; it stays. Otherwise the
        // first of them in the deck is shown.
        if (report.together) {
          const position = id => after.slides.findIndex(slide => slide.id === id);
          focus = touched.some(id => after.selectedSlideIds?.includes(id)) ? undefined : [...touched].sort((left, right) => position(left) - position(right))[0];
        }
        if (focus) await jsApi.PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([focus]); await ctx.sync(); }).catch(() => {});
        return { applied: steps.length, slideIds: touched, ...(Object.keys(refs).length && { refs }), undoable: undoStack.at(-1)?.editId ?? null };
      } finally {
        applying = false;
        overlay.busy(false);
      }
    },
    async undo({ editId }) {
      if (applying) throw agentError('busy', 'Another edit is still running.');
      const top = undoStack.at(-1);
      if (!top) throw agentError('undo-unavailable', 'There is no AI edit to undo in this editor (the user may have edited it by hand since; they can undo in the editor with Cmd/Ctrl+Z).');
      if (top.editId !== editId) throw agentError('undo-unavailable', undoStack.some(entry => entry.editId === editId) ? 'Only the most recent AI edit can be undone; undo the later edits first.' : 'That edit is no longer in this editor.');
      applying = true;
      overlay.busy(true, labels.saving);
      try {
        if (!rollBack(top.baseline)) throw agentError('editor-failed', 'The editor could not return to the state before that edit; reload the document.');
        undoStack.pop();
        return { undone: editId, undoable: undoStack.at(-1)?.editId ?? null };
      } finally { applying = false; overlay.busy(false); }
    },
    dispose() { overlay.dispose(); },
  };
}

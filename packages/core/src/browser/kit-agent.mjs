import { presentationSelectionMethods } from './presentation-selection.mjs';
import { unionRects } from './selection-edit.mjs';
import { deckProgress, plannedLayout, showViewTab, switchDeckView } from './presentation-generation.mjs';
import { PPTX_LIMITS, validatePptxOps } from '../agent/pptx.mjs';
import { LABELS, NORMAL, SORTER, createOverlay, inspectDeck, rollback, runSteps, screenRect } from './presentation-agent.mjs';

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
  let progress;
  async function leaveSorter(first) {
    progress?.dispose();
    progress = undefined;
    overlay.busy(false);
    await switchDeckView(api(), document, NORMAL);
    if (first) await api().PowerPoint.run(async ctx => { ctx.presentation.setSelectedSlides([first]); await ctx.sync(); });
  }
  async function preparePages(count, roles, signal) {
    if (!Number.isInteger(count) || count < 1 || count > PPTX_LIMITS.maxPlannedSlides) throw agentError('invalid-edit', 'Invalid slide count.');
    const jsApi = api();
    const deck = await inspectDeck(jsApi);
    if (deck.slides.length || progress) throw agentError('invalid-edit', 'A slide plan can only initialize a new, empty deck.');
    const baseline = store().encodeBinary();
    try {
      signal?.throwIfAborted();
      await switchDeckView(jsApi, document, SORTER);
      overlay.busy(true);
      const layoutIds = Array.from({length:count}, (_, i) => plannedLayout(deck.layouts,
        roles?.[i] ?? (i === 0 ? 'cover' : i === count - 1 ? 'closing' : 'content')));
      const slides = await jsApi.PowerPoint.run(async ctx => {
        const list = ctx.presentation.slides;
        for (const layoutId of layoutIds) list.add(layoutId ? {layoutId} : undefined);
        list.load('items/id');
        await ctx.sync();
        // Kit may insert several new pages after the same active page. Restore the requested
        // layout sequence before exposing IDs, even when the native insertion order differs.
        const items = list.items.map(slide => ({id:slide.id, layout:slide.layout}));
        for (const {layout} of items) layout.load('id');
        await ctx.sync();
        const remaining = items.slice();
        const ids = layoutIds.map(id => {
          const index = id ? remaining.findIndex(item => item.layout.id === id) : 0;
          if (index < 0 || !remaining.length) throw agentError('editor-failed', 'The editor did not apply the planned layout.');
          return remaining.splice(index, 1)[0].id;
        });
        const order = items.map(item => item.id);
        for (const [index, id] of ids.entries()) {
          if (id === order[index]) continue;
          list.getItem(id).moveTo(index);
          order.splice(index, 0, order.splice(order.indexOf(id),1)[0]);
        }
        await ctx.sync();
        return ids;
      });
      if (slides.length !== count) throw agentError('editor-failed', 'The editor did not create the planned slides.');
      progress = deckProgress(document, slides, locale, deck.slideSize);
      // The sorter mounts its thumbnail grid after the document mutation. Start the
      // creation gesture clock from that paint, including on cold editor startup.
      const pause = wait ?? (ms => new Promise(resolve => setTimeout(resolve, ms)));
      await Promise.race([progress.firstThumbnail, pause(1800)]);
      overlay.busy(false);
      showViewTab(document);
      if (animate()) await pause(320);
      signal?.throwIfAborted();
      progress.phase(slides, 'waiting');
      return {slideIds:slides};
    } catch (error) {
      rollBack(baseline);
      await leaveSorter().catch(() => {});
      throw error;
    }
  }
  const selection=presentationSelectionMethods({api,document,animate,locale,wait,locked:async run=>{
    if(applying) throw agentError('busy','Another edit is still running.');
    applying=true;
    try {const result=await run();undoStack.length=0;return result;} finally {applying=false;}
  }});
  return {
    ...selection,
    /** True while an AI edit or undo is writing: its changes are not the user's. */
    get applying() { return applying; },
    async endGeneration() { if (progress) await leaveSorter(); },
    /** The user edited by hand: AI undo would also take that back, so it ends here. */
    handOff() { undoStack.length = 0; },
    get undoable() { return undoStack.at(-1)?.editId ?? null; },
    inspect: () => inspectDeck(api()),
    locate(target) {
      return target?.kind==='document-context' ? unionRects((target.geometry ?? []).map(g=>screenRect(document,g,target.slide)).filter(Boolean)) : selection.locate(target);
    },
    async apply({ ops, editId, prepareSlideCount, slideLayouts, completedSlideIds }, signal) {
      if (applying) throw agentError('busy', 'Another edit is still running.');
      applying = true;
      const plan = progress;
      if (prepareSlideCount === undefined) overlay.busy(true, plan ? undefined : labels.editing);
      try {
        if (prepareSlideCount !== undefined) return await preparePages(prepareSlideCount, slideLayouts, signal);
        const jsApi = api();
        const baseline = store().encodeBinary();
        const refs = {};
        const deck = await inspectDeck(jsApi);
        const { steps } = validatePptxOps(ops, deck);
        const targets = [...new Set(steps.map(step => step.slideId))];
        if (plan && steps.some(step => !plan.pages.has(step.slideId) || ['addSlide','moveSlide','deleteSlide'].includes(step.op))) {
          throw agentError('invalid-edit', 'The planned slides already exist. Use their slideIds to fill them.');
        }
        const finished = completedSlideIds ?? targets;
        if (plan && (!Array.isArray(finished) || finished.some(id => !targets.includes(id)))) {
          throw agentError('invalid-edit', 'completedSlideIds must name slides written in this batch.');
        }
        const previous = plan && targets.map(id => [id, plan.pages.get(id).phase]);
        const geometry = new Map(deck.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape])));
        // runSteps says whether the slides were written at the same time.
        const report = {};
        let after;
        try {
          // v1.1 (spec §7): select before write, text growing in a few real writes, page by page with agent tracks.
          const moving = animate();
          await runSteps(jsApi, steps, { overlay: plan ? undefined : overlay, planned: plan, document, signal, labels, wait, animate: moving, typing: moving ? 'chunks' : 'overlay', locale, refs, report, geometryOf: (slideId, shapeId) => geometry.get(`${slideId}/${shapeId}`) });
          after = await inspectDeck(jsApi);
          const texts = new Map(after.slides.flatMap(slide => slide.shapes.map(shape => [`${slide.id}/${shape.id}`, shape.text])));
          for (const step of steps) if (step.op === 'setText' && texts.get(`${step.slideId}/${step.shapeId}`) !== step.text) throw agentError('editor-failed', 'The presentation did not take the planned text.');
        } catch (error) {
          // The editor's own history takes the deck back; those undo steps reach Kit as changes too.
          if (plan) for (const [id, phase] of previous) plan.phase([id],phase);
          if (plan && signal?.aborted) await leaveSorter().catch(() => {});
          if (!rollBack(baseline)) throw agentError('editor-failed', 'The edit failed and the editor could not return to the state before it; reload the document.');
          throw signal?.aborted ? agentError('cancelled', 'The edit was cancelled; nothing was changed.') : error;
        }
        if (EDIT_ID.test(editId ?? '')) { undoStack.push({ editId, baseline }); if (undoStack.length > MAX_UNDO) undoStack.shift(); }
        const touched = [...new Set(steps.map(step => step.createdId && step.op === 'addSlide' ? step.createdId : step.slideId).filter(Boolean))];
        let focus = touched.find(id => steps.every(step => !(step.op === 'deleteSlide' && step.slideId === id)));
        // Several slides written at once: the user may have picked one of them meanwhile; it stays. Otherwise the
        // first of them in the deck is shown.
        if (plan) {
          focus = undefined;
          plan.phase(targets, 'partial');
          plan.phase(finished, 'done');
          if ([...plan.pages.values()].every(page => page.phase === 'done')) await leaveSorter(after.slides[0]?.id);
        } else if (report.together) {
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
    dispose() { progress?.dispose(); selection.dispose(); overlay.dispose(); },
  };
}

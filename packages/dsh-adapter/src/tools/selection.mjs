import { defineTool } from '@deepseek-ai/dsh-tools';
import { SELECTION_LIMITS } from '../../../core/src/index.mjs';

/**
 * selection_rewrite (spec 261109-selection-ai-edit §4): the user selected text in a Word document (or cells in a
 * spreadsheet) in the side panel and sent a request about it; their message names the selection (`sel-…`). The model
 * answers with the replacement only: there is no position to give, so nothing outside the selection can change. The
 * user asked for exactly this change in the editor, so it applies without an approval card, once per selection; they
 * can undo it from the panel.
 */
const json = { schema: { type: 'json' }, render: (_args, value) => [{ type: 'text', text: JSON.stringify(value) }] };

export const CLARIFY_EDIT = 'If essential information is missing or a problem has multiple user-facing resolutions, use ask_user_question when available. Offer specific feasible choices and allow a custom answer. Wait for the answer, then resume the same target. On skip, leave that target untouched. Proceed directly with unambiguous edits.';

export const SELECTION_REWRITE_DESCRIPTION = `Replace the text (or cells) the user selected in a document open in the side panel. Use it when the user's message says they selected something in a file and names a selection id like "sel-1a2b3c4d" with what to change (their message quotes the whole selection and their request). ${CLARIFY_EDIT} Answer with the replacement only — do not call document_read or document_edit for it: give \`text\` for a Word or PowerPoint text selection (the new text for exactly the selected part, paragraphs separated by \\n; keep the text around the selection out of it) or \`values\` for a spreadsheet selection (rows of cell values, exactly as many rows and columns as the selection; text starting with "=" is a formula). Write in the language and tone the request asks for and keep the meaning unless asked otherwise. In ask mode it returns pendingReview: true; report the candidate as awaiting confirmation and do not save it or call another edit tool. In always-apply mode it applies directly. Use only once per selection: to change it again, the user selects again. Only text is replaced: requests about formatting (bold, colours, fonts) cannot be done with it — tell the user plainly instead. Paragraphs rewritten as a whole take their first formatting. Errors: selection-unknown (not in this conversation or expired), selection-stale (the user changed the text since; ask them to select again), selection-used (already rewritten), cancelled (the user cancelled it in the editor: do not change anything). Limits: ${SELECTION_LIMITS.maxRewrite} characters of text, ${SELECTION_LIMITS.maxCells} cells. Result: { applied, selection, path, before, after, unsaved, pendingReview? }; the change stays unsaved in the editor like other AI edits, and the user can keep or undo it there.`;

export function registerSelectionTools(ctx, { service, failure }) {
  ctx.tools.register(defineTool({
    name: 'selection_rewrite',
    description: SELECTION_REWRITE_DESCRIPTION,
    parameters: {
      selection: { type: 'string', required: true, description: 'The selection id from the user\'s message, like "sel-1a2b3c4d".' },
      text: { type: 'string', description: 'Word or PowerPoint selection: the replacement for the selected text; paragraphs separated by \\n.' },
      values: { type: 'array', items: { type: 'array', items: { type: 'json' } }, description: 'Spreadsheet selection: rows of new cell values (text, numbers, true/false or null), the same size as the selection.' },
    },
    output: json,
    async execute(args, exec) {
      const sessionId = exec.agent?.session?.id;
      try {
        if (!sessionId) throw new Error('selection_rewrite needs an agent Session.');
        return await service.aiRewriteSelection(sessionId, { selection: args.selection, ...(args.text !== undefined && { text: args.text }), ...(args.values !== undefined && { values: args.values }) }, exec.signal);
      } catch (error) { throw failure(error); }
    },
  }));
}

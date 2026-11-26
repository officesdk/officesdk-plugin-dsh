import { runTemplateScript } from './template-python.mjs';
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { SuiteError } from '../../../core/src/index.mjs';


export async function blankExistingDeck({ stylesRoot, style, dataRoot, signal, reportingData }) {
  const base = join(dataRoot, 'existing-deck-generation');
  await mkdir(base, { recursive: true, mode: 0o700 });
  const dir = await mkdtemp(join(base, 'template-'));
  try {
    const out = join(dir, 'prepared.pptx');
    const args = ['--source', join(stylesRoot, style, 'template.pptx'), '--out', out];
    let script = 'prepare_existing_deck.py';
    if (style === 'slide-002') {
      // Seed only the native chart and data-driven shapes before opening. The
      // slide copy is written via presentation_edit in the visible sorter.
      const { chart, bars, statuses } = reportingData ?? {};
      const spec = join(dir, 'data.json');
      await writeFile(spec, JSON.stringify({ chart, bars, statuses, slides: Array.from({ length: 8 }, () => ({})) }));
      args.push('--spec', spec);
      script = 'fill_reporting_deck.py';
    }
    try {
      await runTemplateScript('shimo-pptx', script, args, { cwd: dir, signal });
    } catch (error) {
      const validation = error.stderr?.match(/^ValueError: (.+)$/m);
      if (validation && style === 'slide-002') throw new SuiteError('invalid-edit', validation[1]);
      throw error;
    }
    return await readFile(out);
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

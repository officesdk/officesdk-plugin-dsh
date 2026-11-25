const MiB = 1024 * 1024;
const minute = 60_000;
export const DEFAULT_LIMITS = Object.freeze({
  maxInputBytes: 32 * MiB, maxOutputBytes: 100 * MiB, maxStageBytes: 256 * MiB,
  maxRetainedBytes: 1024 * MiB, maxConcurrent: 2, maxQueued: 8, maxLeases: 64, maxReaders: 16,
  maxOutputFiles: 10_000, maxDrafts: 16, maxDraftBytes: 512 * MiB, draftTtlMs: 24 * 60 * minute,
  leaseTtlMs: 15 * minute, parkTtlMs: 24 * 60 * minute, autosaveMs: 15_000, exportTtlMs: minute, maxRssBytes: 2048 * MiB,
  timeoutMs: Object.freeze({
    import: Object.freeze({ docx: minute, xlsx: minute, pptx: 3 * minute }),
    export: Object.freeze({ docx: 3 * minute, xlsx: 2 * minute, pptx: 3 * minute }),
  }),
});

const positive = (name, value) => {
  if (!Number.isSafeInteger(value) || value < 1) throw new TypeError(`Invalid limit ${name}.`);
  return value;
};

/** Every limit is a positive safe integer; nested timeouts merge per operation and format. */
export function normalizeLimits(overrides = {}) {
  const limits = {};
  for (const [name, value] of Object.entries(DEFAULT_LIMITS)) {
    if (name === 'timeoutMs') continue;
    limits[name] = positive(name, overrides[name] ?? value);
  }
  for (const name of Object.keys(overrides)) if (!Object.hasOwn(DEFAULT_LIMITS, name)) throw new TypeError(`Unknown limit ${name}.`);
  const timeouts = overrides.timeoutMs ?? {};
  if (typeof timeouts !== 'object') throw new TypeError('Invalid limit timeoutMs.');
  limits.timeoutMs = {};
  for (const [operation, formats] of Object.entries(DEFAULT_LIMITS.timeoutMs)) {
    limits.timeoutMs[operation] = {};
    for (const [format, value] of Object.entries(formats)) {
      limits.timeoutMs[operation][format] = positive(`timeoutMs.${operation}.${format}`, timeouts[operation]?.[format] ?? value);
    }
  }
  return limits;
}

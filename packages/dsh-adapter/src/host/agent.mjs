import { randomUUID } from 'node:crypto';
import { SuiteError } from '../../../core/src/index.mjs';

/**
 * Commands from AI tools to the Client showing a Session. Tools wait on a reply; Clients collect
 * commands with a long poll. A command nobody collects, or a Client that stops reporting, times out.
 */
export class AgentHub {
  /** `platform`: the Host's `process.platform`, put on every command so the Client spells its path the Host's way. */
  constructor({ pollMs = 20_000, pickupMs = 60_000, aliveMs = 45_000, scheduler = globalThis, platform } = {}) {
    Object.assign(this, { pollMs, pickupMs, aliveMs, scheduler, platform });
    this.queues = new Map();
    this.pollers = new Map();
    this.waiting = new Map();
  }

  queue(sessionId) {
    let queue = this.queues.get(sessionId);
    if (!queue) this.queues.set(sessionId, queue = []);
    return queue;
  }

  deliver(sessionId, command) {
    if (this.platform) command = { ...command, platform: this.platform };
    const pollers = this.pollers.get(sessionId);
    const poller = pollers?.shift();
    if (poller) poller(command);
    else this.queue(sessionId).push(command);
  }

  /**
   * Sends one command and resolves with the Client's result. `pickupMs` bounds how long the command
   * waits for a Client. After pickup the wait has no total limit (the user may be deciding), but
   * the Client must report it is still working every `aliveMs`: a lost command or a closed window
   * would otherwise leave the tool waiting forever.
   */
  request(sessionId, command, signal) {
    const requestId = randomUUID();
    return new Promise((resolve, reject) => {
      const entry = { sessionId, resolve, reject, picked: false };
      const cleanup = () => { this.scheduler.clearTimeout(entry.timer); signal?.removeEventListener('abort', onAbort); this.waiting.delete(requestId); };
      entry.settle = (error, value) => { cleanup(); error ? reject(error) : resolve(value); };
      entry.timer = this.scheduler.setTimeout(() => {
        if (entry.picked) return;
        this.withdraw(sessionId, requestId);
        entry.settle(new SuiteError('no-editor', 'No DSH window is showing this conversation, so the document could not be opened. Ask the user to open the conversation, then try again.'));
      }, this.pickupMs);
      const onAbort = () => {
        if (entry.picked) this.deliver(sessionId, { kind: 'cancel', target: requestId });
        else this.withdraw(sessionId, requestId);
        entry.settle(new SuiteError('cancelled', 'The document command was cancelled.'));
      };
      if (signal?.aborted) return onAbort();
      signal?.addEventListener('abort', onAbort, { once: true });
      this.waiting.set(requestId, entry);
      this.deliver(sessionId, { ...command, requestId });
    });
  }

  /** A one-way message for the Client (no reply); dropped if no Client collects it within `pickupMs`. */
  notify(sessionId, command) {
    const requestId = randomUUID();
    this.deliver(sessionId, { ...command, requestId });
    // Housekeeping only: it must not keep the Host process alive.
    this.scheduler.setTimeout(() => this.withdraw(sessionId, requestId), this.pickupMs)?.unref?.();
  }

  withdraw(sessionId, requestId) {
    const queue = this.queues.get(sessionId);
    const index = queue?.findIndex(command => command.requestId === requestId) ?? -1;
    if (index >= 0) queue.splice(index, 1);
  }

  /** Long poll: the next command for this Session, or null after `pollMs`. */
  next(sessionId, signal) {
    const queued = this.queues.get(sessionId)?.shift();
    if (queued) return Promise.resolve(this.pick(queued));
    return new Promise(resolve => {
      const pollers = this.pollers.get(sessionId) ?? [];
      this.pollers.set(sessionId, pollers);
      const finish = command => { this.scheduler.clearTimeout(timer); signal?.removeEventListener('abort', abort); const index = pollers.indexOf(finish); if (index >= 0) pollers.splice(index, 1); resolve(command ? this.pick(command) : null); };
      const abort = () => finish(null);
      const timer = this.scheduler.setTimeout(() => finish(null), this.pollMs);
      signal?.addEventListener('abort', abort, { once: true });
      pollers.push(finish);
    });
  }

  pick(command) {
    const entry = this.waiting.get(command.requestId);
    if (entry) { entry.picked = true; this.watch(entry); }
    return command;
  }

  watch(entry) {
    this.scheduler.clearTimeout(entry.timer);
    entry.timer = this.scheduler.setTimeout(() => entry.settle(new SuiteError('no-editor',
      'The DSH window stopped responding while working on the document. Ask the user to check the side panel, then try again.')), this.aliveMs);
  }

  /** A Client still working on a picked command; only the Session it went to may report. */
  alive(sessionId, requestId) {
    const entry = this.waiting.get(requestId);
    if (!entry || entry.sessionId !== sessionId || !entry.picked) return false;
    this.watch(entry);
    return true;
  }

  /** A Client's answer; only the Session the command went to may answer it. */
  reply(sessionId, requestId, { ok, result, code, message }) {
    const entry = this.waiting.get(requestId);
    if (!entry || entry.sessionId !== sessionId) return false;
    if (ok) entry.settle(undefined, result ?? null);
    else entry.settle(new SuiteError(knownCode(code), typeof message === 'string' && message ? message.slice(0, 2000) : 'The editor could not do that.'));
    return true;
  }

  dispose() {
    for (const entry of [...this.waiting.values()]) entry.settle(new SuiteError('unavailable', 'The document service stopped.'));
    for (const pollers of this.pollers.values()) for (const finish of [...pollers]) finish(null);
  }
}

const CLIENT_CODES = new Set(['invalid-edit', 'read-only', 'busy', 'cancelled', 'editor-failed', 'no-editor', 'source-changed', 'edit-conflict', 'unsupported-operation', 'forbidden', 'not-found', 'unavailable', 'timeout', 'lease-unavailable']);
const knownCode = code => CLIENT_CODES.has(code) ? code : 'editor-failed';

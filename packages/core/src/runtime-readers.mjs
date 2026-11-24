import { fail } from './errors.mjs';

/** SDK module/icon bursts must wait for disk capacity, not poison module imports with 429. */
export class RuntimeReaders {
  constructor(limit = 16, maxQueued = 1024) {
    this.limit = limit;
    this.maxQueued = maxQueued;
    this.active = 0;
    this.waiters = [];
  }

  async read(signal, operation) {
    signal.throwIfAborted();
    if (this.active < this.limit) this.active++;
    else {
      if (this.waiters.length >= this.maxQueued) fail('busy', 'The runtime queue is full.');
      await new Promise((resolve, reject) => {
        const waiter = { resolve, reject, signal, abort: () => {
          this.waiters.splice(this.waiters.indexOf(waiter), 1);
          reject(signal.reason);
        } };
        this.waiters.push(waiter);
        signal.addEventListener('abort', waiter.abort, { once: true });
      });
    }
    try {
      signal.throwIfAborted();
      return await operation();
    } finally {
      const next = this.waiters.shift();
      if (next) {
        next.signal.removeEventListener('abort', next.abort);
        next.resolve(); // Transfer the occupied slot before admitting new arrivals.
      } else this.active--;
    }
  }
}

import { SuiteError } from '../../../core/src/index.mjs';

/** Session scope comes from Typert's workspaceFileScope lookup; non-owners go through remoteAccess. */
export function createAuthorizer(ctx) {
  return {
    async authorize(peer, scope, signal, { action, path, revision }) {
      signal.throwIfAborted();
      if (peer.signal?.aborted) throw new SuiteError('forbidden', 'The document peer is no longer available.');
      const lookup = ctx.typert.lookups.get('workspaceFileScope');
      const current = await lookup?.resolve(scope.sessionId);
      if (!current || current.sessionId !== scope.sessionId || current.workspaceRoot !== scope.workspaceRoot) throw new SuiteError('forbidden', 'The Session is no longer available.');
      if (peer.remote?.principal?.kind !== 'owner') {
        const decision = ctx.get('remoteAccess')?.decide({ principal: peer.remote?.principal, action: `shimoPreview/${action}`,
          resources: [{ kind: 'session', id: scope.sessionId }], args: { workspaceFileScopeId: scope.sessionId, path, revision } });
        if (decision?.kind !== 'allow') throw new SuiteError('forbidden', 'Access to the document was denied.');
      }
      signal.throwIfAborted();
    },
  };
}

import { fromWorkspaceFile } from './errors.mjs';

/** Core's SourceProvider is WorkspaceFiles' own shape; only failures need translating. */
export function createSource(workspaceFiles) {
  return {
    async stat(scope, path, signal) {
      try { return await workspaceFiles.stat(scope, path, signal); } catch (error) { throw fromWorkspaceFile(error); }
    },
    async readBytes(scope, path, options, signal) {
      try { return await workspaceFiles.readBytes(scope, path, options, signal); } catch (error) { throw fromWorkspaceFile(error); }
    },
  };
}

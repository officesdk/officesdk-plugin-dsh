import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { SuiteError } from '../../../core/src/index.mjs';

/**
 * Runtime copy of scripts/platform-package.mjs's PLATFORMS (lib/host cannot import build scripts; a test keeps the
 * names and suffixes equal). Platforms a package can carry office-kit's Kit for (spec §4.6 B′: one complete package per platform, its Kit in
 * `resources/kit`). Windows builds name their executables `.exe`; the shared files (modoc.js, mow_wasm*,
 * bos-pin.txt) sit next to them because Kit checks its install layout from its own directory.
 */
export const KIT_PLATFORMS = Object.freeze({
  'darwin-arm64': { exe: '' },
  'win32-x64': { exe: '.exe' },
});
const RUNTIME = 'offline-kit-runtime';
const WORKER = 'compose-worker';
const SHARED = ['modoc.js'];

export const platformKey = ({ platform = process.platform, arch = process.arch } = {}) => `${platform}-${arch}`;

/**
 * The Kit executable for this machine under `kitRoot`, or why there is none: `unsupported` names the platform,
 * the reason (`platform`: no Kit is built for it; `mismatch`: the package carries another platform's Kit;
 * `missing`/`incomplete`: the files are not all there) and the user-facing message.
 * Only names are checked; launching a wrong binary is what this avoids, not a corrupted one.
 */
export function resolveKit(kitRoot, host = {}, exists = existsSync) {
  const platform = platformKey(host);
  const entry = KIT_PLATFORMS[platform];
  const has = name => exists(join(kitRoot, name));
  if (!entry) return { unsupported: { platform, reason: 'platform' } };
  const binary = RUNTIME + entry.exe;
  if (!has(binary)) {
    // An extension-less runtime on Windows, or an .exe elsewhere, is another platform's package.
    const foreign = [RUNTIME, `${RUNTIME}.exe`].find(name => name !== binary && has(name));
    return { unsupported: { platform, reason: foreign ? 'mismatch' : 'missing', file: binary } };
  }
  const absent = [WORKER + entry.exe, ...SHARED].filter(name => !has(name));
  if (absent.length) return { unsupported: { platform, reason: 'incomplete', file: absent.join(', ') } };
  return { binary: join(kitRoot, binary) };
}

// A universal package names the platforms its kit-manifest.json has (`supported`, test ones marked); a complete
// package the ones it can be built for.
const supportedList = (supported, test) => (supported ?? Object.keys(KIT_PLATFORMS).map(platform => ({ platform }))).map(item => (item.status === 'test' ? `${item.platform}${test}` : item.platform));
// The Host does not know the Client's language: Chinese first (the users'), then English.
const REASONS = {
  platform: ({ platform, supported }) => [`OfficeSDK 目前只支持 ${supportedList(supported, '（测试）').join('、')}，当前系统是 ${platform}。`, `OfficeSDK supports ${supportedList(supported, ' (test)').join(', ')} only; this is ${platform}.`],
  mismatch: ({ platform }) => [`插件包里的文档服务是其他平台的版本，请安装 ${platform} 版本的插件。`, `The package carries another platform's document service; install the ${platform} package.`],
  missing: ({ platform, file }) => [`插件包里没有 ${platform} 的文档服务（${file}），请安装 ${platform} 版本的插件。`, `The package has no ${platform} document service (${file}); install the ${platform} package.`],
  incomplete: ({ file }) => [`插件包里的文档服务不完整（缺少 ${file}），请重新安装插件。`, `The package's document service is incomplete (missing ${file}); reinstall the plugin.`],
};

/** What opening a document says when this machine has no Kit (spec §4.1: no generic unavailable). */
export function unsupportedError(unsupported) {
  const [zh, en] = REASONS[unsupported.reason](unsupported);
  return new SuiteError('unsupported-operation', `本平台暂不支持打开 Office 文档：${zh} This platform cannot open Office documents yet: ${en}`, { platform: unsupported.platform, reason: unsupported.reason });
}

/** Stands in for Core's converter: the package ships none, so Core converts and opens nothing. */
export const NO_CONVERTER = Object.freeze({
  capabilities: () => ({}),
  import: () => Promise.reject(new SuiteError('unavailable', 'This package has no Core converter.')),
  export: () => Promise.reject(new SuiteError('unavailable', 'This package has no Core converter.')),
});

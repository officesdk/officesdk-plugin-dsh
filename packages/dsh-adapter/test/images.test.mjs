import test from 'node:test';
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fileURLToPath } from 'node:url';

// shimo-pptx's image scripts: find_images.py (search, fetch, fit, credits), theme_art.py, and brand_deck.py's credits.
// No network: the scripts' one HTTP function is replaced with fixtures.
const SCRIPTS = fileURLToPath(new URL('../skills/shimo-pptx/scripts/', import.meta.url));
const PYTHON = [process.env.SHIMO_PYTHON, '/Applications/DeepSeek Harness.app/Contents/Resources/runtime/primary-runtime/dependencies/python/bin/python3.12', 'python3']
  .filter(Boolean).find(candidate => spawnSync(candidate, ['-c', 'import pptx, lxml, PIL'], { stdio: 'ignore' }).status === 0);
const skip = !PYTHON && 'no python with python-pptx and Pillow';

/** Runs `body` (python) with find_images loaded as `fi` and its HTTP replaced by FIXTURES; prints JSON results. */
function runPython(cwd, body) {
  const prelude = `
import importlib.util, io, json, os, sys, contextlib
sys.dont_write_bytecode = True
from PIL import Image
def load(name):
    spec = importlib.util.spec_from_file_location(name, os.path.join(${JSON.stringify(SCRIPTS)}, name + '.py'))
    m = importlib.util.module_from_spec(spec); spec.loader.exec_module(m); return m
fi = load('find_images')
def png(w, h, color):
    b = io.BytesIO(); Image.new('RGB', (w, h), color).save(b, 'PNG'); return b.getvalue()
OPENVERSE = {'results': [
  {'title': 'Skyline A', 'creator': 'ann', 'license': 'by', 'license_version': '2.0', 'license_url': 'https://cc/by/2.0', 'url': 'https://img/a.jpg', 'foreign_landing_url': 'https://page/a', 'width': 3000, 'height': 2000, 'source': 'flickr', 'tags': [{'name': 'city'}]},
  {'title': 'No commercial', 'creator': 'x', 'license': 'by-nc', 'license_version': '2.0', 'url': 'https://img/nc.jpg', 'width': 3000, 'height': 2000},
  {'title': 'No derivatives', 'creator': 'x', 'license': 'by-nd', 'license_version': '4.0', 'url': 'https://img/nd.jpg', 'width': 3000, 'height': 2000},
  {'title': 'Tall one', 'creator': 'y', 'license': 'cc0', 'url': 'https://img/tall.jpg', 'width': 1000, 'height': 3000},
  {'title': 'Hongya Cave.jpg', 'creator': 'z', 'license': 'by-sa', 'license_version': '4.0', 'url': 'https://img/dup-openverse.jpg', 'width': 4000, 'height': 2000, 'source': 'wikimedia'},
  {'title': 'Old bridge', 'creator': 'flickr person', 'license': 'by', 'license_version': '2.0', 'url': 'https://img/other-bridge.jpg', 'width': 2000, 'height': 1500, 'source': 'flickr'},
]}
def meta(lic, artist='<a href="u">Bob</a>'):
    return {'LicenseShortName': {'value': lic}, 'Artist': {'value': artist}, 'LicenseUrl': {'value': 'https://cc/' + lic}, 'ImageDescription': {'value': '<b>River</b> at night'}}
COMMONS = {'query': {'pages': {
  '1': {'index': 1, 'title': 'File:Hongya Cave.jpg', 'imageinfo': [{'mime': 'image/jpeg', 'width': 4000, 'height': 2000, 'thumburl': 'https://img/hongya.jpg', 'descriptionurl': 'https://commons/hongya', 'extmetadata': meta('CC BY-SA 4.0')}]},
  '2': {'index': 2, 'title': 'File:Gfdl only.jpg', 'imageinfo': [{'mime': 'image/jpeg', 'width': 4000, 'height': 2000, 'thumburl': 'https://img/gfdl.jpg', 'extmetadata': meta('GFDL')}]},
  '3': {'index': 3, 'title': 'File:Map.svg', 'imageinfo': [{'mime': 'image/svg+xml', 'width': 4000, 'height': 2000, 'thumburl': 'https://img/map.png', 'extmetadata': meta('CC0')}]},
  '4': {'index': 4, 'title': 'File:Old bridge.png', 'imageinfo': [{'mime': 'image/png', 'width': 2400, 'height': 1600, 'url': 'https://img/bridge.png', 'descriptionurl': 'https://commons/bridge', 'extmetadata': meta('Public domain', '')}]},
}}}
IMAGES = {'https://img/a.jpg': png(3000, 2000, (200, 10, 10)), 'https://img/hongya.jpg': png(4000, 2000, (10, 200, 10)),
          'https://img/bridge.png': png(2400, 1600, (10, 10, 200)), 'https://img/tall.jpg': png(1000, 3000, (90, 90, 90))}
DOWN = set()
def fake_get(url, accept='application/json', limit=fi.DOWNLOAD_MAX):
    if any(url.startswith(d) for d in DOWN): raise OSError('source down')
    if url == 'https://img/huge.jpg': raise ValueError('larger than 40 MB')
    if url.startswith('https://api.openverse.org/'): return json.dumps(OPENVERSE).encode()
    if url.startswith('https://commons.wikimedia.org/'): return json.dumps(COMMONS).encode()
    if url in IMAGES: return IMAGES[url]
    raise OSError('404 ' + url)
fi.get = fake_get
def run(*argv):
    out, err = io.StringIO(), io.StringIO()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        try: code = fi.main(list(argv))
        except SystemExit as e: code = e.code
    return {'code': code, 'out': out.getvalue(), 'err': err.getvalue()}
R = {}
`;
  const result = spawnSync(PYTHON, ['-I', '-c', `${prelude}\n${body}\nprint(json.dumps(R, ensure_ascii=False))`], { cwd, encoding: 'utf8' });
  assert.equal(result.status, 0, result.stderr);
  return JSON.parse(result.stdout.trim().split('\n').at(-1));
}

test('find_images search keeps CC0 / public domain / CC BY / CC BY-SA only, filters orientation, drops cross-source duplicates', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-images-'));
  try {
    const r = runPython(dir, `
R['lic'] = [fi.norm_license(*x) for x in [('by', '2.0'), ('by-sa', '4.0'), ('cc0',), ('pdm',), ('CC BY-SA 3.0',), ('CC BY-SA 2.5 cn',), ('CC BY-SA 3.0 IGO',), ('CC BY 4.0 International',), ('Public Domain Mark 1.0',), ('by-nc', '2.0'), ('by-nd', '4.0'), ('CC BY-NC-SA 2.0 de',), ('GFDL',), ('',)]]
R['any'] = run('search', '--query', 'chongqing')
R['cands'] = json.load(open('images/candidates.json'))
R['landscape'] = run('search', '--query', 'chongqing', '--orientation', 'landscape')
R['land'] = [c['title'] for c in json.load(open('images/candidates.json'))]
DOWN.add('https://api.openverse.org/')
R['down'] = run('search', '--query', 'chongqing')
R['downTitles'] = [c['title'] for c in json.load(open('images/candidates.json'))]
DOWN.add('https://commons.wikimedia.org/')
R['none'] = run('search', '--query', 'chongqing')
`);
    assert.deepEqual(r.lic, ['CC BY 2.0', 'CC BY-SA 4.0', 'CC0', 'Public domain', 'CC BY-SA 3.0', 'CC BY-SA 2.5 CN', 'CC BY-SA 3.0 IGO', 'CC BY 4.0', 'Public domain', null, null, null, null, null]);
    assert.deepEqual(r.cands.map(c => [c.title, c.license]), [['Skyline A', 'CC BY 2.0'], ['Tall one', 'CC0'], ['Hongya Cave.jpg', 'CC BY-SA 4.0'], ['Old bridge', 'CC BY 2.0'], ['Old bridge.png', 'Public domain']],
      'NC, ND, GFDL and SVG are dropped; Commons\' "Hongya Cave.jpg" is the Commons file Openverse already listed; a Flickr photo with a similar title stays');
    assert.equal(r.cands[4].creator, '');
    assert.match(r.any.out, /编号只对这次 search 有效/);
    assert.equal(r.cands[0].page, 'https://page/a');
    assert.match(r.any.out, /^1\. Skyline A \| ann \| CC BY 2\.0 \| 3000x2000 \| https:\/\/page\/a/m);
    assert.match(r.any.out, /5\. Old bridge\.png \| 未署名 \| Public domain/);
    assert.deepEqual(r.land, ['Skyline A', 'Hongya Cave.jpg', 'Old bridge', 'Old bridge.png']);
    assert.deepEqual(r.downTitles, ['Hongya Cave.jpg', 'Old bridge.png'], 'one source down: the other still answers');
    assert.match(r.down.err, /warning: openverse「chongqing」：source down/);
    assert.equal(r.down.code, 0);
    assert.equal(r.none.code, 1, 'every source down is not "no match"');
    assert.match(r.none.out, /图库都连不上.*不用再换关键词：直接用 theme_art\.py/);
    const empty = runPython(dir, `
OPENVERSE['results'] = []; COMMONS['query']['pages'] = {}
R['empty'] = run('search', '--query', 'nothing')
`);
    assert.equal(empty.empty.code, 0);
    assert.match(empty.empty.out, /没有找到可商用、可修改授权的图片.*theme_art\.py/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('find_images fetch saves a bounded JPEG, records its source, and refuses the same picture twice', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-images-'));
  try {
    const r = runPython(dir, `
run('search', '--query', 'chongqing')
R['fetch'] = run('fetch', '--pick', '1', '--name', 'skyline', '--pick', '5', '--max-side', '1200')
R['names'] = run('fetch', '--pick', '2', '--name', 'a', '--name', 'b')
R['size'] = list(Image.open('images/skyline.jpg').size)
R['format'] = Image.open('images/skyline.jpg').format
R['again'] = run('fetch', '--pick', '1')
R['range'] = run('fetch', '--pick', '9')
OPENVERSE['results'][0]['url'] = 'https://img/huge.jpg'
run('search', '--query', 'chongqing')
R['huge'] = run('fetch', '--pick', '1')
R['sources'] = json.load(open('images/sources.json'))
`);
    assert.equal(r.fetch.code, 0, r.fetch.out);
    assert.deepEqual(r.size, [1200, 800]);
    assert.equal(r.format, 'JPEG');
    assert.deepEqual(r.sources.map(s => [s.file, s.title, s.creator, s.license, s.page, s.kind]), [
      ['images/skyline.jpg', 'Skyline A', 'ann', 'CC BY 2.0', 'https://page/a', 'photo'],
      ['images/Old-bridge.jpg', 'Old bridge.png', '', 'Public domain', 'https://commons/bridge', 'photo'],
    ]);
    assert.equal(r.again.code, 1);
    assert.match(r.again.out, /与已下载的 images\/skyline\.jpg 是同一张图；同一张图整份只用一次/);
    assert.match(String(r.range.code), /--pick 9 不在候选 1–5 里/);
    assert.match(String(r.names.code), /--name 比 --pick 多/);
    assert.equal(r.huge.code, 1);
    assert.match(r.huge.out, /下载或解码失败：larger than 40 MB/);
    assert.ok(r.sources.every(s => /^[0-9a-f]{64}$/.test(s.sha256) && /^[0-9a-f]{64}$/.test(s.download_sha256) && s.sha256 !== s.download_sha256), 'sha256 is the saved file, download_sha256 the bytes fetched');
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('find_images fit crops to the box around the centre, can darken, records the crop, and allows one crop per picture', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-images-'));
  try {
    const r = runPython(dir, `
run('search', '--query', 'chongqing'); run('fetch', '--pick', '1', '--name', 'skyline')
R['fit'] = run('fit', '--in', 'images/skyline.jpg', '--box', '312x270')
w, h = Image.open('images/skyline-312x270.jpg').size
R['ratio'] = round(w / h, 2)
R['second'] = run('fit', '--in', 'images/skyline.jpg', '--box', '720x405', '--darken', '0.45')
R['redo'] = run('fit', '--in', 'images/skyline.jpg', '--box', '720x405', '--darken', '0.45', '--out', 'images/skyline-312x270.jpg')
img = Image.open('images/skyline-312x270.jpg').convert('RGB')
R['redoRatio'] = round(img.width / img.height, 2)
R['dark'] = img.getpixel((img.width // 2, img.height // 2))
R['fitted'] = json.load(open('images/sources.json'))[0]['fitted']
R['fittedSha'] = json.load(open('images/sources.json'))[0]['fitted_sha256']
R['badBox'] = run('fit', '--in', 'images/skyline.jpg', '--box', 'wide')
R['badDark'] = run('fit', '--in', 'images/skyline.jpg', '--box', '1x1', '--darken', '1')
Image.new('RGB', (100, 50)).save('loose.png')
R['loose'] = run('fit', '--in', 'loose.png', '--box', '1x1')
R['credits'] = run('credits')
R['md'] = open('images/sources.md').read()
`);
    assert.equal(r.fit.code, 0, r.fit.err);
    assert.equal(r.ratio, 1.16);
    assert.match(String(r.second.code), /已裁成 images\/skyline-312x270\.jpg；同一张图整份只用一次/);
    assert.equal(r.redo.code, 0, 'the same output may be fitted again');
    assert.equal(r.redoRatio, 1.78);
    assert.ok(r.dark[0] < 120 && r.dark[1] < 10, `darkened ${r.dark}`);
    assert.deepEqual(r.fitted, ['images/skyline-312x270.jpg']);
    assert.equal(r.fittedSha.length, 1);
    assert.match(String(r.badBox.code), /--box 写成 宽x高/);
    assert.match(String(r.badDark.code), /--darken 取 0 到 1 之间/);
    assert.match(r.loose.err, /不在 images\/sources\.json 里，来源要自己写进交付说明/);
    assert.match(r.md, /^# 图片来源\n\n- skyline\.jpg：「Skyline A」，作者 ann，CC BY 2\.0（https:\/\/cc\/by\/2\.0），来源 https:\/\/page\/a\n$/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('find_images fit --anchor keeps the named edge; omitting it stays centre', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-images-anchor-'));
  try {
    const r = runPython(dir, `
def band(w, h, a, b, axis):
    im = Image.new('RGB', (w, h), b)
    px = im.load()
    if axis == 'x':
        for x in range(w // 2):
            for y in range(h): px[x, y] = a
    else:
        for y in range(h // 2):
            for x in range(w): px[x, y] = a
    im.save('band.png')
band(400, 200, (255, 0, 0), (0, 0, 255), 'x')
R['default'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'd.jpg')
R['center'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'c.jpg', '--anchor', 'center')
R['left'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'l.jpg', '--anchor', 'left')
R['right'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'r.jpg', '--anchor', 'right')
R['dpx'] = list(Image.open('d.jpg').getpixel((10, 10)))
R['cpx'] = list(Image.open('c.jpg').getpixel((10, 10)))
R['lpx'] = list(Image.open('l.jpg').getpixel((10, 10)))
R['rpx'] = list(Image.open('r.jpg').getpixel((10, 10)))
band(200, 400, (255, 0, 0), (0, 0, 255), 'y')
R['top'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 't.jpg', '--anchor', 'top')
R['bottom'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'b.jpg', '--anchor', 'bottom')
R['tpx'] = list(Image.open('t.jpg').getpixel((10, 10)))
R['bpx'] = list(Image.open('b.jpg').getpixel((10, 10)))
R['bad'] = run('fit', '--in', 'band.png', '--box', '1x1', '--anchor', 'middle')
# A recorded picture re-cropped to the same --out with another anchor replaces the crop; a second --out is refused.
band(400, 200, (255, 0, 0), (0, 0, 255), 'x')
json.dump([{'file': 'band.png'}], open('sources.json', 'w'))
R['first'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'same.jpg', '--anchor', 'left', '--dir', '.')
R['again'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'same.jpg', '--anchor', 'right', '--dir', '.')
R['againpx'] = list(Image.open('same.jpg').getpixel((10, 10)))
R['other'] = run('fit', '--in', 'band.png', '--box', '1x1', '--out', 'other.jpg', '--dir', '.')
R['fitted'] = json.load(open('sources.json'))[0]['fitted']
`);
    // JPEG rounds colours a little, so check which colour dominates rather than exact bytes.
    const red = px => px[0] > 200 && px[2] < 60;
    const blue = px => px[2] > 200 && px[0] < 60;
    assert.equal(r.default.code, 0, r.default.err);
    assert.equal(r.center.code, 0);
    assert.deepEqual(r.dpx, r.cpx, 'default is centre');
    assert.ok(red(r.lpx), `left keeps the red left half: ${r.lpx}`);
    assert.ok(blue(r.rpx), `right keeps the blue right half: ${r.rpx}`);
    assert.ok(red(r.tpx), `top keeps the red top half: ${r.tpx}`);
    assert.ok(blue(r.bpx), `bottom keeps the blue bottom half: ${r.bpx}`);
    assert.notEqual(r.bad.code, 0);
    assert.match(r.bad.err + r.bad.out, /--anchor/);
    assert.equal(r.first.code, 0, r.first.err);
    assert.equal(r.again.code, 0, r.again.err);
    assert.ok(blue(r.againpx), `re-cropping to the same --out replaces it: ${r.againpx}`);
    assert.notEqual(r.other.code, 0);
    assert.match(String(r.other.code), /已裁成 same\.jpg/);
    assert.deepEqual(r.fitted, ['same.jpg']);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('theme_art draws a picture of the asked size in the style\'s colours and records it as generated', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-art-'));
  try {
    const art = args => spawnSync(PYTHON, ['-I', join(SCRIPTS, 'theme_art.py'), ...args], { cwd: dir, encoding: 'utf8' });
    const made = art(['--motif', 'mountains,skyline,bridge,river', '--mood', 'night', '--size', '640x360', '--out', 'images/cover.png', '--note', '重庆：山城夜景', '--seed', '7']);
    assert.equal(made.status, 0, made.stderr);
    const again = art(['--motif', 'mountains,skyline,bridge,river', '--mood', 'night', '--size', '640x360', '--out', 'images/cover.png', '--note', '重庆：山城夜景', '--seed', '7']);
    assert.equal(again.status, 0);
    assert.equal(JSON.parse(made.stdout).sha256, JSON.parse(again.stdout).sha256, 'the same seed draws the same picture');
    const check = spawnSync(PYTHON, ['-I', '-c', `from PIL import Image; import json
im = Image.open('images/cover.png').convert('RGB')
print(json.dumps({'size': im.size, 'colours': len(set(im.getdata()))}))`], { cwd: dir, encoding: 'utf8' });
    const info = JSON.parse(check.stdout);
    assert.deepEqual(info.size, [640, 360]);
    assert.ok(info.colours > 50, 'a picture, not a flat panel');
    const sources = JSON.parse(readFileSync(join(dir, 'images', 'sources.json'), 'utf8'));
    assert.deepEqual(sources.map(s => [s.file, s.kind, s.title]), [['images/cover.png', 'generated', '重庆：山城夜景']], 'drawing the same file again replaces its record');
    const tiny = art(['--motif', 'grid,network', '--size', '16x16', '--out', 'images/tiny.jpg']);
    assert.equal(tiny.status, 0, tiny.stderr);
    for (const [args, message] of [
      [['--motif', 'dragons', '--out', 'x.png'], /未知母题 dragons；可用：sky, sun, mountains/],
      [['--motif', 'sky', '--style', '../styles/default', '--out', 'x.png'], /未知样式 \.\.\/styles\/default；可用样式：(?:[a-z0-9-]+, )*default\b/],
      [['--motif', 'sky', '--palette', 'color-9', '--out', 'x.png'], /没有配色 color-9；可用：color-1/],
      [['--motif', 'sky', '--size', 'big', '--out', 'x.png'], /--size 写成 宽x高/],
      [['--motif', 'sky', '--size', '10x99999', '--out', 'x.png'], /每边 16–4096/],
      [['--motif', 'sky', '--out', 'images/cover'], /--out 要以 \.png 或 \.jpg 结尾/],
    ]) {
      const bad = art(args);
      assert.notEqual(bad.status, 0);
      assert.match(bad.stderr, message);
    }
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('brand_deck writes recorded image credits into the slide notes and asks for a picture where one is missing', { skip }, () => {
  const dir = mkdtempSync(join(tmpdir(), 'officesdk-credits-'));
  try {
    runPython(dir, `
run('search', '--query', 'chongqing'); run('fetch', '--pick', '1', '--name', 'skyline')
`);
    const art = spawnSync(PYTHON, ['-I', join(SCRIPTS, 'theme_art.py'), '--motif', 'network,grid', '--size', '400x400', '--out', 'images/net.png', '--note', '城市数据网络'], { cwd: dir, encoding: 'utf8' });
    assert.equal(art.status, 0, art.stderr);
    writeFileSync(join(dir, 'deck.json'), JSON.stringify({ title: '重庆\n山城', subtitle: '城市介绍', brand: { logo_missing_reason: '测试' }, slides: [
      { layout: 'cover' },
      { layout: 'image-statement', image: 'images/renamed.jpg', title: '两江\n交汇', text: '说明', notes: '讲两江' },
      { layout: 'numbered', lead: '导语', tag: '看点', title: '名片', image: 'images/net.png', items: ['一', '二'] },
      { layout: 'statement', title: '山城', items: [{ heading: '山', text: '说明' }], side: { heading: '要点', items: ['一'] } },
      { layout: 'risks', title: '风险\n提示', lead: '导语', intro: '说明', tag: '注意', items: ['一'], image: 'images/net-copy.png' },
      { layout: 'closing' },
    ] }));
    // Credits follow the content: the recorded photo under another name keeps its credit; another picture named like
    // a recorded one gets none.
    const shuffle = spawnSync(PYTHON, ['-I', '-c', `import shutil; from PIL import Image
shutil.copy('images/skyline.jpg', 'images/renamed.jpg'); Image.new('RGB', (300, 200), (1, 2, 3)).save('images/net-copy.png')`], { cwd: dir, encoding: 'utf8' });
    assert.equal(shuffle.status, 0, shuffle.stderr);
    const built = spawnSync(PYTHON, ['-I', join(SCRIPTS, 'brand_deck.py'), 'build', '--spec', 'deck.json', '--out', 'cq.pptx'], { cwd: dir, encoding: 'utf8' });
    assert.equal(built.status, 0, built.stderr);
    const report = JSON.parse(built.stdout);
    assert.deepEqual(report.image_credits.map(c => c.slide), ['第2页(image-statement)', '第3页(numbered)']);
    assert.match(report.warnings.join('\n'), /第4页\(statement\)：未提供 image，图位暂时改为色面。请按 shimo-pptx 补图：先用 scripts\/find_images\.py 上网找，找不到合适的用 scripts\/theme_art\.py/);
    const notes = spawnSync(PYTHON, ['-I', '-c', `from pptx import Presentation; import json
p = Presentation('cq.pptx')
print(json.dumps([s.notes_slide.notes_text_frame.text if s.has_notes_slide else '' for s in p.slides], ensure_ascii=False))`], { cwd: dir, encoding: 'utf8' });
    const text = JSON.parse(notes.stdout);
    assert.equal(text[1], '讲两江\n图片：renamed.jpg：「Skyline A」，作者 ann，CC BY 2.0（https://cc/by/2.0），来源 https://page/a');
    assert.equal(text[2], '图片：net.png：按主题生成的示意配图（城市数据网络），无版权限制');
    assert.equal(text[3], '');
    assert.equal(text[4], '', 'a picture that is not in sources.json gets no credit');
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

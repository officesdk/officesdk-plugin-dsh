#!/usr/bin/env python3
"""PPTX / PDF → one PNG per page, on this Mac (style-distill tooling; developer machines only).

    python3 -I pptx2png.py <deck.pptx | deck.pdf> --out <dir> [--dpi 96] [--timeout 180]

PPTX goes through Microsoft PowerPoint for Mac (AppleScript "save as PDF"), PDF through poppler's pdftoppm.
Writes <dir>/page-01.png … and prints their paths as JSON. Standard library only.

Requirements and how failures look:
- Microsoft PowerPoint for Mac in /Applications. Missing → exit 2.
- poppler (`brew install poppler`) for pdftoppm. Missing → exit 2.
- PowerPoint is sandboxed: the deck is copied into its container
  (~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/style-distill/) and the copy removed afterwards.
- The first run asks for Automation permission (System Settings › Privacy & Security › Automation: let your terminal
  control Microsoft PowerPoint). Someone must be there to allow it. A denial (-1743) or a hang past --timeout exits 3.
Only the presentation this script opened is closed; PowerPoint itself keeps running.
"""
import argparse
import json
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

POWERPOINT = Path('/Applications/Microsoft PowerPoint.app')
CONTAINER = Path.home() / 'Library/Containers/com.microsoft.Powerpoint/Data/Documents/style-distill'
AUTOMATION_HINT = ('请在 系统设置 › 隐私与安全性 › 自动化 里允许当前终端控制 Microsoft PowerPoint；'
                   '也可能是 PowerPoint 弹出了对话框（授权、恢复文档等）挡住了脚本，处理后重试。')


class ConvertError(Exception):
    def __init__(self, message, code):
        super().__init__(message)
        self.code = code


def _as_applescript(text):
    return '"' + text.replace('\\', '\\\\').replace('"', '\\"') + '"'


def pptx_to_pdf(pptx, pdf, timeout):
    if not POWERPOINT.exists():
        raise ConvertError(f'没有找到 {POWERPOINT}：PPTX 转图需要 Microsoft PowerPoint for Mac。', 2)
    CONTAINER.mkdir(parents=True, exist_ok=True)
    stem = f'deck-{uuid.uuid4().hex[:10]}'
    copy, out = CONTAINER / f'{stem}.pptx', CONTAINER / f'{stem}.pdf'
    shutil.copyfile(pptx, copy)
    script = f'''
set src to POSIX file {_as_applescript(str(copy))}
set dst to {_as_applescript(str(out))}
tell application "Microsoft PowerPoint"
    open src
    set pres to missing value
    repeat 50 times
        repeat with i from 1 to (count of presentations)
            if (name of presentation i) is {_as_applescript(copy.name)} then set pres to presentation i
        end repeat
        if pres is not missing value then exit repeat
        delay 0.2
    end repeat
    if pres is missing value then error "PowerPoint did not open the deck"
    save pres in (POSIX file dst) as save as PDF
end tell
repeat 150 times
    if (do shell script "test -s " & quoted form of dst & " && echo y || echo n") is "y" then exit repeat
    delay 0.2
end repeat
delay 0.5
tell application "Microsoft PowerPoint" to close pres saving no
'''
    try:
        run = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=timeout)
        if run.returncode != 0:
            denied = '-1743' in run.stderr or 'Not authorized' in run.stderr
            raise ConvertError(f'PowerPoint 导出 PDF 失败：{run.stderr.strip()}' + ('\n' + AUTOMATION_HINT if denied else ''), 3)
        if not out.exists():
            raise ConvertError('PowerPoint 没有写出 PDF。' + AUTOMATION_HINT, 3)
        shutil.copyfile(out, pdf)
    except subprocess.TimeoutExpired:
        raise ConvertError(f'PowerPoint 导出超过 {timeout} 秒。' + AUTOMATION_HINT, 3) from None
    finally:
        _close(copy.name)
        for leftover in (copy, out, CONTAINER / ('~$' + copy.name)):
            leftover.unlink(missing_ok=True)


def _close(name):
    """Close our copy if a failed run left it open (no-op otherwise); never quits PowerPoint."""
    script = f'''
tell application "Microsoft PowerPoint"
    repeat with i from (count of presentations) to 1 by -1
        if (name of presentation i) is {_as_applescript(name)} then close presentation i saving no
    end repeat
end tell
'''
    try:
        subprocess.run(['osascript', '-e', script], capture_output=True, text=True, timeout=30)
    except subprocess.TimeoutExpired:
        pass


def _pdftoppm():
    tool = shutil.which('pdftoppm') or next((p for p in ('/opt/homebrew/bin/pdftoppm', '/usr/local/bin/pdftoppm') if Path(p).exists()), None)
    if not tool:
        raise ConvertError('没有找到 pdftoppm：请先 brew install poppler。', 2)
    return tool


def pdf_to_png(pdf, out_dir, dpi):
    tool = _pdftoppm()
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob('page-*.png'):
        old.unlink()
    run = subprocess.run([tool, '-r', str(dpi), '-png', str(pdf), str(out_dir / 'page')], capture_output=True, text=True)
    if run.returncode != 0:
        raise ConvertError(f'pdftoppm 失败：{run.stderr.strip()}', 3)
    pages = sorted(out_dir.glob('page-*.png'))
    # pdftoppm pads page numbers to the page count's width; rename to two digits for stable sorting.
    renamed = []
    for i, page in enumerate(pages, 1):
        target = out_dir / f'page-{i:02d}.png'
        if page != target:
            page.rename(target)
        renamed.append(target)
    return renamed


def convert(src, out_dir, dpi=96, timeout=180):
    """PNG paths, one per page, for a .pptx or .pdf."""
    src = Path(src).resolve()
    if src.suffix.lower() == '.pdf':
        return pdf_to_png(src, Path(out_dir), dpi)
    if src.suffix.lower() != '.pptx':
        raise ConvertError(f'只支持 .pptx 或 .pdf：{src}', 2)
    _pdftoppm()   # check both tools before PowerPoint is driven
    if not POWERPOINT.exists():
        raise ConvertError(f'没有找到 {POWERPOINT}：PPTX 转图需要 Microsoft PowerPoint for Mac。', 2)
    with tempfile.TemporaryDirectory(prefix='pptx2png-') as tmp:
        pdf = Path(tmp) / 'deck.pdf'
        pptx_to_pdf(src, pdf, timeout)
        return pdf_to_png(pdf, Path(out_dir), dpi)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('src', type=Path)
    ap.add_argument('--out', type=Path, required=True)
    ap.add_argument('--dpi', type=int, default=96)
    ap.add_argument('--timeout', type=int, default=180)
    a = ap.parse_args()
    try:
        pages = convert(a.src, a.out, a.dpi, a.timeout)
    except ConvertError as e:
        print(str(e), file=sys.stderr)
        return e.code
    print(json.dumps([str(p) for p in pages], ensure_ascii=False, indent=1))
    return 0


if __name__ == '__main__':
    sys.exit(main())

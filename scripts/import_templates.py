"""Import the five approved source sets without changing the originals.

Run with the bundled Python and pass the bundled soffice executable explicitly.
Legacy .doc justifications are converted once, not during document generation.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
import unicodedata


ROOT = Path(__file__).resolve().parents[1]


def find_normalized(parent: Path, name: str) -> Path:
    wanted = unicodedata.normalize('NFC', name)
    matches = [p for p in parent.iterdir() if unicodedata.normalize('NFC', p.name) == wanted]
    if len(matches) != 1:
        raise ValueError(f'Expected exactly one source: {parent / name}')
    return matches[0]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    parser.add_argument('--soffice', required=True, type=Path)
    args = parser.parse_args()
    catalog = json.loads((ROOT / 'templates/catalog.json').read_text())
    for mode, config in catalog.items():
        folder = find_normalized(args.source, config['source_folder'])
        destination = ROOT / 'templates' / mode
        destination.mkdir(exist_ok=True)
        for filename, original in config['source_files'].items():
            source = find_normalized(folder, original)
            target = destination / filename
            if target.exists():
                raise FileExistsError(f'Refusing to overwrite an imported template: {target}')
            if source.suffix.lower() != '.doc':
                shutil.copy2(source, target)
            else:
                with TemporaryDirectory(prefix='qujat-template-import-') as temporary:
                    temporary = Path(temporary)
                    local_source = temporary / 'source.doc'
                    shutil.copy2(source, local_source)
                    env = dict(os.environ, HOME=str(temporary))
                    subprocess.run([
                        str(args.soffice), '--headless',
                        f'-env:UserInstallation={(temporary / "profile").as_uri()}',
                        '--convert-to', 'docx', '--outdir', str(temporary), str(local_source),
                    ], env=env, check=True, capture_output=True, text=True, timeout=90)
                    shutil.copy2(temporary / 'source.docx', target)
            print(f'{mode}: {source.name} -> {target.relative_to(ROOT)}', flush=True)


if __name__ == '__main__':
    main()

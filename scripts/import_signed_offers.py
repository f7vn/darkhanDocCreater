"""Import signed PDF replacements, retaining originals and their SHA-256 provenance."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import unicodedata

from import_templates import find_normalized


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('source', type=Path)
    args = parser.parse_args()
    catalog_path = ROOT / 'templates/catalog.json'
    catalog = json.loads(catalog_path.read_text(encoding='utf-8'))
    planned = []
    matched = set()
    for mode, config in catalog.items():
        folder = find_normalized(args.source, config['source_folder'])
        for supplier in config['suppliers']:
            original = config['source_files'][Path(supplier['offer']).name]
            wanted = unicodedata.normalize('NFC', Path(original).with_suffix('.pdf').name)
            sources = [p for p in folder.glob('*.pdf') if unicodedata.normalize('NFC', p.name) == wanted]
            if not sources:
                continue  # Suppliers without an update keep the previous approved template.
            if len(sources) != 1:
                raise ValueError(f'Ambiguous PDF replacement: {wanted}')
            source = sources[0]
            relative = f'{mode}/signed-{supplier["key"]}.pdf'
            destination = ROOT / 'templates' / relative
            digest = hashlib.sha256(source.read_bytes()).hexdigest()
            if destination.exists() and hashlib.sha256(destination.read_bytes()).hexdigest() != digest:
                raise FileExistsError(f'Refusing to overwrite a different signed template: {destination}')
            supplier['signed_offer'] = relative
            supplier['signed_source'] = str(Path(args.source.name) / source.relative_to(args.source))
            supplier['signed_sha256'] = digest
            planned.append((source, destination))
            matched.add(source)
    unmatched = set(args.source.rglob('*.pdf')) - matched
    if unmatched:
        raise ValueError(f'Unmapped signed offers: {sorted(map(str, unmatched))}')
    if not planned:
        raise ValueError('No signed offers found')
    for source, destination in planned:
        if not destination.exists():
            shutil.copy2(source, destination)
        print(f'{source.relative_to(args.source)} -> {destination.relative_to(ROOT)}')
    catalog_path.write_text(json.dumps(catalog, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()

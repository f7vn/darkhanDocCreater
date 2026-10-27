from decimal import Decimal
import hashlib
from pathlib import Path
import re
from tempfile import TemporaryDirectory
import unittest

import pymupdf as fitz

import app


EXPECTED = {
    'aesthetic': {'best', 'nalibai'},
    'finishing': {'abdykhalykov', 'sarsenova'},
    'stands': {'abdicul'},
    'graphics': {'best', 'nalibai'},
    'combined': {'best', 'nalibai'},
}


def image_inventory(document):
    return sorted(hashlib.sha256(document.extract_image(item[0])['image']).hexdigest()
                  for item in document[0].get_images())


class SignedOfferTests(unittest.TestCase):
    def test_unknown_signature_fields_are_never_silently_removed(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / 'signature.pdf'
            with fitz.open() as document:
                page = document.new_page()
                widget = fitz.Widget()
                widget.field_name = 'Real signature field'
                widget.field_type = fitz.PDF_WIDGET_TYPE_SIGNATURE
                widget.rect = fitz.Rect(100, 100, 200, 150)
                page.add_widget(widget)
                document.save(source)
            with self.assertRaisesRegex(ValueError, 'электронная подпись'):
                app.create_signed_offer(source, Path(directory) / 'out.pdf', 'aesthetic', 'best',
                                        Decimal('10'), Decimal('25000'), 0, Decimal(0))

    def test_all_nine_updates_have_exact_source_hashes(self):
        for mode, providers in EXPECTED.items():
            signed = {s['key'] for s in app.KTRU[mode]['suppliers'] if s.get('signed_offer')}
            self.assertEqual(signed, providers)
            for supplier in app.KTRU[mode]['suppliers']:
                if supplier['key'] not in signed:
                    continue
                source = app.TEMPLATE_DIR / supplier['signed_offer']
                self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), supplier['signed_sha256'])
                self.assertEqual(source.parts[-2], mode)
                self.assertIn('КП с подписями и печатями', supplier['signed_source'])

    def test_every_signed_offer_keeps_artwork_and_fills_all_prices(self):
        cases = [
            (Decimal('32.5'), Decimal('21800.75'), 3, Decimal('90000.50')),
            (Decimal('12.34'), Decimal('12345678.90'), 2, Decimal('87654321')),
            (Decimal('100000'), Decimal('100000000'), 10000, Decimal('100000000')),
        ]
        with TemporaryDirectory() as directory:
            for mode, providers in EXPECTED.items():
                for supplier in app.KTRU[mode]['suppliers']:
                    if supplier['key'] not in providers:
                        continue
                    source_path = app.TEMPLATE_DIR / supplier['signed_offer']
                    source_bytes = source_path.read_bytes()
                    with fitz.open(source_path) as source:
                        images = image_inventory(source)
                        original_size = source[0].rect
                        # Crop the original horizontal rules next to editable
                        # cells. White redaction outlines used to erase these.
                        rule = None
                        if supplier['key'] == 'nalibai':
                            y = {'aesthetic': 277.6, 'graphics': 291.4, 'combined': 434.4}[mode]
                            rule = fitz.Rect(450, y, 520, y + 0.3)
                        elif supplier['key'] == 'abdykhalykov':
                            rule = fitz.Rect(340, 508.95, 520, 509.2)
                        elif supplier['key'] == 'abdicul':
                            rule = fitz.Rect(120, 454, 540, 454.3)
                        rule_pixels = source[0].get_pixmap(clip=rule, matrix=fitz.Matrix(4, 4)).samples if rule else None
                    for index, (area, rate, count, stand_rate) in enumerate(cases):
                        with self.subTest(mode=mode, provider=supplier['key'], case=index):
                            area = Decimal(0) if mode == 'stands' else area
                            count = count if mode in ('combined', 'stands') else 0
                            destination = Path(directory) / f'{mode}-{supplier["key"]}-{index}.pdf'
                            app.create_supplier_offer(mode, supplier, destination, area, rate, count, stand_rate)
                            with fitz.open(destination) as actual:
                                self.assertEqual(len(actual), 1)
                                self.assertEqual(actual[0].rect, original_size)
                                self.assertEqual(image_inventory(actual), images)
                                self.assertGreaterEqual(len(images), 2)
                                self.assertFalse(list(actual[0].widgets()))
                                if rule:
                                    self.assertEqual(actual[0].get_pixmap(clip=rule, matrix=fitz.Matrix(4, 4)).samples, rule_pixels)
                                text = actual[0].get_text()
                                self.assertNotIn('SIGN', text)
                                # BEST's original signature baseline is made
                                # of underscores and must remain under the ink.
                                if supplier['key'] == 'best':
                                    text = text.replace('__________________', '')
                                self.assertNotIn('____', text)
                                compact = re.sub(r'\s', '', text)
                                total = app.money(area * rate) + app.money(count * stand_rate)
                                for value in ([area, rate] if area else []) + ([Decimal(count), stand_rate] if count else []) + [total]:
                                    self.assertIn(app.decimal_label(value).replace(' ', ''), compact)
                                if supplier['key'] in ('nalibai', 'abdykhalykov'):
                                    strip = actual[0].get_pixmap(clip=fitz.Rect(1, 0, 30, 800), alpha=False)
                                    self.assertEqual(set(strip.samples), {255})
                                if supplier['key'] == 'best':
                                    top = 474 if mode == 'combined' else 444
                                    samples = actual[0].get_pixmap(clip=fitz.Rect(46, top, 565, top + (59 if count else 45)), alpha=False).samples
                                    self.assertFalse(any(r > 220 and g > 220 and b < 50
                                                         for r, g, b in zip(samples[::3], samples[1::3], samples[2::3])))
                            self.assertEqual(source_path.read_bytes(), source_bytes)


if __name__ == '__main__':
    unittest.main()

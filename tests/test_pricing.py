import unittest
from decimal import Decimal
from unittest.mock import patch

import app


class PricingTests(unittest.TestCase):
    def test_default_competitor_percentages(self):
        actual = app.supplier_rates(Decimal('100000'), (Decimal('7'), Decimal('15')), ('base', 'first', 'second'))
        self.assertEqual(actual, {'base': Decimal('100000'), 'first': Decimal('107000'), 'second': Decimal('115000')})

    def test_round_half_up_to_cents(self):
        actual = app.supplier_rates(Decimal('0.10'), (Decimal('5'), Decimal('15')), ('base', 'first', 'second'))
        self.assertEqual(actual['first'], Decimal('0.11'))
        self.assertEqual(actual['second'], Decimal('0.12'))

    def test_stands_need_no_area_and_use_stand_suppliers(self):
        with patch.object(app, 'generate_package') as generate:
            app.generate_from_payload({'ktru': 'stands', 'standCount': '3', 'standPrice': '100 000'})
        mode, area, count, rates, _ = generate.call_args.args
        self.assertEqual((mode, area, count), ('stands', Decimal(0), 3))
        self.assertEqual(rates, {'abdicul': Decimal('100000'), 'trade': Decimal('107000'), 'andrey': Decimal('115000')})

    def test_custom_markup_applies_to_both_components(self):
        with patch.object(app, 'generate_package') as generate:
            app.generate_from_payload({'ktru': 'combined', 'area': '12,5', 'areaPrice': '20000', 'standCount': '2', 'standPrice': '100000', 'markup1': '8,5', 'markup2': '20'})
        _, area, count, stands, walls = generate.call_args.args
        self.assertEqual((area, count), (Decimal('12.5'), 2))
        self.assertEqual(stands['nalibai'], Decimal('108500'))
        self.assertEqual(walls['nalibai'], Decimal('21700'))
        self.assertEqual(stands['archi'], Decimal('120000'))
        self.assertEqual(walls['archi'], Decimal('24000'))

    def test_invalid_prices_and_percentages(self):
        for value in ('NaN', 'Infinity', '-1', '0', '0.001', '100000001'):
            with self.subTest(price=value), self.assertRaises(ValueError):
                app.parse_money(value, 'цену за стенд')
        for value in ('NaN', 'Infinity', '-1', '1001', '1.111'):
            with self.subTest(markup=value), self.assertRaises(ValueError):
                app.parse_markup(value)


if __name__ == '__main__':
    unittest.main()

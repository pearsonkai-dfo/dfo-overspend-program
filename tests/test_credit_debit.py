import unittest

import main


class CreditDebitTests(unittest.TestCase):
    def test_debit_keeps_amount(self):
        self.assertEqual(main.resolve_matching_amount_for_transaction_code("D", 125.50), 125.50)

    def test_credit_negates_amount(self):
        self.assertEqual(main.resolve_matching_amount_for_transaction_code("C", 125.50), -125.50)

    def test_unrecognized_code_returns_none(self):
        self.assertIsNone(main.resolve_matching_amount_for_transaction_code("X", 125.50))


if __name__ == "__main__":
    unittest.main()

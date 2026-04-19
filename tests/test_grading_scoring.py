import unittest

from src.quiz_pool.main import earns_full_credit


class EarnsFullCreditTests(unittest.TestCase):
    def test_exact_match_earns_full_credit(self) -> None:
        self.assertTrue(earns_full_credit(["A", "C"], ["C", "A"]))

    def test_missing_correct_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit(["A"], ["A", "C"]))

    def test_extra_wrong_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit(["A", "C", "D"], ["A", "C"]))

    def test_blank_answer_earns_no_credit(self) -> None:
        self.assertFalse(earns_full_credit([], ["A", "C"]))


if __name__ == "__main__":
    unittest.main()

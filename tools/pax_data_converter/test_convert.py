import datetime as dt
import unittest

from convert import (
    build_pax_rows,
    compute_age,
    guess_gender,
    parse_travel_date,
    pax_type_for_age,
    split_name,
)


class SplitNameTests(unittest.TestCase):
    def test_two_parts(self):
        self.assertEqual(split_name("TUYCHIEVA DILNOZA"), ("DILNOZA", "TUYCHIEVA"))

    def test_multi_word_given_name(self):
        self.assertEqual(
            split_name("GANIYUT ABDUVALI OGLI"), ("ABDUVALI OGLI", "GANIYUT")
        )

    def test_single_word(self):
        self.assertEqual(split_name("MADONNA"), ("MADONNA", "MADONNA"))


class GuessGenderTests(unittest.TestCase):
    def test_female_surname(self):
        self.assertEqual(guess_gender("DILNOZA", "TUYCHIEVA"), "Female")

    def test_male_surname(self):
        self.assertEqual(guess_gender("IKHTIYOR", "DJUMAKULOV"), "Male")

    def test_non_declining_surname_uses_first_name(self):
        self.assertEqual(guess_gender("NARGIZA", "GANIYUT"), "Female")
        self.assertEqual(guess_gender("ULUGBEK", "GANIYUT"), "Male")


class AgeAndPaxTypeTests(unittest.TestCase):
    def test_compute_age_before_birthday(self):
        self.assertEqual(compute_age(dt.date(2020, 8, 20), dt.date(2026, 8, 14)), 5)

    def test_compute_age_after_birthday(self):
        self.assertEqual(compute_age(dt.date(2020, 8, 1), dt.date(2026, 8, 14)), 6)

    def test_pax_type_thresholds(self):
        self.assertEqual(pax_type_for_age(1), "INF")
        self.assertEqual(pax_type_for_age(2), "CHD")
        self.assertEqual(pax_type_for_age(11), "CHD")
        self.assertEqual(pax_type_for_age(12), "ADT")


class ParseTravelDateTests(unittest.TestCase):
    def test_explicit_date(self):
        self.assertEqual(
            parse_travel_date("14.08.2026", "whatever.xlsx", None),
            dt.date(2026, 8, 14),
        )

    def test_from_filename(self):
        self.assertEqual(
            parse_travel_date(None, "TASBUS_14.08.xlsx", 2026), dt.date(2026, 8, 14)
        )

    def test_missing_date_raises(self):
        with self.assertRaises(ValueError):
            parse_travel_date(None, "no_date_here.xlsx", 2026)


class BuildPaxRowsTests(unittest.TestCase):
    def test_full_row(self):
        input_rows = [{
            "full_name": "TUYCHIEVA DILNOZA",
            "dob": dt.datetime(2001, 11, 19),
            "passport_no": "FA6450108",
            "expiry": dt.datetime(2027, 9, 28),
            "room": "TWIN",
        }]
        rows = build_pax_rows(input_rows, dt.date(2026, 8, 14))
        self.assertEqual(len(rows), 1)
        row = rows[0]
        self.assertEqual(row["Pax Type"], "ADT")
        self.assertEqual(row["Title"], "mrs")
        self.assertEqual(row["First Name"], "DILNOZA")
        self.assertEqual(row["Last Name"], "TUYCHIEVA")
        self.assertEqual(row["Nationality"], "Uzbekistan")


if __name__ == "__main__":
    unittest.main()

import unittest

from excel_splitter import utils


class SheetNameRegistryTests(unittest.TestCase):
    def test_collisions_match_existing_names_and_skip_reserved_suffixes(self):
        registry = utils.SheetNameRegistry()
        names = ["Sales - 2", "Sales", "sales", "Sales", "Sales - 2", "x/y", "x:y"]
        self.assertEqual(
            [utils.sanitize_sheet_name(name, registry) for name in names],
            ["Sales - 2", "Sales", "sales - 3", "Sales - 4", "Sales - 2 - 2", "x-y", "x-y - 2"],
        )

    def test_unicode_and_long_names_remain_unique(self):
        registry = utils.SheetNameRegistry()
        self.assertEqual(utils.sanitize_sheet_name("Straße", registry), "Straße")
        self.assertEqual(utils.sanitize_sheet_name("STRASSE", registry), "STRASSE - 2")
        prefix = "a" * 31
        result = [utils.sanitize_sheet_name(prefix + str(i), registry) for i in range(1000)]
        self.assertEqual(result[0], prefix)
        self.assertEqual(result[-1], "a" * 24 + " - 1000")
        self.assertEqual(len({name.casefold() for name in result}), 1000)
        self.assertTrue(all(len(name) <= 31 for name in result))

    def test_plain_sets_remain_supported(self):
        used = {"Sales"}
        self.assertEqual(utils.sanitize_sheet_name("sales", used), "sales - 2")
        self.assertEqual(used, {"Sales", "sales - 2"})

    def test_case_variants_and_unicode_truncation_use_first_available_suffix(self):
        registry = utils.SheetNameRegistry()
        short = "ß" * 15 + "a"
        long = "ss" * 15 + "A"
        self.assertEqual(utils.sanitize_sheet_name(short, registry), short)
        self.assertEqual(utils.sanitize_sheet_name(long, registry), "s" * 27 + " - 2")
        self.assertEqual(utils.sanitize_sheet_name(short, registry), short + " - 2")

        registry = utils.SheetNameRegistry()
        variants = [format(i, "031b").replace("0", "a").replace("1", "A") for i in range(1000)]
        results = [utils.sanitize_sheet_name(name, registry) for name in variants]
        self.assertEqual(results[0], "a" * 31)
        self.assertEqual(results[-1][-7:], " - 1000")
        self.assertEqual(len({name.casefold() for name in results}), 1000)


if __name__ == "__main__":
    unittest.main()

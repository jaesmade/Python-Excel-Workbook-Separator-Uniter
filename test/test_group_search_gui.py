"""Group search and bulk selection across paginated results."""

import unittest
from tkinter import TclError

from excel_splitter.gui.main_window import ExcelSplitterApp


class GroupSearchTests(unittest.TestCase):
    def setUp(self):
        try:
            self.app = ExcelSplitterApp()
        except TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.addCleanup(self.app.destroy)
        self.app.groups = {(f"Team {i:03d}", "North"): [(i,)] for i in range(120)}
        self.app.groups[("Other", "South")] = [(999,)]
        self.app._render_groups()

    def test_search_filters_labels_and_resets_page_without_losing_selection(self):
        app = self.app
        app.group_checkboxes[("Team 000", "North")].toggle()
        app._change_group_page(2)
        app.group_search_var.set("TEAM 00")
        self.assertEqual(app.group_page, 0)
        self.assertEqual(len(app.group_checkboxes), 10)
        self.assertNotIn(("Team 000", "North"), app.selected_group_keys)
        app.group_search_var.set("south")
        self.assertEqual(set(app.group_checkboxes), {("Other", "South")})
        app.group_search_var.set("no matching group")
        self.assertEqual(app.group_checkboxes, {})
        self.assertEqual(app.group_next.cget("state"), "disabled")
        app.group_search_var.set("")
        self.assertNotIn(("Team 000", "North"), app.selected_group_keys)
        self.assertEqual(len(app.selected_group_keys), 120)

    def test_bulk_selection_applies_to_all_matching_pages_only(self):
        app = self.app
        app.group_search_var.set("north")
        app.group_deselect_all.invoke()
        self.assertEqual(app.selected_group_keys, {("Other", "South")})
        app._change_group_page(2)
        self.assertTrue(all(not checkbox.get() for checkbox in app.group_checkboxes.values()))
        app.group_select_all.invoke()
        self.assertEqual(len(app.selected_group_keys), 121)
        app.group_search_var.set("south")
        app.group_deselect_all.invoke()
        self.assertEqual(len(app.selected_group_keys), 120)
        self.assertNotIn(("Other", "South"), app.selected_group_keys)

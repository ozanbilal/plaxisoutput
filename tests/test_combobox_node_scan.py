import unittest

import export_plaxis_data as core


class FakeItem:
    def __init__(self, text):
        self.text = text
        self.clicked = False

    def window_text(self):
        return self.text

    def click_input(self):
        self.clicked = True


class FakeCombo:
    def __init__(self, pages, fail_keys=False):
        self.pages = [[FakeItem(text) for text in page] for page in pages]
        self.page = 0
        self.fail_keys = fail_keys
        self.expanded = False
        self.collapsed = False

    def descendants(self, control_type=None):
        self._assert_list_item(control_type)
        return self.pages[self.page]

    def type_keys(self, key, set_foreground=True):
        if self.fail_keys:
            raise RuntimeError("keys unavailable")
        if key == "{HOME}":
            self.page = 0
        elif key == "{PGDN}" and self.page < len(self.pages) - 1:
            self.page += 1

    def expand(self):
        self.expanded = True
        self.collapsed = False

    def collapse(self):
        self.collapsed = True
        self.expanded = False

    def select(self, _text):
        raise RuntimeError("force scroll fallback")

    @staticmethod
    def _assert_list_item(control_type):
        if control_type != "ListItem":
            raise AssertionError(f"unexpected control type: {control_type}")


class ComboboxNodeScanTests(unittest.TestCase):
    def test_collects_nodes_across_visible_pages(self):
        combo = FakeCombo(
            [
                ["Node 1", "Node 2"],
                ["Node 3", "Node 4"],
                ["Node 5"],
            ]
        )

        names = core._collect_combobox_node_names(combo, wait_sec=0.0)

        self.assertEqual(names, ["Node 1", "Node 2", "Node 3", "Node 4", "Node 5"])

    def test_selects_node_across_visible_pages(self):
        combo = FakeCombo(
            [
                ["Node 1", "Node 2"],
                ["Node 3", "Node 4"],
                ["Node 5"],
            ]
        )

        mode = core._select_combobox_node_by_scroll(combo, "Node 4", wait_sec=0.0)

        self.assertEqual(mode, "combobox:number-match-scroll")
        self.assertTrue(combo.pages[1][1].clicked)
        self.assertTrue(combo.collapsed)

    def test_collect_fails_instead_of_returning_visible_page_when_keys_fail(self):
        combo = FakeCombo([["Node 1", "Node 2"]], fail_keys=True)

        with self.assertRaisesRegex(RuntimeError, "focus/rewind"):
            core._collect_combobox_node_names(combo, wait_sec=0.0)

    def test_select_collapses_dropdown_when_rewind_fails(self):
        combo = FakeCombo([["Node 1", "Node 2"]], fail_keys=True)

        with self.assertRaisesRegex(RuntimeError, "focus/rewind"):
            core._select_combobox_node_by_scroll(combo, "Node 2", wait_sec=0.0)

        self.assertTrue(combo.collapsed)


if __name__ == "__main__":
    unittest.main()

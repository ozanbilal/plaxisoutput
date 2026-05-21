import unittest

import pandas as pd

import export_plaxis_data as core


class NodeTimeHistoryWideTests(unittest.TestCase):
    def test_builds_one_wide_sheet_per_direction_phase(self):
        time_df = pd.DataFrame(
            [
                {
                    "Direction": direction,
                    "Phase": phase,
                    "CurvePointId": series,
                    "Series": series,
                    "NodeName": series,
                    "Step": step,
                    "DynamicTime": step * 0.01,
                    "Acceleration": accel,
                }
                for direction, phase in [("X", "DD2_X_A [Phase_1]"), ("Y", "DD2_Y_A [Phase_2]")]
                for step in [1, 2, 3]
                for series, accel in [("Node 1", step * 10.0), ("Node 2", step * 20.0)]
            ]
        )

        sheets = core._build_node_timehistory_wide_sheets(time_df, "DynamicTime")

        self.assertEqual([name for name, _ in sheets], ["NodeTH_X_Phase_1", "NodeTH_Y_Phase_2"])
        first = sheets[0][1]
        self.assertEqual(first.columns.tolist(), ["Step", "DynamicTime", "Node 1", "Node 2"])
        self.assertEqual(len(first), 3)
        self.assertEqual(first["Node 1"].tolist(), [10.0, 20.0, 30.0])
        self.assertEqual(first["Node 2"].tolist(), [20.0, 40.0, 60.0])


if __name__ == "__main__":
    unittest.main()

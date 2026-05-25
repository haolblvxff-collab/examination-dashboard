import io
import unittest

import openpyxl

import app


def make_workbook(rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["姓名", "学校", "考试", "语文", "数学", "英语", "物理", "化学", "生物"])
    for row in rows:
        sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


class ParseExcelTests(unittest.TestCase):
    def setUp(self):
        app.DATA.update({"students": {}, "exams": [], "schools": [], "filename": ""})

    def test_students_with_same_name_in_different_schools_are_distinct(self):
        result = app.parse_excel(
            make_workbook(
                [
                    ["张三", "一中", "一模", 10, 20, 30, 40, 50, 60],
                    ["张三", "二中", "一模", 60, 50, 40, 30, 20, 10],
                ]
            )
        )
        self.assertEqual(len(result["students"]), 2)
        self.assertEqual(result["schools"], ["一中", "二中"])

    def test_zero_scores_are_valid_and_included_in_stats(self):
        result = app.parse_excel(
            make_workbook(
                [
                    ["甲", "一中", "一模", 0, 0, 0, 0, 0, 0],
                    ["乙", "一中", "一模", 60, 60, 60, 60, 60, 60],
                ]
            )
        )
        app.DATA.update(result)
        stats = app.get_stats()
        self.assertEqual(stats["n_students"], 2)
        self.assertEqual(stats["avg_total"], 180.0)
        self.assertEqual(stats["subject_avg"]["语文"], 30.0)

    def test_exam_order_preserves_known_mocks_before_custom_names(self):
        result = app.parse_excel(
            make_workbook(
                [
                    ["甲", "一中", "月考", 1, 1, 1, 1, 1, 1],
                    ["甲", "一中", "二模", 1, 1, 1, 1, 1, 1],
                    ["甲", "一中", "一模", 1, 1, 1, 1, 1, 1],
                ]
            )
        )
        self.assertEqual(result["exams"], ["一模", "二模", "月考"])


if __name__ == "__main__":
    unittest.main()

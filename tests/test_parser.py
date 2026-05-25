import unittest

from sotong_parser import parse_input


class ParseInputTest(unittest.TestCase):
    def test_excel_tab_pair(self):
        self.assertEqual(
            parse_input('충주중학교\t홍길동'),
            [{'org': '충주중학교', 'name': '홍길동'}],
        )

    def test_hwp_alternating_lines(self):
        self.assertEqual(
            parse_input('충주중학교\n홍길동\n청주고등학교\n김철수'),
            [
                {'org': '충주중학교', 'name': '홍길동'},
                {'org': '청주고등학교', 'name': '김철수'},
            ],
        )

    def test_name_only_multiple_people(self):
        self.assertEqual(
            parse_input('홍길동 김철수'),
            [
                {'org': '', 'name': '홍길동'},
                {'org': '', 'name': '김철수'},
            ],
        )

    def test_school_abbreviation_and_full_name(self):
        self.assertEqual(
            parse_input('충주중 홍길동'),
            [{'org': '충주중학교', 'name': '홍길동'}],
        )
        self.assertEqual(
            parse_input('충주중학교 홍길동'),
            [{'org': '충주중학교', 'name': '홍길동'}],
        )


if __name__ == '__main__':
    unittest.main()

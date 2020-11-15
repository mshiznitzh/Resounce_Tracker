import unittest
import datetime
import main

class MyTestCase(unittest.TestCase):
    def test_find_the_counstrunction_Season(self):

        self.assertEqual(main.find_the_counstrunction_Season(datetime.date(2020, 7, 1)), datetime.date(2020, 12, 31))
        self.assertEqual(main.find_the_counstrunction_Season(datetime.date(2020, 1, 1)), datetime.date(2020, 6, 30))


if __name__ == '__main__':
    unittest.main()

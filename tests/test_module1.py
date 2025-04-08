import unittest
from my_python_project.module1 import function1

class TestModule1(unittest.TestCase):
    def test_function1(self):
        self.assertEqual(function1(), "Hello from module1")

if __name__ == '__main__':
    unittest.main()
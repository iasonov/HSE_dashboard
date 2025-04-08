import unittest
from my_python_project.module2 import function2

class TestModule2(unittest.TestCase):
    def test_function2(self):
        self.assertEqual(function2(), "Hello from module2")

if __name__ == '__main__':
    unittest.main()
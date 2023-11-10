import unittest
import main


class TestGame(unittest.TestCase):
    def test_imput(self):
        result = main.guess_function(1, 4, 3, 3)
        self.assertTrue(result) 

    def test_imput_wrong_guess(self):
        result = main.guess_function(1, 4, 3, 1)
        self.assertFalse(result) 

    def test_imput_wrong_number(self):
        result = main.guess_function(1, 4, 5, 1)
        self.assertFalse(result) 

    def test_imput_wrong_imput(self):
        result = main.guess_function(1, 4, "string", 1)
        self.assertFalse(result)     

if __name__ == '__main__':
    unittest.main()
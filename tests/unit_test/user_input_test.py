"""
    [summary]

[extended_summary]
"""
import unittest
from unittest import mock
import sys

from src.user_input import UserInput

class UserInputTests(unittest.TestCase):

    user_input = UserInput()

    def setUp(self) -> None:
        self.patcher = mock.patch('src.user_input.utils.MetaGetVariable',return_value = "True")
        self.patcher.start()

    def test_get_user_input_is_from_gui(self):

        # Act
        self.user_input.get_user_input_from_gui()

        # Assert
        self.assertEqual(self.user_input.metadb_2d_input,"True")

    #Arrange
    @mock.patch.object(sys,"platform",new = "win32")
    @mock.patch.object(UserInput,"continue_in_windows_cmd")
    def test_windows_interactive_mode_called(self,mock_method):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    # #Arrange
    # @mock.patch.object(sys,"platform",new = "win32")
    # @mock.patch.object(UserInput,"continue_in_windows_cmd")
    # def test_windows_interactive_mode_called_with(self,mock_method):

    #     # Act
    #     self.user_input.get_user_input_from_interactive_mode()

    #     # Assert
    #     mock_method.assert_called_with( )

    #Arrange
    @mock.patch.object(sys,"platform",new = "linux")
    @mock.patch.object(UserInput,"continue_in_linux_cmd")
    def test_linux_interactice_mode_called(self,mock_method):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    #Arrange
    @mock.patch.object(sys,"platform",new = "win32")
    @mock.patch('src.user_input.UserInput.continue_in_windows_cmd',return_value = 0)
    @mock.patch.object(UserInput,"get_user_input_from_gui")
    def test_get_user_input_from_gui_called(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    #Arrange
    @mock.patch.object(sys,"platform",new = "win32")
    @mock.patch('src.user_input.UserInput.continue_in_windows_cmd',return_value = -1)
    @mock.patch.object(UserInput,"run_interactive_mode")
    def test_run_interactive_mode_called(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    #Arrange
    @mock.patch.object(sys,"platform",new = "win32")
    @mock.patch('src.user_input.UserInput.continue_in_windows_cmd',return_value = 0)
    @mock.patch('builtins.print')
    def test_print_called(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        assert mock_method.mock_calls == [mock.call('Script is running in GUI mode')]

    #Arrange
    @mock.patch.object(sys,"platform",new = "win32")
    @mock.patch('src.user_input.UserInput.continue_in_windows_cmd',return_value = -1)
    @mock.patch('src.user_input.UserInput.run_interactive_mode',return_value = None)
    @mock.patch('builtins.print')
    def test_non_gui_print_called(self,mock_method1,mock_method2,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        assert mock_method1.mock_calls == [mock.call('Script is running in NON GUI mode')]

    #Arrange
    @mock.patch.object(sys,"platform",new = "linux2")
    @mock.patch('src.user_input.UserInput.continue_in_windows_cmd',return_value = 0)
    @mock.patch.object(UserInput,"get_user_input_from_gui")
    def test_get_user_input_from_gui_called_in_linux(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    #Arrange
    @mock.patch.object(sys,"platform",new = "linux2")
    @mock.patch('src.user_input.UserInput.continue_in_linux_cmd',return_value = -1)
    @mock.patch.object(UserInput,"run_interactive_mode")
    def test_run_interactive_mode_called_in_linux(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        mock_method.assert_called()

    #Arrange
    @mock.patch.object(sys,"platform",new = "linux2")
    @mock.patch('src.user_input.UserInput.continue_in_linux_cmd',return_value = 0)
    @mock.patch('builtins.print')
    def test_print_called_in_linux(self,mock_method,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        assert mock_method.mock_calls == [mock.call('Script is running in GUI mode')]

    #Arrange
    @mock.patch.object(sys,"platform",new = "linux2")
    @mock.patch('src.user_input.UserInput.continue_in_linux_cmd',return_value = -1)
    @mock.patch('src.user_input.UserInput.run_interactive_mode',return_value = None)
    @mock.patch('builtins.print')
    def test_non_gui_print_called_in_linux(self,mock_method1,mock_method2,mock_return):

        # Act
        self.user_input.get_user_input_from_interactive_mode()

        # Assert
        assert mock_method1.mock_calls == [mock.call('Script is running in NON GUI mode')]

    def tearDown(self) -> None:
        self.patcher.stop()

if __name__ == '__main__':
    unittest.main()

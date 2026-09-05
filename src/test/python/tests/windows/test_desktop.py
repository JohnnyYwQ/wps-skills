import unittest
from unittest.mock import Mock


class DesktopWindowTests(unittest.TestCase):
    def test_bound_window_is_detected_without_a_document_in_its_title(self):
        from wps_skills.windows.desktop import visible_document_window
        user = Mock()
        user.IsWindow.return_value = True
        user.GetAncestor.return_value = 42
        user.IsWindowVisible.return_value = True
        user.IsIconic.return_value = False
        user.GetForegroundWindow.return_value = 42
        def owner(hwnd, pointer):
            pointer._obj.value = 123
            return 1
        def bounds(hwnd, pointer):
            pointer._obj.left = 0
            pointer._obj.top = 0
            pointer._obj.right = 1280
            pointer._obj.bottom = 800
            return True
        user.GetWindowThreadProcessId.side_effect = owner
        user.GetWindowRect.side_effect = bounds
        reference = {'hwnd': 7, 'processId': 123}
        result = visible_document_window(reference, user32=user)
        self.assertEqual(42, result['hwnd'])
        self.assertEqual(7, result['documentHwnd'])
        self.assertEqual(1280, result['width'])
        user.EnumWindows.assert_not_called()
        user.GetWindowTextW.assert_not_called()
        user.IsWindow.return_value = False
        self.assertIsNone(visible_document_window(reference, user32=user))
        user.IsWindow.return_value = True
        self.assertIsNone(visible_document_window(dict(reference, processId=456), user32=user))


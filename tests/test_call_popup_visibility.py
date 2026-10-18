import gc
import tkinter as tk
import unittest
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.call_session import IncomingCallSessionTracker, MissedCall
from sms_ui.app_ui_namespace_runtime import center_window_namespace_runtime
from sms_ui.call_popup_namespace_runtime import (
    close_call_popup_namespace_runtime,
    close_missed_call_popup_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_missed_call_popup_namespace_runtime,
    show_call_popup_namespace_runtime,
    show_missed_call_popup_namespace_runtime,
)
from sms_ui.window_icon_runtime import install_window_icon_runtime


class CallPopupVisibilityTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest('Tk display unavailable')
        self.root.withdraw()
        self.root.attributes('-alpha', 0)
        self.root.geometry('640x440+250+150')
        self.root.deiconify()
        self.root.update()
        self.errors = []
        self.root.report_callback_exception = lambda *args: self.errors.append(str(args))
        module = SimpleNamespace(Toplevel=tk.Toplevel)
        boxes = SimpleNamespace(**{name: lambda *a, **k: None for name in (
            'showinfo', 'showwarning', 'showerror', 'askyesno')})
        install_window_icon_runtime(self.root, module, boxes,
            icon_path=str(Path(__file__).resolve().parents[1] / 'icon.ico'),
            path_exists=lambda p: Path(p).is_file(), log_error=self.errors.append)

        def create(*args, **kwargs):
            window = module.Toplevel(*args, **kwargs)
            window.attributes('-alpha', 0)
            if self.move_on_map:
                def reposition(event):
                    if event.widget is window:
                        window.unbind('<Map>', binding)
                        window.after_idle(lambda: window.geometry('+0+0'))
                binding = window.bind('<Map>', reposition, add='+')
            return window

        self.move_on_map = False
        self.patch = patch('tkinter.Toplevel', create)
        self.patch.start()
        self.namespace = {
            'root': self.root, 'current_call_popup': None, 'current_dial_popup': None,
            'current_missed_call_popup': None, 'serial_lock': None, 'serial_obj': None,
            'INCOMING_CALL_SESSION': IncomingCallSessionTracker(),
            'center_window': lambda w, p: center_window_namespace_runtime({}, w, p),
            'port_ui': lambda *a: None, 'set_status': lambda *a: None,
            'show_window': lambda: None,
            'run_on_ui_thread': lambda cb, _post: cb(), 'ui_post': lambda cb: cb(),
            'log_file_only': self.errors.append,
        }
        self.namespace['close_call_popup'] = lambda: close_call_popup_namespace_runtime(self.namespace)

    def tearDown(self):
        self.patch.stop()
        self.root.destroy()
        self.namespace.clear()
        del self.root
        gc.collect()

    def incoming(self, generation='local-1'):
        self.namespace['INCOMING_CALL_SESSION'].start('10086')
        show_call_popup_namespace_runtime(self.namespace, '10086', generation)
        self.root.update()
        return self.namespace['current_call_popup']

    def missed(self):
        show_missed_call_popup_namespace_runtime(
            self.namespace, MissedCall('10086', datetime(2026, 1, 1)))
        self.root.update()
        return self.namespace['current_missed_call_popup']

    def assert_incoming_above_missed(self, incoming, missed):
        stack = tuple(map(str, self.root.tk.splitlist(self.root.tk.call('wm', 'stackorder', '.'))))
        self.assertTrue(incoming.winfo_ismapped())
        self.assertTrue(missed.winfo_ismapped())
        self.assertGreater(stack.index(str(incoming)), stack.index(str(missed)))
        self.assertEqual(self.errors, [])

    def test_late_missed_notice_cannot_cover_active_incoming_window(self):
        incoming = self.incoming()
        missed = self.missed()
        self.assert_incoming_above_missed(incoming, missed)

    def test_refreshing_old_missed_notice_keeps_incoming_visible(self):
        missed = self.missed()
        incoming = self.incoming()
        self.missed()
        self.assert_incoming_above_missed(incoming, missed)

    def test_focusing_missed_notice_cannot_cover_active_incoming_window(self):
        missed = self.missed()
        incoming = self.incoming()
        missed.focus_force()
        self.root.update()
        self.assert_incoming_above_missed(incoming, missed)
        self.assertIsNone(self.root.grab_current())
        self.assertFalse(self.root.tk.call('after', 'info'))

    def test_hiding_and_remapping_incoming_updates_missed_notice_priority(self):
        missed = self.missed()
        incoming = self.incoming()
        self.assertFalse(missed.attributes('-topmost'))

        incoming.withdraw()
        self.root.update()
        self.assertTrue(missed.attributes('-topmost'))
        self.assertIs(self.namespace['current_call_popup'], incoming)

        incoming.deiconify()
        self.root.update()
        missed.focus_force()
        self.root.update()
        self.assertFalse(missed.attributes('-topmost'))
        self.assert_incoming_above_missed(incoming, missed)

    def test_destroying_old_window_does_not_raise_missed_above_new_call(self):
        missed = self.missed()
        old = self.incoming('local-old')
        set_call_popup_namespace_runtime(self.namespace, None)
        self.namespace['INCOMING_CALL_SESSION'].finish()
        incoming = self.incoming('local-new')

        old.destroy()
        self.root.update()
        missed.focus_force()
        self.root.update()
        self.assertIs(self.namespace['current_call_popup'], incoming)
        self.assertFalse(missed.attributes('-topmost'))
        self.assert_incoming_above_missed(incoming, missed)

    def test_closing_incoming_restores_missed_notice_priority(self):
        missed = self.missed()
        incoming = self.incoming()
        self.assertFalse(missed.attributes('-topmost'))

        close_call_popup_namespace_runtime(self.namespace)
        self.root.update()
        self.assertFalse(incoming.winfo_exists())
        self.assertIs(self.namespace['current_missed_call_popup'], missed)
        self.assertTrue(missed.winfo_ismapped())
        self.assertTrue(missed.attributes('-topmost'))
        self.assertEqual(self.errors, [])

    def test_destroying_incoming_restores_missed_notice_priority(self):
        missed = self.missed()
        incoming = self.incoming()
        self.assertFalse(missed.attributes('-topmost'))

        incoming.destroy()
        self.root.update()
        self.assertTrue(missed.winfo_ismapped())
        self.assertTrue(missed.attributes('-topmost'))
        self.assertEqual(self.errors, [])

    def test_closing_missed_notice_keeps_incoming_visible(self):
        incoming = self.incoming()
        missed = self.missed()
        self.assertFalse(missed.attributes('-topmost'))
        missed.focus_force()
        self.root.update()
        self.assert_incoming_above_missed(incoming, missed)

        close_missed_call_popup_namespace_runtime(self.namespace)
        self.root.update()
        self.assertFalse(missed.winfo_exists())
        self.assertIs(self.namespace['current_call_popup'], incoming)
        self.assertTrue(incoming.winfo_ismapped())
        self.assertTrue(incoming.attributes('-topmost'))
        self.assertEqual(self.errors, [])

    def test_native_stacking_error_is_logged(self):
        missed = self.missed()
        self.incoming()
        with patch.object(missed, 'attributes', side_effect=tk.TclError('window operation failed')):
            set_missed_call_popup_namespace_runtime(self.namespace, missed)
        self.assertEqual(len(self.errors), 1)
        self.assertIn('Sync phone popup stacking failed', self.errors[0])

    def test_center_is_restored_after_first_map_repositions_window(self):
        self.move_on_map = True
        incoming = self.incoming()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - incoming.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - incoming.winfo_height()) // 2
        self.assertEqual((incoming.winfo_x(), incoming.winfo_y()), (x, y))
        self.assertEqual(self.errors, [])

    def test_destroy_during_map_cancels_pending_placement(self):
        self.namespace['INCOMING_CALL_SESSION'].start('10086')
        show_call_popup_namespace_runtime(self.namespace, '10086', 'local-1')
        window = self.namespace['current_call_popup']
        window.bind('<Map>', lambda event: window.destroy() if event.widget is window else None, add='+')
        self.root.update()
        self.assertFalse(window.winfo_exists())
        self.assertEqual(self.errors, [])
        self.assertFalse(self.root.tk.call('after', 'info'))

    def test_destroy_during_placement_does_not_raise_tcl_error(self):
        original_center = self.namespace['center_window']
        centers = []

        def center(window, parent):
            centers.append(window)
            if len(centers) == 2:
                window.after_idle(window.destroy)
            original_center(window, parent)

        self.namespace['center_window'] = center
        window = self.incoming()
        self.assertFalse(window.winfo_exists())
        self.assertEqual(self.errors, [])

    def test_repeated_redial_keeps_center_and_active_window_above_notice(self):
        for i in range(30):
            self.root.geometry(f'640x440+{250 + i % 3 * 50}+150')
            self.root.update()
            incoming = self.incoming(f'local-{i}')
            missed = self.missed()
            self.assert_incoming_above_missed(incoming, missed)
            self.assertEqual(incoming.winfo_x(), self.root.winfo_rootx()
                             + (self.root.winfo_width() - incoming.winfo_width()) // 2)
            close_call_popup_namespace_runtime(self.namespace)
            self.namespace['INCOMING_CALL_SESSION'].finish()
            self.root.update()
        self.assertEqual(self.errors, [])

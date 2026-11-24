"""Window ownership regressions use an invented library and no desktop applications."""

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

SKILLS = Path(__file__).resolve().parents[1] / "skills"
sys.path.insert(0, str(SKILLS / "cite-from-pdfs/scripts"))
import sync_to_endnote as sync
import endnote_window as windows


class EndNoteWindowTests(unittest.TestCase):
    def setUp(self):
        self.process = Mock(pid=41)
        self.process.poll.return_value = None
        self.target = windows.TargetWindow(self.process, Path("target.enl"), 101, "Afx:Fixture", 12345)
        self.main = {101: ("EndNote 21 - target.enl", "Afx:Fixture", 0)}
        self.view = dict(self.main)
        self.start = patch.object(windows, "_process_start", return_value=12345)
        self.holders = patch.object(windows, "_holders", return_value={(41, 12345)})
        self.enum = patch.object(windows, "_windows", side_effect=lambda pid: self.view.copy())
        self.processes = {41}
        self.peers = patch.object(windows, "_session_endnote_pids", side_effect=lambda pid: self.processes.copy())
        for context in (self.start, self.holders, self.enum, self.peers):
            context.start()
            self.addCleanup(context.stop)

    def test_hidden_second_process_blocks_action_even_without_another_window(self):
        self.processes.add(44)
        with self.assertRaises(windows.WindowInterference):
            self.target.require_exclusive()

    def test_different_library_owning_a_same_named_dialog_is_rejected(self):
        self.view.update({102: ("EndNote 21 - foreign.enl", "Afx:Fixture", 0),
                          103: ("Import File", "#32770", 102)})
        with self.assertRaises(windows.WindowInterference):
            self.target.wait_dialog((), lambda hwnd, name: name == "Import File", .5)

    def test_same_title_with_wrong_pid_start_or_missing_database_handle(self):
        with patch.object(windows, "_process_start", return_value=34567):
            with self.assertRaises(windows.WindowInterference):
                self.target.require_exclusive()
        with patch.object(windows, "_holders", return_value={(99, 12345)}):
            with self.assertRaises(windows.WindowInterference):
                self.target.require_exclusive()

    def test_dialog_requires_exact_owner_even_if_the_title_and_pid_match(self):
        self.view[103] = ("Import File", "#32770", 101)
        self.view[104] = ("Open", "#32770", 103)
        self.target.require_exclusive((103, 104))
        self.view[104] = ("Open", "#32770", 101)
        with self.assertRaises(windows.WindowInterference):
            self.target.require_exclusive((103, 104))

    def test_only_cancel_on_verified_task_owned_promotion(self):
        import win32gui
        self.view[106] = (windows.PROMOTION_TITLE, "#32770", 101)
        with (patch.object(windows, "_promotion_cancel", return_value=202),
              patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.pop(106)) as click):
            self.target.require_exclusive()
        click.assert_called_once_with(202, 0x00F5, 0, 0, 0x0002, 2000)

    def test_announcement_controls_can_finish_painting_after_its_frame(self):
        import win32gui
        self.view[106] = (windows.PROMOTION_TITLE, "#32770", 101)
        with (patch.object(windows, "_promotion_cancel", side_effect=[None, 202, 202]),
              patch.object(windows.time, "sleep") as waited,
              patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.pop(106)) as click):
            self.target.require_exclusive()
        waited.assert_called_once()
        click.assert_called_once_with(202, 0x00F5, 0, 0, 0x0002, 2000)

    def test_startup_rechecks_delayed_windows_before_dispatch(self):
        import win32gui
        def late_popup(_):
            self.view[106] = (windows.PROMOTION_TITLE, "#32770", 101)
        with (patch.object(windows.time, "monotonic", side_effect=[0, .1, .2, .3, 2.6]),
              patch.object(windows.time, "sleep", side_effect=late_popup),
              patch.object(windows, "_promotion_cancel", return_value=202),
              patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.pop(106)) as click):
            self.target.wait_for_startup(10)
        click.assert_called_once_with(202, 0x00F5, 0, 0, 0x0002, 2000)

    def test_only_previously_cancelled_hidden_owner_of_import_can_bridge(self):
        self.view[108] = ("Import File", "#32770", 106)
        with patch.object(windows, "_is_hidden_promotion", return_value=True) as matches:
            with self.assertRaises(windows.WindowInterference):
                self.target.require_exclusive((108,))
            matches.assert_not_called()
            self.target.cancelled_promotion = (106, 202)
            self.target.require_exclusive((108,))
            matches.assert_called_once_with(106, 202, 101, 41)
            self.view[108] = ("Save Changes", "#32770", 106)
            with self.assertRaises(windows.WindowInterference):
                self.target.require_exclusive((108,))
        self.view[108] = ("Import File", "#32770", 106)
        with patch.object(windows, "_is_hidden_promotion", return_value=False):
            with self.assertRaises(windows.WindowInterference):
                self.target.require_exclusive((108,))

    def test_promotion_requires_all_three_button_ids_and_embedded_page(self):
        import win32gui
        titles = {501: "Learn More...", 502: "Cancel", 503: "Ignore This"}
        classes = {501: "Button", 502: "Button", 503: "Button", 504: "Internet Explorer_Server"}
        with (patch.object(win32gui, "EnumChildWindows", side_effect=lambda h, f, p: [f(c, p) for c in classes]),
              patch.object(win32gui, "GetClassName", side_effect=classes.get),
              patch.object(win32gui, "GetParent", return_value=106),
              patch.object(win32gui, "GetDlgCtrlID", side_effect={501: 3617, 502: 2, 503: 3619}.get),
              patch.object(win32gui, "GetWindowText", side_effect=titles.get),
              patch.object(win32gui, "IsWindowVisible", return_value=True),
              patch.object(win32gui, "IsWindowEnabled", return_value=True)):
            self.assertEqual(windows._promotion_cancel(106), 502)
            titles[502] = "Discard changes"
            self.assertIsNone(windows._promotion_cancel(106))
            titles[502] = "Cancel"
            del classes[504]
            self.assertIsNone(windows._promotion_cancel(106))

    def test_other_announcements_libraries_owners_or_wrong_database_are_untouched(self):
        import win32gui
        for wrong in ("title", "owner", "other library", "buttons", "database"):
            with self.subTest(wrong=wrong):
                self.view = self.main.copy()
                self.view[106] = (windows.PROMOTION_TITLE if wrong != "title" else "Other Announcement",
                                  "#32770", 101 if wrong != "owner" else 901)
                if wrong == "other library":
                    self.view[109] = ("foreign.enl", "Afx:Fixture", 0)
                with (patch.object(windows, "_promotion_cancel", return_value=None if wrong == "buttons" else 202),
                      patch.object(windows, "_holders", return_value=set() if wrong == "database" else {(41, 12345)}),
                      patch.object(win32gui, "SendMessageTimeout") as click):
                    with self.assertRaises(windows.WindowInterference):
                        self.target.require_exclusive()
                    click.assert_not_called()

    def test_foreign_library_appears_after_filter_before_committing(self):
        import win32gui
        self.target.wait_dialog = Mock(side_effect=[103, 104])
        self.target.wait_dismissed = Mock(side_effect=lambda chain, hwnd, timeout: self.view.pop(hwnd))
        messages = []
        def post(hwnd, msg, wparam, lparam):
            messages.append((hwnd, msg))
            if hwnd == 101:
                self.view[103] = ("Import File", "#32770", 101)
            if hwnd == 1097:
                self.view[104] = ("Open", "#32770", 103)
        def open_foreign(*args):
            self.view[102] = ("EndNote 21 - foreign.enl", "Afx:Fixture", 0)
        enw = Path("fixture.enw").resolve()
        with (patch.object(sync, "_menu_has_import", return_value=True),
              patch.object(sync, "_control", side_effect=lambda dlg, cid, cls: cid if dlg == 103 else 3001),
              patch.object(sync, "_read_text", side_effect=["Import", str(enw), enw.name]),
              patch.object(sync, "_filename_edit", return_value=1152),
              patch.object(sync, "_send"),
              patch.object(sync, "_select_import_option", side_effect=open_foreign),
              patch.object(win32gui, "PostMessage", side_effect=post),
              patch.object(win32gui, "SendMessageTimeout") as close):
            with self.assertRaises(windows.WindowInterference):
                sync._targeted_import(self.target, enw, 1)
            with self.assertRaises(windows.WindowInterference):
                self.target.close()
        self.assertEqual([h for h, _ in messages], [101, 1097, 3001])
        close.assert_not_called()
        self.assertIn(102, self.view)

    def test_selected_filename_can_arrive_after_picker_closes(self):
        self.view[103] = ("Import File", "#32770", 101)
        file = Path("fixture.enw").resolve()
        with (patch.object(sync, "_control", return_value=1100),
              patch.object(sync, "_read_text", side_effect=["", file.name]),
              patch.object(sync.time, "sleep") as waited):
            sync._wait_selected_import(self.target, 103, file, 2)
        waited.assert_called_once()
        with (patch.object(sync, "_control", return_value=1100),
              patch.object(sync, "_read_text", return_value="different.enw"),
              patch.object(sync.time, "sleep") as waited):
            with self.assertRaises(windows.WindowInterference):
                sync._wait_selected_import(self.target, 103, file, 2)
        waited.assert_not_called()

    def test_import_progress_is_observed_to_completion_without_clicking(self):
        import win32gui
        progress = {**self.main, 106: ("Importing...", "#32770", 101)}
        with (patch.object(windows, "_windows", side_effect=[self.main, progress, self.main]),
              patch.object(sync, "_children", return_value=[201, 202]),
              patch.object(win32gui, "IsWindowVisible", return_value=True),
              patch.object(win32gui, "GetClassName", side_effect=["Button", "msctls_progress32"]),
              patch.object(win32gui, "GetDlgCtrlID", side_effect=[2, 3000]),
              patch.object(win32gui, "GetWindowText", side_effect=["Cancel", ""]),
              patch.object(win32gui, "PostMessage") as message):
            self.target.wait_dismissed((), 103, 1, progress=sync._is_import_progress)
        message.assert_not_called()

    def test_progress_with_hidden_advertisement_owner_still_checks_controls(self):
        progress_view = {**self.main, 109: ("Importing...", "#32770", 106)}
        self.target.cancelled_promotion = (106, 202)
        with (patch.object(windows, "_windows", side_effect=[progress_view, self.main]),
              patch.object(windows, "_is_hidden_promotion", return_value=True)):
            signature = Mock(return_value=True)
            self.target.wait_dismissed((), 103, 2, progress=signature)
            signature.assert_called_once_with(109, "Importing...")

    def test_lookalike_progress_missing_real_progress_control_is_rejected(self):
        import win32gui
        self.view[106] = ("Importing...", "#32770", 101)
        with (patch.object(sync, "_children", return_value=[201]),
              patch.object(win32gui, "IsWindowVisible", return_value=True),
              patch.object(win32gui, "GetClassName", return_value="Button"),
              patch.object(win32gui, "GetDlgCtrlID", return_value=2),
              patch.object(win32gui, "GetWindowText", return_value="Cancel")):
            with self.assertRaises(windows.WindowInterference):
                self.target.wait_dismissed((), 103, 1, progress=sync._is_import_progress)

    def test_unexpected_owned_title_rejected_without_inspecting_controls(self):
        self.view[106] = ("Confirm Record Changes", "#32770", 101)
        with patch.object(sync, "_children") as children:
            with self.assertRaises(windows.WindowInterference):
                self.target.wait_dismissed((), 103, 1, progress=sync._is_import_progress)
        children.assert_not_called()

    def test_even_matching_progress_requires_target_owner_and_no_other_library(self):
        matches = Mock(return_value=True)
        for extra in ({106: ("Importing...", "#32770", 99)},
                      {106: ("Importing...", "#32770", 101), 107: ("foreign.enl", "Afx:Fixture", 0)}):
            self.view = {**self.main, **extra}
            with self.assertRaises(windows.WindowInterference):
                self.target.wait_dismissed((), 103, 1, progress=matches)
        matches.assert_not_called()

    def test_released_database_waits_for_original_process_to_exit(self):
        import win32gui
        def exit_naturally(_):
            self.processes.clear()
            self.process.poll.return_value = 0
        with (patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.clear()) as close,
              patch.object(windows, "_holders", side_effect=lambda db: {(41, 12345)} if self.view else set()),
              patch.object(windows.time, "sleep", side_effect=exit_naturally) as waited):
            self.assertTrue(self.target.close(timeout=2))
        waited.assert_called_once()
        close.assert_called_once_with(101, 0x0010, 0, 0, 0x0002, 3000)
        self.process.terminate.assert_not_called()
        self.process.kill.assert_not_called()

    def test_second_window_or_process_during_shutdown_fails_without_touching_it(self):
        import win32gui
        for foreign in ("window", "hidden process"):
            with self.subTest(foreign=foreign):
                self.view = self.main.copy()
                self.processes = {41}
                def open_foreign(_):
                    if foreign == "window":
                        self.view[102] = ("foreign.enl", "Afx:Fixture", 0)
                    else:
                        self.processes.add(44)
                with (patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.clear()) as close,
                      patch.object(windows, "_holders", side_effect=lambda db: {(41, 12345)} if self.view else set()),
                      patch.object(windows.time, "sleep", side_effect=open_foreign)):
                    with self.assertRaises(windows.WindowInterference):
                        self.target.close(timeout=2)
                close.assert_called_once_with(101, 0x0010, 0, 0, 0x0002, 3000)
                self.process.terminate.assert_not_called()
                self.process.kill.assert_not_called()

    def test_released_process_that_never_exits_times_out_without_termination(self):
        import win32gui
        with (patch.object(win32gui, "SendMessageTimeout", side_effect=lambda *args: self.view.clear()) as close,
              patch.object(windows, "_holders", side_effect=lambda db: {(41, 12345)} if self.view else set()),
              patch.object(windows.time, "monotonic", side_effect=[0, .1, .2, 2]),
              patch.object(windows.time, "sleep") as waited):
            self.assertFalse(self.target.close(timeout=1))
        self.assertEqual(waited.call_count, 2)
        close.assert_called_once_with(101, 0x0010, 0, 0, 0x0002, 3000)
        self.process.terminate.assert_not_called()
        self.process.kill.assert_not_called()

    def test_no_application_level_import_of_an_enw(self):
        with tempfile.TemporaryDirectory() as td:
            base = Path(td).resolve()
            library, batch = base / "target.enl", base / "delta.enw"
            library.touch()
            batch.touch()
            with (patch.object(sync, "require_closed_endnote"),
                  patch.object(sync, "resolve_endnote_exe", return_value="EndNote.exe"),
                  patch.object(sync.subprocess, "Popen", return_value=self.process) as popen,
                  patch.object(sync.TargetWindow, "bind", return_value=self.target),
                  patch.object(sync, "_targeted_import") as native,
                  patch.object(self.target, "close", return_value=True)):
                result = sync.run_endnote_import(library, batch)
        popen.assert_called_once_with(["EndNote.exe", str(library)])
        native.assert_called_once_with(self.target, batch, 40.0)
        self.assertTrue(result["cleanly_closed"])

    def test_matches_library_title_across_endnote_versions(self):
        self.assertTrue(windows._matches_library_title("EndNote 21 - target.enl", "target.enl"))
        self.assertTrue(windows._matches_library_title("EndNote X9 - target.enl", "target.enl"))
        self.assertTrue(windows._matches_library_title("EndNote 20 - target", "target.enl"))
        self.assertTrue(windows._matches_library_title("EndNote 2025 - target.enl", "target.enl"))
        self.assertTrue(windows._matches_library_title("target.enl", "target.enl"))
        self.assertFalse(windows._matches_library_title("EndNote 21 - other.enl", "target.enl"))
        self.assertFalse(windows._matches_library_title("New Library", "target.enl"))

    def test_enum_windows_sandbox_failure_raises_window_interference(self):
        import win32gui
        self.enum.stop()
        with patch.object(win32gui, "EnumWindows", side_effect=OSError(126, "EnumWindows")):
            with self.assertRaisesRegex(windows.WindowInterference, "interactive Windows desktop"):
                windows._windows(1234)
        with patch.object(win32gui, "EnumWindows", return_value=None):
            with self.assertRaisesRegex(windows.WindowInterference, "No interactive desktop windows enumerated"):
                windows._windows(1234)

    def test_holders_resets_lasterror_234_probe_and_real_empty_desktop_windows(self):
        import ctypes
        import tempfile
        with tempfile.NamedTemporaryFile() as tmp:
            windows._holders(tmp.name)
            self.assertEqual(ctypes.WinDLL("kernel32").GetLastError(), 0)

"""Formatting on a simulated multi-library Windows desktop; no Office programs are started."""

from contextlib import ExitStack, redirect_stdout
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, Mock, patch

SCRIPTS = Path(__file__).resolve().parents[1] / "skills/endnote-cwyw/scripts"
sys.path.insert(0, str(SCRIPTS))
import format_cwyw as formatter


class FormatWindowTests(unittest.TestCase):
    def test_report_cannot_replace_the_docx(self):
        with tempfile.TemporaryDirectory() as tmp:
            source, destination = Path(tmp) / "source.docx", Path(tmp) / "new.docx"
            source.touch()
            with patch.object(sys, "argv", ["format", "--input", str(source), "--output", str(destination),
                                            "--report", str(destination), "--preflight-only"]):
                with self.assertRaisesRegex(ValueError, "separate paths"):
                    formatter.main()
            self.assertFalse(destination.exists())

    def test_target_lost_during_bind_does_not_close_an_unknown_window(self):
        with tempfile.TemporaryDirectory() as tmp:
            lib = Path(tmp) / "proposed.enl"
            lib.touch()
            with (patch.object(formatter, "list_process_pids", return_value=[]),
                  patch.object(formatter, "resolve_endnote_exe", return_value=Path("EndNote.exe")),
                  patch.object(formatter.subprocess, "Popen") as start,
                  patch.object(formatter.TargetWindow, "bind", side_effect=RuntimeError("two library windows"))):
                with self.assertRaisesRegex(RuntimeError, "two library windows"):
                    formatter.ensure_endnote_library_open(lib)
                start.return_value.terminate.assert_not_called()
                start.return_value.kill.assert_not_called()

    def test_foreign_dialog_same_pid_is_not_the_style_dialog(self):
        gui = Mock()
        con = Mock(GW_OWNER=4)
        processes = Mock()
        processes.GetWindowThreadProcessId.return_value = (1, 101)
        gui.IsWindowVisible.return_value = True
        gui.GetClassName.return_value = "#32770"
        owners = {30: 20, 20: 0, 31: 10}
        gui.GetWindow.side_effect = lambda hwnd, _: owners[hwnd]
        self.assertFalse(formatter.owned_word_dialog(30, 10, 101, gui, processes, con))
        self.assertTrue(formatter.owned_word_dialog(31, 10, 101, gui, processes, con))

    def test_same_pid_foreign_library_during_macro_prevents_save_and_close(self):
        self.run_simulated_format(foreign_during_macro=True, close_ok=True)

    def test_failed_target_close_does_not_publish_saved_candidate(self):
        self.run_simulated_format(foreign_during_macro=False, close_ok=False)

    def test_audited_candidate_published_after_verified_close(self):
        self.run_simulated_format(foreign_during_macro=False, close_ok=True)

    def run_simulated_format(self, *, foreign_during_macro, close_ok):
        with tempfile.TemporaryDirectory() as tmp, ExitStack() as stack:
            root = Path(tmp).resolve()
            source, output = root / "source.docx", root / "delivered.docx"
            lib, report = root / "test.enl", root / "report.json"
            source.write_bytes(b"frozen source")
            lib.touch()
            doc = MagicMock(FullName=str(source))
            doc.ActiveWindow.Hwnd = 100
            app = MagicMock()
            app.ActiveDocument = doc
            app.Documents.Open.return_value = doc
            app.Documents.Count = 0
            app.Documents.Add.return_value.Windows.Item.return_value.Hwnd = 99
            target = Mock()
            windows_of_owned_pid = [10]

            def require_target():
                if windows_of_owned_pid != [10]:
                    raise RuntimeError("Second library in owned EndNote process")

            target.require_exclusive.side_effect = require_target
            target.close.return_value = close_ok

            def run_macro(name):
                self.assertEqual(name, "ENFormatBibliography")
                if foreign_during_macro:
                    windows_of_owned_pid.append(20)

            app.Run.side_effect = run_macro
            staged = []

            def save(_doc, path, _expected, _timeout, _pythoncom, isolation):
                isolation.check()
                self.assertNotEqual(path, output)
                self.assertFalse(output.exists())
                path.write_bytes(b"native candidate")
                staged.append(path)
                return {}

            modules = {key: MagicMock() for key in
                       ("pythoncom", "win32com", "win32com.client", "win32con", "win32gui", "win32process")}
            modules["win32com"].client = modules["win32com.client"]
            modules["win32com.client"].DispatchEx.return_value = app
            modules["win32process"].GetWindowThreadProcessId.return_value = (1, 101)
            stack.enter_context(patch.dict(sys.modules, modules))
            stack.enter_context(patch.object(formatter, "list_process_pids", return_value=[]))
            stack.enter_context(patch.object(formatter, "inspect_modal_dialogs", return_value=[]))
            stack.enter_context(patch.object(formatter, "ensure_endnote_library_open", return_value=(
                {"pid": 102, "enl": str(lib)}, target)))
            stack.enter_context(patch.object(formatter, "temporary_state", return_value={"citation_groups": 1, "record_ids": [2]}))
            stack.enter_context(patch.object(formatter, "audit", return_value={"status": "PASS", "fields": {}}))
            stack.enter_context(patch.object(formatter, "save_if_complete", side_effect=save))
            stack.enter_context(patch.object(sys, "argv", [
                "format", "--enl", str(lib), "--input", str(source), "--output", str(output), "--report", str(report)
            ]))
            with io.TextIOWrapper(io.BytesIO(), encoding="utf-8") as stream, redirect_stdout(stream):
                result = formatter.main()
            detail = json.loads(report.read_text(encoding="utf-8"))
            self.assertEqual(source.read_bytes(), b"frozen source")
            self.assertEqual(bool(output.exists()), close_ok and not foreign_during_macro)
            self.assertEqual(result, 0 if output.exists() else 1)
            if foreign_during_macro:
                self.assertIn("Second library", detail["worker"]["isolation_error"])
                self.assertEqual(windows_of_owned_pid, [10, 20])
                self.assertEqual(staged, [])
                target.close.assert_not_called()
            else:
                target.close.assert_called_once()
                if close_ok:
                    self.assertEqual(output.read_bytes(), b"native candidate")
                else:
                    self.assertEqual(staged[0].read_bytes(), b"native candidate")

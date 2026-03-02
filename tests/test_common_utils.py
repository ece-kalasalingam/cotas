import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from common.utils import (
    app_settings_path,
    coerce_excel_number,
    from_portable_path,
    get_last_saved_dir,
    remember_dialog_dir,
    normalize,
    resolve_dialog_start_path,
    resource_path,
    set_last_saved_dir,
    to_portable_path,
)


class TestNormalize(unittest.TestCase):
    def test_none(self) -> None:
        self.assertEqual(normalize(None), "")

    def test_trim_and_lower(self) -> None:
        self.assertEqual(normalize("  TeSt  "), "test")

    def test_whitespace_only(self) -> None:
        self.assertEqual(normalize("   "), "")

    def test_numeric_and_bool(self) -> None:
        self.assertEqual(normalize(123), "123")
        self.assertEqual(normalize(12.50), "12.5")
        self.assertEqual(normalize(True), "true")
        self.assertEqual(normalize(False), "false")

    def test_control_chars(self) -> None:
        self.assertEqual(normalize("\n\t  A\t\n"), "a")

    def test_unicode(self) -> None:
        self.assertEqual(normalize(" \u00c4\u00d6\u00dc "), "\u00e4\u00f6\u00fc")


class TestResourcePath(unittest.TestCase):
    def test_relative_path(self) -> None:
        expected_base = os.path.dirname(
            os.path.dirname(os.path.abspath("common/utils.py"))
        )
        rel = os.path.join("assets", "help.svg")
        self.assertEqual(resource_path(rel), os.path.join(expected_base, rel))
        self.assertTrue(os.path.isabs(resource_path(rel)))

    def test_empty_and_dot_paths(self) -> None:
        expected_base = os.path.dirname(
            os.path.dirname(os.path.abspath("common/utils.py"))
        )
        self.assertEqual(resource_path(""), os.path.join(expected_base, ""))
        self.assertEqual(resource_path("."), os.path.join(expected_base, "."))
        self.assertEqual(resource_path(".."), os.path.join(expected_base, ".."))

    def test_absolute_input_behavior(self) -> None:
        expected_base = os.path.dirname(
            os.path.dirname(os.path.abspath("common/utils.py"))
        )
        abs_input = os.path.abspath(expected_base)
        self.assertEqual(resource_path(abs_input), abs_input)

    def test_meipass_mode(self) -> None:
        had_meipass = hasattr(sys, "_MEIPASS")
        old_meipass = getattr(sys, "_MEIPASS", None)
        try:
            sys._MEIPASS = r"C:\temp\bundle_root"
            self.assertEqual(
                resource_path("assets/icon.svg"),
                os.path.join(r"C:\temp\bundle_root", "assets/icon.svg"),
            )
        finally:
            if had_meipass:
                sys._MEIPASS = old_meipass
            else:
                delattr(sys, "_MEIPASS")


class TestCoerceExcelNumber(unittest.TestCase):
    def test_none_and_empty(self) -> None:
        self.assertIsNone(coerce_excel_number(None))
        self.assertEqual(coerce_excel_number("   "), "")

    def test_preserve_bool(self) -> None:
        self.assertIs(coerce_excel_number(True), True)
        self.assertIs(coerce_excel_number(False), False)

    def test_numeric_python_types(self) -> None:
        self.assertEqual(coerce_excel_number(12), 12)
        self.assertEqual(coerce_excel_number(12.0), 12)
        self.assertEqual(coerce_excel_number(12.25), 12.25)

    def test_numeric_strings(self) -> None:
        self.assertEqual(coerce_excel_number("12"), 12)
        self.assertEqual(coerce_excel_number(" 12.0 "), 12)
        self.assertEqual(coerce_excel_number("12.50"), 12.5)
        self.assertEqual(coerce_excel_number("1e3"), 1000)
        self.assertEqual(coerce_excel_number("1,234"), 1234)

    def test_non_numeric_strings(self) -> None:
        self.assertEqual(coerce_excel_number("abc"), "abc")
        self.assertEqual(coerce_excel_number("  A12 "), "A12")


class TestSettingsHelpers(unittest.TestCase):
    @patch("common.utils._runtime_base_dir", return_value=Path(r"D:\portable\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=False)
    def test_app_settings_path_portable_or_dev(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(path, Path(r"D:\portable\FOCUS\settings.json"))

    @patch("common.utils._runtime_base_dir", return_value=Path(r"C:\Program Files\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch.dict(os.environ, {"APPDATA": r"C:\Users\alice\AppData\Roaming"}, clear=False)
    @patch("common.utils.sys.platform", "win32")
    def test_app_settings_path_installed_windows(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(
            path,
            Path(r"C:\Users\alice\AppData\Roaming\FOCUS\settings.json"),
        )

    @patch("common.utils._runtime_base_dir", return_value=Path("/Applications/FOCUS.app/Contents/MacOS"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch("common.utils.sys.platform", "darwin")
    @patch("pathlib.Path.home", return_value=Path("/Users/alice"))
    def test_app_settings_path_installed_macos(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(
            str(path).replace("\\", "/"),
            "/Users/alice/Library/Application Support/FOCUS/settings.json",
        )

    @patch("common.utils._runtime_base_dir", return_value=Path("/usr/local/bin"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch("common.utils.sys.platform", "linux")
    @patch.dict(os.environ, {"XDG_CONFIG_HOME": "/home/alice/.config"}, clear=False)
    def test_app_settings_path_installed_linux(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(
            str(path).replace("\\", "/"),
            "/home/alice/.config/FOCUS/settings.json",
        )

    @patch("common.utils.sys.frozen", True, create=True)
    @patch("common.utils.sys.platform", "win32")
    @patch.dict(
        os.environ,
        {"FOCUS_PORTABLE": "1", "ProgramFiles": r"C:\Program Files"},
        clear=False,
    )
    def test_portable_override_on_installed_exe(self) -> None:
        from common.utils import _is_installed_exe

        self.assertFalse(_is_installed_exe(Path(r"C:\Program Files\FOCUS")))

    def test_portable_path_roundtrip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            folder = base / "reports"
            folder.mkdir()

            serialized = to_portable_path(str(folder))
            self.assertIn("/", serialized)
            restored = from_portable_path(serialized)
            self.assertEqual(os.path.normcase(restored), os.path.normcase(str(folder)))

    def test_set_and_get_last_saved_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            target = run_base / "excel"
            target.mkdir()

            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ):
                settings_path = set_last_saved_dir(str(target), app_name="FOCUS")
                self.assertTrue(settings_path.exists())
                self.assertEqual(settings_path, run_base / "settings.json")

                loaded = get_last_saved_dir(app_name="FOCUS")
                self.assertEqual(
                    os.path.normcase(loaded or ""),
                    os.path.normcase(str(target)),
                )

    def test_resolve_dialog_start_path_prefers_json_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            target = run_base / "excel"
            target.mkdir()
            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ):
                set_last_saved_dir(str(target), app_name="FOCUS")
                start_path = resolve_dialog_start_path("FOCUS", "demo.xlsx")
                self.assertEqual(
                    os.path.normcase(start_path),
                    os.path.normcase(str(target / "demo.xlsx")),
                )

    def test_remember_dialog_dir_updates_json(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            target = run_base / "excel"
            target.mkdir()
            selected = target / "out.xlsx"

            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ):
                remember_dialog_dir(str(selected), "FOCUS")
                loaded = get_last_saved_dir("FOCUS")
                self.assertEqual(
                    os.path.normcase(loaded or ""),
                    os.path.normcase(str(target)),
                )


if __name__ == "__main__":
    unittest.main()

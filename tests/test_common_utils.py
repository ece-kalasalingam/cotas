import os
import sys
import tempfile
import unittest
from pathlib import Path, PurePosixPath
from typing import Any, cast
from unittest.mock import Mock, patch

import pytest

from common.exceptions import AppSystemError, ValidationError
from common.utils import (
    _reset_runtime_storage_cache_for_tests,
    app_secrets_dir,
    app_settings_path,
    coerce_excel_number,
    create_app_runtime_sqlite_file,
    emit_user_status,
    from_portable_path,
    get_last_saved_dir,
    get_ui_language_preference,
    log_process_message,
    normalize,
    remember_dialog_dir,
    resolve_dialog_start_path,
    resource_path,
    set_last_saved_dir,
    set_ui_language_preference,
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
            cast_sys = sys  # alias for test-only dynamic attribute usage
            setattr(cast_sys, "_MEIPASS", r"C:\temp\bundle_root")
            self.assertEqual(
                resource_path("assets/icon.svg"),
                os.path.join(r"C:\temp\bundle_root", "assets/icon.svg"),
            )
        finally:
            if had_meipass:
                setattr(sys, "_MEIPASS", old_meipass)
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

    def test_other_types_return_as_is(self) -> None:
        payload = {"a": 1}
        self.assertEqual(coerce_excel_number(payload), payload)


@pytest.mark.windows_acl_compat
class TestSettingsHelpers(unittest.TestCase):
    def setUp(self) -> None:
        _reset_runtime_storage_cache_for_tests()

    @patch("common.utils._runtime_base_dir", return_value=Path(r"D:\portable\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=False)
    @patch("common.utils._is_storage_dir_usable", return_value=True)
    def test_app_settings_path_portable_or_dev(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(path, Path(r"D:\portable\FOCUS\settings.json"))

    @patch("common.utils._runtime_base_dir", return_value=Path(r"C:\Program Files\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch("common.utils._is_storage_dir_usable", return_value=True)
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
    @patch("common.utils._is_storage_dir_usable", return_value=True)
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
    @patch("common.utils._is_storage_dir_usable", return_value=True)
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
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
            ):
                settings_path = set_last_saved_dir(str(target), app_name="FOCUS")
                self.assertTrue(settings_path.exists())
                self.assertEqual(settings_path, run_base / "settings.json")

                loaded = get_last_saved_dir(app_name="FOCUS")
                self.assertEqual(
                    os.path.normcase(loaded or ""),
                    os.path.normcase(str(target)),
                )

    def test_ui_language_preference_defaults_to_english(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
            ):
                self.assertEqual(get_ui_language_preference(app_name="FOCUS"), "en")

    def test_ui_language_auto_aliases_normalize_to_english(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
            ):
                set_ui_language_preference(app_name="FOCUS", ui_language="auto")
                self.assertEqual(get_ui_language_preference(app_name="FOCUS"), "en")

    def test_resolve_dialog_start_path_prefers_json_dir(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            target = run_base / "excel"
            target.mkdir()
            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
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
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
            ):
                remember_dialog_dir(str(selected), "FOCUS")
                loaded = get_last_saved_dir("FOCUS")
                self.assertEqual(
                    os.path.normcase(loaded or ""),
                    os.path.normcase(str(target)),
                )

    def test_read_settings_payload_decode_and_oserror_paths(self) -> None:
        from common.utils import _read_settings_payload

        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "settings.json"
            p.write_text("{bad", encoding="utf-8")
            self.assertEqual(_read_settings_payload(p), {})

        with patch("pathlib.Path.exists", return_value=True), patch(
            "pathlib.Path.read_text", side_effect=OSError("denied")
        ):
            self.assertEqual(_read_settings_payload(Path("x")), {})

    def test_get_last_saved_dir_invalid_and_remember_dir_guards(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            run_base = Path(tmp)
            with patch("common.utils._runtime_base_dir", return_value=run_base), patch(
                "common.utils._is_installed_exe", return_value=False
            ), patch(
                "common.utils._is_storage_dir_usable", return_value=True
            ):
                settings_file = run_base / "settings.json"
                settings_file.write_text('{"last_saved_dir": 1}', encoding="utf-8")
                self.assertIsNone(get_last_saved_dir("FOCUS"))

        remember_dialog_dir("", "FOCUS")
        with patch("os.path.expanduser", side_effect=OSError("bad path")):
            remember_dialog_dir("~/x", "FOCUS")
        with tempfile.TemporaryDirectory() as tmp:
            remember_dialog_dir(str(Path(tmp) / "missing.xlsx"), "FOCUS")
            remember_dialog_dir(str(Path(tmp) / "missing_parent" / "x.xlsx"), "FOCUS")

    def test_resolve_dialog_start_path_downloads_and_home_fallbacks(self) -> None:
        with patch("common.utils.get_last_saved_dir", return_value=None), patch(
            "pathlib.Path.home", return_value=Path("/home/test")
        ), patch(
            "common.utils.Path.exists", side_effect=[True, False]
        ), patch(
            "common.utils.Path.is_dir", return_value=True
        ):
            out = resolve_dialog_start_path("FOCUS", "a.xlsx")
            self.assertIn("Downloads", out)

        with patch("common.utils.get_last_saved_dir", return_value=None), patch(
            "pathlib.Path.home", return_value=Path("/home/test")
        ), patch("common.utils.Path.exists", return_value=False), patch(
            "common.utils.Path.is_dir", return_value=False
        ):
            out = resolve_dialog_start_path("FOCUS", "a.xlsx")
            self.assertTrue(out.endswith("a.xlsx"))

    def test_runtime_base_and_join_path_posix_branch(self) -> None:
        from common.utils import _join_path, _runtime_base_dir

        with patch("common.utils.sys.frozen", True, create=True), patch(
            "common.utils.sys.executable", "/tmp/app/bin/app"
        ):
            base = _runtime_base_dir()
            self.assertTrue(str(base).lower().endswith("app\\bin") or str(base).lower().endswith("app/bin"))

        self.assertEqual(
            str(_join_path(cast(Any, PurePosixPath("/tmp/base")), "child")).replace("\\", "/"),
            "/tmp/base/child",
        )

    @patch("common.utils._runtime_base_dir", return_value=Path(r"D:\portable\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=False)
    @patch("common.utils._is_storage_dir_usable", return_value=True)
    def test_app_secrets_dir_portable_or_dev(self, *_mocks) -> None:
        path = app_secrets_dir("FOCUS")
        self.assertEqual(path, Path(r"D:\portable\FOCUS\secrets"))

    @patch("common.utils._runtime_base_dir", return_value=Path(r"C:\Program Files\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch.dict(os.environ, {"PROGRAMDATA": r"C:\ProgramData"}, clear=False)
    @patch("common.utils.sys.platform", "win32")
    def test_app_secrets_dir_installed_windows_uses_programdata(self, *_mocks) -> None:
        path = app_secrets_dir("FOCUS")
        self.assertEqual(path, Path(r"C:\ProgramData\FOCUS\secrets"))

    @patch("common.utils._runtime_base_dir", return_value=Path("/Applications/FOCUS.app/Contents/MacOS"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch("common.utils.sys.platform", "darwin")
    def test_app_secrets_dir_installed_macos_uses_users_shared(self, *_mocks) -> None:
        path = app_secrets_dir("FOCUS")
        self.assertEqual(str(path).replace("\\", "/"), "/Users/Shared/FOCUS/secrets")

    @patch("common.utils._runtime_base_dir", return_value=Path("/usr/local/bin"))
    @patch("common.utils._is_installed_exe", return_value=True)
    @patch("common.utils.sys.platform", "linux")
    @patch.dict(os.environ, {"XDG_CONFIG_HOME": "/home/alice/.config"}, clear=False)
    def test_app_secrets_dir_installed_linux_uses_installed_storage_root(self, *_mocks) -> None:
        path = app_secrets_dir("FOCUS")
        self.assertEqual(str(path).replace("\\", "/"), "/home/alice/.config/FOCUS/secrets")

    @patch("common.utils._runtime_base_dir", return_value=Path(r"D:\portable\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=False)
    @patch("common.utils._is_storage_dir_usable", return_value=False)
    @patch("common.utils.tempfile.mkdtemp", return_value=r"D:\tmp\focus_runtime_123")
    def test_app_settings_path_falls_back_to_temp_when_primary_not_usable(self, *_mocks) -> None:
        path = app_settings_path("FOCUS")
        self.assertEqual(path, Path(r"D:\tmp\focus_runtime_123\settings.json"))

    @patch("common.utils._runtime_base_dir", return_value=Path(r"D:\portable\FOCUS"))
    @patch("common.utils._is_installed_exe", return_value=False)
    @patch("common.utils._is_storage_dir_usable", return_value=False)
    def test_create_app_runtime_sqlite_file_uses_runtime_temp_fallback(self, *_mocks) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with patch("common.utils.tempfile.mkdtemp", return_value=tmp):
                expected = str(Path(tmp) / "sqlite" / "a.sqlite3")
                with patch("common.utils.tempfile.mkstemp", return_value=(12, expected)):
                    fd, path = create_app_runtime_sqlite_file("FOCUS", prefix="a_", suffix=".sqlite3")
            self.assertEqual(fd, 12)
            self.assertEqual(path, expected)


class TestLogProcessMessage(unittest.TestCase):
    def test_success_message(self) -> None:
        logger = Mock()
        notify = Mock()

        result = log_process_message(
            "validating workbook",
            logger=logger,
            notify=notify,
            success_message="Workbook validated.",
        )

        self.assertTrue(result)
        logger.info.assert_called_once()
        self.assertEqual(logger.info.call_args.args[0], "Workbook validated.")
        notify.assert_called_once_with("Workbook validated.", "success")

    def test_validation_error_shows_detailed_message(self) -> None:
        logger = Mock()
        notify = Mock()
        error = ValidationError("Row 4: CO value is missing.")

        result = log_process_message(
            "validating workbook",
            logger=logger,
            error=error,
            notify=notify,
        )

        self.assertFalse(result)
        logger.warning.assert_called_once()
        self.assertEqual(logger.warning.call_args.args[:2], ("%s failed due to data error.", "validating workbook"))
        notify.assert_called_once_with("Row 4: CO value is missing.", "error")

    def test_validation_error_prefers_code_mapping_message(self) -> None:
        logger = Mock()
        notify = Mock()
        error = ValidationError(
            "fallback",
            code="WORKBOOK_NOT_FOUND",
            context={"workbook": "sample.xlsx"},
        )

        result = log_process_message(
            "validating workbook",
            logger=logger,
            error=error,
            notify=notify,
        )

        self.assertFalse(result)
        notify.assert_called_once()
        self.assertIn("sample.xlsx", notify.call_args.args[0])
        self.assertEqual(notify.call_args.args[1], "error")

    def test_validation_error_empty_detail_fallback_message(self) -> None:
        logger = Mock()
        notify = Mock()
        error = ValidationError("   ")
        result = log_process_message("process", logger=logger, error=error, notify=notify)
        self.assertFalse(result)
        notify.assert_called_once_with("Validation failed due to invalid data.", "error")

    def test_app_system_error_logs_generic_english_message(self) -> None:
        logger = Mock()
        notify = Mock()
        error = AppSystemError("உருவாக்கம் தோல்வியடைந்தது")

        result = log_process_message(
            "generating final report",
            logger=logger,
            error=error,
            notify=notify,
        )

        self.assertFalse(result)
        logger.error.assert_called_once()
        self.assertEqual(
            logger.error.call_args.args[:2],
            ("%s failed due to a system/application error.", "generating final report"),
        )
        notify.assert_called_once_with(
            "Error happened while generating final report.",
            "error",
        )

    def test_system_error_shows_process_scoped_message(self) -> None:
        logger = Mock()
        notify = Mock()
        error = RuntimeError("disk write failed")

        result = log_process_message(
            "generating final report",
            logger=logger,
            error=error,
            notify=notify,
        )

        self.assertFalse(result)
        logger.exception.assert_called_once()
        self.assertEqual(
            logger.exception.call_args.args[:2],
            ("%s failed due to a system/application error.", "generating final report"),
        )
        self.assertIs(logger.exception.call_args.kwargs.get("exc_info"), error)
        notify.assert_called_once_with(
            "Error happened while generating final report.",
            "error",
        )


class TestEmitUserStatus(unittest.TestCase):
    def test_emits_when_signal_has_emit(self) -> None:
        signal = Mock()
        logger = Mock()

        emit_user_status(signal, "Done", logger=logger)

        signal.emit.assert_called_once_with("Done")
        logger.exception.assert_not_called()

    def test_ignores_missing_emit(self) -> None:
        signal = object()
        logger = Mock()

        emit_user_status(signal, "Done", logger=logger)

        logger.debug.assert_called_once()


if __name__ == "__main__":
    unittest.main()


def test_runtime_min_free_bytes_handles_invalid_and_negative_env() -> None:
    from common import utils

    with patch.dict(os.environ, {utils.RUNTIME_MIN_FREE_BYTES_ENV_VAR: 'not-a-number'}, clear=False):
        assert utils._runtime_min_free_bytes() == utils.DEFAULT_RUNTIME_MIN_FREE_BYTES

    with patch.dict(os.environ, {utils.RUNTIME_MIN_FREE_BYTES_ENV_VAR: '-5'}, clear=False):
        assert utils._runtime_min_free_bytes() == 0


def test_directory_helpers_return_false_on_oserror() -> None:
    from common import utils

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp)
        with patch('builtins.open', side_effect=OSError('blocked')):
            assert utils._directory_is_writable(path) is False

    with patch('common.utils.shutil.disk_usage', side_effect=OSError('denied')):
        assert utils._directory_has_free_space(Path('.'), min_free_bytes=1) is False


def test_remember_dialog_dir_safe_logs_warning_on_oserror() -> None:
    from common import utils

    logger = Mock()
    with patch('common.utils.remember_dialog_dir', side_effect=OSError('cannot write')):
        utils.remember_dialog_dir_safe('C:/tmp/out.xlsx', app_name='FOCUS', logger=logger)

    logger.warning.assert_called_once()


def test_safe_extra_formatter_backfills_job_and_step_fields() -> None:
    import logging

    from common import utils

    formatter = utils._SafeExtraFormatter(fmt=utils.LOG_FORMAT, datefmt=utils.LOG_DATE_FORMAT)
    record = logging.LogRecord(
        name='test.logger',
        level=logging.INFO,
        pathname=__file__,
        lineno=1,
        msg='hello',
        args=(),
        exc_info=None,
    )

    rendered = formatter.format(record)

    assert 'job=-' in rendered
    assert 'step=-' in rendered


def test_emit_user_status_ignores_blank_and_none_signal_and_logs_emit_failures() -> None:
    logger = Mock()

    signal = Mock()
    emit_user_status(signal, '   ', logger=logger)
    signal.emit.assert_not_called()

    emit_user_status(None, 'hello', logger=logger)

    class _BadSignal:
        def emit(self, _message: str) -> None:
            raise RuntimeError('emit failed')

    emit_user_status(_BadSignal(), 'hello', logger=logger)

    logger.exception.assert_called_once()


def test_is_installed_exe_windows_and_unix_roots() -> None:
    from common import utils

    with patch('common.utils.sys.frozen', True, create=True), patch('common.utils.sys.platform', 'win32'), patch.dict(
        os.environ,
        {
            'ProgramFiles': r'C:\\Program Files',
            'ProgramFiles(x86)': r'C:\\Program Files (x86)',
            'LOCALAPPDATA': r'C:\\Users\\alice\\AppData\\Local',
            'FOCUS_PORTABLE': '0',
        },
        clear=False,
    ):
        assert utils._is_installed_exe(Path(r'C:\\Program Files\\FOCUS')) is True
        assert utils._is_installed_exe(Path(r'D:\\portable\\FOCUS')) is False

    with patch('common.utils.sys.frozen', True, create=True), patch('common.utils.sys.platform', 'darwin'), patch.dict(
        os.environ,
        {'FOCUS_PORTABLE': '0'},
        clear=False,
    ):
        assert utils._is_installed_exe(Mock(resolve=lambda: '/Applications/FOCUS.app/Contents/MacOS')) is True
        assert utils._is_installed_exe(Mock(resolve=lambda: '/Users/alice/FOCUS')) is False

    with patch('common.utils.sys.frozen', True, create=True), patch('common.utils.sys.platform', 'linux'), patch.dict(
        os.environ,
        {'FOCUS_PORTABLE': '0'},
        clear=False,
    ):
        assert utils._is_installed_exe(Mock(resolve=lambda: '/usr/bin/focus')) is True
        assert utils._is_installed_exe(Mock(resolve=lambda: '/home/alice/focus')) is False

def test_runtime_mode_portable_branch() -> None:
    from common import utils

    with patch('common.utils._is_installed_exe', return_value=False), patch('common.utils.sys.frozen', True, create=True):
        assert utils._runtime_mode(Path('C:/portable')) == 'portable'


def test_join_path_posix_branch() -> None:
    from common import utils

    assert utils._join_path(Path('/tmp/base'), 'child').as_posix().endswith('/tmp/base/child')


def test_app_runtime_storage_dir_returns_cached_temp_dir_directly() -> None:
    from common import utils

    utils._reset_runtime_storage_cache_for_tests()
    temp_path = Path('C:/tmp/focus_runtime_cached')
    utils._RUNTIME_STORAGE_DIR_CACHE['FOCUS'] = temp_path
    utils._RUNTIME_TEMP_DIRS.add(temp_path)

    with patch('common.utils.app_primary_storage_dir', side_effect=AssertionError('should not be called')):
        assert utils.app_runtime_storage_dir('FOCUS') == temp_path

    utils._reset_runtime_storage_cache_for_tests()


def test_create_app_runtime_sqlite_file_falls_back_when_primary_mkstemp_fails() -> None:
    from common import utils

    with tempfile.TemporaryDirectory() as tmp1, tempfile.TemporaryDirectory() as tmp2:
        primary = Path(tmp1)
        fallback = Path(tmp2)

        def _mkstemp_side_effect(*_args, **kwargs):
            if str(kwargs.get('dir', '')).endswith('sqlite') and str(primary / 'sqlite') in str(kwargs['dir']):
                raise OSError('primary fail')
            return (44, str(Path(kwargs['dir']) / 'ok.sqlite3'))

        with patch('common.utils.app_runtime_storage_dir', return_value=primary), patch(
            'common.utils._new_runtime_temp_dir', return_value=fallback
        ), patch('common.utils.tempfile.mkstemp', side_effect=_mkstemp_side_effect):
            fd, out = utils.create_app_runtime_sqlite_file('FOCUS', prefix='x_', suffix='.sqlite3')

        assert fd == 44
        assert str(fallback / 'sqlite') in out.replace('\\', '/') or str((fallback / 'sqlite')).replace('\\', '/') in out.replace('\\', '/')


def test_app_log_path_uses_settings_parent() -> None:
    from common import utils

    with patch('common.utils.app_settings_path', return_value=Path('C:/base/settings.json')):
        assert str(utils.app_log_path('FOCUS')).replace('\\', '/').endswith('C:/base/focus.log')


def test_configure_app_logging_configures_once_and_reuses_guard() -> None:
    import logging

    from common import utils

    if hasattr(utils.configure_app_logging, '_configured'):
        delattr(utils.configure_app_logging, '_configured')

    log_path = Path('C:/tmp/focus.log')

    class _FakeHandler:
        def __init__(self, *_args, **_kwargs) -> None:
            self.level = None
            self.formatter = None

        def setLevel(self, level: int) -> None:  # noqa: N802
            self.level = level

        def setFormatter(self, formatter) -> None:  # noqa: N802
            self.formatter = formatter

    root_logger = Mock()

    with patch('common.utils.app_log_path', return_value=log_path), patch('common.utils.RotatingFileHandler', _FakeHandler), patch(
        'common.utils.logging.getLogger', return_value=root_logger
    ):
        returned = utils.configure_app_logging('FOCUS', level=logging.INFO, max_bytes=123, backup_count=2)

    assert returned == log_path
    root_logger.setLevel.assert_called_once_with(logging.INFO)
    root_logger.addHandler.assert_called_once()

    with patch('common.utils.app_log_path', return_value=log_path), patch('common.utils.RotatingFileHandler', side_effect=AssertionError('should not construct handler')):
        returned_again = utils.configure_app_logging('FOCUS')

    assert returned_again == log_path

    if hasattr(utils.configure_app_logging, '_configured'):
        delattr(utils.configure_app_logging, '_configured')


